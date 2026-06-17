#requires -Version 7
<#
.SYNOPSIS
  Nasdaq(RSS) + NYSE(CSV) 현재 거래정지를 수집해 data/halts.json 갱신,
  거래정지 내역을 data/halts_history.json 에 누적(중복제거·보존기간 정리).

.NOTES
  - self-hosted 러너(미국 사이트 도달 가능)에서 실행 전제. (Azure 클라우드는 503 봇차단 위험)
  - 시각은 원본 그대로 ET(미국 동부) 유지. KST 변환은 화면(index.html)에서 수행.
  - ⚠️ 필드 매핑은 best-effort. 첫 실행 시 콘솔의 "[진단]" 로그로 실제 필드명을 확인하고
    필요하면 아래 매핑 구간을 조정하세요.
#>
[CmdletBinding()]
param(
  [string]$DataDir = (Join-Path $PSScriptRoot '..' 'data'),
  [int]$HistoryRetentionDays = 90
)
$ErrorActionPreference = 'Stop'
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
$NASDAQ_RSS = 'http://www.nasdaqtrader.com/rss.aspx?feed=tradehalts'
$NASDAQ_SRC = 'https://www.nasdaqtrader.com/trader.aspx?id=tradehalts'
$NYSE_CSV   = 'https://www.nyse.com/api/trade-halts/current/download'
$NYSE_SRC   = 'https://www.nyse.com/trade/trading-halts'

# Nasdaq/UTP 거래정지 사유 코드 → 설명 (필요 시 보강)
$REASONS = @{
  'T1'='News Pending'; 'T2'='News Released'; 'T3'='News and Resumption Times';
  'T5'='Single Stock Trading Pause (Volatility)'; 'T6'='Extraordinary Market Activity';
  'T8'='ETF Component Prices Not Available'; 'T12'='Additional Information Requested by Nasdaq';
  'H4'='Non-compliance'; 'H9'='Not Current in Filings'; 'H10'='SEC Trading Suspension';
  'H11'='Regulatory Concern'; 'O1'='Operations Halt, Contact Market Operations';
  'IPO1'='IPO Issue Not Yet Trading'; 'IPOQ'='IPO Quotes Released for Quotation';
  'M'='Volatility Trading Pause (LULD)'; 'LUDP'='Volatility Trading Pause (LULD)';
  'LUDS'='Volatility Trading Pause - Straddle Condition';
  'MWC1'='Market Wide Circuit Breaker Level 1'; 'MWC2'='Market Wide Circuit Breaker Level 2';
  'MWC3'='Market Wide Circuit Breaker Level 3'; 'MWCQ'='Market Wide Circuit Breaker Resumption';
  'R1'='New Issue Available'; 'R4'='Qualifications Halt Ended - Maint. Req. Met';
  'R9'='Filing Requirement Satisfied'; 'D'='Security Deletion'
}

function Format-Date([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){return ''}
  $s=$s.Trim()
  foreach($f in 'MM/dd/yyyy','M/d/yyyy','yyyy-MM-dd','MM-dd-yyyy','yyyyMMdd'){
    $d=[datetime]::MinValue
    if([datetime]::TryParseExact($s,$f,[cultureinfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$d)){return $d.ToString('yyyy-MM-dd')}
  }
  $d=[datetime]::MinValue
  if([datetime]::TryParse($s,[cultureinfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$d)){return $d.ToString('yyyy-MM-dd')}
  return $s
}
function Format-Time([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){return ''}
  if($s -match '(\d{1,2}):(\d{2}):(\d{2})'){ return ('{0:00}:{1}:{2}' -f [int]$matches[1],$matches[2],$matches[3]) }
  if($s -match '(\d{1,2}):(\d{2})'){ return ('{0:00}:{1}:00' -f [int]$matches[1],$matches[2]) }
  return $s.Trim()
}
function Get-Field($hash,[string[]]$names){
  foreach($n in $names){ foreach($k in $hash.Keys){ if($k -ieq $n){ return $hash[$k] } } }
  foreach($n in $names){ foreach($k in $hash.Keys){ if($k -imatch [regex]::Escape($n)){ return $hash[$k] } } }
  return ''
}
function Remove-Bom([string]$s){ if($s){ return $s.TrimStart([char]0xFEFF,[char]0x200B) } return $s }
function Normalize-Market([string]$m){
  if([string]::IsNullOrWhiteSpace($m)){ return 'Nasdaq' }
  $m=$m.Trim()
  if($m -ieq 'NASDAQ'){ return 'Nasdaq' }
  return $m   # NYSE, NYSE American, NYSE Arca, BATS, IEX 등은 원문 유지(화면 배지가 'nyse' 포함 여부로 색 구분)
}

function Get-NasdaqHalts {
  Write-Host "[Nasdaq] $NASDAQ_RSS 수집..."
  try{
    $resp = Invoke-WebRequest -Uri $NASDAQ_RSS -Headers @{ 'User-Agent'=$UA } -TimeoutSec 30 -UseBasicParsing
    $xml = [xml](Remove-Bom $resp.Content)   # BOM 제거 후 파싱
  }catch{ Write-Warning "[Nasdaq] 수집 실패: $($_.Exception.Message)"; return @() }

  $items = $xml.SelectNodes('//item')
  if(-not $items -or $items.Count -eq 0){ Write-Warning "[Nasdaq] item 없음 (장 마감/정지 없음 가능)"; return @() }

  # [진단] 첫 item의 실제 자식 필드명 출력 — 매핑 검증용
  $first=@{}; foreach($c in $items[0].ChildNodes){ if($c.NodeType -eq 'Element'){ $first[$c.LocalName]=$c.InnerText.Trim() } }
  Write-Host "[진단][Nasdaq] item 필드: $($first.Keys -join ', ')"

  $out=@()
  foreach($it in $items){
    $h=@{}; foreach($c in $it.ChildNodes){ if($c.NodeType -eq 'Element'){ $h[$c.LocalName]=$c.InnerText.Trim() } }
    $sym = Get-Field $h @('IssueSymbol','Symbol','NCAIssueSymbol')
    if([string]::IsNullOrWhiteSpace($sym)){ continue }
    $code = Get-Field $h @('ReasonCode','HaltReasonCode')
    $out += [ordered]@{
      symbol              = $sym
      name                = Get-Field $h @('IssueName','CompanyName','Name')
      exchange            = Normalize-Market (Get-Field $h @('Market'))   # UTP 피드는 전 거래소 포함
      haltDate            = Format-Date (Get-Field $h @('HaltDate'))
      haltTime            = Format-Time (Get-Field $h @('HaltTime'))
      reasonCode          = $code
      reason              = $(if($REASONS.ContainsKey($code)){ $REASONS[$code] } else { $code })
      resumptionDate      = Format-Date (Get-Field $h @('ResumptionDate'))
      resumptionQuoteTime = Format-Time (Get-Field $h @('ResumptionQuoteTime'))
      resumptionTradeTime = Format-Time (Get-Field $h @('ResumptionTradeTime'))
      sourceUrl           = $NASDAQ_SRC
    }
  }
  Write-Host "[Nasdaq] $($out.Count)건"
  return $out
}

function Get-NyseHalts {
  Write-Host "[NYSE] $NYSE_CSV 수집..."
  try{
    $resp = Invoke-WebRequest -Uri $NYSE_CSV -Headers @{ 'User-Agent'=$UA } -TimeoutSec 30 -UseBasicParsing
    $raw = $resp.Content
    # .Content가 byte[]로 오는 경우(charset 미지정) 문자열로 디코딩
    if($raw -is [byte[]]){ $content = [System.Text.Encoding]::UTF8.GetString($raw) } else { $content = [string]$raw }
    $content = Remove-Bom $content
    $preview = ($content.Substring(0,[Math]::Min(160,$content.Length)) -replace '\r?\n',' ')
    Write-Host "[진단][NYSE] 응답길이=$($content.Length), 미리보기: $preview"
    $rows = $content | ConvertFrom-Csv
  }catch{ Write-Warning "[NYSE] 수집 실패: $($_.Exception.Message)"; return @() }

  if(-not $rows){ Write-Warning "[NYSE] 행 없음 (정지 없음 또는 CSV 아님 — 위 미리보기 확인)"; return @() }
  Write-Host "[진단][NYSE] CSV 컬럼: $((($rows[0].PSObject.Properties.Name)) -join ' | ')"

  $out=@()
  foreach($r in $rows){
    $h=@{}; foreach($p in $r.PSObject.Properties){ $h[$p.Name]=("$($p.Value)").Trim() }
    $sym = Get-Field $h @('Symbol','Ticker','Symbol ')
    if([string]::IsNullOrWhiteSpace($sym)){ continue }
    $out += [ordered]@{
      symbol              = $sym
      name                = Get-Field $h @('Name','Company','Security Name','Issuer')
      exchange            = Normalize-Market (Get-Field $h @('Exchange','Market','Listing Market'))   # CSV의 실제 상장 거래소 사용(과거 'NYSE' 하드코딩 버그 수정)
      haltDate            = Format-Date (Get-Field $h @('Halt Date','HaltDate','Date'))
      haltTime            = Format-Time (Get-Field $h @('Halt Time','HaltTime','Time'))
      reasonCode          = Get-Field $h @('Reason Code','ReasonCode','Halt Reason Code')
      reason              = Get-Field $h @('Reason','Halt Reason','Reason Description')
      resumptionDate      = Format-Date (Get-Field $h @('Resume Date','Resumption Date','ResumptionDate'))
      resumptionQuoteTime = Format-Time (Get-Field $h @('Resume Quote Time','Resumption Quote Time','Quote Time'))
      resumptionTradeTime = Format-Time (Get-Field $h @('Resume Trade Time','Resumption Trade Time','Trade Time','Resume Time'))
      sourceUrl           = $NYSE_SRC
    }
  }
  Write-Host "[NYSE] $($out.Count)건"
  return $out
}

# ---------- 수집 ----------
$current = @()
$current += Get-NasdaqHalts   # UTP 통합 피드 (전 거래소 커버)
$current += Get-NyseHalts     # 보조 (실패 시 무시)

# 소스 간 중복 제거 (symbol|haltDate|haltTime|exchange)
$seenCur=@{}; $dedup=@()
foreach($h in $current){
  $k="$($h.symbol)|$($h.haltDate)|$($h.haltTime)"
  if(-not $seenCur.ContainsKey($k)){ $seenCur[$k]=$true; $dedup+=$h }
}
# 최신 정지 순으로 정렬 (화면 상단에 오늘 건이 오도록)
$current = $dedup | Sort-Object @{Expression={$_.haltDate};Descending=$true}, @{Expression={$_.haltTime};Descending=$true}

$nowUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
if(-not (Test-Path $DataDir)){ New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }

# ---------- halts.json (현재) ----------
$haltsObj = [ordered]@{
  lastUpdated = $nowUtc
  sources     = [ordered]@{ nasdaq=$NASDAQ_SRC; nyse=$NYSE_SRC }
  halts       = @($current)
}
$haltsPath = Join-Path $DataDir 'halts.json'
($haltsObj | ConvertTo-Json -Depth 8) | Set-Content -Path $haltsPath -Encoding UTF8
Write-Host "→ $haltsPath ($($current.Count)건)"

# ---------- halts_history.json (내역 누적) ----------
$histPath = Join-Path $DataDir 'halts_history.json'
$existing = @()
if(Test-Path $histPath){
  try{ $existing = @((Get-Content $histPath -Raw | ConvertFrom-Json).halts) }catch{ $existing=@() }
}
# 현재건을 내역 스키마(resumptionTime 단일)로 변환
$currentHist = foreach($h in $current){
  [ordered]@{
    symbol=$h.symbol; name=$h.name; exchange=$h.exchange
    haltDate=$h.haltDate; haltTime=$h.haltTime
    reasonCode=$h.reasonCode; reason=$h.reason
    resumptionDate=$h.resumptionDate
    resumptionTime=$(if($h.resumptionTradeTime){$h.resumptionTradeTime}else{$h.resumptionQuoteTime})
    sourceUrl=$h.sourceUrl
  }
}
# 병합 + 중복제거 (symbol|haltDate|haltTime|exchange)
$seen=@{}; $merged=@()
foreach($h in @($existing)+@($currentHist)){
  if($null -eq $h){continue}
  $k = "$($h.symbol)|$($h.haltDate)|$($h.haltTime)"
  if(-not $seen.ContainsKey($k)){ $seen[$k]=$true; $merged += $h }
  else {
    # 기존 항목의 재개정보가 비어있고 새 항목에 있으면 갱신
    if($h.resumptionTime){ ($merged | Where-Object {"$($_.symbol)|$($_.haltDate)|$($_.haltTime)" -eq $k} | Select-Object -First 1).resumptionTime = $h.resumptionTime }
  }
}
# 보존기간 정리 + 정렬(haltDate, haltTime 내림차순)
$cut = (Get-Date).AddDays(-$HistoryRetentionDays).ToString('yyyy-MM-dd')
$merged = $merged | Where-Object { [string]::IsNullOrWhiteSpace($_.haltDate) -or $_.haltDate -ge $cut }
$merged = $merged | Sort-Object @{Expression={$_.haltDate};Descending=$true}, @{Expression={$_.haltTime};Descending=$true}

$histObj = [ordered]@{
  lastUpdated = $nowUtc
  sources     = [ordered]@{ nasdaq=$NASDAQ_SRC; nyse='https://www.nyse.com/api/trade-halts/historical/download' }
  halts       = @($merged)
}
($histObj | ConvertTo-Json -Depth 8) | Set-Content -Path $histPath -Encoding UTF8
Write-Host "→ $histPath (누적 $($merged.Count)건, 보존 ${HistoryRetentionDays}일)"
Write-Host "완료: $nowUtc"
