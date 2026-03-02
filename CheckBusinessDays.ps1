# 必要な変数を初期化
$today = Get-Date
$targetDays = @(3, 5)

# TLS 1.2を有効化（HTTPS通信のため）
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 祝日CSVのURL（内閣府）
$holidayUrl = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"

# 一時ファイルを用意してCSVを保存
$tmpFile = "$env:TEMP\syukujitsu.csv"
$tmpFileUtf8 = "$env:TEMP\shukujitsu_utf8.csv"
Invoke-WebRequest -Uri $holidayUrl -OutFile $tmpFile

# Shift-JIS → UTF-8 に変換してから読み込み
$sjis = [System.Text.Encoding]::GetEncoding('shift_jis')
$content = [System.IO.File]::ReadAllText($tmpFile, $sjis)
[System.IO.File]::WriteAllText($tmpFileUtf8, $content, [System.Text.Encoding]::UTF8)

$holidays = @()
$csvData = Import-Csv -Path $tmpFileUtf8 -Encoding UTF8

# カラム名を動的に取得（1列目が日付列）
$dateColumn = ($csvData[0].PSObject.Properties | Select-Object -First 1).Name
$csvData | ForEach-Object {
    $dateStr = $_.$dateColumn
    if (![string]::IsNullOrWhiteSpace($dateStr)) {
        try {
            # フォーマットを明示してパース
            $date = [datetime]::ParseExact($dateStr, 'yyyy/M/d', $null)
            $holidays += $date.Date
        } catch {
            Write-Warning "日付変換に失敗: $dateStr"
        }
    }
}

# 営業日を計算する関数
function Get-BusinessDate {
    param (
        [datetime]$startDate,
        [int]$businessDays
    )
    $currentDate = $startDate
    $count = 0
    while ($count -lt $businessDays) {
        $currentDate = $currentDate.AddDays(1)
        if ($currentDate.DayOfWeek -in 'Saturday', 'Sunday') { continue }
        if ($holidays -contains $currentDate.Date) { continue }
        $count++
    }
    return $currentDate
}

# 結果を表示
foreach ($days in $targetDays) {
    $resultDate = Get-BusinessDate -startDate $today -businessDays $days
    Write-Output "$days 営業日後: $($resultDate.ToString('yyyy年MM月dd日（ddd）'))"
}

# 今月の祝日を表示
$currentMonth = $today.Month
$currentYear = $today.Year
$nameColumn = ($csvData[0].PSObject.Properties | Select-Object -Skip 1 -First 1).Name
$thisMonthHolidays = $csvData | Where-Object {
    $dateStr = $_.$dateColumn
    if (![string]::IsNullOrWhiteSpace($dateStr)) {
        try {
            $d = [datetime]::ParseExact($dateStr, 'yyyy/M/d', $null)
            return ($d.Year -eq $currentYear -and $d.Month -eq $currentMonth)
        } catch { return $false }
    }
    return $false
}

Write-Host ""
Write-Host "--- 今月（${currentYear}年${currentMonth}月）の祝日 ---" -ForegroundColor Cyan
if ($thisMonthHolidays) {
    foreach ($h in $thisMonthHolidays) {
        $d = [datetime]::ParseExact($h.$dateColumn, 'yyyy/M/d', $null)
        $name = $h.$nameColumn
        Write-Host "  $($d.ToString('MM/dd（ddd）')) $name" -ForegroundColor Green
    }
} else {
    Write-Host "  今月は祝日がありません。" -ForegroundColor Gray
}
Write-Host ""

Write-Host "Enterキーを押すと終了します..." -ForegroundColor Yellow
Read-Host
