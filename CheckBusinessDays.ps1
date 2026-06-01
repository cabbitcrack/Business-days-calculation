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

# 今月／翌月の祝日を表示
$currentMonth = $today.Month
$currentYear = $today.Year
$todayDate = $today.Date
$nameColumn = ($csvData[0].PSObject.Properties | Select-Object -Skip 1 -First 1).Name

# 指定した年月の祝日レコードを取得するヘルパー
function Get-HolidaysInMonth {
    param(
        [int]$year,
        [int]$month
    )
    return $csvData | Where-Object {
        $dateStr = $_.$dateColumn
        if (![string]::IsNullOrWhiteSpace($dateStr)) {
            try {
                $d = [datetime]::ParseExact($dateStr, 'yyyy/M/d', $null)
                return ($d.Year -eq $year -and $d.Month -eq $month)
            } catch { return $false }
        }
        return $false
    }
}

$thisMonthHolidays = Get-HolidaysInMonth -year $currentYear -month $currentMonth

# 今月の祝日のうち、まだ過ぎていないものがあるか確認
$upcomingThisMonth = $thisMonthHolidays | Where-Object {
    $d = [datetime]::ParseExact($_.$dateColumn, 'yyyy/M/d', $null)
    $d.Date -ge $todayDate
}

if ($upcomingThisMonth) {
    # 今月にまだ過ぎていない祝日がある場合 → 今月の祝日を表示
    Write-Host ""
    Write-Host "--- 今月（${currentYear}年${currentMonth}月）の祝日 ---" -ForegroundColor Cyan
    foreach ($h in $thisMonthHolidays) {
        $d = [datetime]::ParseExact($h.$dateColumn, 'yyyy/M/d', $null)
        $name = $h.$nameColumn
        Write-Host "  $($d.ToString('MM/dd（ddd）')) $name" -ForegroundColor Green
    }
    Write-Host ""
} else {
    # 今月の祝日が全て過ぎている（または無い）場合 → 月末の週なら翌月の祝日を表示
    # 月末の週の判定: 今日が属する週（月曜～日曜）に月末が含まれるか
    $daysToSunday = (7 - [int]$today.DayOfWeek) % 7
    $endOfWeek = $today.Date.AddDays($daysToSunday)
    $lastDayOfMonth = (Get-Date -Year $currentYear -Month $currentMonth -Day 1).AddMonths(1).AddDays(-1).Date
    $isLastWeekOfMonth = $endOfWeek -ge $lastDayOfMonth

    if ($isLastWeekOfMonth) {
        $nextMonthDate = $today.AddMonths(1)
        $nextMonth = $nextMonthDate.Month
        $nextYear = $nextMonthDate.Year
        $nextMonthHolidays = Get-HolidaysInMonth -year $nextYear -month $nextMonth

        Write-Host ""
        Write-Host "--- 翌月（${nextYear}年${nextMonth}月）の祝日 ---" -ForegroundColor Cyan
        if ($nextMonthHolidays) {
            foreach ($h in $nextMonthHolidays) {
                $d = [datetime]::ParseExact($h.$dateColumn, 'yyyy/M/d', $null)
                $name = $h.$nameColumn
                Write-Host "  $($d.ToString('MM/dd（ddd）')) $name" -ForegroundColor Green
            }
        } else {
            Write-Host "  翌月に祝日はありません" -ForegroundColor Yellow
        }
        Write-Host ""
    }
}

# 結果を表示したままにする
Write-Host "Enterキーを押すと終了します..." -ForegroundColor Yellow
Read-Host
