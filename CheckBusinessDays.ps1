# 必要な変数を初期化
$today = Get-Date
$targetDays = @(3, 5)

# 祝日CSVのURL（内閣府）
$holidayUrl = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"

# 一時ファイルを用意してCSVを保存
$tmpFile = "$env:TEMP\syukujitsu.csv"
Invoke-WebRequest -Uri $holidayUrl -OutFile $tmpFile

# Shift-JISでCSVを読み込みし、日付に変換
$holidays = @()
Import-Csv -Path $tmpFile -Encoding Default | ForEach-Object {
    if (![string]::IsNullOrWhiteSpace($_.日付)) {
        try {
            $date = [datetime]::Parse($_.日付)
            $holidays += $date.Date
        } catch {
            Write-Warning "日付変換に失敗: $_.日付"
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

# 結果を表示したままにする
Write-Host "Enterキーを押すと終了します..." -ForegroundColor Yellow
Read-Host
