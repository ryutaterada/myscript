# Description: イベントログの内容を比較し、既知のエラー以外を抽出する
# Author: Ryuta Terada (JBS)
# Last Modified: 2023/12/20
# Environment: Windows Server 2016 Standard
# PowerShell: 5.1.17763.1490
# Usage: PowerShellで実行する
# Note: 既知のエラーは未知のエラーからCSVファイルに転記する
#      既知のエラーのCSVファイルは以下の形式で記載する
#      Id,ProviderName,説明
#      1000,KerberosError,認証エラーだが動作に影響はない
#      1001,Microsoft-Windows-User Profiles Service,ユーザープロファイルの削除に失敗したが、動作に影響はない

# 変数
# 作業フォルダパス
$path = "E:\work\python"
# 時刻情報
$time = Get-Date -Format "yyyyMMddHHmmss"
# トランスクリプトファイル名
$transcriptName = "Transcript_$($time).txt"
# トランスクリプトパス
$transcriptPath = "$($path)\Transcript\$($transcriptName)"
# 既知のイベントログリスト名
$csvName = "EventList.csv"
# 既知のイベントログリストパス
$csvPath = "$($path)\$($csvName)"
# 未知のイベントログリスト名
$unknownLogName = "UnknownEventList_$($time).csv"
# 未知のイベントログリストパス
$unknownLogPath = "$($path)\$($unknownLogName)"
# 未知のイベントログリスト
$unknownLogList = [System.Collections.ArrayList]@()
# イベントログ取得件数
$eventCount = 10000
# イベントログ名
$logName = "Application", "System"
# イベントログレベル
$logLevel = 3

# トランスクリプト用フォルダがなければ作成
if (!(Test-Path -Path "$($path)\Transcript")) {
    New-Item -Path "$($path)\Transcript" -ItemType Directory
}
# トランスクリプト開始
Start-Transcript -Path $transcriptPath

# イベントログ取得
$eventLogs = Get-WinEvent -LogName $logName -MaxEvents $eventCount | Where-Object { $_.Level -le $logLevel }

# 既知のイベントログリストを読み込む
$csvData = Import-Csv -Path $csvPath -Encoding utf8

# イベントログの内容を比較
foreach ($log in $eventLogs) {
    # イベントIDおよびソースが一致しないログを抽出
    if (!($log.Id -in $csvData.Id -and $log.ProviderName -in $csvData.ProviderName) -and !($log.Id -in $unknownLogList.Id -and $log.ProviderName -in $unknownLogList.ProviderName)) {
        # 未知のエラー
        Write-Host "未知のエラー: イベントID $($log.Id) - ソース $($log.ProviderName)"
        $unknownLogList.Add($log) | Out-Null
    }
    elseif ($log.Id -in $unknownLogList.Id -and $log.ProviderName -in $unknownLogList.ProviderName) {
        # 登録済みのエラー
        # Write-Host "登録済みのエラー: イベントID $($log.Id) - ソース $($log.ProviderName)"
    }
    # elseif ($log.Id -in $csvData.Id -and $log.ProviderName -in $csvData.ProviderName) {
    #     # 既知のエラー
    #     Write-Host "既知のエラー: イベントID $($log.Id) - ソース $($log.ProviderName)"
    # }
    # else {
    #     # その他
    #     Write-Error "不正な条件式(デバッグ用)"
    # }
}

# 未知のエラーをリストとして出力
if ($unknownLogList.Count -gt 0) {
    $unknownLogList | Export-Csv -Path $unknownLogPath -NoTypeInformation -Encoding utf8 -Append
}

# 情報出力
Write-Host "イベントログ取得件数: $($eventCount)"
Write-Host "エラーログ件数: $($eventLogs.Count)"
Write-Host "未知のエラーログ件数: $($unknownLogList.Count)"

# トランスクリプト終了
Stop-Transcript

# 待機画面表示
timeout.exe -1
