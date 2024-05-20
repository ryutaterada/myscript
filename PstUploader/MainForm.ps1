Add-Type -AssemblyName System.Windows.Forms

# 変数
$rootPath = "E:\work\ps\azcopy"
# $rootPath = "C:\work\PSTUploader"
$width = 600
$height = 800
$azCopyPath = "$($rootPath)\azcopy.exe"
$StorageAccountName = "azcopyttest1481"
$SASKey = "sp=racwl&st=2024-03-10T14:04:32Z&se=2024-11-12T22:04:32Z&spr=https&sv=2022-11-02&sr=c&sig=Ld7Nbm9bhwMDRhbGUsGTWhb1BBi%2Fe0h9ydXQSm6eCL4%3D"

# フォルダ作成処理
$worFolderList = ("$($rootPath)\temp", "$($rootPath)\Log", "$($rootPath)\Output")
foreach ($workFolderPath in $worFolderList) {
    if (-Not (Test-Path -Path $workFolderPath)) {
        New-Item -ItemType Directory -Path $workFolderPath > $null
    }
}

# Transcript取得
$transcriptPath = "$($rootPath)\Log\Transcript_$(Get-Date -Format "yyyyMMddHHmmss").txt"
Start-Transcript -Path $transcriptPath

# メールアドレス取得関数を追加
function Get-OutlookEmailAddress {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject "Outlook.Application"
    $namespace = $outlook.GetNamespace("MAPI")
    $accounts = $namespace.Accounts
    $emailAddresses = @()
    foreach ($account in $accounts) {
        $emailAddresses += $account.SmtpAddress
    }
    $outlook.Quit()
    return $emailAddresses
}

# ローカルおよびBlobStorageからアップロード済みのファイルを検索する処理
function Get-UploadedPstFile {
    # メールアドレス分解
    $address = ($emailTextBox1.Text).Split("@")[0]

    # 検索対象フォルダを指定
    $localFolders = @(
        "C:\",
        "D:\"
    )

    # PSTカウント変数
    $totalCount = 0
    $totalSize = 0
    $LocalPSTInfoList = [System.Collections.ArrayList]@()

    # 出力フォルダパス
    $localPSTListPath = "$($rootPath)\Output\PCData.txt"
    $blobPstFilePath = "$($rootPath)\Output\BlobPstFileList_$($address).txt"

    # ユーザーリスト読み込み
    $userListLocalPath = "$($rootPath)\temp\UserList.csv"
    $userListBlobPath = "01_Manage/00_Group/UserList.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($userListBlobPath)?$($SASKey)"
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 移行対象ユーザーリストを取得します。"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($userListLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 移行対象ユーザーリストの取得完了。"
    $userList = Import-Csv -Path $userListLocalPath -Encoding UTF8
    $group = $userList | Where-Object { $_.Mail -eq $($emailTextBox1.Text) }

    # 所属グループが存在しない場合 エラーメッセージを表示し、フォルダ検索処理は継続する。
    if ($null -eq $group.Group) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - あなたのメールアドレスが移行対象者として登録されていません。"
        Set-ErrorMessage -Message "あなたのメールアドレスが移行対象者として登録されていません。"
    }

    # PSTファイルを検索
    foreach ($localFolder in $localFolders) {
        # 検索処理
        $localPSTFiles = Get-ChildItem -Path $localFolder -Filter "*.pst" -File -Recurse -ErrorAction SilentlyContinue

        # それぞれのPSTファイルの情報を取得
        foreach ($localPSTFile in $localPSTFiles) {
            # リストに追加
            $LocalPSTInfoList.Add([PSCustomObject]@{
                    Name     = $localPSTFile.Name
                    FilePath = $localPSTFile.FullName
                    Size     = $localPSTFile.Length
                    Owner    = (Get-Acl -Path $localPSTFile.FullName).Owner
                }) > $null

            # カウント変数の更新
            $totalCount++
            $totalSize += $size
        }

        Clear-Variable -Name localPSTFiles
    }

    # ローカルPCデータの出力
    $LocalPSTList = @"
$hostname
$ipAddress
$totalCount
$totalSize
"@
    foreach ($LocalPSTInfo in $LocalPSTInfoList) {
        $LocalPSTList += "$($LocalPSTInfo.Name) $($LocalPSTInfo.FolderPath) $($LocalPSTInfo.Size) $($LocalPSTInfo.Owner)`n"
    }
    $LocalPSTList | Out-File -FilePath $localPSTListPath -Encoding UTF8

    # azcopy listコマンドを実行してアップロード済みのPSTファイルを取得
    $destinationPath = "00_User/$($group.Group)/$($address)/00_UserUpload/"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($destinationPath)?$($SASKey)"
    $uploadedFileListPath = "$($rootPath)\temp\UploadedFileList.txt"
    Start-Process -FilePath $azCopyPath -ArgumentList "list $($SASURL) --running-tally --machine-readable" -NoNewWindow -Wait -RedirectStandardOutput $uploadedFileListPath
    $output = Get-Content -Path $uploadedFileListPath
    $blobPstFile = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt $output.Count - 3; $i++) {
        # ファイル情報を取得
        $blobUploadDate = $output[$i].Split(":")[1].Split("/")[0].Substring(1)
        $blobUploadFileName = $output[$i].Split(":")[1].Substring(16, $output[$i].Split(":")[1].LastIndexOf(";") - 16)
        $blobUploadFileSize = $output[$i].Split(":")[2].Substring(1)

        # 時間の修正
        $blobUploadDate = ([DateTime]::ParseExact($blobUploadDate, "yyyyMMddHHmmss", $null)).ToString("yyyy/MM/dd HH:mm:ss")

        # Blobのファイル名とローカルのファイル名を比較
        # $LocalPSTIndex = $LocalPSTInfoList.Name.IndexOf($blobUploadFileName)

        # # アップロード済みかどうかを判定
        # if (
        #     $blobUploadFileName -in $LocalPSTInfoList.Name -and
        #     $blobUploadFileSize -in $LocalPSTInfoList[$LocalPSTIndex].Size -and
        #     $LocalPSTIndex -ne -1) {
        #     $flag = "クラウドにアップロード済み"
        #     $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTIndex].FilePath
        # }
        # elseif (
        #     $blobUploadFileName -in $LocalPSTInfoList.Name -and
        #     [Int64]$blobUploadFileSize -lt [Int64]$LocalPSTInfoList[$LocalPSTIndex].Size -and
        #     $LocalPSTIndex -ne -1) {
        #     $flag = "クラウドにアップロード済み&容量増加"
        #     $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTIndex].FilePath
        # }
        # elseif (
        #     $blobUploadFileName -in $LocalPSTInfoList.Name -and
        #     [Int64]$blobUploadFileSize -gt [Int64]$LocalPSTInfoList[$LocalPSTIndex].Size -and
        #     $LocalPSTIndex -ne -1) {
        #     $flag = "クラウドにアップロード済み&容量減少(別ファイルまたはバックアップファイル)"
        #     $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTIndex].FilePath
        # }
        # else {
        #     $flag = "クラウドにアップロード済み&PC内にファイルが存在しない"
        #     $blobUploadFilePath = "PC内にファイルが存在しない"
        # }

        # アップロード済みかどうかを判定
        if ($blobUploadFileName -in $LocalPSTInfoList.Name) {
            $flag = "クラウドにアップロード済み"
            $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTInfoList.Name.IndexOf($blobUploadFileName)].FilePath
        }
        else {
            $flag = "クラウドにアップロード済み&PC内にファイルが存在しない"
            $blobUploadFilePath = "PC内にファイルが存在しない"
        }

        # ファイル情報を追加
        $blobPstFile.Add([PSCustomObject]@{
                メールアドレス  = $address
                ステータス    = $flag
                ファイル名    = $blobUploadFileName
                ファイルパス   = $blobUploadFilePath
                ファイル容量   = "{0:N1}" -f ([Int64]$blobUploadFileSize / 1024 / 1024) + "MB"
                アップロード日時 = $blobUploadDate
            }) > $null

        Clear-Variable -Name flag, blobUploadFilePath
    }

    # ローカルにあってクラウドにないファイルの突合
    foreach ($LocalPSTInfo in $LocalPSTInfoList) {
        if (-Not($LocalPSTInfo.Name -in $blobPstFile.Name)) {
            $blobPstFile.Add([PSCustomObject]@{
                    メールアドレス  = $address
                    ステータス    = "未アップロード"
                    ファイル名    = $LocalPSTInfo.Name
                    ファイルパス   = $LocalPSTInfo.FilePath
                    ファイル容量   = "{0:N1}" -f ([Int64]$LocalPSTInfo.Size / 1024 / 1024) + "MB"
                    アップロード日時 = ""
                }) > $null
        }
    }

    # テキストとして保存&表示
    $blobPstFile | Out-File -FilePath $blobPstFilePath -Encoding UTF8
    Invoke-Item -Path $blobPstFilePath

    $maxTotalPstFileSize = 90000000000
    if ([Int64]$output[$output.Count - 1].Split(":")[2].Substring(1) -gt $maxTotalPstFileSize) {
        $str = @"
アップロード可能なファイルサイズ(90GB)を超えています。
アップロード済みのファイルを削除する場合は、移行担当者へ連絡をお願いします。
"@
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
        Set-ErrorMessage -Message $str
        return
    }

    Set-SystemMessage -Message "アップロード状況確認処理が完了しました。"
    return
}

# PSTファイル取得関数を追加
function Get-PstFile {
    # サイズ制限設定ファイル読み込み
    $sizeLimitConfigLocalPath = "$($rootPath)\temp\SizeLimitConfig"
    $sizeLimitConfigBlobPath = "01_Manage/01_SizeLimit/SizeLimitConfig"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($sizeLimitConfigBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($sizeLimitConfigLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    $sizeLimitConfig = Get-Content -Path $sizeLimitConfigLocalPath -Encoding UTF8

    # ファイル選択ダイアログを表示
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "ファイルを選択"
    $openFileDialog.Filter = "PSTファイル (*.pst)|*.pst"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments") + "\Outlook データファイル"

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        # ファイル情報を取得
        $fileInfo = Get-Item $openFileDialog.FileName
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 選択されたファイル: $($openFileDialog.FileName)"

        # PSTファイル以外が選択された場合
        if ($openFileDialog.FileName -notlike "*.pst") {
            # PSTファイル以外が選択された場合。新規ファイル作成などの操作により実行可能。
            $str = "PSTファイル以外が選択されました。メールが保存されているPSTファイルを選択してください。"
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
            Set-ErrorMessage -Message $str
            return
        }

        # サイズ制限確認
        if ([Int64]$fileInfo.Length -gt ([Int64]$sizeLimitConfig * 1024 * 1024 * 1024)) {
            $str = "ファイルサイズが制限を超えています。$($sizeLimitConfig)GB以上のPSTファイルは本ツールでの移行対象外です。"
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
            Set-ErrorMessage -Message $str
            return
        }

        # ファイル名、ファイル容量、ファイルパスを追加 "{0:N1}" -f ([Int64]$fileInfo.Length / 1024 / 1024) + "MB"
        $listViewItem = New-Object System.Windows.Forms.ListViewItem($fileInfo.Name)
        $listViewItem.SubItems.Add("{0:N1}" -f ([Int64]$fileInfo.Length / 1024 / 1024) + "MB")
        $listViewItem.SubItems.Add($fileInfo.FullName)
        $listView.Items.Add($listViewItem)
        Set-SystemMessage -Message "ファイルが選択されました。"
        return
    }
    else {
        # ファイル選択画面を閉じた場合
        $str = "「ファイル選択ボタン」が押されましたが、ファイルが選択されませんでした。"
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
        Set-ErrorMessage -Message $str
        return
    }
}

# PSTアップロード関数を追加
function Invoke-PstFileUpload {
    param (
        [System.Windows.Forms.ListView+ListViewItemCollection]$ItemList
    )

    # メールアドレス分解
    $address = ($emailTextBox1.Text).Split("@")[0]

    # アップロードログの保存
    $uploadedPstLogFilePath = "$($rootPath)\Log\UploadedPstLog_$($address).csv"

    # ユーザーリスト読み込み
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 設定ファイルを読み込みます。"
    $userListLocalPath = "$($rootPath)\temp\UserList.csv"
    $userListBlobPath = "01_Manage/00_Group/UserList.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($userListBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($userListLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    $userList = Import-Csv -Path $userListLocalPath -Encoding UTF8
    $group = $userList | Where-Object { $_.Mail -eq $($emailTextBox1.Text) }

    # 所属グループが存在しない場合
    if ($null -eq $group.Group) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - あなたのメールアドレスが移行対象者として登録されていません。アップロード処理を中断します。"
        Set-ErrorMessage -Message @"
あなたのメールアドレスが移行対象者として登録されていません。
アップロード処理を中断します。
"@
        return
    }

    # 帯域制限リスト読み込み
    $trafficListLocalPath = "$($rootPath)\temp\TrafficControl.csv"
    $trafficListBlobPath = "01_Manage/00_Group/TrafficControl.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($trafficListBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($trafficListLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    $trafficControlList = Import-Csv -Path $trafficListLocalPath -Encoding UTF8
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 設定ファイルの読み込み完了。"
    # 帯域制限設定
    $bpsRate = $trafficControlList | Where-Object { $_.Group -eq $group.Group }
    $NetQoSPolicyName = "AzCopyPolicy01"
    $DSCPAction = 1
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - ネットワーク事前設定を実行。"
    if (-Not(Get-NetQosPolicy -Name $NetQoSPolicyName -ErrorAction SilentlyContinue)) {
        New-NetQosPolicy -Name $NetQoSPolicyName -AppPathNameMatchCondition $azCopyPath -DSCPAction $DSCPAction -ThrottleRateActionBitsPerSecond $bpsRate -Precedence 0
    }

    # ファイルアップロード処理
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 選択されたファイルのアップロード処理開始。"
    foreach ($item in $ItemList) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($item.SubItems[0].Text)のアップロードを開始。"
        $filePath = $item.SubItems[2].Text
        $time = Get-Date -Format "yyyyMMddHHmmss"
        Start-Sleep -Seconds 1
        $destinationPath = "00_User/$($group.Group)/$($address)/00_UserUpload/$($time)/"
        $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($destinationPath)?$($SASKey)"
        try {
            Start-Process -FilePath $azCopyPath -ArgumentList "copy $($filePath) $($SASURL)" -Wait
            [PSCustomObject]@{
                Address  = $address
                Time     = $time
                FilePath = $item.SubItems[2].Text
                FileSize = $item.SubItems[1].Text
            } | Export-Csv -Path $uploadedPstLogFilePath -Append -Encoding UTF8 -NoTypeInformation
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($item.SubItems[0].Text)のアップロードが完了。"
        }
        catch {
            # エラーが発生した場合 ×ボタンによるキャンセルかどうかの判定処理を追加
            if ($_.Exception.Message -eq "The user closed the window.") {
                Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - ファイルのアップロードがキャンセルされました。"
                Set-SystemMessage -Message "ファイルのアップロードがキャンセルされました。"
                return
            }
            else {
                $errorMessage = $_.Exception.Message
                Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - エラーが発生しました: $($errorMessage)"
                Set-ErrorMessage -Message "アップロードエラーが発生しました。"
                return
            }
        }
    }
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 選択されたファイルのアップロード処理完了。"
    Set-SystemMessage -Message "選択されたファイルのアップロード処理が完了しました。"

    # ネットワーク設定削除
    Remove-NetQosPolicy -Name $NetQoSPolicyName -Confirm:$false

    # リストクリア
    $listView.Items.Clear()
}

# メールアドレス一致チェック関数を追加
function Test-EmailAddress {
    # メールアドレスが一致しているかどうかを確認
    if ($emailTextBox1.Text -ne $emailTextBox2.Text) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - メールアドレスが一致していません。処理をスキップします。"
        # エラーメッセージを表示
        Set-ErrorMessage -Message "メールアドレスが一致していません。メールアドレス入力欄を修正してください。"
        return $false
    }
    return $true
}

# エラーメッセージの設定
function Set-ErrorMessage {
    param (
        [string]$Message
    )
    $errorLabel.ForeColor = [System.Drawing.Color]::Red
    $errorLabel.Text = $Message
}

# システムメッセージの設定
function Set-SystemMessage {
    param (
        [string]$Message
    )
    $errorLabel.ForeColor = [System.Drawing.Color]::Black
    $errorLabel.Text = $Message
}

function Test-OutlookRunning {
    # Get-ProcessコマンドレットでOutlookのプロセスを取得
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($null -ne $outlookProcess) {
        $str = "Outlookが実行中です。終了してください。"
        [System.Windows.Forms.MessageBox]::Show($str, "警告", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Set-ErrorMessage -Message $str
        return $false
    }
    # プロセスが存在しない場合は$trueを、存在する場合は$falseを返す
    return $true
}

# 起動時Outlook確認
Test-OutlookRunning > $null

# フォームを作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "新PCメールデータ移行ツール"
$form.Size = New-Object System.Drawing.Size($width, $height)

# 説明文章を追加
$heightColumn = 10
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "①移行を希望するメールアドレスを入力してください。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.Width = 300
$form.Controls.Add($descriptionLabel)

# メールアドレス入力欄1
$heightColumn += 25
$emailLabel1 = New-Object System.Windows.Forms.Label
$emailLabel1.Text = "メールアドレス："
$emailLabel1.Location = New-Object System.Drawing.Point(10, $heightColumn)
$emailLabel1.Width = 140
$form.Controls.Add($emailLabel1)

$emailTextBox1 = New-Object System.Windows.Forms.TextBox
$emailTextBox1.Location = New-Object System.Drawing.Point(150, $heightColumn)
$emailTextBox1.Text = Get-OutlookEmailAddress # Set initial value
$emailTextBox1.Width = 390 # Set the width of the text box
$emailTextBox1.Enabled = $false # Make the text box read-only by default
$form.Controls.Add($emailTextBox1)

# メールアドレス入力欄2
$heightColumn += 25
$emailLabel2 = New-Object System.Windows.Forms.Label
$emailLabel2.Text = "メールアドレス(確認用)："
$emailLabel2.Location = New-Object System.Drawing.Point(10, $heightColumn)
$emailLabel2.Width = 140
$form.Controls.Add($emailLabel2)

$emailTextBox2 = New-Object System.Windows.Forms.TextBox
$emailTextBox2.Location = New-Object System.Drawing.Point(150, $heightColumn)
$emailTextBox2.Text = $emailTextBox1.Text # Set initial value
$emailTextBox2.Width = 390 # Set the width of the text box
$emailTextBox2.Enabled = $false # Make the text box read-only by default
$form.Controls.Add($emailTextBox2)

# チェックボックス
$heightColumn += 25
$checkBoxLabel = New-Object System.Windows.Forms.Label
$checkBoxLabel.Text = "メールアドレス編集ボタン："
$checkBoxLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$checkBoxLabel.Width = 140 # Set the width of the label
$form.Controls.Add($checkBoxLabel)

$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Location = New-Object System.Drawing.Point(150, $heightColumn)
$form.Controls.Add($checkBox)

# チェックボックスのイベントハンドラ
$checkBox.Add_CheckedChanged({
        if ($checkBox.Checked) {
            $emailTextBox1.Enabled = $true
            $emailTextBox2.Enabled = $true
        }
        else {
            $emailTextBox1.Enabled = $false
            $emailTextBox2.Enabled = $false
        }
    })

# 警告ラベル
$heightColumn += 25
$warningLabel = New-Object System.Windows.Forms.Label
$warningLabel.Text = "エラー：メールアドレスが一致しません"
$warningLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$warningLabel.AutoSize = $true
$warningLabel.ForeColor = [System.Drawing.Color]::Red
$warningLabel.Visible = $false
$form.Controls.Add($warningLabel)

# メールアドレスの一致チェックのイベントハンドラ
$emailTextBox1.Add_TextChanged({
        if ($emailTextBox1.Text -ne $emailTextBox2.Text) {
            $warningLabel.Visible = $true
        }
        else {
            $warningLabel.Visible = $false
        }
    })

$emailTextBox2.Add_TextChanged({
        if ($emailTextBox1.Text -ne $emailTextBox2.Text) {
            $warningLabel.Visible = $true
        }
        else {
            $warningLabel.Visible = $false
        }
    })


# 説明文章を追加
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "②「確認」を押すとアップロード済みのPSTファイルおよびPC内のPSTファイルを確認できます。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# 確認ボタン
$heightColumn += 25
$confirmButton = New-Object System.Windows.Forms.Button
$confirmButton.Text = "確認"
$confirmButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($confirmButton)
# ボタンがクリックされたときの処理
$confirmButton.Add_Click({
        # メールアドレス一致確認 & Outlook実行確認
        if (Test-EmailAddress -eq $true -and Test-OutlookRunning -eq $true) {
            # 関数を呼び出す処理を追加
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 確認ボタンがクリックされました。"
            Get-UploadedPstFile
        }
    })

# 説明文章を追加
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "③「ファイル選択」を押すとアップロードするPSTファイルを選択することができます。" + "`r`n" + "1回のアップロードで同時に10ファイルまで可能。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# ファイル選択ボタン
$heightColumn += 40
$pstButton = New-Object System.Windows.Forms.Button
$pstButton.Text = "ファイル選択"
$pstButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($pstButton)
# ボタンがクリックされたときの処理
$pstButton.Add_Click({
        # メールアドレス一致確認 & Outlook実行確認
        if (Test-EmailAddress -eq $true -and Test-OutlookRunning -eq $true) {
            # 関数を呼び出す処理を追加
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - ファイル選択ボタンがクリックされました。"
            Get-PstFile
        }
    })

# クリアボタン
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "クリア"
$clearButton.Location = New-Object System.Drawing.Point(100, $heightColumn)
$form.Controls.Add($clearButton)
# ボタンがクリックされたときの処理
$clearButton.Add_Click({
        $listView.Items.Clear()
        Set-SystemMessage -Message "選択されたファイルがクリアされました。"
    })

# １０行のリスト
$heightColumn += 25
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, $heightColumn)
$listView.Size = New-Object System.Drawing.Size(560, 150)
$listView.View = [System.Windows.Forms.View]::Details
$listView.Columns.Add("ファイル名", 200) > $null
$listView.Columns.Add("ファイル容量", 100) > $null
$listView.Columns.Add("ファイルパス", 250) > $null
$form.Controls.Add($listView)

# 説明文章を追加
$heightColumn += 180
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "④「アップロード」を押すと選択されたPSTファイルをアップロードすることができます。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# アップロードボタン
$heightColumn += 25
$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "アップロード"
$uploadButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($uploadButton)
# ボタンがクリックされたときの処理
$uploadButton.Add_Click({
        # メールアドレス一致確認 & Outlook実行確認
        if (Test-EmailAddress -eq $true -and Test-OutlookRunning -eq $true) {
            # 関数を呼び出す処理を追加
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - アップロードボタンがクリックされました。"
            Invoke-PstFileUpload -ItemList $listView.Items
        }
    })

# 説明文章を追加
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "○メッセージ表示欄"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# エラーメッセージを表示
$heightColumn += 25
$errorLabel = New-Object System.Windows.Forms.Label
$errorLabel.Text = ""
$errorLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$errorLabel.AutoSize = $true
# $errorLabel.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($errorLabel)

# フォームが閉じられたことを検知するイベントハンドラ
$form.Add_FormClosed({
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - フォームが閉じられました。"
        Stop-Transcript
        Stop-Process -Id $PID
    })

# AzCopyの存在確認
if ( (Test-Path -Path $azCopyPath)) {
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - AzCopyモジュールが見つかりません。"
    Set-ErrorMessage -Message "AzCopyモジュールが見つかりません。作業が実施できません。"
}

# フォームを表示
$form.ShowDialog()
