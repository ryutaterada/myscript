Add-Type -AssemblyName System.Windows.Forms

# 変数
$rootPath = "E:\work\ps\azcopy"
$width = 600
$height = 800
$azCopyPath = "$($rootPath)\azcopy.exe"
$StorageAccountName = "azcopyttest1481"
$SASKey = ""

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

# BlobStorageからアップロード済みのファイルを収録する関数
function Get-UploadedPstFile {
    # ローカルのPSTファイルを取得

    # ローカルのPSTファイルアップロードログを取得

    # azcopy listコマンドを実行してアップロード済みのPSTファイルを取得
    $destinationPath = "00_User/$($group.Group)/$($address)/00_UserUpload/"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($destinationPath)?$($SASKey)"
    $outputFilePath = "$($rootPath)\temp\output.txt"
    Start-Process -FilePath $azCopyPath -ArgumentList "list `"$SASURL`" --running-tally --machine-readable" -NoNewWindow -Wait -RedirectStandardOutput $outputFilePath
    $output = Get-Content -Path $outputFilePath
    $uploadPstFile = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt $output.Count - 3; $i++) {
        # ファイル情報を取得
        $UploadDate = $output[$i].Split(":")[1].Split("/")[0].Substring(1)
        $UploadFileName = $output[$i].Split(":")[1].Substring(16, $output[$i].Split(":")[1].LastIndexOf(";") - 16)

        # 時間の修正
        $UploadDate = ([DateTime]::ParseExact($UploadDate, "yyyyMMddHHmmss", $null)).ToString("yyyy/MM/dd HH:mm:ss")

        # 情報の突合

        # ファイル情報を追加
        $uploadPstFile.Add([PSCustomObject]@{
                アップロード日時 = $UploadDate
                ファイル名    = $UploadFileName
            })
    }

    # テキストとして表示
    # テキストとして表示
    $uploadPstFile | Out-File -FilePath "C:\path\to\output.txt"
    Invoke-Item -Path "C:\path\to\output.txt"

    $output[$output.Count - 2] # INFO: File count: 2
    $output[$output.Count - 1] # INFO: Total file size: 175
    $maxTotalPstFileSize = 90000000000
    if ([Int64]$output[$output.Count - 1].Split(":")[2].Substring(1) -gt $maxTotalPstFileSize) {
        $str = "アップロード可能なファイルサイズを超えています。\nアップロード済みのファイルを削除するために、移行担当者に連絡してください。\n連絡先 : test@example.com"
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
        Set-ErrorMessage -Message $str
        # return
    }
    Remove-Item -Path $outputFilePath -Force

    return $uploadedPstFile
}

# PSTファイル取得関数を追加
function Get-PstFile {
    # param (
    #     OptionalParameters
    # )
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "ファイルを選択"
    $openFileDialog.Filter = "PSTファイル (*.pst)|*.pst"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        # $selectedFile = $openFileDialog.FileName
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - 選択されたファイル: $($openFileDialog.FileName)"

        if ($openFileDialog.FileName -notlike "*.pst") {
            # PSTファイル以外が選択された場合。新規ファイル作成などの操作により実行可能。
            $str = "PSTファイル以外が選択されました。メールが保存されているPSTファイルを選択してください。"
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
            Set-ErrorMessage -Message $str
            return
        }

        # ファイル情報を取得
        $fileInfo = Get-Item $openFileDialog.FileName

        # ファイル名、ファイル容量、ファイルパスを追加
        $listViewItem = New-Object System.Windows.Forms.ListViewItem($fileInfo.Name)
        $listViewItem.SubItems.Add($fileInfo.Length)
        $listViewItem.SubItems.Add($fileInfo.FullName)
        $listView.Items.Add($listViewItem)
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
    Write-Host "PSTファイルをアップロードします。"

    # メールアドレス分解
    $address = ($emailTextBox1.Text).Split("@")[0]
    Write-Host $address

    # ユーザーリスト読み込み
    $userListLocalPath = "$($rootPath)\temp\UserList.csv"
    $userListBlobPath = "01_Manage/00_Group/UserList.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($userListBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy `"$SASURL`" `"$userListLocalPath`" " -NoNewWindow -Wait
    $userList = Import-Csv -Path $userListLocalPath -Encoding UTF8

    # 帯域制限リスト読み込み
    $trafficListLocalPath = "$($rootPath)\temp\TrafficControl.csv"
    $trafficListBlobPath = "01_Manage/00_Group/TrafficControl.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($trafficListBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy `"$SASURL`" `"$trafficListLocalPath`" " -NoNewWindow -Wait
    $trafficControlList = Import-Csv -Path $trafficListLocalPath -Encoding UTF8

    # ユーザーリストと帯域制限リストを結合
    $group = $userList | Where-Object { $_.Mail -eq $($emailTextBox1.Text) }
    $bpsRate = $trafficControlList | Where-Object { $_.Group -eq $group.Group }
    if ($bpsRate.Speed -ne 0) {
        $NetQoSPolicyName = "AzCopyPolicy01"
        $DSCPAction = 1
        New-NetQosPolicy -Name $NetQoSPolicyName -AppPathNameMatchCondition $azCopyPath -DSCPAction $DSCPAction -ThrottleRateActionBitsPerSecond $bpsRate -Precedence 0
    }

    # ファイルアップロード処理
    foreach ($item in $ItemList) {
        $filePath = $item.SubItems[2].Text
        $time = Get-Date -Format "yyyyMMddHHmmss"
        Start-Sleep -Seconds 1
        $destinationPath = "00_User/$($group.Group)/$($address)/00_UserUpload/$($time)/"
        $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($destinationPath)?$($SASKey)"
        Start-Process -FilePath $azCopyPath -ArgumentList "copy `"$filePath`" `"$SASURL`"" -Wait
    }

    # 作業ファイルおよび設定削除
    Remove-Item -Path $userListLocalPath -Force
    Remove-Item -Path $trafficListLocalPath -Force
    if ($bpsRate.Speed -ne 0) {
        Remove-NetQosPolicy -Name $NetQoSPolicyName -Confirm:$false
    }
}

# メールアドレス一致チェック関数を追加
function Test-EmailAddress {
    # メールアドレスが一致しているかどうかを確認
    if ($emailTextBox1.Text -ne $emailTextBox2.Text) {
        Write-Host "メールアドレスが一致していません。処理をスキップします。"
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
    $errorLabel.Text = $Message
}

function Test-OutlookRunning {
    # Get-ProcessコマンドレットでOutlookのプロセスを取得
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($true) {
        # if ($outlookProcess) {
        $result = [System.Windows.Forms.MessageBox]::Show("Outlookプロセスが実行中です。強制終了しますか？", "警告", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $outlookProcess | ForEach-Object { $_.Kill() }
        }
    }
    # プロセスが存在する場合は$trueを、存在しない場合は$falseを返す
    return ($null -ne $outlookProcess)
}

# フォームを作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "新PC移行用PSTアップロード"
$form.Size = New-Object System.Drawing.Size($width, $height)

# 説明文章を追加
$heightColumn = 10
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "①移行を希望するメールアドレスを入力してください。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# メールアドレス入力欄1
$heightColumn += 25
$emailLabel1 = New-Object System.Windows.Forms.Label
$emailLabel1.Text = "メールアドレス："
$emailLabel1.Location = New-Object System.Drawing.Point(10, $heightColumn)
$emailLabel2.AutoSize = $true
$form.Controls.Add($emailLabel1)

$emailTextBox1 = New-Object System.Windows.Forms.TextBox
$emailTextBox1.Location = New-Object System.Drawing.Point(140, $heightColumn)
$emailTextBox1.Text = "test@example.com" # Set initial value
$emailTextBox1.Width = 390 # Set the width of the text box
$emailTextBox1.Enabled = $false # Make the text box read-only by default
$form.Controls.Add($emailTextBox1)

# メールアドレス入力欄2
$heightColumn += 25
$emailLabel2 = New-Object System.Windows.Forms.Label
$emailLabel2.Text = "メールアドレス(確認用)："
$emailLabel2.Location = New-Object System.Drawing.Point(10, $heightColumn)
$emailLabel2.AutoSize = $true
$form.Controls.Add($emailLabel2)

$emailTextBox2 = New-Object System.Windows.Forms.TextBox
$emailTextBox2.Location = New-Object System.Drawing.Point(140, $heightColumn)
$emailTextBox2.Text = "test@example.com" # Set initial value
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
        # メールアドレスが一致しているかどうかを確認
        if (Test-EmailAddress -eq $true) {
            # 関数を呼び出す処理を追加
            Write-Host "確認ボタンがクリックされました。"
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
        # メールアドレスが一致しているかどうかを確認
        if (Test-EmailAddress -eq $true) {
            # 関数を呼び出す処理を追加
            Write-Host "PST選択ボタンがクリックされました。"
            Get-PstFile
        }
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
        # メールアドレスが一致しているかどうかを確認
        if (Test-EmailAddress -eq $true) {
            # 関数を呼び出す処理を追加
            Write-Host "アップロードボタンがクリックされました。"
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
$errorLabel.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($errorLabel)

# フォームを表示
$form.ShowDialog()
