Add-Type -AssemblyName System.Windows.Forms

# 変数
$width = 600
$height = 800
# $widthColumn = 10
$heightColumn = 10

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


# PSTファイル取得関数を追加
function Get-PstFile {
    # param (
    #     OptionalParameters
    # )
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "ファイルを選択"
    $openFileDialog.Filter = "PSTファイル (*.pst)|*.pst"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Documents")
    $openFileDialog.MaxFileSize = 25 * 1024 * 1024 * 1024 # 25GB

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $selectedFile = $openFileDialog.FileName
        Write-Host "選択されたファイル: $selectedFile"

        # ファイル情報を取得
        $fileInfo = Get-Item $selectedFile

        # ファイル名、ファイル容量、ファイルパスを追加
        $listViewItem = New-Object System.Windows.Forms.ListViewItem($fileInfo.Name)
        $listViewItem.SubItems.Add($fileInfo.Length)
        $listViewItem.SubItems.Add($fileInfo.FullName)
        $listView.Items.Add($listViewItem)
    }
}

# PSTアップロード関数を追加
function Invoke-PstFileUpload {
    param (
        [System.Windows.Forms.ListView+ListViewItemCollection]$ItemList
    )
    Write-Host "PSTファイルをアップロードします。"

    # メールアドレス分解
    $mail = "test01@example.com"
    $address = $mail.Split("@")[0]
    # ネットワーク帯域制御確認
    # ファイルを読み込む処理を追加
    $azCopyPath = "E:\work\ps\azcopy\azcopy.exe"
    $sourcePath = "E:\work\ps\azcopy\folder\User02\UserList.csv"
    $destinationPath = "Manage/Network/UserList.csv"
    $SASURL = "https://azcopyttest1481.blob.core.windows.net/migrationwiz/$($destinationPath)?sp=racwl&st=2024-03-10T14:04:32Z&se=2024-11-12T22:04:32Z&spr=https&sv=2022-11-02&sr=c&sig=Ld7Nbm9bhwMDRhbGUsGTWhb1BBi%2Fe0h9ydXQSm6eCL4%3D"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy `"$SASURL`" `"$sourcePath`" " -NoNewWindow -Wait
    $sourcePath = "E:\work\ps\azcopy\folder\User02\TrafficControl.csv"
    $destinationPath = "Manage/Network/TrafficControl.csv"
    $SASURL = "https://azcopyttest1481.blob.core.windows.net/migrationwiz/$($destinationPath)?sp=racwl&st=2024-03-10T14:04:32Z&se=2024-11-12T22:04:32Z&spr=https&sv=2022-11-02&sr=c&sig=Ld7Nbm9bhwMDRhbGUsGTWhb1BBi%2Fe0h9ydXQSm6eCL4%3D"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy `"$SASURL`" `"$sourcePath`" " -NoNewWindow -Wait

    # ユーザーリストと帯域制限リストを取得
    $userList = Import-Csv -Path "E:\work\ps\azcopy\folder\User02\UserList.csv" -Encoding UTF8
    $trafficControlList = Import-Csv -Path "E:\work\ps\azcopy\folder\User02\TrafficControl.csv" -Encoding UTF8

    # ユーザーリストと帯域制限リストを結合
    $group = $userList | Where-Object { $_.Mail -eq $mail }
    $bpsRate = $trafficControlList | Where-Object { $_.Group -eq $group.Group }
    if ($bpsRate -ne 0) {
        $NetQoSPolicyName = "AzCopyPolicy01"
        $DSCPAction = 1
        New-NetQosPolicy -Name $NetQoSPolicyName -AppPathNameMatchCondition "E:\work\ps\azcopy\azcopy.exe" -DSCPAction $DSCPAction -ThrottleRateActionBitsPerSecond $bpsRate -Precedence 0
    }

    foreach ($item in $ItemList) {
        $filePath = $item.SubItems[2].Text
        $filePath = "E:\work\ps\azcopy\folder\User01\TestPST.pst"
        $destinationPath = "$($address)/00_UserUpload/" # Replace with your desired destination path
        $SASURL = "https://azcopyttest1481.blob.core.windows.net/migrationwiz/$($destinationPath)?sp=racwl&st=2024-03-10T14:04:32Z&se=2024-11-12T22:04:32Z&spr=https&sv=2022-11-02&sr=c&sig=Ld7Nbm9bhwMDRhbGUsGTWhb1BBi%2Fe0h9ydXQSm6eCL4%3D"
        Start-Process -FilePath "E:\work\ps\azcopy\azcopy" -ArgumentList "copy `"$filePath`" `"$SASURL`"" -NoNewWindow -Wait
    }

    # 作業ファイルおよび設定削除
    Remove-Item -Path "E:\work\ps\azcopy\folder\User02\UserList.csv" -Force
    Remove-Item -Path "E:\work\ps\azcopy\folder\User02\TrafficControl.csv" -Force
    if ($bpsRate -ne 0) {
        Remove-NetQosPolicy -Name $NetQoSPolicyName -Confirm:$false
    }
}

# フォームを作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "新PC移行用PSTアップロード"
$form.Size = New-Object System.Drawing.Size($width, $height)

# 説明文章を追加
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
$form.Controls.Add($emailLabel1)

$emailTextBox1 = New-Object System.Windows.Forms.TextBox
$emailTextBox1.Location = New-Object System.Drawing.Point(120, $heightColumn)
$emailTextBox1.Text = "test@example.com" # Set initial value
$emailTextBox1.Width = 390 # Set the width of the text box
$emailTextBox1.Enabled = $false # Make the text box read-only by default
$form.Controls.Add($emailTextBox1)

# メールアドレス入力欄2
$heightColumn += 25
$emailLabel2 = New-Object System.Windows.Forms.Label
$emailLabel2.Text = "メールアドレス："
$emailLabel2.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($emailLabel2)

$emailTextBox2 = New-Object System.Windows.Forms.TextBox
$emailTextBox2.Location = New-Object System.Drawing.Point(120, $heightColumn)
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
        # 関数を呼び出す処理を追加
        Write-Host "確認ボタンがクリックされました。"
    })

# 説明文章を追加
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "③「ファイル選択」を押すとアップロードするPSTファイルを選択することができます。" + "`r`n" + "1回のアップロードで同時に10ファイルまで可能。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# PST選択ボタン
$heightColumn += 40
$pstButton = New-Object System.Windows.Forms.Button
$pstButton.Text = "PST選択"
$pstButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($pstButton)
# ボタンがクリックされたときの処理
$pstButton.Add_Click({
        # 関数を呼び出す処理を追加
        Write-Host "PST選択ボタンがクリックされました。"
        Get-PstFile
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
        # 関数を呼び出す処理を追加
        Write-Host "アップロードボタンがクリックされました。"
        Invoke-PstFileUpload -ItemList $listView.Items
    })

# 説明文章を追加
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "⑤「移行」を押すと移行するPSTファイルを確定することができます。"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# 移行ボタン
$heightColumn += 25
$migrationButton = New-Object System.Windows.Forms.Button
$migrationButton.Text = "移行"
$migrationButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($migrationButton)
# ボタンがクリックされたときの処理
$migrationButton.Add_Click({
        # 関数を呼び出す処理を追加
        Write-Host "移行ボタンがクリックされました。"
        # Get-PstFile
    })

# フォームを表示
$form.ShowDialog()
