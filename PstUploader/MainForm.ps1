Add-Type -AssemblyName System.Windows.Forms

# 変数
$width = 600
$height = 600
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

# メールアドレス入力欄
$heightColumn += 25
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "メールアドレス:"
$emailLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($emailLabel)

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(120, $heightColumn)
$emailTextBox.Text = "test@example.com" # Set initial value
# $emailTextBox.Text = Get-OutlookEmailAddress # Set initial value
$emailTextBox.Width = 390 # Set the width of the text box
$form.Controls.Add($emailTextBox)

# 説明文章を追加
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "②「確認」を押すとアップロード済みのPSTファイルを確認できます。"
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
        # Get-PstFile
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
