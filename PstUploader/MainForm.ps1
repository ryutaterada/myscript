Add-Type -AssemblyName System.Windows.Forms

# フォームを作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "メインフォーム"
$form.Size = New-Object System.Drawing.Size(400, 300)

# メールアドレス入力欄
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "メールアドレス:"
$emailLabel.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($emailLabel)

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(120, 10)
$emailTextBox.Text = "test@example.com" # Set initial value
$emailTextBox.Width = 200 # Set the width of the text box
$form.Controls.Add($emailTextBox)

# メールアドレス取得関数を追加

# １０行のリスト
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 40)
$listView.Size = New-Object System.Drawing.Size(380, 150)
$listView.View = [System.Windows.Forms.View]::Details
$listView.Columns.Add("ファイル名")
$listView.Columns.Add("ファイル容量")
$listView.Columns.Add("ファイルパス")
$form.Controls.Add($listView)

# PST選択ボタン
$pstButton = New-Object System.Windows.Forms.Button
$pstButton.Text = "PST選択"
$pstButton.Location = New-Object System.Drawing.Point(10, 200)
$form.Controls.Add($pstButton)
# ボタンがクリックされたときの処理
$pstButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "ファイルを選択"
        $openFileDialog.Filter = "PSTファイル (*.pst)|*.pst"
        $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
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
    })

# アップロードボタン
$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "アップロード"
$uploadButton.Location = New-Object System.Drawing.Point(100, 200)
$form.Controls.Add($uploadButton)

# 確認ボタン
$confirmButton = New-Object System.Windows.Forms.Button
$confirmButton.Text = "確認"
$confirmButton.Location = New-Object System.Drawing.Point(190, 200)
$form.Controls.Add($confirmButton)
# ボタンがクリックされたときの処理
$confirmButton.Add_Click({
        # 関数を呼び出す処理を追加
        Write-Host "確認ボタンがクリックされました。"
        New-Item -Path "E:\work\ps\azcopy\Form" -Name "test.pst" -ItemType "file" -WhatIf
    })

# 移行ボタン
$migrationButton = New-Object System.Windows.Forms.Button
$migrationButton.Text = "移行"
$migrationButton.Location = New-Object System.Drawing.Point(280, 200)
$form.Controls.Add($migrationButton)

# フォームを表示
$form.ShowDialog()
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "ファイルを選択"
$openFileDialog.Filter = "すべてのファイル (*.*)|*.*"

if ($openFileDialog.ShowDialog() -eq 'OK') {
    $selectedFile = $openFileDialog.FileName
    Write-Host "選択されたファイル: $selectedFile"
}


# アップロードボタン
$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "アップロード"
$uploadButton.Location = New-Object System.Drawing.Point(100, 200)
$form.Controls.Add($uploadButton)

# 確認ボタン
$confirmButton = New-Object System.Windows.Forms.Button
$confirmButton.Text = "確認"
$confirmButton.Location = New-Object System.Drawing.Point(190, 200)
$form.Controls.Add($confirmButton)
# ボタンがクリックされたときの処理
$confirmButton.Add_Click({
        # 関数を呼び出す処理を追加
        Write-Host "確認ボタンがクリックされました。"
        New-Item -Path "E:\work\ps\azcopy\Form" -Name "test.pst" -ItemType "file" -WhatIf
    })

# 移行ボタン
$migrationButton = New-Object System.Windows.Forms.Button
$migrationButton.Text = "移行"
$migrationButton.Location = New-Object System.Drawing.Point(280, 200)
$form.Controls.Add($migrationButton)

# フォームを表示
$form.ShowDialog()
