Add-Type -AssemblyName System.Windows.Forms

# �ϐ�
$rootPath = "E:\work\ps\azcopy"
# $rootPath = "C:\work\PSTUploader"
$width = 600
$height = 800
$azCopyPath = "$($rootPath)\azcopy.exe"
$StorageAccountName = "azcopyttest1481"
$SASKey = "sp=racwl&st=2024-03-10T14:04:32Z&se=2024-11-12T22:04:32Z&spr=https&sv=2022-11-02&sr=c&sig=Ld7Nbm9bhwMDRhbGUsGTWhb1BBi%2Fe0h9ydXQSm6eCL4%3D"

# �t�H���_�쐬����
$worFolderList = ("$($rootPath)\temp", "$($rootPath)\Log", "$($rootPath)\Output")
foreach ($workFolderPath in $worFolderList) {
    if (-Not (Test-Path -Path $workFolderPath)) {
        New-Item -ItemType Directory -Path $workFolderPath > $null
    }
}

# Transcript�擾
$transcriptPath = "$($rootPath)\Log\Transcript_$(Get-Date -Format "yyyyMMddHHmmss").txt"
Start-Transcript -Path $transcriptPath

# ���[���A�h���X�擾�֐���ǉ�
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

# ���[�J�������BlobStorage����A�b�v���[�h�ς݂̃t�@�C�����������鏈��
function Get-UploadedPstFile {
    # ���[���A�h���X����
    $address = ($emailTextBox1.Text).Split("@")[0]

    # �����Ώۃt�H���_���w��
    $localFolders = @(
        "C:\",
        "D:\"
    )

    # PST�J�E���g�ϐ�
    $totalCount = 0
    $totalSize = 0
    $LocalPSTInfoList = [System.Collections.ArrayList]@()

    # �o�̓t�H���_�p�X
    $localPSTListPath = "$($rootPath)\Output\PCData.txt"
    $blobPstFilePath = "$($rootPath)\Output\BlobPstFileList_$($address).txt"

    # ���[�U�[���X�g�ǂݍ���
    $userListLocalPath = "$($rootPath)\temp\UserList.csv"
    $userListBlobPath = "01_Manage/00_Group/UserList.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($userListBlobPath)?$($SASKey)"
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �ڍs�Ώۃ��[�U�[���X�g���擾���܂��B"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($userListLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �ڍs�Ώۃ��[�U�[���X�g�̎擾�����B"
    $userList = Import-Csv -Path $userListLocalPath -Encoding UTF8
    $group = $userList | Where-Object { $_.Mail -eq $($emailTextBox1.Text) }

    # �����O���[�v�����݂��Ȃ��ꍇ �G���[���b�Z�[�W��\�����A�t�H���_���������͌p������B
    if ($null -eq $group.Group) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - ���Ȃ��̃��[���A�h���X���ڍs�Ώێ҂Ƃ��ēo�^����Ă��܂���B"
        Set-ErrorMessage -Message "���Ȃ��̃��[���A�h���X���ڍs�Ώێ҂Ƃ��ēo�^����Ă��܂���B"
    }

    # PST�t�@�C��������
    foreach ($localFolder in $localFolders) {
        # ��������
        $localPSTFiles = Get-ChildItem -Path $localFolder -Filter "*.pst" -File -Recurse -ErrorAction SilentlyContinue

        # ���ꂼ���PST�t�@�C���̏����擾
        foreach ($localPSTFile in $localPSTFiles) {
            # ���X�g�ɒǉ�
            $LocalPSTInfoList.Add([PSCustomObject]@{
                    Name     = $localPSTFile.Name
                    FilePath = $localPSTFile.FullName
                    Size     = $localPSTFile.Length
                    Owner    = (Get-Acl -Path $localPSTFile.FullName).Owner
                }) > $null

            # �J�E���g�ϐ��̍X�V
            $totalCount++
            $totalSize += $size
        }

        Clear-Variable -Name localPSTFiles
    }

    # ���[�J��PC�f�[�^�̏o��
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

    # azcopy list�R�}���h�����s���ăA�b�v���[�h�ς݂�PST�t�@�C�����擾
    $destinationPath = "00_User/$($group.Group)/$($address)/00_UserUpload/"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($destinationPath)?$($SASKey)"
    $uploadedFileListPath = "$($rootPath)\temp\UploadedFileList.txt"
    Start-Process -FilePath $azCopyPath -ArgumentList "list $($SASURL) --running-tally --machine-readable" -NoNewWindow -Wait -RedirectStandardOutput $uploadedFileListPath
    $output = Get-Content -Path $uploadedFileListPath
    $blobPstFile = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt $output.Count - 3; $i++) {
        # �t�@�C�������擾
        $blobUploadDate = $output[$i].Split(":")[1].Split("/")[0].Substring(1)
        $blobUploadFileName = $output[$i].Split(":")[1].Substring(16, $output[$i].Split(":")[1].LastIndexOf(";") - 16)
        $blobUploadFileSize = $output[$i].Split(":")[2].Substring(1)

        # ���Ԃ̏C��
        $blobUploadDate = ([DateTime]::ParseExact($blobUploadDate, "yyyyMMddHHmmss", $null)).ToString("yyyy/MM/dd HH:mm:ss")

        # Blob�̃t�@�C�����ƃ��[�J���̃t�@�C�������r
        # $LocalPSTIndex = $LocalPSTInfoList.Name.IndexOf($blobUploadFileName)

        # # �A�b�v���[�h�ς݂��ǂ����𔻒�
        # if (
        #     $blobUploadFileName -in $LocalPSTInfoList.Name -and
        #     $blobUploadFileSize -in $LocalPSTInfoList[$LocalPSTIndex].Size -and
        #     $LocalPSTIndex -ne -1) {
        #     $flag = "�N���E�h�ɃA�b�v���[�h�ς�"
        #     $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTIndex].FilePath
        # }
        # elseif (
        #     $blobUploadFileName -in $LocalPSTInfoList.Name -and
        #     [Int64]$blobUploadFileSize -lt [Int64]$LocalPSTInfoList[$LocalPSTIndex].Size -and
        #     $LocalPSTIndex -ne -1) {
        #     $flag = "�N���E�h�ɃA�b�v���[�h�ς�&�e�ʑ���"
        #     $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTIndex].FilePath
        # }
        # elseif (
        #     $blobUploadFileName -in $LocalPSTInfoList.Name -and
        #     [Int64]$blobUploadFileSize -gt [Int64]$LocalPSTInfoList[$LocalPSTIndex].Size -and
        #     $LocalPSTIndex -ne -1) {
        #     $flag = "�N���E�h�ɃA�b�v���[�h�ς�&�e�ʌ���(�ʃt�@�C���܂��̓o�b�N�A�b�v�t�@�C��)"
        #     $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTIndex].FilePath
        # }
        # else {
        #     $flag = "�N���E�h�ɃA�b�v���[�h�ς�&PC���Ƀt�@�C�������݂��Ȃ�"
        #     $blobUploadFilePath = "PC���Ƀt�@�C�������݂��Ȃ�"
        # }

        # �A�b�v���[�h�ς݂��ǂ����𔻒�
        if ($blobUploadFileName -in $LocalPSTInfoList.Name) {
            $flag = "�N���E�h�ɃA�b�v���[�h�ς�"
            $blobUploadFilePath = $LocalPSTInfoList[$LocalPSTInfoList.Name.IndexOf($blobUploadFileName)].FilePath
        }
        else {
            $flag = "�N���E�h�ɃA�b�v���[�h�ς�&PC���Ƀt�@�C�������݂��Ȃ�"
            $blobUploadFilePath = "PC���Ƀt�@�C�������݂��Ȃ�"
        }

        # �t�@�C������ǉ�
        $blobPstFile.Add([PSCustomObject]@{
                ���[���A�h���X  = $address
                �X�e�[�^�X    = $flag
                �t�@�C����    = $blobUploadFileName
                �t�@�C���p�X   = $blobUploadFilePath
                �t�@�C���e��   = "{0:N1}" -f ([Int64]$blobUploadFileSize / 1024 / 1024) + "MB"
                �A�b�v���[�h���� = $blobUploadDate
            }) > $null

        Clear-Variable -Name flag, blobUploadFilePath
    }

    # ���[�J���ɂ����ăN���E�h�ɂȂ��t�@�C���̓ˍ�
    foreach ($LocalPSTInfo in $LocalPSTInfoList) {
        if (-Not($LocalPSTInfo.Name -in $blobPstFile.Name)) {
            $blobPstFile.Add([PSCustomObject]@{
                    ���[���A�h���X  = $address
                    �X�e�[�^�X    = "���A�b�v���[�h"
                    �t�@�C����    = $LocalPSTInfo.Name
                    �t�@�C���p�X   = $LocalPSTInfo.FilePath
                    �t�@�C���e��   = "{0:N1}" -f ([Int64]$LocalPSTInfo.Size / 1024 / 1024) + "MB"
                    �A�b�v���[�h���� = ""
                }) > $null
        }
    }

    # �e�L�X�g�Ƃ��ĕۑ�&�\��
    $blobPstFile | Out-File -FilePath $blobPstFilePath -Encoding UTF8
    Invoke-Item -Path $blobPstFilePath

    $maxTotalPstFileSize = 90000000000
    if ([Int64]$output[$output.Count - 1].Split(":")[2].Substring(1) -gt $maxTotalPstFileSize) {
        $str = @"
�A�b�v���[�h�\�ȃt�@�C���T�C�Y(90GB)�𒴂��Ă��܂��B
�A�b�v���[�h�ς݂̃t�@�C�����폜����ꍇ�́A�ڍs�S���҂֘A�������肢���܂��B
"@
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
        Set-ErrorMessage -Message $str
        return
    }

    Set-SystemMessage -Message "�A�b�v���[�h�󋵊m�F�������������܂����B"
    return
}

# PST�t�@�C���擾�֐���ǉ�
function Get-PstFile {
    # �T�C�Y�����ݒ�t�@�C���ǂݍ���
    $sizeLimitConfigLocalPath = "$($rootPath)\temp\SizeLimitConfig"
    $sizeLimitConfigBlobPath = "01_Manage/01_SizeLimit/SizeLimitConfig"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($sizeLimitConfigBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($sizeLimitConfigLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    $sizeLimitConfig = Get-Content -Path $sizeLimitConfigLocalPath -Encoding UTF8

    # �t�@�C���I���_�C�A���O��\��
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "�t�@�C����I��"
    $openFileDialog.Filter = "PST�t�@�C�� (*.pst)|*.pst"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments") + "\Outlook �f�[�^�t�@�C��"

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        # �t�@�C�������擾
        $fileInfo = Get-Item $openFileDialog.FileName
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �I�����ꂽ�t�@�C��: $($openFileDialog.FileName)"

        # PST�t�@�C���ȊO���I�����ꂽ�ꍇ
        if ($openFileDialog.FileName -notlike "*.pst") {
            # PST�t�@�C���ȊO���I�����ꂽ�ꍇ�B�V�K�t�@�C���쐬�Ȃǂ̑���ɂ����s�\�B
            $str = "PST�t�@�C���ȊO���I������܂����B���[�����ۑ�����Ă���PST�t�@�C����I�����Ă��������B"
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
            Set-ErrorMessage -Message $str
            return
        }

        # �T�C�Y�����m�F
        if ([Int64]$fileInfo.Length -gt ([Int64]$sizeLimitConfig * 1024 * 1024 * 1024)) {
            $str = "�t�@�C���T�C�Y�������𒴂��Ă��܂��B$($sizeLimitConfig)GB�ȏ��PST�t�@�C���͖{�c�[���ł̈ڍs�ΏۊO�ł��B"
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
            Set-ErrorMessage -Message $str
            return
        }

        # �t�@�C�����A�t�@�C���e�ʁA�t�@�C���p�X��ǉ� "{0:N1}" -f ([Int64]$fileInfo.Length / 1024 / 1024) + "MB"
        $listViewItem = New-Object System.Windows.Forms.ListViewItem($fileInfo.Name)
        $listViewItem.SubItems.Add("{0:N1}" -f ([Int64]$fileInfo.Length / 1024 / 1024) + "MB")
        $listViewItem.SubItems.Add($fileInfo.FullName)
        $listView.Items.Add($listViewItem)
        Set-SystemMessage -Message "�t�@�C�����I������܂����B"
        return
    }
    else {
        # �t�@�C���I����ʂ�����ꍇ
        $str = "�u�t�@�C���I���{�^���v��������܂������A�t�@�C�����I������܂���ł����B"
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($str)"
        Set-ErrorMessage -Message $str
        return
    }
}

# PST�A�b�v���[�h�֐���ǉ�
function Invoke-PstFileUpload {
    param (
        [System.Windows.Forms.ListView+ListViewItemCollection]$ItemList
    )

    # ���[���A�h���X����
    $address = ($emailTextBox1.Text).Split("@")[0]

    # �A�b�v���[�h���O�̕ۑ�
    $uploadedPstLogFilePath = "$($rootPath)\Log\UploadedPstLog_$($address).csv"

    # ���[�U�[���X�g�ǂݍ���
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �ݒ�t�@�C����ǂݍ��݂܂��B"
    $userListLocalPath = "$($rootPath)\temp\UserList.csv"
    $userListBlobPath = "01_Manage/00_Group/UserList.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($userListBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($userListLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    $userList = Import-Csv -Path $userListLocalPath -Encoding UTF8
    $group = $userList | Where-Object { $_.Mail -eq $($emailTextBox1.Text) }

    # �����O���[�v�����݂��Ȃ��ꍇ
    if ($null -eq $group.Group) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - ���Ȃ��̃��[���A�h���X���ڍs�Ώێ҂Ƃ��ēo�^����Ă��܂���B�A�b�v���[�h�����𒆒f���܂��B"
        Set-ErrorMessage -Message @"
���Ȃ��̃��[���A�h���X���ڍs�Ώێ҂Ƃ��ēo�^����Ă��܂���B
�A�b�v���[�h�����𒆒f���܂��B
"@
        return
    }

    # �ш搧�����X�g�ǂݍ���
    $trafficListLocalPath = "$($rootPath)\temp\TrafficControl.csv"
    $trafficListBlobPath = "01_Manage/00_Group/TrafficControl.csv"
    $SASURL = "https://$($StorageAccountName).blob.core.windows.net/migrationwiz/$($trafficListBlobPath)?$($SASKey)"
    Start-Process -FilePath $azCopyPath -ArgumentList "copy $($SASURL) $($trafficListLocalPath)" -NoNewWindow -Wait -RedirectStandardOutput "NUL"
    $trafficControlList = Import-Csv -Path $trafficListLocalPath -Encoding UTF8
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �ݒ�t�@�C���̓ǂݍ��݊����B"
    # �ш搧���ݒ�
    $bpsRate = $trafficControlList | Where-Object { $_.Group -eq $group.Group }
    $NetQoSPolicyName = "AzCopyPolicy01"
    $DSCPAction = 1
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �l�b�g���[�N���O�ݒ�����s�B"
    if (-Not(Get-NetQosPolicy -Name $NetQoSPolicyName -ErrorAction SilentlyContinue)) {
        New-NetQosPolicy -Name $NetQoSPolicyName -AppPathNameMatchCondition $azCopyPath -DSCPAction $DSCPAction -ThrottleRateActionBitsPerSecond $bpsRate -Precedence 0
    }

    # �t�@�C���A�b�v���[�h����
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �I�����ꂽ�t�@�C���̃A�b�v���[�h�����J�n�B"
    foreach ($item in $ItemList) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($item.SubItems[0].Text)�̃A�b�v���[�h���J�n�B"
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
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - $($item.SubItems[0].Text)�̃A�b�v���[�h�������B"
        }
        catch {
            # �G���[�����������ꍇ �~�{�^���ɂ��L�����Z�����ǂ����̔��菈����ǉ�
            if ($_.Exception.Message -eq "The user closed the window.") {
                Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �t�@�C���̃A�b�v���[�h���L�����Z������܂����B"
                Set-SystemMessage -Message "�t�@�C���̃A�b�v���[�h���L�����Z������܂����B"
                return
            }
            else {
                $errorMessage = $_.Exception.Message
                Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �G���[���������܂���: $($errorMessage)"
                Set-ErrorMessage -Message "�A�b�v���[�h�G���[���������܂����B"
                return
            }
        }
    }
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �I�����ꂽ�t�@�C���̃A�b�v���[�h���������B"
    Set-SystemMessage -Message "�I�����ꂽ�t�@�C���̃A�b�v���[�h�������������܂����B"

    # �l�b�g���[�N�ݒ�폜
    Remove-NetQosPolicy -Name $NetQoSPolicyName -Confirm:$false

    # ���X�g�N���A
    $listView.Items.Clear()
}

# ���[���A�h���X��v�`�F�b�N�֐���ǉ�
function Test-EmailAddress {
    # ���[���A�h���X����v���Ă��邩�ǂ������m�F
    if ($emailTextBox1.Text -ne $emailTextBox2.Text) {
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - ���[���A�h���X����v���Ă��܂���B�������X�L�b�v���܂��B"
        # �G���[���b�Z�[�W��\��
        Set-ErrorMessage -Message "���[���A�h���X����v���Ă��܂���B���[���A�h���X���͗����C�����Ă��������B"
        return $false
    }
    return $true
}

# �G���[���b�Z�[�W�̐ݒ�
function Set-ErrorMessage {
    param (
        [string]$Message
    )
    $errorLabel.ForeColor = [System.Drawing.Color]::Red
    $errorLabel.Text = $Message
}

# �V�X�e�����b�Z�[�W�̐ݒ�
function Set-SystemMessage {
    param (
        [string]$Message
    )
    $errorLabel.ForeColor = [System.Drawing.Color]::Black
    $errorLabel.Text = $Message
}

function Test-OutlookRunning {
    # Get-Process�R�}���h���b�g��Outlook�̃v���Z�X���擾
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($null -ne $outlookProcess) {
        $str = "Outlook�����s���ł��B�I�����Ă��������B"
        [System.Windows.Forms.MessageBox]::Show($str, "�x��", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Set-ErrorMessage -Message $str
        return $false
    }
    # �v���Z�X�����݂��Ȃ��ꍇ��$true���A���݂���ꍇ��$false��Ԃ�
    return $true
}

# �N����Outlook�m�F
Test-OutlookRunning > $null

# �t�H�[�����쐬
$form = New-Object System.Windows.Forms.Form
$form.Text = "�VPC���[���f�[�^�ڍs�c�[��"
$form.Size = New-Object System.Drawing.Size($width, $height)

# �������͂�ǉ�
$heightColumn = 10
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "�@�ڍs����]���郁�[���A�h���X����͂��Ă��������B"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.Width = 300
$form.Controls.Add($descriptionLabel)

# ���[���A�h���X���͗�1
$heightColumn += 25
$emailLabel1 = New-Object System.Windows.Forms.Label
$emailLabel1.Text = "���[���A�h���X�F"
$emailLabel1.Location = New-Object System.Drawing.Point(10, $heightColumn)
$emailLabel1.Width = 140
$form.Controls.Add($emailLabel1)

$emailTextBox1 = New-Object System.Windows.Forms.TextBox
$emailTextBox1.Location = New-Object System.Drawing.Point(150, $heightColumn)
$emailTextBox1.Text = Get-OutlookEmailAddress # Set initial value
$emailTextBox1.Width = 390 # Set the width of the text box
$emailTextBox1.Enabled = $false # Make the text box read-only by default
$form.Controls.Add($emailTextBox1)

# ���[���A�h���X���͗�2
$heightColumn += 25
$emailLabel2 = New-Object System.Windows.Forms.Label
$emailLabel2.Text = "���[���A�h���X(�m�F�p)�F"
$emailLabel2.Location = New-Object System.Drawing.Point(10, $heightColumn)
$emailLabel2.Width = 140
$form.Controls.Add($emailLabel2)

$emailTextBox2 = New-Object System.Windows.Forms.TextBox
$emailTextBox2.Location = New-Object System.Drawing.Point(150, $heightColumn)
$emailTextBox2.Text = $emailTextBox1.Text # Set initial value
$emailTextBox2.Width = 390 # Set the width of the text box
$emailTextBox2.Enabled = $false # Make the text box read-only by default
$form.Controls.Add($emailTextBox2)

# �`�F�b�N�{�b�N�X
$heightColumn += 25
$checkBoxLabel = New-Object System.Windows.Forms.Label
$checkBoxLabel.Text = "���[���A�h���X�ҏW�{�^���F"
$checkBoxLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$checkBoxLabel.Width = 140 # Set the width of the label
$form.Controls.Add($checkBoxLabel)

$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Location = New-Object System.Drawing.Point(150, $heightColumn)
$form.Controls.Add($checkBox)

# �`�F�b�N�{�b�N�X�̃C�x���g�n���h��
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

# �x�����x��
$heightColumn += 25
$warningLabel = New-Object System.Windows.Forms.Label
$warningLabel.Text = "�G���[�F���[���A�h���X����v���܂���"
$warningLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$warningLabel.AutoSize = $true
$warningLabel.ForeColor = [System.Drawing.Color]::Red
$warningLabel.Visible = $false
$form.Controls.Add($warningLabel)

# ���[���A�h���X�̈�v�`�F�b�N�̃C�x���g�n���h��
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


# �������͂�ǉ�
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "�A�u�m�F�v�������ƃA�b�v���[�h�ς݂�PST�t�@�C�������PC����PST�t�@�C�����m�F�ł��܂��B"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# �m�F�{�^��
$heightColumn += 25
$confirmButton = New-Object System.Windows.Forms.Button
$confirmButton.Text = "�m�F"
$confirmButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($confirmButton)
# �{�^�����N���b�N���ꂽ�Ƃ��̏���
$confirmButton.Add_Click({
        # ���[���A�h���X��v�m�F & Outlook���s�m�F
        if (Test-EmailAddress -eq $true -and Test-OutlookRunning -eq $true) {
            # �֐����Ăяo��������ǉ�
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �m�F�{�^�����N���b�N����܂����B"
            Get-UploadedPstFile
        }
    })

# �������͂�ǉ�
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "�B�u�t�@�C���I���v�������ƃA�b�v���[�h����PST�t�@�C����I�����邱�Ƃ��ł��܂��B" + "`r`n" + "1��̃A�b�v���[�h�œ�����10�t�@�C���܂ŉ\�B"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# �t�@�C���I���{�^��
$heightColumn += 40
$pstButton = New-Object System.Windows.Forms.Button
$pstButton.Text = "�t�@�C���I��"
$pstButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($pstButton)
# �{�^�����N���b�N���ꂽ�Ƃ��̏���
$pstButton.Add_Click({
        # ���[���A�h���X��v�m�F & Outlook���s�m�F
        if (Test-EmailAddress -eq $true -and Test-OutlookRunning -eq $true) {
            # �֐����Ăяo��������ǉ�
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �t�@�C���I���{�^�����N���b�N����܂����B"
            Get-PstFile
        }
    })

# �N���A�{�^��
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "�N���A"
$clearButton.Location = New-Object System.Drawing.Point(100, $heightColumn)
$form.Controls.Add($clearButton)
# �{�^�����N���b�N���ꂽ�Ƃ��̏���
$clearButton.Add_Click({
        $listView.Items.Clear()
        Set-SystemMessage -Message "�I�����ꂽ�t�@�C�����N���A����܂����B"
    })

# �P�O�s�̃��X�g
$heightColumn += 25
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, $heightColumn)
$listView.Size = New-Object System.Drawing.Size(560, 150)
$listView.View = [System.Windows.Forms.View]::Details
$listView.Columns.Add("�t�@�C����", 200) > $null
$listView.Columns.Add("�t�@�C���e��", 100) > $null
$listView.Columns.Add("�t�@�C���p�X", 250) > $null
$form.Controls.Add($listView)

# �������͂�ǉ�
$heightColumn += 180
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "�C�u�A�b�v���[�h�v�������ƑI�����ꂽPST�t�@�C�����A�b�v���[�h���邱�Ƃ��ł��܂��B"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# �A�b�v���[�h�{�^��
$heightColumn += 25
$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "�A�b�v���[�h"
$uploadButton.Location = New-Object System.Drawing.Point(10, $heightColumn)
$form.Controls.Add($uploadButton)
# �{�^�����N���b�N���ꂽ�Ƃ��̏���
$uploadButton.Add_Click({
        # ���[���A�h���X��v�m�F & Outlook���s�m�F
        if (Test-EmailAddress -eq $true -and Test-OutlookRunning -eq $true) {
            # �֐����Ăяo��������ǉ�
            Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �A�b�v���[�h�{�^�����N���b�N����܂����B"
            Invoke-PstFileUpload -ItemList $listView.Items
        }
    })

# �������͂�ǉ�
$heightColumn += 40
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "�����b�Z�[�W�\����"
$descriptionLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$descriptionLabel.AutoSize = $true
$form.Controls.Add($descriptionLabel)

# �G���[���b�Z�[�W��\��
$heightColumn += 25
$errorLabel = New-Object System.Windows.Forms.Label
$errorLabel.Text = ""
$errorLabel.Location = New-Object System.Drawing.Point(10, $heightColumn)
$errorLabel.AutoSize = $true
# $errorLabel.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($errorLabel)

# �t�H�[��������ꂽ���Ƃ����m����C�x���g�n���h��
$form.Add_FormClosed({
        Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - �t�H�[���������܂����B"
        Stop-Transcript
        Stop-Process -Id $PID
    })

# AzCopy�̑��݊m�F
if ( (Test-Path -Path $azCopyPath)) {
    Write-Host "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") - AzCopy���W���[����������܂���B"
    Set-ErrorMessage -Message "AzCopy���W���[����������܂���B��Ƃ����{�ł��܂���B"
}

# �t�H�[����\��
$form.ShowDialog()
