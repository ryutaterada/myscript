# Get the hostname
$hostname = $env:COMPUTERNAME

# Get the IP address
$ipAddress = (Test-Connection -ComputerName $hostname -Count 1).IPv4Address.IPAddressToString

# Specify the folders to search for .pst files
$folders = @(
    "C:\Folder1",
    "C:\Folder2",
    "C:\Folder3"
)

# Initialize variables for total count and size
$totalCount = 0
$totalSize = 0

# Initialize an array to store file information
$fileInfo = @()

# Iterate through each folder
foreach ($folder in $folders) {
    # Get .pst files in the folder
    $pstFiles = Get-ChildItem -Path $folder -Filter "*.pst" -File -Recurse

    # Iterate through each .pst file
    foreach ($pstFile in $pstFiles) {
        # Get file information
        $name = $pstFile.Name
        $size = $pstFile.Length
        $owner = (Get-Acl -Path $pstFile.FullName).Owner

        # Add file information to the array
        $fileInfo += [PSCustomObject]@{
            Name  = $name
            Size  = $size
            Owner = $owner
        }

        # Update total count and size
        $totalCount++
        $totalSize += $size
    }

    Clear-Variable -Name pstFiles
}

# Save the output to a text file
$outputPath = "C:\Output.txt"
$output = @"
$hostname
$ipAddress
$totalCount
$totalSize
"@

# Append file information to the output
foreach ($file in $fileInfo) {
    $output += "$($file.Name) $($file.Size) $($file.Owner)`n"
}

$output | Out-File -FilePath $outputPath -Encoding UTF8
