# Specify the file size in bytes
$sizeInBytes = 1024

# Specify the path and name of the file to be created
$filePath = "C:\path\to\your\file.txt"

# Calculate the number of clusters needed to achieve the desired file size
$clusters = [math]::Ceiling($sizeInBytes / 4096)

# Use fsutil to create the file with the specified size
fsutil file createnew $filePath $clusters
