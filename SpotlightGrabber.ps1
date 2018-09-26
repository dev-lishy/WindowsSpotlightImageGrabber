Add-Type -AssemblyName System.Drawing

# Change destination folder name here
$IMAGE_DESTINATION_FOLDER = 'Spotlight Images'
$LANDSCAPE_PATH_REL = '.\landscape'
$PORTRAIT_PATH_REL = '.\portrait'

# Sets up paths
Set-Variable -Name USER_PATH -Value $env:HOMEDRIVE$env:HOMEPATH
Set-Variable -Name SCRIPT_PATH -Value $PSScriptRoot
Set-Variable -Name IMAGE_DESTINATION_PATH -Value $USER_PATH'\Pictures\'$IMAGE_DESTINATION_FOLDER
Set-Variable -Name IMAGE_SOURCE_PATH -Value $USER_PATH'\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets'

# Create destination folder if necessary
If(!(Test-Path $IMAGE_DESTINATION_PATH))
{
    New-Item -ItemType directory -Path $IMAGE_DESTINATION_PATH
}

# Copy over all files from Assets folder
Copy-Item $IMAGE_SOURCE_PATH\* -Destination $IMAGE_DESTINATION_PATH

# Set script location to destination folder
Set-Location -Path $IMAGE_DESTINATION_PATH

# Create folders if necessary
If(!(Test-Path $LANDSCAPE_PATH_REL'\old'))
{
    New-Item -ItemType directory -Path $LANDSCAPE_PATH_REL'\old'
}
If(!(Test-Path $PORTRAIT_PATH_REL'\old'))
{
    New-Item -ItemType directory -Path $PORTRAIT_PATH_REL'\old'
}

# Copy previous images into 'old' folders
Move-Item $LANDSCAPE_PATH_REL'\*' $LANDSCAPE_PATH_REL'\old' -Force -EA 0
Move-Item $PORTRAIT_PATH_REL'\*' $PORTRAIT_PATH_REL'\old' -Force -EA 0

# Add .jpg extension to file names
Get-ChildItem -File | Rename-Item -NewName { [io.path]::ChangeExtension($_.name, "jpg") }

# Find .jpg files with correct height dimensions and move to landscape and portrait folders
Get-ChildItem $IMAGE_DESTINATION_PATH -Filter *.jpg |
ForEach-Object -Process {
    $FILENAME = $_.FullName
    $FILE = New-Object System.Drawing.Bitmap $FILENAME
    $HEIGHT = $FILE.Height
    $FILE.Dispose()
    If($HEIGHT -eq 1080)
    {
        Move-Item $FILENAME $LANDSCAPE_PATH_REL
    }
    ElseIf($HEIGHT -eq 1920)
    {
        Move-Item $FILENAME $PORTRAIT_PATH_REL
    }
}

# Remove all files initially copied from Assets folder
Get-ChildItem -File | ForEach-Object { Remove-Item $_.FullName -Force }