<# The purpose of this script is to do the folling
 Import a CSV file with server name and share information
 Collect the share permissions for each share in the CSV and log it to a file
 Set the share permissions to read only
 Attempt to close any files open on the noted share

The CSV should have the following headers:server,share
    Server column should contain just the server name
    Share column should contain just the share name

#>

# Getting the desktop path of the user launching the script. Just as a default starting path
$DesktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)

# Prompt the user to select the CSV file for the script
Function Get-FileName($initialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.Filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.Title = "Select GET SHARE ACCESS CSV"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

$FilePath  = Get-FileName -initialDirectory $DesktopPath
$csv      = @()
$csv      = Import-Csv -Path $FilePath 
$results = @()

# Set count for activity status
$i = 0

#Loop through all items in the CSV 
ForEach ($item In $csv) 
{

    # Put the objects into string variables, because new-dfsnfoldertarget likes strings
    [string]$server = $item.server
    [string]$share = $item.share
    
    # This little diddy will provide a progress bar!
    $i = $i+1
    Write-Progress -Activity "Checking \\$server\$share" -Status "Progress:" -PercentComplete ($i/$csv.Count*100)
    
    # Getting SMB share permissions for each server/share in the CSV
    $results += get-SmbShareAccess -CimSession $server -Name $share
}

write-host -ForegroundColor Green "Saving share-results.csv to desktop"
$results | export-csv -path $desktopPath+"\share-results.csv" -NoTypeInformation

# clear variables to be used in setting share permissions
Clear-Variable i
Clear-Variable item
Clear-Variable server
Clear-Variable share

#Loop through all items in the the results (data is fresh
ForEach ($item In $results) 
{

    # Put the objects into string variables, because new-dfsnfoldertarget likes strings
    [string]$server = $item.PSComputerName
    [string]$share = $item.name
    [string]$account = $item.AccountName
    [string]$permissions = "Read"
    
    # This little diddy will provide a progress bar!
    $i = $i+1
    Write-Progress -Activity "Updating \\$server\$share permissions to $permissions for $account" -Status "Progress:" -PercentComplete ($i/$csv.Count*100)
    
    # Setting the SMB share permissions to READ for each server/share in the results (created from initial CSV)
    Grant-SmbShareAccess -CimSession $server -Name $share -AccountName $account -AccessRight $permissions -force
}

# clear variables to be used in kicking users
Clear-Variable i
Clear-Variable item
Clear-Variable server
Clear-Variable share

#Loop through all items in the CSV 
ForEach ($item In $csv) 
{

    # Put the objects into string variables, because new-dfsnfoldertarget likes strings
    [string]$server = $item.server
    [string]$share = $item.share
    
    # This little diddy will provide a progress bar!
    $i = $i+1
    Write-Progress -Activity "Checking \\$server\$share. Ignore any errors you see on the screen" -Status "Progress:" -PercentComplete ($i/$csv.Count*100)
  
    # Doing the work
    # Get's the file path of the share from the server
    $fp = Get-SmbShare -CimSession $server -Name $share | Select-Object -Property path

    # convert the file path because powershell fails on single slash paths in this situation
    $efp = [regex]::Escape($fp.path)

    # Find files open under the share path and close them!
    Get-SmbOpenFile -CimSession $server | Where-Object -Property Path -match $efp | Close-SmbOpenFile -Force
    
}

write-host "All done!" -BackgroundColor Green
