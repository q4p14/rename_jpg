# Get parameters from command line
Param ($source,$dest,$delsrc)
# Add the .NET Framework Class System.Drawing to the PowerShell Session
Add-Type -AssemblyName System,System.Drawing

$datecurrent = Get-Date -Format "yyyy-MM-dd_hh-mm-ss"

# Read the destination and source from CLI if the weren't already
# specified
if(!$source){
    $source = Read-Host -Prompt "Quelle?"
}

if(!$dest){
    $dest = Read-Host -Prompt "Ziel?"
}

# Set the Source and Destination paths
#$source = "C:\Users\bob\Documents\Test"
#$dest = "C:\Users\bob\Documents\Fotos"

# Ask the user if he wants to delete the original files after copying them
$delsrc = Read-Host -Prompt "Sollen die Ursprungsdateien gelöscht werden? [y/n]"
# Notify user and set default option parameter to "n" if not specified
switch ($delsrc){
    y {Write-Host "Ursprungsdateien werden entfernt."}
    n {Write-Host "Ursprungsdateien werden nicht entfernt."}
    default {Write-Host "Ursprungsdateien werden nicht entfernt."; $delsrc = "n"}
}

# Counting variable for renaming files
$counter = 0

Get-ChildItem -Path "$source" -Recurse -Filter *.jpg | ForEach-Object {
    # Initialize a new bitmap class object from the given image file
    $imagefile = New-Object System.Drawing.Bitmap($_.FullName)
    # Get the images date values with a built in method, select from
    # index 0 to 9
    $imagedate = $imagefile.GetPropertyItem(36867).Value[0..18]

    # Get the values for year, month and day from the corresponding positions
    # in the $imagedate array and cast them to the char type
    $dateyear = [Char]$imagedate[0]+[Char]$imagedate[1]+[Char]$imagedate[2]+`
    [Char]$imagedate[3]
    $datemonth = [Char]$imagedate[5]+[Char]$imagedate[6]
    $dateday = [Char]$imagedate[8]+[Char]$imagedate[9]

    # Construct the $datetaken meta data variable
    $datetaken = "$dateyear" + "." + "$datemonth" + "." + "$dateday"

    # In case the destination already exists with a certain event, this
    # has to be catched
    $copydest = Get-ChildItem -Path $dest\* -Filter "$datetaken*"
    # If no destination was found, defaults to destination path followed
    # by DateTaken
    if(!$copydest) {
        $copydest = "$dest\$datetaken"
    }

    # Check if the folder for the given day already exists, if it doesn't
    # create it in the destination folder
    if(Test-Path "$copydest"){
    } else {
        $dateevent = Read-Host -Prompt "Anlass zum Datum $datetaken"
        New-Item -Path $dest -ItemType Directory -Name "$datetaken $dateevent" `
        | Out-Null
    }

    # If the destination was modified by user input, catch the changes
    $copydest = Get-ChildItem -Path $dest\* -Filter "$datetaken*"

    # Get the values for hour, minute and second from the corresponding positions
    # in the $imagedate array and cast them to the char type
    $datehour = [Char]$imagedate[11]+[Char]$imagedate[12]
    $dateminute = [Char]$imagedate[14]+[Char]$imagedate[15]
    $datesecond = [Char]$imagedate[17]+[Char]$imagedate[18]

    # Construct the $datetaken meta data variable
    $timetaken = "$datehour" + "." + "$dateminute" + "." + "$datesecond"

    #$lockp = CMD /C "openfiles /query /fo table | find /I """""
    #Write-Host $lockp

    # Copy the photo to the destination
    Copy-Item -Path $_.FullName -Destination $copydest

    # Rename files to a temporary file name
    Rename-Item -Path "$copydest\$_" -NewName "$counter $datecurrent $timetaken.jpg"

    # Increment the counter for renaming
    $counter++
}

# Get all newly created folders in the destination path and start a loop
# for each of them
Get-ChildItem -Path "$dest\*" | ForEach-Object {
    # Reset the renaming index counter
    $counter = 0
    # Get all images in the current folder and start renaming them
    Get-ChildItem -Path $_ | Where-Object {$_.Name -match "^[0-9]+\s$datecurrent\s"}`
    | ForEach-Object {
        # Get the pure name of the file without the temporary index
        $newname = $_.Name -replace "^[0-9]+\s$datecurrent\s",''
        $tmppath = $_.FullName -replace "$_.Name","Nr$counter $newname"
        # Rename the current file to get the new index in front of the name
        if(Test-Path $tmppath) {
            Rename-Item -Path $_.FullName -NewName "Nr$counter $datecurrent $newname"
            Write-Host "Duplicate files detected, renaming files to current date"
        }else{
            Rename-Item -Path $_.FullName -NewName "Nr$counter $newname"
            Write-Host "Test"
        }

        # Increment the counter
        $counter++
    }
}

# If specified by the user, delete the original files
 if($delsrc -eq "y" ){
    Get-ChildItem -Path "$source" -Recurse -Filter *.jpg | Remove-Item -Force
  }
