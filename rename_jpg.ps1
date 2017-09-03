# Get parameters from command line
Param ($global:source,$global:dest,$delsrc)
# Add the .NET Framework Class System.Drawing to the PowerShell Session

$datecurrent = Get-Date -Format "yyyy-MM-dd_hh-mm-ss"

# Read the destination and source from CLI if the weren't already
# specified
if(!$global:source){
    $global:source = Read-Host -Prompt "Source?"
}

if(!$global:dest){
    $global:dest = Read-Host -Prompt "Destination?"
}

# Set the Source and Destination paths
#$global:source = "C:\Users\bob\Documents\Test"
#$global:dest = "C:\Users\bob\Documents\Fotos"

# Ask the user if he wants to delete the original files after copying them
$delsrc = Read-Host -Prompt "Should the original files be removed? [y/n]"
# Notify user and set default option parameter to "n" if not specified
switch ($delsrc){
    y {Write-Host "Original files will be removed!"}
    n {Write-Host "Original files will be preserved!"}
    default {Write-Host "Original files will be preserved!."; $delsrc = "n"}

}

# Counting variable for renaming files
$counter = 0
Add-Type -AssemblyName System,System.Drawing
Get-ChildItem -Path "$global:source" -Recurse -Filter *.jpg | ForEach-Object {
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

    # Construct the $global:datetaken meta data variable
    $global:datetaken = "$dateyear" + "." + "$datemonth" + "." + "$dateday"

    # In case the destination already exists with a certain event, this
    # has to be catched
    $copydest = Get-ChildItem -Path $global:dest\* -Filter "$global:datetaken*"
    # If no destination was found, defaults to destination path followed
    # by DateTaken
    if(!$copydest) {
        $copydest = "$global:dest\$global:datetaken"
    }

    # Check if the folder for the given day already exists, if it doesn't
    # create it in the destination folder
    if(Test-Path "$copydest"){
    } else {
        $dateevent = Read-Host -Prompt "State the ocassion to $global:datetaken"
        New-Item -Path $global:dest -ItemType Directory -Name "$global:datetaken $dateevent" `
        | Out-Null
    }

    # If the destination was modified by user input, catch the changes
    $copydest = Get-ChildItem -Path $global:dest\* -Filter "$global:datetaken*"

    # Get the values for hour, minute and second from the corresponding positions
    # in the $imagedate array and cast them to the char type
    $datehour = [Char]$imagedate[11]+[Char]$imagedate[12]
    $dateminute = [Char]$imagedate[14]+[Char]$imagedate[15]
    $datesecond = [Char]$imagedate[17]+[Char]$imagedate[18]

    # Construct the $global:datetaken meta data variable
    $global:timetaken = "$datehour" + "." + "$dateminute" + "." + "$datesecond"

    #$lockp = CMD /C "openfiles /query /fo table | find /I """""
    #Write-Host $lockp

    # Copy the photo to the destination
    Copy-Item -Path $_.FullName -Destination $copydest

    # Rename files to a temporary file name
    Rename-Item -Path "$copydest\$_" -NewName "$counter $datecurrent $global:timetaken.jpg"

    # Increment the counter for renaming
    $counter++
}

# Get all newly created folders in the destination path and start a loop
# for each of them
Get-ChildItem -Path "$global:dest\*" | ForEach-Object {
    # Reset the renaming index counter
    $counter = 1
    # Get all images in the current folder and start renaming them
    Get-ChildItem -Path $_ | Where-Object {$_.Name -match "^[0-9]+\s$datecurrent\s"}`
    | ForEach-Object {
        # Get the pure name of the file without the temporary index
        $newname = $_.Name -replace "^[0-9]+\s$datecurrent\s",''
        # Get the name of a potential duplicate file
        $tmppath = $_.FullName -replace "$_","Nr$counter $newname"

        # Detect duplicates
        if(Test-Path "$tmppath") {
            # Duplicate detected, name after current date (during runtime), so no
            # collision happens
            Rename-Item -Path $_.FullName -NewName "Nr$counter $datecurrent $newname"
            $global:dupfiles = "true"
        }else{
            # Rename the current file to get the new index in front of the name
            Rename-Item -Path $_.FullName -NewName "Nr$counter $newname"
        }

        # Increment the counter
        $counter++
    }
}

# Duplicate file detection prompt
if ($global:dupfiles){
    Write-Host "Duplicate files detected. Renamed them to avoid collision."
}

Read-Host -Prompt 'Done! Press "Enter" to exit'

# If specified by the user, delete the original files
# if($delsrc -eq "y" ){
#    Get-ChildItem -Path "$global:source" -Recurse -Filter *.jpg | Remove-Item -Force
#  }
