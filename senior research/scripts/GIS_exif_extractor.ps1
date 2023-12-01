# adpted from https://stackoverflow.com/questions/68971176/use-powershell-to-extract-gps-latitude-etc-properties-details-from-an-image

cls
$ErrorActionPreference = 'SilentlyContinue'

$Shell = New-Object -ComObject "WScript.Shell"
#$Button = $Shell.Popup("This script will extract GPS position data from geotagged images and output them to a CSV file in the same folder that the script is located.`nA selection dialog will appear next to select the folder where the images are stored.`n`nPlease ensure that the selected folder only contains geotagged images as the script currently parses ALL images in the selected folder and will inlude bad data in the ouput for files without exif GPS data.`n`n`nR. Olsen 2023", 0, "Instructions", 0)
Write-Host "This script will extract GPS position data from geotagged images and output them to a CSV file in the same folder that the script is located.`nA selection dialog will appear next to select the folder where the images are stored.`n`nPlease ensure that the selected folder only contains geotagged images as the script currently parses ALL images in the selected folder and will inlude bad data in the ouput for files without exif GPS data.`n`nNOTE: Some services such as Googe Drive may remove GPS data`n`nR. Olsen 2023`n"

pause
$OutputArray = @()
#Write-Host "Press enter to continue and CTRL-C to exit ..."
#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
$NoData = -9999
$NoDataAsk = read-host "`nWhat integer value for no data fields you you want to use? (Default is -9999)`nEnter [d] for default or an integer value and press enter"
#$NoDataSet = $NoDataAsk -ne "d"
    if($NoDataAsk -ne "d") {
     $NoData = $NoDataAsk
     Write-Host "NoData changed to $NoData`n"
    }

Function Get-Folder($initialDirectory="")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select the folder where your geotagged images are stored.`n`nClicking Cancel will end the script"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    } else {
	exit
	}
    return $folder
}
$FolderPath = Get-Folder
$OutName = "coord_export.$((Get-Date).ToString('yyyyMMddTHHmmss')).csv"

Write-Host "fileName,`tlon,`tlat,`talt,`tLocalDTG,`tZuluOffset,`tGPSposError,`tGPSsatNum"

Get-ChildItem $FolderPath -Filter *.* | where { ! $_.PSIsContainer } |
	Foreach-Object {
    # Get the file name and path, write it out to screen
    $FileName = $_.FullName
    $Name = (Split-Path -Path $FileName -Leaf).Split(".")[0];

    # Create an ImageFile object and load an image file
    $image = New-Object -ComObject Wia.ImageFile
    $image.LoadFile($FileName)

	#Clear variables for Lat and Lon
	Clear-Variable Lat*
	Clear-Variable Lon*
	Clear-Variable Alt*
    Clear-Variable GPS*

	$LatDEG = $image.Properties.Item('GpsLatitude').Value[1].Value
	$LatMIN = $image.Properties.Item('GpsLatitude').Value[2].Value
	$LatSEC = $image.Properties.Item('GpsLatitude').Value[3].Value 
	$LatREF = $image.Properties.Item('GpsLatitudeRef').Value
	$LatSIN = [string]::Empty
	if ($LatREF -ne 'N`0') {
		$LatSIN = '-'
	}
	$LonDEG = $image.Properties.Item('GpsLongitude').Value[1].Value
	$LonMIN = $image.Properties.Item('GpsLongitude').Value[2].Value
	$LonSEC = $image.Properties.Item('GpsLongitude').Value[3].Value
	$LonREF = $image.Properties.Item('GpsLongitudeRef').Value
	$LatSIN = [string]::Empty
	if ($LonREF -ne 'E`0') {
		$LonSIN = '-'
	}
	$AltVAL = If ($image.Properties.Exists('6'))     { ($image.Properties.Item('6').Value) | select -ExpandProperty Value } Else {$NoData}
#	$AltOUT
#	if ($AltVAL -eq $null ) {
#		$AltOUT = -9999
#	} else {
#		$AltOUT = $AltVAL
#	}

    $LocalDTG = If ($image.Properties.Exists('36867'))     { ($image.Properties.Item('36867').Value)} Else {$NoData}
    $ZuluOffset = If ($image.Properties.Exists('34858'))     { ($image.Properties.Item('34858').Value)} Else {$NoData}
    $GPSposError = If ($image.Properties.Exists('31'))    { $image.Properties.Item('31').Value }                         Else {$NoData}
    $GPSsatNum = If ($image.Properties.Exists('8'))     { $image.Properties.Item('8').Value }                          Else {$NoData}

<#

integrate portions of this for better error handling

			# Source: http://www.exiv2.org/tags.html
                        # Source: https://sno.phy.queensu.ca/~phil/exiftool/TagNames/GPS.html

			            'GPSAltitude'                   = If ($image.Properties.Exists('6'))     { ($image.Properties.Item('6').Value) | select -ExpandProperty Value } Else {" "}
                        'GPSDateStamp'                  = If ($image.Properties.Exists('29'))    { $image.Properties.Item('29').Value }                         Else {" "}
                        'GPSDifferential'               = If ($image.Properties.Exists('30'))    { $gpsdiffer[[int]$image.Properties.Item('30').Value] }        Else {" "}
                        'GPSHPositioningError'          = If ($image.Properties.Exists('31'))    { $image.Properties.Item('31').Value }                         Else {" "}
                        'GPSImgDirection'               = If ($image.Properties.Exists('17'))    { ($image.Properties.Item('17').Value) | select -ExpandProperty Value } Else {" "}
                        'GPSImgDirectionRef'            = If ($image.Properties.Exists('16'))    { $gpsdirect["$($image.Properties.Item('16').Value)"] }        Else {" "}
                        'GPSLatitude'                   = If ($image.Properties.Exists('2'))     { (($image.Properties.Item('2').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'GPSLatitudeRef'                = If ($image.Properties.Exists('1'))     { $image.Properties.Item('1').Value }                          Else {" "}
                        'GPSLongitude'                  = If ($image.Properties.Exists('4'))     { (($image.Properties.Item('4').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'GPSLongitudeRef'               = If ($image.Properties.Exists('3'))     { $image.Properties.Item('3').Value }                          Else {" "}
                        'GPSMapDatum'                   = If ($image.Properties.Exists('18'))    { $image.Properties.Item('18').Value }                         Else {" "}
                        'GPSSatellites'                 = If ($image.Properties.Exists('8'))     { $image.Properties.Item('8').Value }                          Else {" "}
                        'GPSStatus'                     = If ($image.Properties.Exists('9'))     { $gpsstat["$($image.Properties.Item('9').Value)"] }           Else {" "}
                        'GPSTag'                        = If ($image.Properties.Exists('34853')) { $image.Properties.Item('34853').Value }                      Else {" "}
                        'GPSTimeStamp'                  = If ($image.Properties.Exists('7'))     { (($image.Properties.Item('7').Value) | select -ExpandProperty Value) -join (':') } Else {" "}
#>




	
	# Convert them to Degrees Minutes Seconds Ref
	
	$LatSTR = $LatDEG+($LatMIN/60)+($LatSEC/3600)
	$LonSTR = $LonDEG+($LonMIN/60)+($LonSEC/3600)
	$LatSHORT = $LatSTR.ToString('F7')
	$LonSHORT = $LonSTR.ToString('F7')
	
	# Write the full coordinates out
	Write-Host "$Name,`t$LonSIN$LonSHORT,`t$LatSIN$LatSHORT,`t$AltVAL,`t$LocalDTG,`t$ZuluOffset,`t$GPSposError,`t$GPSsatNum"
#	Add-Content -Path "$PSScriptRoot\$OutName" -Value "$Name, $LonSIN$LonSHORT, $LatSIN$LatSHORT, $AltVAL"
#    $OutputArray += "$Name, $LonSIN$LonSHORT, $LatSIN$LatSHORT, $AltVAL, $GPSdate, $GPStime, $GPSposError, $GPSsatNum"
    $OutputArray += "$Name, $LonSIN$LonSHORT, $LatSIN$LatSHORT, $AltVAL, $LocalDTG, $ZuluOffset, $GPSposError, $GPSsatNum"
}
Write-Host
Write-Host

$WriteAsk = read-host "Would you like to write this to a file?`n`nType y to continue or n to abort"
$WriteAbort = $WriteAsk -ne "n"
    if($WriteAbort -ne $true) {
    exit
    } else {
    Write-Host "write file"
    New-Item -Path "$PSScriptRoot\$OutName" -ItemType File -Value "fileName, lon, lat, alt, LocalDTG, ZuluOffset, GPSposError, GPSsatNum`n"
    $OutputArray | ForEach-Object {Add-Content -Path "$PSScriptRoot\$OutName" -Value $PSItem}
    }

Write-Host
Write-Host "$OutName written to $PSScriptRoot"
Write-Host
Write-Host "You may now close this window or wait for it to automatically close in 10 seconds"
Start-Sleep -Seconds 10

