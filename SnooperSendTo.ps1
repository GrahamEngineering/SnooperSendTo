<#

    This script will create a Shortcut to Snooper in the SendTo right-click menu in Explorer
        Path for context-menu SendTo is: %AppData%\Microsoft\Windows\SendTo (or in PS: $env:AppData\Microsoft\Windows\SendTo)
        Path for Snooper is: %ProgramFiles%\Microsoft Lync Server 2013\Debugging Tools\Snooper.exe (or in PS: $env:ProgramFiles\Microsoft Lync Server 2013\Debugging Tools\Snooper.exe)

		Created 5/2/2014 by M. Graham
			
        Updated 7/27/2016
            (support for Skype for Business path)
#>


# Configure Variables
$OriginalVerbosePreference = $VerbosePreference
$VerbosePreference = "Continue"

#
# This is the path to snooper.  If you're running it on your own machine and not something with the Debugging tools installed, you can change as needed.
#
$SnooperPath = $env:ProgramFiles + "\Microsoft Lync Server 2013\Debugging Tools\Snooper.exe"
$SendToPath = $env:AppData + "\Microsoft\Windows\Sendto"


# Check to ensure Snooper exists
if (!(Test-Path $SnooperPath))
{
    # Snooper Does not exist.  Exit.
    Write-Warning "Snooper.exe does not exist at path $SnooperPath"
    $SnooperPath = $env:ProgramFiles + "\Skype for Business Server 2015\Debugging Tools\Snooper.exe"
    if (!(Test-Path $SnooperPath))
    {
        Write-Warning "Snooper.exe does not exist at path $SnooperPath"
        return
    }
    else
    {
        Write-Warning "Snooper found in Skype For Business directory"
    }
    $VerbosePreference = $OriginalVerbosePreference
    
}
else
{
    # Snooper.exe exists
    Write-Verbose "Snooper.exe found at path $SnooperPath"
}

if (!(Test-Path $SendToPath))
{
    # Could not find the path to "SendTo"  Exit.
    Write-Error "SendTo path not found at $SendToPath"
    $VerbosePreference = $OriginalVerbosePreference
    return
}
else
{
    # SendTo exists in expected location
    Write-Verbose "SendTo path exists"
}

# Make sure snooper.lnk does not already exist in sendto
if (Test-Path "$SendToPath\Snooper.lnk")
{
    # Link already exsits
    Write-Warning "Link for Snooper already exists in SendTo folder. Exiting."
    $VerbosePreference = $OriginalVerbosePreference
    return
}
else
{
    Write-Verbose "No existing snooper.lnk file found.  Creating new shortcut"
}

# Create a new shortcut to Snooper.exe in the SendTo Folder
$WshShell = New-Object -ComObject WScript.shell
$Shortcut = $WshShell.CreateShortcut("$SendToPath\Snooper.lnk")
$Shortcut.TargetPath = $SnooperPath
$Shortcut.Save()

if (Test-Path "$SendToPath\Snooper.lnk")
{
    Write-Output "Link created successfully"
}
else
{
    Write-Warning "Link creation was unsuccessful"
}
$VerbosePreference = $OriginalVerbosePreference

