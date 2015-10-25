#
# Executes the vsvars32.bat file for the visual studio selected in powershell
#
# From Chris Tavares via Scott Hanelman:
# http://www.hanselman.com/blog/AwesomeVisualStudioCommandPromptAndPowerShellIconsWithOverlays.aspx
function Get-Batchfile ($file) {
        $cmd = "`"$file`" & set"
        cmd /c $cmd | Foreach-Object {
                $p, $v = $_.split('=')
                Set-Item -path env:$p -value $v
        }
}

# VsVars with dynamic completion for VS Versions
function VsVars {
    [CmdletBinding()]
    Param(
        # Any other parameters can go here
    )
 
    DynamicParam {

			([intptr]::size -eq 8)
			{
			    $registryKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\VisualStudio"
			}
			else
			{
			    $registryKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio"
			}

            # Set the dynamic parameters' name
            $ParameterName = 'version'
            
            # Create the dictionary 
            $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

            # Create the collection of attributes
            $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            
            # Create and set the parameters' attributes
            $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ParameterAttribute.Mandatory = $true
            $ParameterAttribute.Position = 1

            # Add the attributes to the attributes collection
            $AttributeCollection.Add($ParameterAttribute)

            # Generate and set the ValidateSet 
            $arrSet = gci -Path $registryKey | where-object {$_.Name -match "[0-9][0-9]\.[0-9]" }  | Select pschildname | foreach { "$($_.PSChildName)" }
            $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

            # Add the ValidateSet to the attributes collection
            $AttributeCollection.Add($ValidateSetAttribute)

            # Create and return the dynamic parameter
            $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
            $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
            return $RuntimeParameterDictionary
    }

    begin {
        # Bind the parameter to a friendly variable
        $version = $PsBoundParameters[$ParameterName]
    }

    process {
		([intptr]::size -eq 8)
		{
		    $registryKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\" + $version
		}
		else
		{
		    $registryKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\" + $version
		}
	        $VsKey = get-ItemProperty $registryKey

	        $VsInstallPath = [System.IO.Path]::GetDirectoryName($VsKey.InstallDir)
	        $VsToolsDir = [System.IO.Path]::GetDirectoryName($VsInstallPath)
	        $VsToolsDir = [System.IO.Path]::Combine($VsToolsDir, "Tools")
	        $BatchFile = [System.IO.Path]::Combine($VsToolsDir, "vsvars32.bat")
	        Get-Batchfile $BatchFile
	        [System.Console]::Title = "Visual Studio " + $version + " Windows Powershell"
	        #add a call to set-consoleicon as seen below...hm...!
    }
}

##############################################################################
## Script: Set-ConsoleIcon.ps1
## By: Aaron Lerch, tiny tiny mods by Hanselman
## Website: www.aaronlerch.com/blog 
## Set the icon of the current console window to the specified icon
## Dot-Source first, like . .\set-consoleicon.ps1
## Usage:  Set-ConsoleIcon [string]
## PS:1 > Set-ConsoleIcon "C:\Icons\special_powershell_icon.ico" 
##############################################################################
 
$WM_SETICON = 0x80
$ICON_SMALL = 0
 
function Set-ConsoleIcon
{
    param(
        [string] $iconFile
    )
 
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null
$iconFile
    # Verify the file exists
    if ([System.IO.File]::Exists($iconFile) -eq $TRUE)
    {
        $icon = new-object System.Drawing.Icon($iconFile) 
 
        if ($icon -ne $null)
        {
            $consoleHandle = GetConsoleWindow
            SendMessage $consoleHandle $WM_SETICON $ICON_SMALL $icon.Handle 
        }
    }
    else
    {
        Write-Host "Icon file not found"
    }
}
 
 
## Invoke a Win32 P/Invoke call.
## From: Lee Holmes
## http://www.leeholmes.com/blog/GetTheOwnerOfAProcessInPowerShellPInvokeAndRefOutParameters.aspx
function Invoke-Win32([string] $dllName, [Type] $returnType, 
   [string] $methodName, [Type[]] $parameterTypes, [Object[]] $parameters) 
{
   ## Begin to build the dynamic assembly
   $domain = [AppDomain]::CurrentDomain
   $name = New-Object Reflection.AssemblyName 'PInvokeAssembly'
   $assembly = $domain.DefineDynamicAssembly($name, 'Run') 
   $module = $assembly.DefineDynamicModule('PInvokeModule')
   $type = $module.DefineType('PInvokeType', "Public,BeforeFieldInit")
 
   ## Go through all of the parameters passed to us.  As we do this, 
   ## we clone the user's inputs into another array that we will use for
   ## the P/Invoke call.  
   $inputParameters = @()
   $refParameters = @()
   
   for($counter = 1; $counter -le $parameterTypes.Length; $counter++) 
   {
      ## If an item is a PSReference, then the user 
      ## wants an [out] parameter.
      if($parameterTypes[$counter - 1] -eq [Ref])
      {
         ## Remember which parameters are used for [Out] parameters 
         $refParameters += $counter
 
         ## On the cloned array, we replace the PSReference type with the 
         ## .Net reference type that represents the value of the PSReference, 
         ## and the value with the value held by the PSReference. 
         $parameterTypes[$counter - 1] = 
            $parameters[$counter - 1].Value.GetType().MakeByRefType()
         $inputParameters += $parameters[$counter - 1].Value
      }
      else
      {
         ## Otherwise, just add their actual parameter to the
         ## input array.
         $inputParameters += $parameters[$counter - 1]
      }
   }
 
   ## Define the actual P/Invoke method, adding the [Out] 
   ## attribute for any parameters that were originally [Ref] 
   ## parameters.
   $method = $type.DefineMethod($methodName, 'Public,HideBySig,Static,PinvokeImpl', 
      $returnType, $parameterTypes) 
   foreach($refParameter in $refParameters)
   {
      $method.DefineParameter($refParameter, "Out", $null)
   }
 
   ## Apply the P/Invoke constructor
   $ctor = [Runtime.InteropServices.DllImportAttribute].GetConstructor([string])
   $attr = New-Object Reflection.Emit.CustomAttributeBuilder $ctor, $dllName
   $method.SetCustomAttribute($attr)
 
   ## Create the temporary type, and invoke the method.
   $realType = $type.CreateType() 
   $realType.InvokeMember($methodName, 'Public,Static,InvokeMethod', $null, $null, 
      $inputParameters)
 
   ## Finally, go through all of the reference parameters, and update the
   ## values of the PSReference objects that the user passed in. 
   foreach($refParameter in $refParameters)
   {
      $parameters[$refParameter - 1].Value = $inputParameters[$refParameter - 1]
   }
}
 
function SendMessage([IntPtr] $hWnd, [Int32] $message, [Int32] $wParam, [Int32] $lParam) 
{
    $parameterTypes = [IntPtr], [Int32], [Int32], [Int32]
    $parameters = $hWnd, $message, $wParam, $lParam
 
    Invoke-Win32 "user32.dll" ([Int32]) "SendMessage" $parameterTypes $parameters 
}
 
function GetConsoleWindow()
{
    Invoke-Win32 "kernel32" ([IntPtr]) "GetConsoleWindow"
}
