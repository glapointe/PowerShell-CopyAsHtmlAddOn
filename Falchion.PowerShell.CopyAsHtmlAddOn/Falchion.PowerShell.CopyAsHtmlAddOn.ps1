function Get-RegistryValue 
{
  param($key, $name)
  
  $key = $key -replace ':',''
  $regkey = "Registry::$key"
  
  Get-ItemProperty -Path $regkey -Name $name | Select-Object -ExpandProperty $name
}

function Get-CopyAsHtmlInstallPath
{
    try
    {
        Get-RegistryValue -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Falchion Consulting, LLC\ISECopyAsHtml\' -Name InstallPath
    }
    catch
    {
        try
        {
            Get-RegistryValue -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Falchion Consulting, LLC\ISECopyAsHtml\' -Name InstallPath
        }
        catch
        {
        }
    }
}

$ErrorActionPreference = "Stop"
$installPath = Get-CopyAsHtmlInstallPath
$ErrorActionPreference = "Continue"

Add-Type -Path "$installPath\Falchion.PowerShell.CopyAsHtmlAddOn.dll"

$copyAsHtmlAddOn = New-Object Falchion.PowerShell.CopyAsHtmlAddOn.CopyAsHtml
$copyAsHtmlAddOn.HostObject = $psISE

$psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Copy as HTML", { $copyAsHtmlAddOn.Copy() }, "CTRL+SHIFT+C") | Out-Null