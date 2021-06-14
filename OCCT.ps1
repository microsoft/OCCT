#***********************************************************************
#
# Copyright (c) 2018 Microsoft Corporation. All rights reserved.
#
# THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
#**********************************************************************

# region Parameter

Param(
    [Parameter(Mandatory = $false)]
    [String]$tenant
)

#endregion Parameter
#region Functions

function Clear-OlkAutodiscoverCache {
    $result = 0
    try {
        # get Outlook AutoDiscover xml cache files
        $cachePath = $env:LOCALAPPDATA + "\Microsoft\Outlook\16"
        $autoDFiles = Get-ChildItem -Path $cachePath -Filter *.xml -Force -ErrorAction Stop
        $autoDFilesCount = ($autoDFiles | Measure-Object -ErrorAction Stop).Count
    } catch {
        Publish-EventLog -EventId 230 -EntryType Error -Message ($global:resources["EventID-230"])
    }

    if ($autoDFilesCount -gt 0) {
        Publish-EventLog -EventId 136 -EntryType Information -Message ([string]::Format($global:resources["EventID-136"], $autoDFilesCount))

        # Remove files
        $filesRemoved = 0
        $retries = 0

        while (($filesRemoved -lt $autoDFilesCount) -and ($retries -le 2)) {
            foreach ($file in $autoDFiles) {
                try {
                    if (Test-Path $($cachePath + "\" + $file.Name) ) {
                        Remove-Item $($cachePath + "\" + $file.Name) -Force -Confirm:$false -ErrorAction Stop
                        $filesRemoved++
                    }
                } catch {
                    Publish-EventLog -EventId 231 -EntryType Error -Message ([string]::Format($global:resources["EventID-231"], $($file.Name)))
                }
            }
            $retries++
            if (($filesRemoved -lt $autoDFilesCount) -and ($retries -le 2)) {
                Start-Sleep -Seconds 5
            } elseif ($filesRemoved -lt $autoDFilesCount) {
                Publish-EventLog -EventId 234 -EntryType Error -Message ($global:resources["EventID-234"])
                $result = 2
            } else {
                $result = 1
            }
        }
        Publish-EventLog -EventId 137 -EntryType Information -Message ([string]::Format($global:resources["EventID-137"], $filesRemoved))
    }
    return $result
}


function Clear-OLSFederationProvider {
    try {
        $OLSFederationProvider = Get-Item -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Licensing\OLSFederationProvider\" -ErrorAction Stop | Select-Object -ExpandProperty Property
    } catch {
        return $false
    }

    if ($OLSFederationProvider) {
        # get property values
        try {
            $OLSFederationProviderValues = Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Licensing\OLSFederationProvider\" -ErrorAction Stop
        } catch {
            return $false
        }
        
        if ($OLSFederationProviderValues) {
            $clearedOLSFederationProviders = 0
            foreach ($key in $OLSFederationProvider) {
                if ($OLSFederationProviderValues.$key -like "*microsoftonline.de*") {
                    try {
                        Set-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Licensing\OLSFederationProvider\" -Name $key -Value "" -Type String -ErrorAction Stop
                        $clearedOLSFederationProviders++
                    } catch {
                        Publish-EventLog -EventId 233 -EntryType Error -Message ($global:resources["EventID-233"])
                        return $false
                    }
                }
            }
            Publish-EventLog -EventId 139 -EntryType Information -Message ([string]::Format($global:resources["EventID-139"], $clearedOLSFederationProviders))
            return $true
        }
    }
}


function Get-CleanupStatus {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $key
    )

    try {
        $occtConfig = Get-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -ErrorAction Stop
    } catch {
        # OCCT hive does not exist. It is only required for custom configuration, ignore.
        return $false
    }
    
    if ($occtConfig) {
        if ($occtConfig.$key -and $occtConfig.$key -gt 0) {
            return $true
        }
    }
    return $false
}


function Get-MaxOffset {
    $defaultMaxOffset = 900
    try {
        $occtConfig = Get-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -ErrorAction Stop
    } catch {
        # no custom max offset defined, use default max offset
        return $defaultMaxOffset
    }
    
    if ($occtConfig) {
        try {
            [int]$offset = $occtConfig.MaxOffset
        } catch {
            # Offset is not numeric (int), use default value
            return $defaultMaxOffset
        }        
        if ($offset -gt 0 -and $offset -lt 59) {
            Publish-EventLog -EventId 120 -EntryType Information -Message ([string]::Format($global:resources["EventID-120"], $offset.tostring()))
            try {
                $offsetSec = ($offset * 60)
                return $offsetSec
            } catch {
                return $defaultMaxOffset
            }
        }
    }
    return $defaultMaxOffset
}


function Get-SleepTime {
    Param(
        [parameter(Mandatory = $true)]
        [Int32]
        $max
    )
    
    $sleepTime = Get-Random -Minimum 0 -Maximum $max -ErrorAction Stop
    Publish-EventLog -EventId 118 -EntryType Information -Message ([string]::Format($global:resources["EventID-118"], ($sleepTime)))
    return $sleepTime;
}


function Initialize-Resources() {
    $global:resources = New-Object System.Collections.Generic.Dictionary"[String,String]"

    # Information messages
    $global:resources.Add("EventID-100", "OCCT run started")
    $global:resources.Add("EventID-101", "Computer name: {0}{1}User-SID: {2}")
    $global:resources.Add("EventID-102", "Tenant: {0}")
    $global:resources.Add("EventID-103", "Tenant is migrated, check FederationProvider.")
    $global:resources.Add("EventID-104", "Tenant is not migrated.")
    $global:resources.Add("EventID-105", "Running Office apps found, ask user to close apps.")
    $global:resources.Add("EventID-106", "Reset of FederationProvider was successful.")
    $global:resources.Add("EventID-107", "{0} is still running.")
    $global:resources.Add("EventID-108", "{0} Office identities found.")
    $global:resources.Add("EventID-109", "Working on identity {0}")
    $global:resources.Add("EventID-110", "FederationProvider found: {0}")
    $global:resources.Add("EventID-111", "BF FederationProvider found.")
    $global:resources.Add("EventID-112", "FederationProvider does not point to BF.")
    $global:resources.Add("EventID-113", "No Office identities found.")
    $global:resources.Add("EventID-114", "Countdown expired, kill {0}")
    $global:resources.Add("EventID-115", "Updated LastRunTimeSuccess and status in registry.")
    $global:resources.Add("EventID-116", "Amount of internet connection tries: {0}")
    $global:resources.Add("EventID-117", "Client cutover is already done.")
    $global:resources.Add("EventID-118", "Sleep for {0} seconds to avoid throttling.")
    $global:resources.Add("EventID-119", "Querying OIDC/HRD document from: {0}")
    $global:resources.Add("EventID-120", "Found MaxOffset in registry: {0} minutes.")
    $global:resources.Add("EventID-121", "No Office identities found for this user.")
    $global:resources.Add("EventID-122", "OCCT is disabled by registry key.")
    $global:resources.Add("EventID-123", "ForceMode enabled, skip throttling sleep.")
    $global:resources.Add("EventID-124", "BF FederationProvider found in Office identity root hive.")
    $global:resources.Add("EventID-125", "FederationProvider found in Office identity root hive: {0}")
    $global:resources.Add("EventID-126", "FederationProvider in Office identity root hive does not point to BF.")
    $global:resources.Add("EventID-127", "FederationProvider in Office identity root hive was removed successfully.")
    $global:resources.Add("EventID-128", "Office app automatically reopened: {0}")
    $global:resources.Add("EventID-129", "Check for tenant cutover against Office HRD endpoint.")
    $global:resources.Add("EventID-130", "Check for tenant cutover against Azure AD OIDC endpoint.")
    $global:resources.Add("EventID-131", "Office identity hive removed.")
    $global:resources.Add("EventID-132", "Device is AAD joined.")
    $global:resources.Add("EventID-133", "{0} OneDrive for Business accounts found.")
    $global:resources.Add("EventID-134", "Set NextEmailHRDUpdate for OneDrive account {0}.")
    $global:resources.Add("EventID-135", "Updated NextEmailHRDUpdate for {0} of {1} OneDrive accounts. Restart OneDrive client.")
    $global:resources.Add("EventID-136", "{0} Outlook AutoDiscover cache files found.")
    $global:resources.Add("EventID-137", "{0} Outlook AutoDiscover cache files removed.")
    $global:resources.Add("EventID-138", "Signed out from {0} BF WAM accounts.")
    $global:resources.Add("EventID-139", "Cleared {0} BF OLSFederationProviders.")

    # Error messages
    $global:resources.Add("EventID-201", "User-SID could not be retrieved.")
    $global:resources.Add("EventID-202", "No internet connectivity present.")
    $global:resources.Add("EventID-203", "Existing Process already running.")
    $global:resources.Add("EventID-204", "Function only allows one or two parameters.")
    $global:resources.Add("EventID-205", "Reset of FederationProvider was not successful.")
    $global:resources.Add("EventID-206", "Error removing BF FederationProvider.")
    $global:resources.Add("EventID-207", "Failed to update LastRunTimeSuccess and status in registry.")
    $global:resources.Add("EventID-208", "Failed to create LastRunTimeSuccess and status in registry.")
    $global:resources.Add("EventID-209", "Error message: {0}")
    $global:resources.Add("EventID-210", "Log file could not be published.")
    $global:resources.Add("EventID-211", "Error closing app: {0}")
    $global:resources.Add("EventID-212", "Countdown expired, unable to set DialogResult.")
    $global:resources.Add("EventID-213", "Continue button clicked, unable to set DialogResult.")
    $global:resources.Add("EventID-214", "Dialog closed, unable to stop countdown.")
    $global:resources.Add("EventID-215", "Error checking process {0}")
    $global:resources.Add("EventID-216", "Error checking FederationProvider.")
    $global:resources.Add("EventID-217", "OIDC lookup failed. Error Message: {0}{1}")
    $global:resources.Add("EventID-218", "Dialog: XmlNodeReader could not be initialized.")
    $global:resources.Add("EventID-219", "Dialog: XamlReader could not be initialized.")
    $global:resources.Add("EventID-220", "Dialog: Timer could not be initialized.")
    $global:resources.Add("EventID-221", "Error checking root FederationProvider.")
    $global:resources.Add("EventID-222", "Error removing BF FederationProvider in Office identity root hive.")
    $global:resources.Add("EventID-223", "Removal of FederationProvider in Office identity root hive was not successful.")
    $global:resources.Add("EventID-224", "Office app could not be started: {0}")
    $global:resources.Add("EventID-225", "Setting status registry key failed.")
    $global:resources.Add("EventID-226", "Office identity registry hive not found.")
    $global:resources.Add("EventID-227", "Office identity registry hive could not be removed.")
    $global:resources.Add("EventID-228", "Error getting AAD Join status.")
    $global:resources.Add("EventID-229", "Failed to set NextEmailHRDUpdate for OneDrive account {0}.")
    $global:resources.Add("EventID-230", "Failed to get Outlook AutoDiscover cache files.")
    $global:resources.Add("EventID-231", "Failed to remove Outlook AutoDiscover cache file: {0}.")
    $global:resources.Add("EventID-232", "Failed to remove BF WAM accounts.")
    $global:resources.Add("EventID-233", "Failed to clear OLSFederationProvider.")
    $global:resources.Add("EventID-234", "Some Outlook AutoDiscover Cache files could not be removed.")

    $global:resources.Add("LogFilePath", "C:\Temp\OCCT\OCCT_{0}.xml")
}


function Publish-EventLog {
    Param(
        [parameter(Mandatory = $true)]
        [Int32]
        $eventId,
        [parameter(Mandatory = $true)]
        [String]
        $entryType,
        [parameter(Mandatory = $true)]
        [String]
        $message
    )

    try {
        $logfile = $env:TEMP + "\OCCT.log"
        $EntryDay = Get-Date -Format [yyyy-MM-dd]
        $EntryTime = Get-Date -Format HH:mm:ss.ffff
    
        Write-Output "$EntryDay `t $EntryTime `t [$entryType] `t [$eventId] `t $message" | Out-File $logfile -Append
    } catch {
        # The script is running in a context that does not have permissions to create the OCCT Event Log category - thus the catch is empty
    }
}


function Remove-BFWamAccount {
    function Test-BFWAMAccount($accountProperties) {
        foreach ($keyvaluepair in $accountProperties) {
            if ($keyvaluepair.key -eq "Authority") {
                if ($keyvaluepair.value -like "*microsoftonline.de*") {
                    return $true
                }
            }
        }
        return $false
    }
    function AwaitAction($WinRtAction) {
        $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod })[0]
        $netTask = $asTask.Invoke($null, @($WinRtAction))
        $netTask.Wait(-1) | Out-Null
    }
    function Await($WinRtTask, $ResultType) {
        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }

    if (-not [Windows.Foundation.Metadata.ApiInformation, Windows, ContentType = WindowsRuntime]::IsMethodPresent("Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager", "FindAllAccountsAsync")) {
        # This script is not supported on this Windows version
        return $false
    }
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $signedOutBFAccounts = 0

    try {
        $provider = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAccountProviderAsync("https://login.microsoft.com", "organizations")) ([Windows.Security.Credentials.WebAccountProvider, Windows, ContentType = WindowsRuntime])
        $accounts = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows, ContentType = WindowsRuntime]::FindAllAccountsAsync($provider, "d3590ed6-52b3-4102-aeff-aad2292ab01c")) ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult, Windows, ContentType = WindowsRuntime])
        $accounts.Accounts | ForEach-Object { 
            if (Test-BFWAMAccount($_.Properties)) {
                # sign out only if it is a BF WAM account
                AwaitAction ($_.SignOutAsync("d3590ed6-52b3-4102-aeff-aad2292ab01c"))
                $signedOutBFAccounts++
            }
        }
        Publish-EventLog -EventId 138 -EntryType Information -Message ([string]::Format($global:resources["EventID-138"], $signedOutBFAccounts))
        return $true
    } catch {
        Publish-EventLog -EventId 232 -EntryType Error -Message ($global:resources["EventID-232"])
        return $false
    }
}


function Remove-OfficeIdentityHive {
    $regOfficeIdentityPath = "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\"

    try {
        $identity = Get-Item $regOfficeIdentityPath -ErrorAction Stop
    } catch {
        Publish-EventLog -EventId 226 -EntryType Error -Message ($global:resources["EventID-226"])
        return $false
    }

    if ($identity) {
        try {
            Remove-Item $regOfficeIdentityPath -Confirm:$false -ErrorAction Stop
            Publish-EventLog -EventId 131 -EntryType Information -Message ($global:resources["EventID-131"])
            return $true
        } catch {
            Publish-EventLog -EventId 227 -EntryType Error -Message ($global:resources["EventID-227"])
            return $false
        }
    } 

    return $false
}


function Reset-FederationProvider {
    try {
        $identities = Get-ChildItem -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\Identities\" -ErrorAction Stop
        $identityCount = ($identities | Measure-Object).count
        Publish-EventLog -EventId 108 -EntryType Information -Message ([string]::Format($global:resources["EventID-108"], $identityCount))
    } catch {
        $identityCount = 0
        $identities = $null
    }

    if ($identityCount -ne 0) {
        foreach ($identity in $identities) {
            Publish-EventLog -EventId 109 -EntryType Information -Message ([string]::Format($global:resources["EventID-109"], $identity))
            $federationProvider = $null
            try {
                $properties = Get-ItemProperty -Path "Registry::$($identity.Name)" -ErrorAction Stop
                $federationProvider = $properties.FederationProvider
            } catch {
                return $false
            }

            # BF Federation provider found, clean
            Publish-EventLog -EventId 111 -EntryType Information -Message ($global:resources["EventID-111"])
            try {
                Publish-EventLog -EventId 110 -EntryType Information -Message ([string]::Format($global:resources["EventID-110"], $federationProvider))
                Set-ItemProperty -Path "Registry::$($identity.Name)" -Name "FederationProvider" -Value "Error" -ErrorAction Stop
                
                return $true
            } catch {
                Publish-EventLog -EventId 206 -EntryType Error -Message ($global:resources["EventID-206"])
                return $false
            }
        }
    } else {
        # no office identities
        Publish-EventLog -EventId 113 -EntryType Information -Message ($global:resources["EventID-113"])
        return $false
    }

    return $false
}


function Reset-RootFederationProvider {
    if (CheckRootFederationProvider) {
        # BF Federation provider found, clean
        Publish-EventLog -EventId 124 -EntryType Information -Message ($global:resources["EventID-124"])
        try {
            $identity = Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\" -ErrorAction Stop
            $federationProvider = $identity.FederationProvider
            Publish-EventLog -EventId 125 -EntryType Information -Message ([string]::Format($global:resources["EventID-125"], $federationProvider))
            Remove-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\" -Name "FederationProvider" -ErrorAction Stop

            return $true
        } catch {
            Publish-EventLog -EventId 222 -EntryType Error -Message ($global:resources["EventID-222"])
            return $false
        }
    } else {
        Publish-EventLog -EventId 126 -EntryType Information -Message ($global:resources["EventID-126"])
    }

    return $false
}


function Set-CleanupStatus {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $key,
        [parameter(Mandatory = $false)]
        [int]
        $result = 1,
        [parameter(Mandatory = $false)]
        [bool]
        $SetLastRuntime = $false
    )

    try {
        $occtConfig = Get-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -ErrorAction Stop
    } catch {
        # OCCT hive does not exist. Create it.
        try {
            New-Item -Path "Registry::HKCU\Software" -Name OCCT -Force
            if ($SetLastRuntime) {
                $time = ([datetime]::UtcNow).tostring("yyyy-MM-dd HH:mm:ss tt")
                Set-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -Name $key -Value $time -Type String -ErrorAction Stop
            } else {
                Set-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -Name $key -Value $result -Type DWord -ErrorAction Stop
            }
        } catch {
            Publish-EventLog -EventId 225 -EntryType Error -Message ($global:resources["EventID-225"])
        }
    }

    if ($occtConfig) {
        try {
            if ($SetLastRuntime) {
                $time = ([datetime]::UtcNow).tostring("yyyy-MM-dd HH:mm:ss tt")
                Set-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -Name $key -Value $time -Type String -ErrorAction Stop
            } else {
                Set-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -Name $key -Value $result -Type DWord -ErrorAction Stop
            }
        } catch {
            Publish-EventLog -EventId 225 -EntryType Error -Message ($global:resources["EventID-225"])
        }
    }
}


function Show-UserPrompt {
    Param(
        [parameter(Mandatory = $true)]
        [PSObject]
        $runningApps
    )
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
    Add-Type -AssemblyName 'PresentationFramework'
    
    # get display language
    try {
        $lang = $PSUICulture
    } catch {
        # use en-US as fallback if user language cannot be retrieved
        $lang = "en-US"
    }

    if ($lang -eq "de-DE") {
        [xml]$WindowXaml = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="OCCT - Office Client Cutover Tool" 
        Height="450" 
        Width="800" 
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="SingleBorderWindow" 
        Topmost="True">
    <Grid RenderTransformOrigin="0.496,0.347">
        <Border BorderBrush="Red" BorderThickness="3" CornerRadius="0" >
        </Border>
        <Button x:Name="_continue" Content="Erneut prüfen" HorizontalAlignment="Left" Margin="610,320,0,0" VerticalAlignment="Top" Height="47" Width="148" />
        <TextBlock HorizontalAlignment="Left" Margin="22,17,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="47" Width="736" FontSize="14" FontWeight="Bold" Foreground="Black"><Run Text="Damit die Office Applikationen weiter fehlerfrei funktionieren, muss eine dringende Anpassung der Office Konfiguration vorgenommen werden. Führen Sie dazu bitte sofort folgende Schritte aus:"/><Run/></TextBlock>
        <ListBox x:Name="applist" Margin="41,198,539,36" BorderBrush="White" IsEnabled="False" Background="Black"/>
        <TextBlock HorizontalAlignment="Left" Margin="35,64,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="89" Width="745" FontWeight="Normal" Foreground="Black"><Run Text=" 1. Speichern Sie alle in den Office Applikationen (z.B. Word, Excel, Outlook, PowerPoint usw.) geöffneten Dateien lokal (z.B. auf dem Desktop, nicht in OneDrive oder SharePoint Online).&#10;
2. Schließen Sie im Anschluss alle unten aufgeführten Office Applikationen.&#10;
3. Warten Sie 10 Sekunden bevor Sie die Office Apps erneut öffnen.&#10;
4. Öffnen Sie die in Schritt 1 lokal gespeicherten Dateien und speichern Sie diese am gewünschten Ort."/></TextBlock>
        <TextBlock x:Name="runningApps" HorizontalAlignment="Left" Margin="36,174,0,0" Text="Momentan geöffnete Office Apps:" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" FontSize="14"/>
        <TextBlock x:Name="countdown" HorizontalAlignment="Right" Margin="0,379,42,0" Text="Die momentan geöffneten Office Apps werden automatisch beendet in 10:00 Minuten." TextWrapping="Wrap" VerticalAlignment="Top" Width="401" Height="19" TextAlignment="Right"/>
    </Grid>
</Window>
'@
    } else {
        [xml]$WindowXaml = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="OCCT - Office Client Cutover Tool" 
        Height="450" 
        Width="800" 
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="SingleBorderWindow" 
        Topmost="True">
    <Grid RenderTransformOrigin="0.496,0.347">
        <Border BorderBrush="Red" BorderThickness="3" CornerRadius="0" >
        </Border>
        <Button x:Name="_continue" Content="Check again" HorizontalAlignment="Left" Margin="610,320,0,0" VerticalAlignment="Top" Height="47" Width="148" />
        <TextBlock HorizontalAlignment="Left" Margin="22,17,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="47" Width="736" FontSize="14" FontWeight="Bold" Foreground="Black"><Run Text="In order to ensure the proper function of the Office applications, the Office configuration must be adjusted urgently. To do this, please carry out the following steps immediately"/><Run/></TextBlock>
        <ListBox x:Name="applist" Margin="41,198,539,36" BorderBrush="White" IsEnabled="False" Background="Black"/>
        <TextBlock HorizontalAlignment="Left" Margin="35,64,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="89" Width="745" FontWeight="Normal" Foreground="Black"><Run Text=" 1. Save all files opened in the Office applications (e.g. Word, Excel, Outlook, PowerPoint etc.) locally (e.g. on your Desktop, not in OneDrive or SharePoint Online).&#10;
2. Then close all Office applications listed below.&#10;
3. Wait 10 seconds before starting the Office apps again.&#10;
4. Open the files saved locally in step 1 again and save it to the desired location."/></TextBlock>
        <TextBlock x:Name="runningApps" HorizontalAlignment="Left" Margin="36,174,0,0" Text="Office apps currently running:" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" FontSize="14"/>
        <TextBlock x:Name="countdown" HorizontalAlignment="Right" Margin="0,379,42,0" Text="Currently opened Office apps will be closed automatically in 10:00 minutes." TextWrapping="Wrap" VerticalAlignment="Top" Width="401" Height="19" TextAlignment="Right"/>
    </Grid>
</Window>
'@
    }
    try {
        $reader = (New-Object System.Xml.XmlNodeReader $WindowXaml)
    } catch {
        Publish-EventLog -EventId 218 -EntryType Error -Message ($global:resources["EventID-218"])
    }
    try {
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $Window.Topmost = $true
    } catch {
        Publish-EventLog -EventId 219 -EntryType Error -Message ($global:resources["EventID-219"])
    }
    
    # button action
    $button = $Window.FindName("_continue")
    $button.Add_Click( {
            $Timer.Stop(); 
            $Window.Close();
            $Script:Timer = $null
            $Script:CountDown = 600
        })

    # add running apps to app list
    $runningApps = $runningApps | Sort-Object -Property DisplayName
    $listbox = $Window.FindName('applist')
    foreach ($app in $runningApps) {
        $listbox.Items.Add($app.DisplayName) | Out-Null
    }

    # get countdown label to update remaining seconds
    $countdownLabel = $Window.FindName('countdown')

    # create Timer
    if (!$Script:Timer) {
        try {
            $Script:Timer = New-Object System.Windows.Forms.Timer
            $Script:Timer.Interval = 1000
        } catch {
            Publish-EventLog -EventId 220 -EntryType Error -Message ($global:resources["EventID-220"])
        }
    }

    function Timer_Tick() {
        $Script:time = ""

        $ts = [timespan]::fromseconds($Script:CountDown)
        $Script:time = "{0:mm:ss}" -f ([datetime]$ts.Ticks)

        if ($lang -eq "de-DE") {
            $countdownLabel.Text = "Die momentan geÃ¶ffneten Office Apps werden automatisch beendet in {0} Minuten." -f $Script:time
        } else {
            $countdownLabel.Text = "Currently opended Office apps will be closed automatically in {0} minutes." -f $Script:time
        }
        --$Script:CountDown
        If ($Script:CountDown -eq 0) {
            $Timer.Stop(); 
            $Window.Close();
            $Script:Timer = $null
            $Script:CountDown = 600
            $Script:killedApplications = KillOfficeApps
        }
    }
    $Script:CountDown = 600
    $Timer.Add_Tick( { Timer_Tick })
    $Timer.Start()	

    # Add closing event
    $Window.Add_Closing( { 
            $Timer.Stop()
            $Timer.Dispose()
            $Script:Timer = $null
            $Script:CountDown = 600
        })
    # Show Window
    $Window.ShowDialog() | Out-Null

    return $Script:killedApplications
}


function Start-OCCT {
    Param(
        [parameter(Mandatory = $false)]
        [String]
        $tenant,
        [parameter(Mandatory = $false)]
        [bool]
        $force = $false,
        [parameter(Mandatory = $false)]
        [bool]
        $resetRootFedProvider = $true,
        [parameter(Mandatory = $false)]
        [bool]
        $reopenOfficeApps = $false,
        [parameter(Mandatory = $false)]
        [bool]
        $officeHRDLookup = $false,
        [parameter(Mandatory = $false)]
        [bool]
        $removeOfficeIdentityHive = $false,
        [parameter(Mandatory = $false)]
        [bool]
        $clearOlkAutodiscoverCache = $false,
        [parameter(Mandatory = $false)]
        [bool]
        $updateODBClient = $true,
        [parameter(Mandatory = $false)]
        [bool]
        $removeBFWamAccount = $true
    )

    $forceMode = $force
    $tenantName = $tenant
    
    Initialize-Resources

    Publish-EventLog -EventId 100 -EntryType Information -Message ($global:resources["EventID-100"])

    # Test if OCCT is already running
    if (Test-ExistingProcess) {
        Publish-EventLog -EventId 203 -EntryType Error -Message ($global:resources["EventID-203"])
        exit
    }
    
    #check if OCCT is disabled
    if (Test-OCCTdisabled) {
        # OCCT is disabled, end program
        Publish-EventLog -EventId 122 -EntryType Information -Message ($global:resources["EventID-122"])
        Exit
    }

    # Read User Sid and computername
    try {
        $currentUser = whoami /user /nh
        $sid = $currentUser.split(" ")[1]
        Publish-EventLog -EventId 101 -EntryType Information -Message ([string]::Format($global:resources["EventID-101"], $env:COMPUTERNAME, [environment]::NewLine, $sid))
    } catch {
        Publish-EventLog -EventId 201 -EntryType Error -Message ($global:resources["EventID-201"])
    }

    # Remove Office identity hive (if requested)
    if ($removeOfficeIdentityHive) {
        Remove-OfficeIdentityHive
        exit
    }

    if($tenantName) {
        # check if client is affected
        $checkFederationProviderResult = Test-FederationProvider
        if ($resetRootFedProvider -eq $true) {
            $checkRootFederationProviderResult = Test-RootFederationProvider
        } else {
            $checkRootFederationProviderResult = $false
        }
        
        if ((($checkFederationProviderResult -eq $false) -and ($checkRootFederationProviderResult -eq $false))) {
            # No BF Federationprovide-r found on client
            Publish-EventLog -EventId 117 -EntryType Information -Message ($global:resources["EventID-117"])
        }

        if (!$forceMode) {
            try {
                $offset = Get-MaxOffset
                Start-Sleep -seconds (Get-SleepTime($offset))
            } catch {
                # Continue without sleep, if sth in sleep function fails
            }
        } else {
            Publish-EventLog -EventId 123 -EntryType Information -Message ($global:resources["EventID-123"])
        }

        # Check internet connectivity.
        # Exit program in case internet connectivity is missing.
        if (!(Test-InternetConnectivity)) {
            Publish-EventLog -EventId 202 -EntryType Error -Message ($global:resources["EventID-202"])
            exit
        }
        
        Publish-EventLog -EventId 102 -EntryType Information -Message ([string]::Format($global:resources["EventID-102"], $tenantName))

        $tenantIsMigrated = $false;
        try {
            if ($officeHRDLookup) {
                Publish-EventLog -EventId 129 -EntryType Information -Message ([string]::Format($global:resources["EventID-129"], $tenantName))
                $tenantIsMigrated = Test-OfficeHRDTenantCutover($tenantName)
            } else {
                Publish-EventLog -EventId 130 -EntryType Information -Message ([string]::Format($global:resources["EventID-130"], $tenantName))
                $tenantIsMigrated = Test-TenantCutover($tenantName)
            }
        } catch {
            Publish-EventLog -EventId 217 -EntryType Error -Message ([string]::Format($global:resources["EventID-217"], [Environment]::NewLine, $_.Exception.Message))
            exit

        }
    } else {
        $tenantIsMigrated = $true
    }
    if ($tenantIsMigrated) {
        Publish-EventLog -EventId 103 -EntryType Information -Message ($global:resources["EventID-103"])
        
        #check if Office apps are running
        $promptCount = 0
        do {
            $runningApps = Test-OfficeProcesses
            if (($runningApps | Measure-object).count -gt 0) {
                Publish-EventLog -EventId 105 -EntryType Information -Message ($global:resources["EventID-105"])
                $killedApps = Show-UserPrompt($runningApps)

                # wait for Office processes to end before checking again
                Start-Sleep -Seconds 4

                # check again for running apps
                $runningApps = Test-OfficeProcesses
            }
            $promptCount++
        } while ((($runningApps | Measure-Object).count -gt 0) -and ($promptCount -lt 10));

        $result = Reset-FederationProvider
        if ($result) {
            Publish-EventLog -EventId 106 -EntryType Information -Message ($global:resources["EventID-106"])
        } else {
            Publish-EventLog -EventId 205 -EntryType Error -Message ($global:resources["EventID-205"])
        }
        
        # Reset root FederationProvider
        if ($resetRootFedProvider -eq $true -and $checkRootFederationProviderResult -eq $true) {
            $resultRootFederationProvider = Reset-RootFederationProvider
            if ($resultRootFederationProvider) {
                Publish-EventLog -EventId 127 -EntryType Information -Message ($global:resources["EventID-127"])
            } else {
                Publish-EventLog -EventId 223 -EntryType Error -Message ($global:resources["EventID-223"])
            }
        }

        # Remove BF WAM Account
        if($removeBFWamAccount) {
            Remove-BFWamAccount
        }
        
        # Clear Outlook AutoDiscover Cache
        if($clearOlkAutodiscoverCache) {
            Clear-OlkAutodiscoverCache
        }
        
        # Update OneDrive Client
        if($updateODBClient) {
            Update-ODBClient
        }

        # Reopen Office apps
        if ($reopenOfficeApps -and ($promptCount -lt 10)) {
            Start-Sleep -Seconds 1
            Start-OfficeApps $killedApps
        }
        

        # Save last successfull runtime in registry
        Set-CleanupStatus -key "LastRunTimeSuccess" -SetLastRuntime $true

        exit
    } else {
        Publish-EventLog -EventId 104 -EntryType Information -Message ($global:resources["EventID-104"])
        
        # Save last successfull runtime in registry
        Set-CleanupStatus -key "LastRunTimeSuccess" -SetLastRuntime $true

        exit
    }
}


function Start-OfficeApps {
    Param(
        [parameter(Mandatory = $true)]
        [PSObject]
        $appsToStart
    )
    foreach ($app in $appsToStart) {
        try {
            if($app.Path) {
                Start-Process $app.Path -ErrorAction Stop
            } else {
                Start-Process $app.ProcessName -ErrorAction Stop
            }
            
            Publish-EventLog -EventId 128 -EntryType Information -Message ([string]::Format($global:resources["EventID-128"], $($app.DisplayName)))
        } catch {
            Publish-EventLog -EventId 224 -EntryType Error -Message ([string]::Format($global:resources["EventID-224"], $($app.DisplayName)))
        }
    }
}


function Stop-OfficeApps {
    $appsToClose = Test-OfficeProcesses
    $killedApps = @()
    foreach ($app in $appsToClose) {
        Publish-EventLog -EventId 114 -EntryType Information -Message ([string]::Format($global:resources["EventID-114"], $app.DisplayName))

        # Ignore, if Get-Process is unable to find the process. Process may have been ended by user.
        foreach ($process in Get-Process $($app.ProcessName) -ErrorAction SilentlyContinue) {
            try {
                # Save path of Office app for restart
                $path = ($process | Select-Object Path).Path
                if (!([string]::IsNullOrEmpty($path))) {
                    $app.Path = $path
                }

                Stop-Process -ID $($process.Id) -Force -Confirm:$false -ErrorAction Stop
                $killedApps += $app
            }
            catch {
                Start-Sleep -Seconds 2
                try {
                    Stop-Process -ID $($process.Id) -Force -Confirm:$false -ErrorAction Stop
                    $killedApps += $app
                }
                catch {
                    Publish-EventLog -EventId 211 -EntryType Error -Message ([string]::Format($global:resources["EventID-211"], $_.Exception.Message))
                }
            }
        }
    }
    return $killedApps
}


function Test-AADJoin {
    try {
        $dsregstatus = dsregcmd /status
    } catch {
        Publish-EventLog -EventId 228 -EntryType Error -Message ($global:resources["EventID-228"])
        return $false
    }

    if ($dsregstatus -match "AzureAdJoined : YES") {
        Publish-EventLog -EventId 132 -EntryType Information -Message ($global:resources["EventID-132"])
        return $true
    } else {
        return $false
    }
}


function Test-ExistingProcess() {
    $me = $PID
    $scriptName = $MyInvocation.MyCommand.Name
    try {
        $result = Get-WmiObject Win32_Process -Filter "Name='PowerShell.EXE'" -ErrorAction Stop |  Where-Object { $_.CommandLine -like "*$($scriptName)*" -and $_.ProcessId -ne $me } -ErrorAction Stop | Measure-Object -ErrorAction Stop
    } catch {
        $result = $null
    }
    if ($result.count -gt 0) {
        return $true;
    } else {
        return $false;
    }
}


function Test-FederationProvider {
    try {
        $identities = Get-ChildItem -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\Identities\" -ErrorAction Stop
        $identityCount = ($identities | Measure-Object).count
        
        if ($identityCount -ne 0) {
            foreach ($identity in $identities) {
            
                $federationProvider = $null
            
                $properties = Get-ItemProperty -Path "Registry::$($identity.Name)" -ErrorAction Stop
                $federationProvider = $properties.FederationProvider
    
                if ($federationProvider -like "*microsoftonline.de*") {
                    return $true
                }
            }
        }
    } catch {
        Publish-EventLog -EventId 216 -EntryType Error -Message ($global:resources["EventID-216"])
        return $false
    }
    return $false
}


function Test-InternetConnectivity() {
    $internetConnectivity = $false
    $counter = 0
    $retryCount = 30
    $internetConnectivityURI = "http://www.msftncsi.com/ncsi.txt"
    $connectivityCheckExpectedResult = "Microsoft NCSI"
    $iCMaximumSleepTime = 10
    $iCRequestTimeOut = 5


    # Check network connectivity.
    do {
        try {
            $webProxy = ([System.Net.WebProxy]::GetDefaultProxy()).getProxy([uri]($internetConnectivityURI)).AbsoluteUri
        } catch {
            $webProxy = $null
            # No web proxy detected, go on without proxy
        }
        $watch = $null
        try {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            if ($webProxy -eq $internetConnectivityURI) {
                $response = Invoke-WebRequest -Uri $internetConnectivityURI -UseBasicParsing -TimeoutSec $iCRequestTimeOut -ErrorAction Stop
            } else {
                $response = Invoke-WebRequest -Uri $internetConnectivityURI -UseBasicParsing -Proxy $webProxy -ProxyUseDefaultCredentials:$true -TimeoutSec $iCRequestTimeOut -ErrorAction Stop
            }
            $watch.Stop()
            if ($response.Content -eq $connectivityCheckExpectedResult -and $response.StatusCode -eq 200 ) {
                $internetConnectivity = $true
            }
        } catch {
            if ($null -ne $watch) {
                $watch.Stop()
            }
        }
        
        $counter++
        if ($internetConnectivity) {
            break
        }
        # Wait 10 seconds
        $sleepInterval = [int]($iCMaximumSleepTime - $watch.Elapsed.Seconds)
        if ($sleepInterval -ge 0) {
            Start-Sleep -Seconds $sleepInterval
        } # No else, because wait time below 0 is not possible
    } while ($counter -lt $retryCount)

    Publish-EventLog -EventId 116 -EntryType Information -Message ([string]::Format($global:resources["EventID-116"], $counter))
    return $internetConnectivity;
}


function Test-OCCTdisabled {
    try {
        $occtConfig = Get-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -ErrorAction Stop
    } catch {
        # OCCT hive does not exist. It is only required for custom configuration, ignore.
        return $false
    }
    
    if ($occtConfig) {
        if ($occtConfig.OCCTPSDisabled -and $occtConfig.OCCTPSDisabled -eq 1) {
            return $true
        }
    }
    return $false
}


function Test-OfficeHRDTenantCutover {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $tenantName
    )

    $OfficeHRDBaseURL = "https://odc.officeapps.live.com/odc/v2.1/federationProvider?domain=<tenant>.onmicrosoft.de"
    $targetEnvironmentName = "Global"
    $targetAuthenticationEndpoint = "https://login.windows.net/common/oauth2/authorize"
    $targetEndpointType = "OrgID"

    $hrdURI = $OfficeHRDBaseURL.Replace("<tenant>", $tenantName)
    Publish-EventLog -EventId 119 -EntryType Information -Message ([string]::Format($global:resources["EventID-119"], $hrdURI))
    
    try {
        $proxy = ([System.Net.WebProxy]::GetDefaultProxy()).getProxy([uri]($hrdURI)).AbsoluteUri
    } catch {
        $proxy = $hrdURI
    }

    try {

        if ($proxy -eq $hrdURI) {
            $response = Invoke-WebRequest -Uri $hrdURI -UseBasicParsing -ErrorAction Stop
        } else {
            $response = Invoke-WebRequest -Uri $hrdURI -UseBasicParsing -Proxy $proxy -ProxyUseDefaultCredentials:$true -ErrorAction Stop
        }

        $hrdResponse = $response.Content | ConvertFrom-Json -ErrorAction Stop

    } catch {
        # For implementing the throtteling check we need an example response from PG - i do not know how an answer looks like
        Publish-EventLog -EventId 209 -EntryType Error -Message ([string]::Format($global:resources["EventID-209"], $_.Exception.Message))
        
        try {
            $e = $_ | ConvertFrom-Json -ErrorAction Stop
            Publish-EventLog -EventId 209 -EntryType Error -Message "$($e.error_description)"
        } catch {
            # No catch action - this is a bonus check to determine whether AAD throws a specific error for the requested tenant. The generic error message is handled below via the exit of the script.
        }

        exit
    }
    
    try {
        $global:fileLog.EventIDs.EventID_209 = 0
    } catch {
        # No catch necessary - try necessary in case $global:fileLog could not be initialized properly
    }

    $hrdAuthEndpoint = $null

    try {
        $hrdAuthEndpoint = $hrdResponse.endpoint | Where-Object { $_.type -eq $targetEndpointType } -ErrorAction Stop | Select-Object -ExpandProperty authentication_endpoint -ErrorAction Stop
    } catch {
        $hrdAuthEndpoint = $null
    }

    # Check whether tenant is post stage 3
    if ($hrdResponse.Environment -eq $targetEnvironmentName -and $hrdAuthEndpoint -eq $targetAuthenticationEndpoint) {
        # tenant is migrated
        return $true;
    } else {
        # tenant is not migrated
        return $false;
    }
}


function Test-OfficeProcesses {
    # get all running Office apps
    $apps = @()
    $runningApps = @()

    $word = [PSCustomObject]@{
        DisplayName = 'Word'
        ProcessName = 'winword'
        Running     = $false
    }
    $apps += $word

    $excel = [PSCustomObject]@{
        DisplayName = 'Excel'
        ProcessName = 'excel'
        Running     = $false
    }
    $apps += $excel

    $access = [PSCustomObject]@{
        DisplayName = 'Access'
        ProcessName = 'msaccess'
        Running     = $false
    }
    $apps += $access

    $onenote = [PSCustomObject]@{
        DisplayName = 'OneNote'
        ProcessName = 'onenote'
        Running     = $false
    }
    $apps += $onenote

    $outlook = [PSCustomObject]@{
        DisplayName = 'Outlook'
        ProcessName = 'outlook'
        Running     = $false
    }
    $apps += $outlook

    $powerpoint = [PSCustomObject]@{
        DisplayName = 'PowerPoint'
        ProcessName = 'powerpnt'
        Running     = $false
    }
    $apps += $powerpoint

    $publisher = [PSCustomObject]@{
        DisplayName = 'Publisher'
        ProcessName = 'mspub'
        Running     = $false
    }
    $apps += $publisher

    $visio = [PSCustomObject]@{
        DisplayName = 'Visio'
        ProcessName = 'visio'
        Running     = $false
    }
    $apps += $visio

    $lync = [PSCustomObject]@{
        DisplayName = 'Skype for Business'
        ProcessName = 'lync'
        Running     = $false
    }
    $apps += $lync

    $project = [PSCustomObject]@{
        DisplayName = 'Project'
        ProcessName = 'winproj'
        Running     = $false
    }
    $apps += $project

    foreach ($app in $apps) {
        try {
            if (Get-Process -Name $app.ProcessName -ErrorAction SilentlyContinue) {
                $app.Running = $true
                $runningApps += $app
            }
        } catch {
            Publish-EventLog -EventId 215 -EntryType Error -Message ([string]::Format($global:resources["EventID-215"], $app.ProcessName))
            $runningApps = @()
        }
        
    }

    return $runningApps
}


function Test-RootFederationProvider {
    try {
        $identity = Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\" -ErrorAction Stop
    } catch {
        Publish-EventLog -EventId 221 -EntryType Error -Message ($global:resources["EventID-221"])
        return $false
    }
    
    if ($identity) {
        $federationProvider = $null
        $federationProvider = $identity.FederationProvider
    
        if ($federationProvider -like "*microsoftonline.de*") {
            return $true
        }
    }
    
    return $false
}


function Test-TenantCutover {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $tenantName
    )
    
    $openIDBaseURL = "https://login.microsoftonline.com/<tenant>.onmicrosoft.de/.well-known/openid-configuration";
    $targetCloudInstanceName = "microsoftonline.com";
    $migrationFlag = "migrating_from_germany_to_global";

    $oidcURI = $openIDBaseURL.Replace("<tenant>", $tenantName)
    Publish-EventLog -EventId 119 -EntryType Information -Message ([string]::Format($global:resources["EventID-119"], $oidcURI))
    
    try {
        $proxy = ([System.Net.WebProxy]::GetDefaultProxy()).getProxy([uri]($oidcURI)).AbsoluteUri
    } catch {
        $proxy = $oidcURI
    }

    try {

        if ($proxy -eq $oidcURI) {
            $response = Invoke-WebRequest -Uri $oidcURI -UseBasicParsing -ErrorAction Stop
        } else {
            $response = Invoke-WebRequest -Uri $oidcURI -UseBasicParsing -Proxy $proxy -ProxyUseDefaultCredentials:$true -ErrorAction Stop
        }

        $oidc = $response.Content | ConvertFrom-Json -ErrorAction Stop

    } catch {
        # For implementing the throtteling check we need an example response from PG - i do not know how an answer looks like
        Publish-EventLog -EventId 209 -EntryType Error -Message ([string]::Format($global:resources["EventID-209"], $_.Exception.Message))

        try {
            $e = $_ | ConvertFrom-Json -ErrorAction Stop
            Publish-EventLog -EventId 209 -EntryType Error -Message "$($e.error_description)"
        } catch {
            # No catch action - this is a bonus check to determine whether AAD throws a specific error for the requested tenant. The generic error message is handled below via the exit of the script.
        }

        exit
    }
    
    try {
        $global:fileLog.EventIDs.EventID_209 = 0
    } catch {
        # No catch necessary - try necessary in case $global:fileLog could not be initialized properly
    }

    # Check whether tenant is post stage 3
    if (($oidc.PSobject.Properties.name -notcontains $migrationFlag) -and ($oidc.cloud_instance_name -eq $targetCloudInstanceName)) {
        # tenant is migrated
        return $true;
    } else {
        # tenant is not migrated
        return $false;
    }
}


function Update-ODBClient {
    $updateDone = $false
    # get all OneDrive accounts
    try {
        $accounts = Get-ChildItem -Path "Registry::HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\" -ErrorAction Stop | Where-Object { $_.Name -like "*Business*" }
        $accountsCount = ($accounts | Measure-Object -ErrorAction Stop).count
        Publish-EventLog -EventId 133 -EntryType Information -Message ([string]::Format($global:resources["EventID-133"], $accountsCount))
    } catch {
        $accountsCount = 0
        $accounts = $null
    }

    if ($accountsCount -ne 0) {
        $successfullyProcessed = 0
        # get epoch time 1sec in past
        [int]$pastEpoch = ([System.DateTimeOffset]::Now.ToUnixTimeSeconds()) - 1

        foreach ($account in $accounts) {
            try {
                Set-ItemProperty -Path $("Registry::" + $($account.Name) + "\AuthenticationURLs") -Name NextEmailHRDUpdate -Value $pastEpoch -Type Qword -ErrorAction Stop
                $successfullyProcessed++
                Publish-EventLog -EventId 134 -EntryType Information -Message ([string]::Format($global:resources["EventID-134"], $($account.Name.split("\")[-1])))
                $updateDone = $true
            } catch {
                Publish-EventLog -EventId 229 -EntryType Error -Message ([string]::Format($global:resources["EventID-229"], $($account.Name.split("\")[-1])))
            }
        }

        Publish-EventLog -EventId 135 -EntryType Information -Message ([string]::Format($global:resources["EventID-135"], $successfullyProcessed, $accountsCount))

        # Restart OneDrive client
        $odb = [PSCustomObject]@{
            DisplayName = 'OneDrive'
            ProcessName = 'OneDrive'
            Path        = $env:LOCALAPPDATA + "\Microsoft\OneDrive\OneDrive.exe"
            Running     = $true
        }
        $stoppedApps = Stop-OfficeApps $odb

        if ($stoppedApps) {
            Start-Sleep -Seconds 1
            Start-OfficeApps $stoppedApps
        }
    }
    return $updateDone
}

#endregion Functions
#region Main

if($tenant) {
    Start-OCCT -Tenant $tenant
} else {
    Start-OCCT
}

#endregion Main
