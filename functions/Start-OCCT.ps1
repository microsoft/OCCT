#***********************************************************************
#
# Copyright (c) 2018 Microsoft Corporation. All rights reserved.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
#**********************************************************************​

function Start-OCCT {
    Param(
        [parameter(Mandatory = $true)]
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
        $clearOlkAutodiscoverCache = $true,
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

    # check if client is affected
    $checkFederationProviderResult = Test-FederationProvider
    if ($resetRootFedProvider -eq $true) {
        $checkRootFederationProviderResult = Test-RootFederationProvider
    } else {
        $checkRootFederationProviderResult = $false
    }
    
    if ((($checkFederationProviderResult -eq $false) -and ($checkRootFederationProviderResult -eq $false))) {
        # No BF Federationprovider found on client
        Publish-EventLog -EventId 117 -EntryType Information -Message ($global:resources["EventID-117"])
    }

    if (!$forceMode) {
        try {
            $offset = GetMaxOffset
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