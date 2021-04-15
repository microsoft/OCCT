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