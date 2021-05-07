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