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