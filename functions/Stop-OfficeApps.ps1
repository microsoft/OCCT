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