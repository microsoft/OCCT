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