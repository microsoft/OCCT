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