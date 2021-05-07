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