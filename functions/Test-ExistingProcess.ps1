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