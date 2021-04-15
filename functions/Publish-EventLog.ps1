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