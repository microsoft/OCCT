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

function Test-AADJoin {
    try {
        $dsregstatus = dsregcmd /status
    } catch {
        Publish-EventLog -EventId 228 -EntryType Error -Message ($global:resources["EventID-228"])
        return $false
    }

    if ($dsregstatus -match "AzureAdJoined : YES") {
        Publish-EventLog -EventId 132 -EntryType Information -Message ($global:resources["EventID-132"])
        return $true
    } else {
        return $false
    }
}