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

function Remove-OfficeIdentityHive {
    $regOfficeIdentityPath = "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\"

    try {
        $identity = Get-Item $regOfficeIdentityPath -ErrorAction Stop
    } catch {
        Publish-EventLog -EventId 226 -EntryType Error -Message ($global:resources["EventID-226"])
        return $false
    }

    if ($identity) {
        try {
            Remove-Item $regOfficeIdentityPath -Confirm:$false -ErrorAction Stop
            Publish-EventLog -EventId 131 -EntryType Information -Message ($global:resources["EventID-131"])
            return $true
        } catch {
            Publish-EventLog -EventId 227 -EntryType Error -Message ($global:resources["EventID-227"])
            return $false
        }
    } 

    return $false
}