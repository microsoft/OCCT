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

function Test-RootFederationProvider {
    try {
        $identity = Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\" -ErrorAction Stop
    } catch {
        Publish-EventLog -EventId 221 -EntryType Error -Message ($global:resources["EventID-221"])
        return $false
    }
    
    if ($identity) {
        $federationProvider = $null
        $federationProvider = $identity.FederationProvider
    
        if ($federationProvider -like "*microsoftonline.de*") {
            return $true
        }
    }
    
    return $false
}