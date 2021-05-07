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

function Test-FederationProvider {
    try {
        $identities = Get-ChildItem -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\Identities\" -ErrorAction Stop
        $identityCount = ($identities | Measure-Object).count
        
        if ($identityCount -ne 0) {
            foreach ($identity in $identities) {
            
                $federationProvider = $null
            
                $properties = Get-ItemProperty -Path "Registry::$($identity.Name)" -ErrorAction Stop
                $federationProvider = $properties.FederationProvider
    
                if ($federationProvider -like "*microsoftonline.de*") {
                    return $true
                }
            }
        }
    } catch {
        Publish-EventLog -EventId 216 -EntryType Error -Message ($global:resources["EventID-216"])
        return $false
    }
    return $false
}