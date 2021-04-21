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

function Reset-FederationProvider {
    try {
        $identities = Get-ChildItem -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\Identities\" -ErrorAction Stop
        $identityCount = ($identities | Measure-Object).count
        Publish-EventLog -EventId 108 -EntryType Information -Message ([string]::Format($global:resources["EventID-108"], $identityCount))
    } catch {
        $identityCount = 0
        $identities = $null
    }

    if ($identityCount -ne 0) {
        foreach ($identity in $identities) {
            Publish-EventLog -EventId 109 -EntryType Information -Message ([string]::Format($global:resources["EventID-109"], $identity))
            $federationProvider = $null
            try {
                $properties = Get-ItemProperty -Path "Registry::$($identity.Name)" -ErrorAction Stop
                $federationProvider = $properties.FederationProvider
            } catch {
                return $false
            }

            # BF Federation provider found, clean
            Publish-EventLog -EventId 111 -EntryType Information -Message ($global:resources["EventID-111"])
            try {
                Publish-EventLog -EventId 110 -EntryType Information -Message ([string]::Format($global:resources["EventID-110"], $federationProvider))
                Set-ItemProperty -Path "Registry::$($identity.Name)" -Name "FederationProvider" -Value "Error" -ErrorAction Stop
                
                return $true
            } catch {
                Publish-EventLog -EventId 206 -EntryType Error -Message ($global:resources["EventID-206"])
                return $false
            }
        }
    } else {
        # no office identities
        Publish-EventLog -EventId 113 -EntryType Information -Message ($global:resources["EventID-113"])
        return $false
    }

    return $false
}