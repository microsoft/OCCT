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

function Clear-OLSFederationProvider {
    try {
        $OLSFederationProvider = Get-Item -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Licensing\OLSFederationProvider\" -ErrorAction Stop | Select-Object -ExpandProperty Property
    } catch {
        return $false
    }

    if ($OLSFederationProvider) {
        # get property values
        try {
            $OLSFederationProviderValues = Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Licensing\OLSFederationProvider\" -ErrorAction Stop
        } catch {
            return $false
        }
        
        if ($OLSFederationProviderValues) {
            $clearedOLSFederationProviders = 0
            foreach ($key in $OLSFederationProvider) {
                if ($OLSFederationProviderValues.$key -like "*microsoftonline.de*") {
                    try {
                        Set-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Licensing\OLSFederationProvider\" -Name $key -Value "" -Type String -ErrorAction Stop
                        $clearedOLSFederationProviders++
                    } catch {
                        Publish-EventLog -EventId 233 -EntryType Error -Message ($global:resources["EventID-233"])
                        return $false
                    }
                }
            }
            Publish-EventLog -EventId 139 -EntryType Information -Message ([string]::Format($global:resources["EventID-139"], $clearedOLSFederationProviders))
            return $true
        }
    }
}