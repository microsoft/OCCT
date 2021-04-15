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

function Reset-RootFederationProvider {
    if (CheckRootFederationProvider) {
        # BF Federation provider found, clean
        Publish-EventLog -EventId 124 -EntryType Information -Message ($global:resources["EventID-124"])
        try {
            $identity = Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\" -ErrorAction Stop
            $federationProvider = $identity.FederationProvider
            Publish-EventLog -EventId 125 -EntryType Information -Message ([string]::Format($global:resources["EventID-125"], $federationProvider))
            Remove-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Office\16.0\Common\Identity\" -Name "FederationProvider" -ErrorAction Stop

            return $true
        } catch {
            Publish-EventLog -EventId 222 -EntryType Error -Message ($global:resources["EventID-222"])
            return $false
        }
    } else {
        Publish-EventLog -EventId 126 -EntryType Information -Message ($global:resources["EventID-126"])
    }

    return $false
}