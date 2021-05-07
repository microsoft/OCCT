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

function Test-OCCTdisabled {
    try {
        $occtConfig = Get-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -ErrorAction Stop
    } catch {
        # OCCT hive does not exist. It is only required for custom configuration, ignore.
        return $false
    }
    
    if ($occtConfig) {
        if ($occtConfig.OCCTPSDisabled -and $occtConfig.OCCTPSDisabled -eq 1) {
            return $true
        }
    }
    return $false
}