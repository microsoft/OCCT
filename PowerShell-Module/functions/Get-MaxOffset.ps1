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

function Get-MaxOffset {
    $defaultMaxOffset = 900
    try {
        $occtConfig = Get-ItemProperty -Path "Registry::HKCU\Software\OCCT\" -ErrorAction Stop
    } catch {
        # no custom max offset defined, use default max offset
        return $defaultMaxOffset
    }
    
    if ($occtConfig) {
        try {
            [int]$offset = $occtConfig.MaxOffset
        } catch {
            # Offset is not numeric (int), use default value
            return $defaultMaxOffset
        }        
        if ($offset -gt 0 -and $offset -lt 59) {
            Publish-EventLog -EventId 120 -EntryType Information -Message ([string]::Format($global:resources["EventID-120"], $offset.tostring()))
            try {
                $offsetSec = ($offset * 60)
                return $offsetSec
            } catch {
                return $defaultMaxOffset
            }
        }
    }
    return $defaultMaxOffset
}