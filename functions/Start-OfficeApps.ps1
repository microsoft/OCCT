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

function Start-OfficeApps {
    Param(
        [parameter(Mandatory = $true)]
        [PSObject]
        $appsToStart
    )
    foreach ($app in $appsToStart) {
        try {
            if($app.Path) {
                Start-Process $app.Path -ErrorAction Stop
            } else {
                Start-Process $app.ProcessName -ErrorAction Stop
            }
            
            Publish-EventLog -EventId 128 -EntryType Information -Message ([string]::Format($global:resources["EventID-128"], $($app.DisplayName)))
        } catch {
            Publish-EventLog -EventId 224 -EntryType Error -Message ([string]::Format($global:resources["EventID-224"], $($app.DisplayName)))
        }
    }
}