$disclaimer = "#***********************************************************************
#
# Copyright (c) 2018 Microsoft Corporation. All rights reserved.
#
# THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
#**********************************************************************"

$functions = Get-ChildItem .\functions | Sort-Object -Property Name

New-Item OCCT.ps1 -Force
$disclaimer | Out-File OCCT.ps1 -Append
"`r`n" | Out-File OCCT.ps1 -Append

foreach ($function in $functions) {
    $code = Get-Content $(".\functions\" + $($function.Name)) | Select-Object -Skip 14 # skip disclaimer
    $code | Out-File OCCT.ps1 -Append
    "`r`n" | Out-File OCCT.ps1 -Append
}

'Start-OCCT -Tenant $tenant' | Out-File OCCT.ps1 -Append