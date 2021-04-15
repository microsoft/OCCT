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

function Test-OfficeHRDTenantCutover {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $tenantName
    )

    $OfficeHRDBaseURL = "https://odc.officeapps.live.com/odc/v2.1/federationProvider?domain=<tenant>.onmicrosoft.de"
    $targetEnvironmentName = "Global"
    $targetAuthenticationEndpoint = "https://login.windows.net/common/oauth2/authorize"
    $targetEndpointType = "OrgID"

    $hrdURI = $OfficeHRDBaseURL.Replace("<tenant>", $tenantName)
    Publish-EventLog -EventId 119 -EntryType Information -Message ([string]::Format($global:resources["EventID-119"], $hrdURI))
    
    try {
        $proxy = ([System.Net.WebProxy]::GetDefaultProxy()).getProxy([uri]($hrdURI)).AbsoluteUri
    } catch {
        $proxy = $hrdURI
    }

    try {

        if ($proxy -eq $hrdURI) {
            $response = Invoke-WebRequest -Uri $hrdURI -UseBasicParsing -ErrorAction Stop
        } else {
            $response = Invoke-WebRequest -Uri $hrdURI -UseBasicParsing -Proxy $proxy -ProxyUseDefaultCredentials:$true -ErrorAction Stop
        }

        $hrdResponse = $response.Content | ConvertFrom-Json -ErrorAction Stop

    } catch {
        # For implementing the throtteling check we need an example response from PG - i do not know how an answer looks like
        Publish-EventLog -EventId 209 -EntryType Error -Message ([string]::Format($global:resources["EventID-209"], $_.Exception.Message))
        
        try {
            $e = $_ | ConvertFrom-Json -ErrorAction Stop
            Publish-EventLog -EventId 209 -EntryType Error -Message "$($e.error_description)"
        } catch {
            # No catch action - this is a bonus check to determine whether AAD throws a specific error for the requested tenant. The generic error message is handled below via the exit of the script.
        }

        exit
    }
    
    try {
        $global:fileLog.EventIDs.EventID_209 = 0
    } catch {
        # No catch necessary - try necessary in case $global:fileLog could not be initialized properly
    }

    $hrdAuthEndpoint = $null

    try {
        $hrdAuthEndpoint = $hrdResponse.endpoint | Where-Object { $_.type -eq $targetEndpointType } -ErrorAction Stop | Select-Object -ExpandProperty authentication_endpoint -ErrorAction Stop
    } catch {
        $hrdAuthEndpoint = $null
    }

    # Check whether tenant is post stage 3
    if ($hrdResponse.Environment -eq $targetEnvironmentName -and $hrdAuthEndpoint -eq $targetAuthenticationEndpoint) {
        # tenant is migrated
        return $true;
    } else {
        # tenant is not migrated
        return $false;
    }
}