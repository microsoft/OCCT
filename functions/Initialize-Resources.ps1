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

function Initialize-Resources() {
    $global:resources = New-Object System.Collections.Generic.Dictionary"[String,String]"

    # Information messages
    $global:resources.Add("EventID-100", "OCCT run started, Version {0}")
    $global:resources.Add("EventID-101", "Computer name: {0}{1}User-SID: {2}")
    $global:resources.Add("EventID-102", "Tenant: {0}")
    $global:resources.Add("EventID-103", "Tenant is migrated, check FederationProvider.")
    $global:resources.Add("EventID-104", "Tenant is not migrated.")
    $global:resources.Add("EventID-105", "Running Office apps found, ask user to close apps.")
    $global:resources.Add("EventID-106", "Reset of FederationProvider was successful.")
    $global:resources.Add("EventID-107", "{0} is still running.")
    $global:resources.Add("EventID-108", "{0} Office identities found.")
    $global:resources.Add("EventID-109", "Working on identity {0}")
    $global:resources.Add("EventID-110", "FederationProvider found: {0}")
    $global:resources.Add("EventID-111", "BF FederationProvider found.")
    $global:resources.Add("EventID-112", "FederationProvider does not point to BF.")
    $global:resources.Add("EventID-113", "No Office identities found.")
    $global:resources.Add("EventID-114", "Countdown expired, kill {0}")
    $global:resources.Add("EventID-115", "Updated LastRunTimeSuccess and status in registry.")
    $global:resources.Add("EventID-116", "Amount of internet connection tries: {0}")
    $global:resources.Add("EventID-117", "Client cutover is already done.")
    $global:resources.Add("EventID-118", "Sleep for {0} seconds to avoid throttling.")
    $global:resources.Add("EventID-119", "Querying OIDC/HRD document from: {0}")
    $global:resources.Add("EventID-120", "Found MaxOffset in registry: {0} minutes.")
    $global:resources.Add("EventID-121", "No Office identities found for this user.")
    $global:resources.Add("EventID-122", "OCCT is disabled by registry key.")
    $global:resources.Add("EventID-123", "ForceMode enabled, skip throttling sleep.")
    $global:resources.Add("EventID-124", "BF FederationProvider found in Office identity root hive.")
    $global:resources.Add("EventID-125", "FederationProvider found in Office identity root hive: {0}")
    $global:resources.Add("EventID-126", "FederationProvider in Office identity root hive does not point to BF.")
    $global:resources.Add("EventID-127", "FederationProvider in Office identity root hive was removed successfully.")
    $global:resources.Add("EventID-128", "Office app automatically reopened: {0}")
    $global:resources.Add("EventID-129", "Check for tenant cutover against Office HRD endpoint.")
    $global:resources.Add("EventID-130", "Check for tenant cutover against Azure AD OIDC endpoint.")
    $global:resources.Add("EventID-131", "Office identity hive removed.")
    $global:resources.Add("EventID-132", "Device is AAD joined.")
    $global:resources.Add("EventID-133", "{0} OneDrive for Business accounts found.")
    $global:resources.Add("EventID-134", "Set NextEmailHRDUpdate for OneDrive account {0}.")
    $global:resources.Add("EventID-135", "Updated NextEmailHRDUpdate for {0} of {1} OneDrive accounts. Restart OneDrive client.")
    $global:resources.Add("EventID-136", "{0} Outlook AutoDiscover cache files found.")
    $global:resources.Add("EventID-137", "{0} Outlook AutoDiscover cache files removed.")
    $global:resources.Add("EventID-138", "Signed out from {0} BF WAM accounts.")
    $global:resources.Add("EventID-139", "Cleared {0} BF OLSFederationProviders.")

    # Error messages
    $global:resources.Add("EventID-201", "User-SID could not be retrieved.")
    $global:resources.Add("EventID-202", "No internet connectivity present.")
    $global:resources.Add("EventID-203", "Existing Process already running.")
    $global:resources.Add("EventID-204", "Function only allows one or two parameters.")
    $global:resources.Add("EventID-205", "Reset of FederationProvider was not successful.")
    $global:resources.Add("EventID-206", "Error removing BF FederationProvider.")
    $global:resources.Add("EventID-207", "Failed to update LastRunTimeSuccess and status in registry.")
    $global:resources.Add("EventID-208", "Failed to create LastRunTimeSuccess and status in registry.")
    $global:resources.Add("EventID-209", "Error message: {0}")
    $global:resources.Add("EventID-210", "Log file could not be published.")
    $global:resources.Add("EventID-211", "Error closing app: {0}")
    $global:resources.Add("EventID-212", "Countdown expired, unable to set DialogResult.")
    $global:resources.Add("EventID-213", "Continue button clicked, unable to set DialogResult.")
    $global:resources.Add("EventID-214", "Dialog closed, unable to stop countdown.")
    $global:resources.Add("EventID-215", "Error checking process {0}")
    $global:resources.Add("EventID-216", "Error checking FederationProvider.")
    $global:resources.Add("EventID-217", "OIDC lookup failed. Error Message: {0}{1}")
    $global:resources.Add("EventID-218", "Dialog: XmlNodeReader could not be initialized.")
    $global:resources.Add("EventID-219", "Dialog: XamlReader could not be initialized.")
    $global:resources.Add("EventID-220", "Dialog: Timer could not be initialized.")
    $global:resources.Add("EventID-221", "Error checking root FederationProvider.")
    $global:resources.Add("EventID-222", "Error removing BF FederationProvider in Office identity root hive.")
    $global:resources.Add("EventID-223", "Removal of FederationProvider in Office identity root hive was not successful.")
    $global:resources.Add("EventID-224", "Office app could not be started: {0}")
    $global:resources.Add("EventID-225", "Setting status registry key failed.")
    $global:resources.Add("EventID-226", "Office identity registry hive not found.")
    $global:resources.Add("EventID-227", "Office identity registry hive could not be removed.")
    $global:resources.Add("EventID-228", "Error getting AAD Join status.")
    $global:resources.Add("EventID-229", "Failed to set NextEmailHRDUpdate for OneDrive account {0}.")
    $global:resources.Add("EventID-230", "Failed to get Outlook AutoDiscover cache files.")
    $global:resources.Add("EventID-231", "Failed to remove Outlook AutoDiscover cache file: {0}.")
    $global:resources.Add("EventID-232", "Failed to remove BF WAM accounts.")
    $global:resources.Add("EventID-233", "Failed to clear OLSFederationProvider.")
    $global:resources.Add("EventID-234", "Some Outlook AutoDiscover Cache files could not be removed.")

    $global:resources.Add("LogFilePath", "C:\Temp\OCCT\OCCT_{0}.xml")
}