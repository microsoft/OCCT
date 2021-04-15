# Office Client Cutover Tool (OCCT)

OCCT is designed for users migrating from Microsoft Cloud Deutschland (MCD) to the global Office 365 (MCI). The migration process is described here: https://docs.microsoft.com/en-us/microsoft-365/enterprise/ms-cloud-germany-transition?view=o365-worldwide

After phase 9 of the migration happened, Office Pro Plus client may stop working. To solve this issue, users have to sign out from all Office applications, restart Office and sign in again. 
OCCT was developed to automate this client-side migration by running the script automatically (e.g. as scheduled task) without need to manually sign out/sign in.

First OCCT checks if phase 9 is done for the tenant. If true, it checks for running Office apps and asks the user to close the apps. After 10 minutes of waiting, OCCT automatically closes all running Office apps. As soon as all Office apps are closed, OCCT updates the client configuration to work with MCI.


## Prerequisites
- Windows 10
- PowerShell 5.0 or higher
- Office Pro Plus

## Installation
1. Download this repository
2. Deploy OCCT to client computers: Copy all files and folder (e.g. to %ProgramFiles%\WindowsPowerShell\Modules\OCCT)
3. Create a scheduled task (e.g. by GPO) to run OCCT regularly (once per hour).
4. Import OCCT PowerShell module
5. Run Start-OCCT -Tenant <tenantname> command with the additional parameters you need (see following section)


## Parameters
| Parameter | Required | Default | Description |
| :------------- |:-------------| :-----| :-----|
| Tenant | Yes | Empty | Name of your tenant. If your tenant domain is contoso.onmicrosoft.de, enter just contoso
| Force | No | False | If true, OCCT runs immediately without throttling protection. Only for testing.
| ResetRootFedProvider | No | True | if true, FederationProvider in the Office identity root hive will also be removed.
| ReopenOfficeApps | No | False | If true, Office Apps closed by OCCT will be re-opened automatically afterwards.
| OfficeHRDLookup | No | False | If true, use alternative way to detect tenant cutover.
| RemoveOfficeIdentityHive | No | False | If true, the complete Office identity hive will be removed. Caution, you might loose some custom configurations.
| ClearOlkAutodiscoverCache | No | True | If true, Outlook AutoDiscovercache will be cleared to remove outdated AutoDiscover information.
| UpdateODBClient | No | True | If true, connection settings of OneDrive for Business client will be updated.
| RemoveBFWamAccount | No | True | If true, accounts from MCD will be removed from WAM.


## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

## Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
