# Office Client Cutover Tool (OCCT)
OCCT is designed for customers migrating from Microsoft Cloud Deutschland to Office 365 services in the new German datacenter regions. OCCT supports the migration of Office Apps on Windows devices after migration phase 9 completed. A detailed description of the migration process is available in [Microsoft Docs](https://docs.microsoft.com/en-us/microsoft-365/enterprise/ms-cloud-germany-transition?view=o365-worldwide).

After phase 9 of the migration process completed, Office 365 Apps for Windows (Word, Excel, PowerPoint, Outlook) may stop working. To solve this issue, users must sign out from all Office applications, close them, open any Office App and sign in again.
OCCT will automate the client-side steps without having to to perform manually sign out/sign in operations. 

The tool is designed to be executed by an affected user once after migration phase 9 has been completed or a scheduled task is added to run the tool periodically. The last option ensures the best user experience in managed environments. Running OCCT does not require local administrator privileges.

While executing OCCT, the tool can examine the migration status for the tenant. If migration phase 9 has been completed, the client must be updated. To successfully run OCCT, all Office Apps must be closed. If the user is running Office Apps, a message prompt will appear and request the user to close all running Office Apps. The message displays which Office Apps are running and need to be closed.

If the user does not respond to the message prompt, OCCT is automatically closing all running Office apps after 10 minutes. Once all Office Apps are closed, OCCT updates the client configuration to work with the new German datacenter region.

## Prerequisites
- Windows 10
- PowerShell 5.0 or higher
- Office Apps for Windows

## Usage
OCCT is built on PowerShell can be executed automated or manually and has been created in two versions - a simple PowerShell script and a PowerShell module. The recommended version for must customers is the script version.

Download the PowerShell script [(`OCCT.ps1`)](https://raw.githubusercontent.com/microsoft/OCCT/main/OCCT.ps1) from this repository.
In case the tenant already passed migration phase 9 and clients encounter issues with Office apps, the user should run `OCCT.ps1` on the client to reconfigure the client to connect to the new German datacenter regions.

Windows 10 does not allow running PowerShell scripts by default. In order to run script, the local execution policy must be modified.
To temporary allow executing PowerShell scripts, open PowerShell and run the command:

`Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser`

Run `OCCT.ps1`, then reset the original state of the execution policy by executing the following command:

`Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope CurrentUser`

IT departments or administrators who are responsible for multiple computers should consider deploying the script file to all Windows clients with Office Apps connected to Microsoft Cloud Deutschland prior to the end of phase 9 and prepare the users.

Either the users execute the script according to the provided instructions after the migration phase 9 is completed, or a scheduled task is created to execute OCCT regularly (e.g. on startup and every 10 minutes).

If the tenant name is provided to OCCT as a parameter (e.g. `.\OCCT.ps1 -tenant MyTenant`), OCCT will check if the migration phase 9 has been completed for the tenant prior to reconfiguring the client. Without this parameter, OCCT will start updating the client configuration without any pre-checks. It is recommended to do the pre-check if OCCT should be executed periodically.

### Summary
1. Download OCCT.ps1 from this repository.
2. Copy this file to all Windows clients with Office Apps connected to Microsoft Cloud Deutschland.
3. Script execution
    - Option 1: Run the script once when migration phase 9 has been completed and your Office Apps are not connected anymore.
    - Option 2: Create a scheduled task (e.g. by GPO) to run OCCT regularly. 

    Ensure the local execution policy of the client allows PowerShell script executions. 

## Advanced Usage
### PowerShell module
This is a PowerShell module with all features provided by OCCT to be used to build your own custom version.
1. Download this repository.
2. Deploy the OCCT module to all client computers: Copy all files and folders to an common PowerShell module folder (e.g. to `%ProgramFiles%\WindowsPowerShell\Modules\OCCT`).
3. Create a scheduled task (e.g. by GPO or manually) to run OCCT regularly.
4. The scheduled task must run PowerShell and the "Start-OCCT -Tenant <tenantname>" command, extended with all additional parameters you want to leverage (see following section).

## Parameters
The parameters are valid for both versions of OCCT, the simple script file and the PowerShell module.
| Parameter | Required | Default | Description |
| :------------- |:-------------| :-----| :-----|
| Tenant | No | Empty | Name of your tenant. If your tenant domain is contoso.onmicrosoft.de, enter contoso only. Adding ".onmicrosoft.de" is not supported. If provided, OCCT will verify if the tenant cutover (phase 9 completed) is already done before performing any actions.
| Force | No | False | If true, OCCT runs immediately without throttling protection. Only for testing.
| ResetRootFedProvider | No | True | If true, FederationProvider in the Office identity root hive will also be removed for the current user.
| ReopenOfficeApps | No | False | If true, all Office Apps which have been closed by OCCT will be restarted automatically after updating the client configuration.
| OfficeHRDLookup | No | False | If true, OCCT will use an alternative way to detect the tenant migration status.
| RemoveOfficeIdentityHive | No | False | If true, the complete Office identity hive will be removed. Caution, you might lose some custom configurations.
| ClearOlkAutodiscoverCache | No | False | If true, Outlook AutoDiscover cache will be cleared to remove outdated AutoDiscover information.
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