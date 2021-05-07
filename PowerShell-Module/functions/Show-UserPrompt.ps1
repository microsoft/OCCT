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

function Show-UserPrompt {
    Param(
        [parameter(Mandatory = $true)]
        [PSObject]
        $runningApps
    )
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
    Add-Type -AssemblyName 'PresentationFramework'
    
    # get display language
    try {
        $lang = $PSUICulture
    } catch {
        # use en-US as fallback if user language cannot be retrieved
        $lang = "en-US"
    }

    if ($lang -eq "de-DE") {
        [xml]$WindowXaml = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="OCCT - Office Client Cutover Tool" 
        Height="450" 
        Width="800" 
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="SingleBorderWindow" 
        Topmost="True">
    <Grid RenderTransformOrigin="0.496,0.347">
        <Border BorderBrush="Red" BorderThickness="3" CornerRadius="0" >
        </Border>
        <Button x:Name="_continue" Content="Erneut prüfen" HorizontalAlignment="Left" Margin="610,320,0,0" VerticalAlignment="Top" Height="47" Width="148" />
        <TextBlock HorizontalAlignment="Left" Margin="22,17,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="47" Width="736" FontSize="14" FontWeight="Bold" Foreground="Black"><Run Text="Damit die Office Applikationen weiter fehlerfrei funktionieren, muss eine dringende Anpassung der Office Konfiguration vorgenommen werden. Führen Sie dazu bitte sofort folgende Schritte aus:"/><Run/></TextBlock>
        <ListBox x:Name="applist" Margin="41,198,539,36" BorderBrush="White" IsEnabled="False" Background="Black"/>
        <TextBlock HorizontalAlignment="Left" Margin="35,64,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="89" Width="745" FontWeight="Normal" Foreground="Black"><Run Text=" 1. Speichern Sie alle in den Office Applikationen (z.B. Word, Excel, Outlook, PowerPoint usw.) geöffneten Dateien lokal (z.B. auf dem Desktop, nicht in OneDrive oder SharePoint Online).&#10;
2. Schließen Sie im Anschluss alle unten aufgeführten Office Applikationen.&#10;
3. Warten Sie 10 Sekunden bevor Sie die Office Apps erneut öffnen.&#10;
4. Öffnen Sie die in Schritt 1 lokal gespeicherten Dateien und speichern Sie diese am gewünschten Ort."/></TextBlock>
        <TextBlock x:Name="runningApps" HorizontalAlignment="Left" Margin="36,174,0,0" Text="Momentan geöffnete Office Apps:" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" FontSize="14"/>
        <TextBlock x:Name="countdown" HorizontalAlignment="Right" Margin="0,379,42,0" Text="Die momentan geöffneten Office Apps werden automatisch beendet in 10:00 Minuten." TextWrapping="Wrap" VerticalAlignment="Top" Width="401" Height="19" TextAlignment="Right"/>
    </Grid>
</Window>
'@
    } else {
        [xml]$WindowXaml = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="OCCT - Office Client Cutover Tool" 
        Height="450" 
        Width="800" 
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="SingleBorderWindow" 
        Topmost="True">
    <Grid RenderTransformOrigin="0.496,0.347">
        <Border BorderBrush="Red" BorderThickness="3" CornerRadius="0" >
        </Border>
        <Button x:Name="_continue" Content="Check again" HorizontalAlignment="Left" Margin="610,320,0,0" VerticalAlignment="Top" Height="47" Width="148" />
        <TextBlock HorizontalAlignment="Left" Margin="22,17,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="47" Width="736" FontSize="14" FontWeight="Bold" Foreground="Black"><Run Text="In order to ensure the proper function of the Office applications, the Office configuration must be adjusted urgently. To do this, please carry out the following steps immediately"/><Run/></TextBlock>
        <ListBox x:Name="applist" Margin="41,198,539,36" BorderBrush="White" IsEnabled="False" Background="Black"/>
        <TextBlock HorizontalAlignment="Left" Margin="35,64,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="89" Width="745" FontWeight="Normal" Foreground="Black"><Run Text=" 1. Save all files opened in the Office applications (e.g. Word, Excel, Outlook, PowerPoint etc.) locally (e.g. on your Desktip, not in OneDrive or SharePoint Online).&#10;
2. Then close all Office applications listed below.&#10;
3. Wait 10 seconds before starting the Office apps again.&#10;
4. Open the files saved locally in step 1 again and save it to the desired location."/></TextBlock>
        <TextBlock x:Name="runningApps" HorizontalAlignment="Left" Margin="36,174,0,0" Text="Office apps currently running:" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Bold" FontSize="14"/>
        <TextBlock x:Name="countdown" HorizontalAlignment="Right" Margin="0,379,42,0" Text="Currently opended Office apps will be closed automatically in 10:00 minutes." TextWrapping="Wrap" VerticalAlignment="Top" Width="401" Height="19" TextAlignment="Right"/>
    </Grid>
</Window>
'@
    }
    try {
        $reader = (New-Object System.Xml.XmlNodeReader $WindowXaml)
    } catch {
        Publish-EventLog -EventId 218 -EntryType Error -Message ($global:resources["EventID-218"])
    }
    try {
        $Window = [Windows.Markup.XamlReader]::Load( $reader )
        $Window.Topmost = $true
    } catch {
        Publish-EventLog -EventId 219 -EntryType Error -Message ($global:resources["EventID-219"])
    }
    
    # button action
    $button = $Window.FindName("_continue")
    $button.Add_Click( {
            $Timer.Stop(); 
            $Window.Close();
            $Script:Timer = $null
            $Script:CountDown = 600
        })

    # add running apps to app list
    $runningApps = $runningApps | Sort-Object -Property DisplayName
    $listbox = $Window.FindName('applist')
    foreach ($app in $runningApps) {
        $listbox.Items.Add($app.DisplayName) | Out-Null
    }

    # get countdown label to update remaining seconds
    $countdownLabel = $Window.FindName('countdown')

    # create Timer
    if (!$Script:Timer) {
        try {
            $Script:Timer = New-Object System.Windows.Forms.Timer
            $Script:Timer.Interval = 1000
        } catch {
            Publish-EventLog -EventId 220 -EntryType Error -Message ($global:resources["EventID-220"])
        }
    }

    function Timer_Tick() {
        $Script:time = ""

        $ts = [timespan]::fromseconds($Script:CountDown)
        $Script:time = "{0:mm:ss}" -f ([datetime]$ts.Ticks)

        if ($lang -eq "de-DE") {
            $countdownLabel.Text = "Die momentan geöffneten Office Apps werden automatisch beendet in {0} Minuten." -f $Script:time
        } else {
            $countdownLabel.Text = "Currently opended Office apps will be closed automatically in {0} minutes." -f $Script:time
        }
        --$Script:CountDown
        If ($Script:CountDown -eq 0) {
            $Timer.Stop(); 
            $Window.Close();
            $Script:Timer = $null
            $Script:CountDown = 600
            $Script:killedApplications = KillOfficeApps
        }
    }
    $Script:CountDown = 600
    $Timer.Add_Tick( { Timer_Tick })
    $Timer.Start()	

    # Add closing event
    $Window.Add_Closing( { 
            $Timer.Stop()
            $Timer.Dispose()
            $Script:Timer = $null
            $Script:CountDown = 600
        })
    # Show Window
    $Window.ShowDialog() | Out-Null

    return $Script:killedApplications
}