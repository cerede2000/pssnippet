
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsFormsIntegration

[XML]$xml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="Counter" Height="119" Width="351.5" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" >
        <Label Name="Label" Content="0" HorizontalAlignment="Left" Margin="16.666,9.333,0,0" VerticalAlignment="Top" FontSize="18"/>
        <Button Name="Button" Content="Start" HorizontalAlignment="Center" VerticalAlignment="Top" Width="75" Margin="123.25,63,123.25,0"/>
        <Button Name="Button2" Content="Popup" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="0,0,0,0"/>
    </Grid>
</Window>
"@

$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Host = $Host
$syncHash.Jobs = @()

$Reader=(New-Object System.Xml.XmlNodeReader $xml)
$syncHash.Window = [Windows.Markup.XamlReader]::Load($Reader)

$syncHash.Label = $syncHash.Window.FindName('Label')
$syncHash.Button = $syncHash.Window.FindName('Button')
$syncHash.Button2 = $syncHash.Window.FindName('Button2')

$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$InitialSessionState.Variables.Add((New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'syncHash', $syncHash, $Null))

$Runspace = [runspacefactory]::CreateRunspacePool(1, $env:NUMBER_OF_PROCESSORS, $InitialSessionState, $host)
$Runspace.ApartmentState = 'STA'
$Runspace.ThreadOptions = 'ReuseThread'
$Runspace.Open()

$syncHash.Button2.Add_Click({

    $powershell = [Powershell]::Create().AddScript({
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName WindowsFormsIntegration
        
        1..10 | %{
          $value = $_
          $syncHash.Host.Runspace.Events.GenerateEvent('RunspaceReportProgess', $null, 'Label', @{Value = $value.ToString() + (get-date).ToString('hh:MM:ss')})
          # $syncHash.Label.Dispatcher.Invoke([action]{ 
              # $syncHash.Label.Content = $value.ToString() + (get-date).ToString('hh:MM:ss')
          # }, "Normal")
            Start-sleep -s 2
        }
        
        [System.Windows.Forms.MessageBox]::Show("PowerShell et WPF")
    })

    $powershell.RunspacePool = $Runspace

    $syncHash.Jobs += [PSCustomObject]@{ Powershell = $PowerShell; Handle = $PowerShell.BeginInvoke() }

    # do {
        # Start-Sleep -Milliseconds 50
        # [System.Windows.Forms.Application]::DoEvents()
    # }while(!$Object.IsCompleted)
})

$syncHash.Button.Add_Click({

    $powershell = [Powershell]::Create().AddScript({
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName WindowsFormsIntegration
        
        [System.Windows.Forms.MessageBox]::Show("Thread popup")
        $syncHash.Host.Runspace.Events.GenerateEvent('RunspaceEvent', 'Sender', 'Args', 'test event')
        
        $connection = Test-Connection -ComputerName google.com -Count 5
        $syncHash.output = [math]::Round(($connection.ResponseTime | Measure-Object -Average).Average)
    })

    $powershell.RunspacePool = $Runspace

    $syncHash.Jobs += [PSCustomObject]@{ Powershell = $PowerShell; Handle = $PowerShell.BeginInvoke() }

    # $powershell.EndInvoke($Object)
    # $powershell.Dispose()

    #$label.Content = $syncHash.output
	#$syncHash.Label.Content = 10
})

$syncHash.timer = New-Object System.Windows.Forms.Timer
$syncHash.timer.Interval = 500
$syncHash.timer.add_tick({
  Write-Host "Tick :: $($syncHash.Jobs.Count)"
  $syncHash.Jobs | ?{ $_.Handle.IsCompleted } | %{
      Write-Host 'Job finish'
      $_.Powershell.EndInvoke($_.Handle)
  }
  $syncHash.Jobs = [Array]($syncHash.Jobs | ?{ -Not $_.Handle.IsCompleted })
})

$syncHash.Window.Add_ContentRendered({
    $syncHash.timer.Start()
})

$syncHash.Window.Add_Closed({
    $syncHash.timer.Stop()
    $syncHash.timer.Dispose()
    
    if($syncHash.Jobs -ne $null) { $syncHash.Jobs | ?{ -Not $_.Handle.IsCompleted } | %{ $_.Powershell.Stop() } }
    $Runspace.Close()
    $Runspace.Dispose()
    
    [System.Windows.Forms.Application]::Exit($null)
})

Register-EngineEvent -SourceIdentifier 'RunspaceEvent' -Action {
  Write-Host 'RunspaceEvent' -fore green
  Write-Host $Event.Sender
  Write-Host $Event.SourceArgs
  Write-Host $Event.MessageData
  Write-Host ($Event | gm)
  $syncHash.Label.Content = 100
}

Register-EngineEvent -SourceIdentifier 'RunspaceReportProgess' -Action {
  $syncHash."$($Event.SourceArgs)".Content = $Event.MessageData.Value
}

[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($syncHash.Window)
$syncHash.Window.Show()
$syncHash.Window.Activate()
$appContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($appContext)

# $syncHash.Window.ShowDialog() | Out-Null

# [System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($syncHash.Window)
# $syncHash.Window.Show()
# $syncHash.Window.Activate()
# $appContext = New-Object System.Windows.Forms.ApplicationContext 
# [void][System.Windows.Forms.Application]::Run($appContext)

# $Window.ShowDialog() | Out-Null
