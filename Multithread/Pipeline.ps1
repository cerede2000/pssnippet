

$runspace1 = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
$runspace1.Open()
$pipeline1 = $runspace1.CreatePipeline()
$pipeline1.Commands.AddScript({
	Param([String[]]$Values)

	$i = 1
	foreach ($Value in $Values) {
		"$i :: $($Value.ToUpper())"
		$i++
		Start-sleep -s 2
	}
})
$pipeline1.Commands[0].Parameters.Add('Values', @('toto', 'tata', 'tutu'))

$evt = Register-ObjectEvent $pipeline1 -EventName StateChanged -Action {
	Write-Host $sender.PipelineState
	Write-Host 'toto'
Write-Host $Event | gm
Write-Host $EventSubscriber | gm
Write-Host $Sender | gm
Write-Host $EventArgs | gm
Write-Host $Args | gm

Write-Host $EventArgs.PipelineStateInfo.State
if($EventArgs.PipelineStateInfo.State -eq 'Completed') {
	Write-Host $Sender.Output.ReadToEnd()
}

}

$pipeline1.InvokeAsync()

$pipeline1.PipelineStateInfo.State

$pipeline1.Output.ReadToEnd()

