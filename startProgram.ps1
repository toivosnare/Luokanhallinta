param(
    [String]$executable,
    [String]$argument,
    [String]$workingDirectory = "C:\",
    [Switch]$runElevated
)

if($argument)
{
    $action = New-ScheduledTaskAction -Execute $executable -WorkingDirectory $workingDirectory -Argument $argument # Create scheduled task action to start the executable 
}
else
{
    $action = New-ScheduledTaskAction -Execute $executable -WorkingDirectory $workingDirectory
}
$user = Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName # Get the user that is logged on the remote computer
if($runElevated)
{
    $principal = New-ScheduledTaskPrincipal -UserId $user -RunLevel Highest # Create scheduled task principal (the user which the executable is to be run as)
}
else
{
    $principal = New-ScheduledTaskPrincipal -UserId $user
}
$task = New-ScheduledTask -Action $action -Principal $principal # Create new scheduled task with the action and principal
$taskname = 'Testi'
try 
{
    $registeredTask = Get-ScheduledTask $taskname -ErrorAction SilentlyContinue # Check if there is already a scheduled task with the same name
} 
catch 
{
    $registeredTask = $null
}
if ($registeredTask)
{
    Unregister-ScheduledTask -InputObject $registeredTask -Confirm:$false # If so remove it
}
$registeredTask = Register-ScheduledTask $taskname -InputObject $task # Register the newly created scheduled task
Start-ScheduledTask -InputObject $registeredTask # Start the scheduled task
Unregister-ScheduledTask -InputObject $registeredTask -Confirm:$false # Remove the scheduled task