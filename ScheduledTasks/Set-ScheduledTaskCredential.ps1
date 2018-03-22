# Set-ScheduledTaskCredential.ps1
# Written by Bill Stewart (bstewart@iname.com)
# http://www.itprotoday.com/management-mobility/updating-scheduled-tasks-credentials

#requires -version 2

<#
.SYNOPSIS
Sets the credentials for one or more scheduled tasks on a computer.

.DESCRIPTION
Sets the credentials for one or more scheduled tasks on a computer.

.PARAMETER TaskName
One or more scheduled task names. Wildcard values are not accepted. This parameter accepts pipeline input.

.PARAMETER TaskCredential
The credentials for the scheduled task. If you don't specify this parameter, you will be prompted for credentials.

.PARAMETER ComputerName
The computer name where the scheduled task(s) reside.

.PARAMETER ConnectionCredential
The credentials to use when connecting to the computer.

.EXAMPLE
PS C:\>Set-ScheduledTaskCredential "My Scheduled Task"
This command will prompt for credentials and configure the specified task using those credentials.

.EXAMPLE
PS C:\>Set-ScheduledTaskCredential "Task 1","Task 2" -ComputerName server1
This command will prompt for credentials and configure the named scheduled tasks on the computer server1.

.EXAMPLE
PS C:\>Set-ScheduledTaskCredential "Task 1","Task 2" -ComputerName server1
This command will prompt for credentials and configure the named scheduled tasks on the computer server1.

.EXAMPLE
PS C:\>Get-Content TaskNames.txt | Set-ScheduledTaskCredential -ConnectionCredential (Get-Credential)
This command will set scheduled task credentials for all tasks named in the file TaskNames.txt. There will be two credential prompts. The first prompt is to specify credentials to connect to the Task Scheduler service, and the second prompt is to specify credentials to use for the scheduled tasks.
#>

[CmdletBinding(SupportsShouldProcess=$TRUE)]
param(
  [parameter(Mandatory=$TRUE,ValueFromPipeline=$TRUE)]
    [String[]] $TaskName,
    [System.Management.Automation.PSCredential] $TaskCredential,
    [String] $ComputerName=$ENV:COMPUTERNAME,
    [System.Management.Automation.PSCredential] $ConnectionCredential
)

begin {
  $PIPELINEINPUT = (-not $PSBOUNDPARAMETERS.ContainsKey("TaskName")) -and (-not $TaskName)
  $TASK_LOGON_PASSWORD = 1
  $TASK_LOGON_S4U = 2
  $TASK_UPDATE = 4
  $MIN_SCHEDULER_VERSION = 0x00010002

  # Try to create the TaskService object on the local computer; throw an error on failure
  try {
    $TaskService = new-object -comobject "Schedule.Service"
  }
  catch [System.Management.Automation.PSArgumentException] {
    throw $_
  }

  # Assume $NULL for the schedule service connection parameters unless -ConnectionCredential used
  $userName = $domainName = $connectPwd = $NULL
  if ($ConnectionCredential) {
    # Get user name, domain name, and plain-text copy of password from PSCredential object
    $userName = $ConnectionCredential.UserName.Split("\")[1]
    $domainName = $ConnectionCredential.UserName.Split("\")[0]
    $connectPwd = $ConnectionCredential.GetNetworkCredential().Password
  }
  try {
    $TaskService.Connect($ComputerName, $userName, $domainName, $connectPwd)
  }
  catch [System.Management.Automation.MethodInvocationException] {
    write-error "Error connecting to '$ComputerName' - '$_'"
    exit
  }

  # Returns a 32-bit unsigned value as a version number (x.y, where x is the
  # most-significant 16 bits and y is the least-significant 16 bits).
  function convertto-versionstr([UInt32] $version) {
    $major = [Math]::Truncate($version / [Math]::Pow(2, 0x10)) -band 0xFFFF
    $minor = $version -band 0xFFFF
    "$($major).$($minor)"
  }

  if ($TaskService.HighestVersion -lt $MIN_SCHEDULER_VERSION) {
    write-error ("Schedule service on '$ComputerName' is version $($TaskService.HighestVersion) " +
      "($(convertto-versionstr($TaskService.HighestVersion))). The Schedule service must " +
      "be version $MIN_SCHEDULER_VERSION ($(convertto-versionstr $MIN_SCHEDULER_VERSION)) " +
      "or higher.")
    exit
  }

  # This prevents a scoping problem--if the $TaskCredential variable
  # doesn't exist, it won't get created in the correct scope--create
  # new variable as a workaround
  $NewTaskCredential = $TaskCredential
  if (-not $NewTaskCredential) {
    $NewTaskCredential = $HOST.UI.PromptForCredential("Task Credentials",
      "Please specify credentials for the scheduled task.", "", "")
    if (-not $NewTaskCredential) {
      write-error "You must specify credentials."
      exit
    }
  }

  function set-scheduledtaskcredential2($taskName) {
    $rootFolder = $TaskService.GetFolder("\")
    try {
      $taskDefinition = $rootFolder.GetTask($taskName).Definition
    }
    catch [System.Management.Automation.MethodInvocationException] {
      write-error "Scheduled task '$taskName' not found on '$computerName'."
      return
    }
    $logonType = $taskDefinition.Principal.LogonType
    # No need to set credentials for tasks that don't have stored credentials.
    if (-not (($logonType -eq $TASK_LOGON_PASSWORD) -or ($logonType -eq $TASK_LOGON_S4U))) {
      write-error "Scheduled task '$taskName' on '$ComputerName' doesn't have stored credentials."
      return
    }
    if (-not $PSCMDLET.ShouldProcess("Task '$taskName' on computer '$ComputerName'",
      "Set scheduled task credentials")) { return }
    try {
      [Void] $rootFolder.RegisterTaskDefinition($taskName, $taskDefinition, $TASK_UPDATE,
        $NewTaskCredential.UserName, $NewTaskCredential.GetNetworkCredential().Password, $logonType)
    }
    catch [System.Management.Automation.MethodInvocationException] {
      write-error "Error updating scheduled task '$taskName' on '$computerName' - '$_'"
    }
  }
}

process {
  if ($PIPELINEINPUT) {
    set-scheduledtaskcredential2 $_
  }
  else {
    $TaskName | foreach-object {
      set-scheduledtaskcredential2 $_
    }
  }
}
