function Get-JmaScheduledTask
{
    <#
        .SYNOPSIS
        Gets the task definition object of a scheduled task that is registered on the local or a remote computer.

        .DESCRIPTION
        The Get-JmaScheduledTask Cmdlet gets the task definition object of a scheduled task that is registered on the local or a remote computer utilizing the windows executable schtasks.exe.
        Mimicing the Microsoft TaskScheduler module you can filter on the taskname with the asterix char (*).

        .PARAMETER ComputerName
		An array of one or more computernames.

        .PARAMETER TaskName
        Specify a path for scheduled tasks in the Task Scheduler namespace.
        Mimicing the (handy but unfortunately only against Windows Server 2012+) Microsoft TaskScheduler module, you can filter the taskname with the asterix char (*)

        .EXAMPLE
        Get all tasks from the local computer .  
        Get-JmaScheduledTask

        .EXAMPLE
        Get specific task from remote computers Server1 and server2. 
        Get-JmaScheduledTask -TaskName '\MyFolder\MyTask' -ComputerName Server1, Server2 -Verbose

        .EXAMPLE
        Search for task(s) from remote computers Server1 and server2 (wildcard).
        Get-JmaScheduledTask -TaskName '\MyFolder\*' -ComputerName Server1, Server2

        .NOTES
        Created by Anders Præstegaard (@aPowerShell).
        19-05-2016	Version 1.0
    #>
	[CmdletBinding()]
	param
	(
		[Parameter(ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[Alias('PSComputerName')]
		[string[]]
		$ComputerName = 'localhost',
		[Parameter(Position = 1)]
		[ValidateNotNullOrEmpty()]
		[string]
		$TaskName
	)
	
	#requires -RunAsAdministrator
	
	begin
	{
		$activity = 'Retrieving all scheduled tasks'
		$codeSearchString = ''
		if ($PSBoundParameters.ContainsKey('TaskName') -and $TaskName -notmatch '\*')
		{
			$activity = "Retrieving scheduled task '$TaskName'"
			$codeTaskNameString = " /tn '$TaskName'"
		}
		else
		{
			$codeSearchString = ''
			if ($TaskName -match '\*')
			{
				$activity = "Searching for scheduled task '$TaskName'"
				$codeSearchString = '| Where-Object -FilterScript { $_.TaskName -like $TaskName }'
			}
			else
			{
				$activity = 'Getting all scheduled tasks'
			}
		}
	}
	process
	{
		
		$ComputerName |
		ForEach-Object -Process {
			$computer = $_
			
			$codeString = 'schtasks.exe /query /s $computer /fo csv /v 2> $null' + $codeTaskNameString + '| ConvertFrom-Csv' + $codeSearchString
			$scriptBlock = [ScriptBlock]::Create($codeString)
			
			Write-Verbose -Message "[$computer] $activity"
			Write-Progress -Activity $activity -CurrentOperation $computer
			
			& $scriptBlock
		}
	}
	end
	{
		Write-Progress -Activity 'Getting scheduled task' -Completed
	}
}