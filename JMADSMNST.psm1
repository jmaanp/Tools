<#
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.119
	 Created on:   	11-05-2016 15:36
	 Created by:   	anp
	 Organization: 	JMA A/S
	 Filename:     	JMADSMNST.psm1
	-------------------------------------------------------------------------
	 Module Name: JMADSMNST
	===========================================================================
#>

function Get-JmaNstInstance
{
<#
	.SYNOPSIS
		This function retreives information from one or more NST-servers

	.DESCRIPTION
		This function retreives the following information:

		-----------------------------------------------------------------------------
		ServerInstance               : DSMNAV_2015_DK
		ServerInstanceServiceAccount : DOMAIN\sa_NSTCLU01
		ServerInstanceVersion        : 8.0.43887.0
		ServerInstanceIsMultiTenant  : True
		ClientServicesPort           : 8100
		ClientServicesCredentialType : Windows
		NstServerInNlbCluster        : True
		NstServer                    : NSTCLU
		CompanyName                  : Andeby A/S
		PSComputerName               : dsmnst02
		State                        : Operational
		DetailedState                :
		Id                           : ank
		DatabaseName                 : ANK
		DatabaseServer               : My-SQL-CLU01
		AlternateId                  : {}
		AllowAppDatabaseWrite        : False
		NasServicesEnabled           : False
		DefaultCompany               :
		DefaultTimeZone              : (UTC+01:00) Brussels, Copenhagen, Madrid, Paris
		-----------------------------------------------------------------------------

	.PARAMETER ComputerName
		Retrives backup results on the specified computers. The default is the local computer.
		Type the NetBIOS name, an IP address, or a fully qualified domain name of one or more computers. To specify the local computer ignore the ComputerName parameter.
		This parameter rely on Windows PowerShell remoting, so your computer has to be configured to run remote commands.

	.PARAMETER Credential
		Specifies a user account that has permission to perform this action. The default is the current user. Type a user name, such as "User01", "Domain01\User01", or User@Contoso.com. Or, enter a PSCredential object, such as an object that is returned by the Get-Credential cmdlet. When you type a user name, you are prompted for a password.

	.PARAMETER NavisionMajorVersion
		Controls the PowerShell module path.
        This parameter is for future usage

	.EXAMPLE
		Get-JmaNstInstance -ComputerName dsmnst02, dsmnst03

	.NOTES
		ANP
		Version 1.0
		14.05.2016
#>

	[CmdletBinding(DefaultParameterSetName='ComputerName')]
	[OutputType([PSCustomObject])]
	param
	(
		[Parameter(ParameterSetName = 'Session')]
		[ValidateNotNullOrEmpty()]
		[System.Management.Automation.Runspaces.PSSession[]]
		${Session},
		[parameter(ParameterSetName='ComputerName')]
		[ValidateNotNullOrEmpty()]
		[string[]]
		${ComputerName} = $env:COMPUTERNAME,
		[parameter(ParameterSetName='ComputerName')]
		[pscredential]
		[System.Management.Automation.CredentialAttribute()]
		${Credential}
	)

	if ($PSBoundParameters.ContainsKey('ComputerName'))
	{
		$PSDefaultParameterValues['Invoke-Command:ComputerName'] = $ComputerName
		if ($PSBoundParameters.ContainsKey('Credential'))
		{
			$PSDefaultParameterValues['Invoke-Command:Credential'] = $Credential
		}
	}
	else
	{
		$PSDefaultParameterValues['Invoke-Command:Session'] = $Session
	}

	$scriptBlock = {
		$navModuleVersion = Get-ChildItem -Path 'C:\Program Files\Microsoft Dynamics NAV\*\Service\NavAdminTool.ps1'

		if ($navModuleVersion -isnot [System.IO.FileInfo])
		{
			if ($navModuleVersion)
			{
				$navModuleVersion = $navModuleVersion | Sort-Object fullname -Descending | Select-Object -First 1
			}
            else
            {
                throw "[$env:COMPUTERNAME] No version of Microsoft Dynamics Nav found"
            }
		}

		try
		{
			$null = Import-Module -Name $navModuleVersion.FullName -DisableNameChecking -Verbose:$false -ErrorAction Stop
			Get-NAVServerInstance |
			Where-Object -Property state -EQ -Value 'Running' |
			Where-Object -FilterScript {
				$_.Version -ge (New-Object -TypeName System.Version -ArgumentList 8, 0)
			} |
			ForEach-Object -Process {
				$serverInstanceConfiguration = Get-NAVServerConfiguration -ServerInstance $_.ServerInstance -ErrorAction SilentlyContinue

				$serverInstanceServiceAccount = $_.ServiceAccount
				$serverInstanceVersion = $_.Version

				$serverInstanceIsMultiTenant = ($serverInstanceConfiguration |
				Where-Object -Property key -EQ -Value 'multitenant' |
				Select-Object -ExpandProperty value) -eq 'true'

				$clientServicesCredentialType = $serverInstanceConfiguration |
				Where-Object -Property key -EQ -Value 'ClientServicesCredentialType' |
				Select-Object -ExpandProperty value

				$clientServicesPort = $serverInstanceConfiguration |
				Where-Object -Property key -EQ -Value 'ClientServicesPort' |
				Select-Object -ExpandProperty value

				try
				{
					#$reg = Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\WLBS\Parameters\Interface\* -ErrorAction Stop
					#$nstServer = $reg.ClusterName
					$nstServer = (Get-NlbCluster -ErrorAction SilentlyContinue).Name
					$nstServerInNlbCluster = $true
				}
				catch
				{
					$nstServer = $env:COMPUTERNAME
					$nstServerInNlbCluster = $false
				}

				Get-NAVTenant -ServerInstance $_.ServerInstance |
				ForEach-Object -Process {
					$tenantInfo = $_ |
					Add-Member -NotePropertyName ServerInstance -NotePropertyValue ($_.ServerInstance -replace 'MicrosoftDynamicsNavServer\$') -Force -PassThru |
					Add-Member -NotePropertyName ServerInstanceServiceAccount -NotePropertyValue $serverInstanceServiceAccount -Force -PassThru |
					Add-Member -NotePropertyName ServerInstanceVersion -NotePropertyValue $serverInstanceVersion -Force -PassThru |
					Add-Member -NotePropertyName ServerInstanceIsMultiTenant -NotePropertyValue $serverInstanceIsMultiTenant -Force -PassThru |
					Add-Member -NotePropertyName ClientServicesPort -NotePropertyValue $clientServicesPort -Force -PassThru |
					Add-Member -NotePropertyName ClientServicesCredentialType -NotePropertyValue $clientServicesCredentialType -Force -PassThru |
					Add-Member -NotePropertyName NstServerInNlbCluster -NotePropertyValue $nstServerInNlbCluster -Force -PassThru |
					Add-Member -NotePropertyName NstServer -NotePropertyValue $nstServer -Force -PassThru

					Get-NAVCompany -ServerInstance $_.ServerInstance -Tenant $_.Id |
					ForEach-Object -Process {
						$tenantInfo |
						Add-Member -NotePropertyName CompanyName -NotePropertyValue $_.CompanyName -Force -PassThru
					}
				}
			}
		}
		catch
		{
			throw "[$env:COMPUTERNAME] $($_.Exception.Message)"
		}
	} # END Scriptblock

	Write-Progress -Activity "Collecting configuration(s) from $(@($ComputerName).Count) NST-servers"

	Invoke-Command -ScriptBlock $scriptBlock |
	Select-Object -Property * -ExcludeProperty RunspaceId

	Write-Progress -Activity "Collecting configurations  $(@($ComputerName).Count) NST-servere" -Completed
}

function Add-JmaShortcut
{
<#
	.SYNOPSIS
		Dette script bruges til at oprette en genvej til et program.

	.DESCRIPTION
		Dette script bruges til at oprette en genvej til program, hvor du kan definere "Description", "WorkDirectory", "Hotkey", "Icon" (sti til et ikon) og endvidere om genvejen skal aktivere admin mode (UAC), når genvejen aktiveres.

	.PARAMETER Path
		 Den fulde sti til genvejen.
		 Hvis den fulde sti ikke er specificeret med en extension, tilføjes ".lnk" som extension.
		 Hvis mappen, som filen skal placeres i ikke findes oprettes den.

	.PARAMETER TargetPath
		Den fulde sti til et program eller en fil.

	.PARAMETER Arguments
		Argumenter som genvejen kalder programmet eller filen med.

	.PARAMETER Description
		Udfylder beskrivesesfeltet på genvejen.

	.PARAMETER HotKey
		Tastaturgenvej.  Valide kombinationer er fx SHIFT+F7, ALT+CTRL+9.

	.PARAMETER WorkingDirectory
		A description of the WorkingDirectory parameter.

	.PARAMETER WindowStyle
		Normal (1), Maximeret (3), eller Minimeret (7).

	.PARAMETER Icon
		 Den fulde sti til en ikonfil.
		 DLL'er, indeholdende flere ikoner, fordrer specificering af et nummer til ikonet,
		 med mindre det første ikon i den DLL-fil ønskes.

	.PARAMETER Force
		A description of the Force parameter.

	.PARAMETER Admin
		Bruges til at lave en genveje, der prompter for admin credentials (UAC'en).

	.PARAMETER Passthru
		A description of the Passthru parameter.

	.PARAMETER WorkDir
		Working directory of the application.  An invalid directory can be specified, but invoking the application from the
		shortcut could fail.

	.EXAMPLE
		[pscustomobject]@{Path = 'C:\Temp\Shortcut\ADoubee.lnk' ;TargetPath = 'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe'} | Add-Shortcut -Admin
		Bruger pipelinen til at fodre Add-Shortcut med en genvejs-konfiguration

	.EXAMPLE
		Add-Shortcut -Path c:\temp\notepad.lnk -TargetPath c:\windows\notepad.exe
		Genererer en simpel genvej med navnet Notepad til programmet Notepad placeret i mappen c:\Temp

	.EXAMPLE
		Add-Shortcut "$($env:Public)\Desktop\Notepad" c:\windows\notepad.exe -WindowStyle 3 -Admin
		Det samme som exemplet ovenfor nu blot specificeret til at åbne maksimeret og med UAC prompt

	.OUTPUTS
		Output'er et PSCustomObject indeholdende alle indstillingerne for genvejen, der er blevet lavet.

	.NOTES
		ANP
		Version 1.0
		14.05.2016

	.INPUTS
		Tager input via pipeline, hvilket vil sige, at du kan pipe et PSCustomObject, som vist i én af eksemplerne.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[Alias('File', 'shortcut')]
		[string]
		$Path,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 1)]
		[Alias('Target')]
		[string]
		$TargetPath,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 2)]
		[Alias('Args', 'Argument')]
		[string]
		$arguments,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 3)]
		[Alias('Desc')]
		[string]
		$Description,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 4)]
		[string]
		$HotKey,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 5)]
		[Alias('WorkDir', 'Directory')]
		[string]
		$workingDirectory,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 6)]
		[int]
		$WindowStyle,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 7)]
		[string]
		$Icon,
		[Parameter(ValueFromPipelineByPropertyName = $true,
				   Position = 8)]
		[switch]
		$Force,
		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[switch]
		$Admin,
		[switch]
		$Passthru
	)

	Process
	{
		if (-not (Test-Path -Path $TargetPath))
		{
			throw "TargetPath '$TargetPath' does not exist which is a requierement by the WScript.Shell COM object"
		}

		if ($Path -notmatch "^.*(\.lnk)$")
		{
			$Path = "$($Path).lnk"
		}

		[System.IO.FileInfo]$Path = $Path
		Try
		{
			if (-not (Test-Path -Path $Path.Directory))
			{
				[void](mkdir -Path $Path.Directory -ErrorAction Stop)
			}
		}
		Catch
		{
			throw "The shortcut could not be created (unable to create directory '$($Path.DirectoryName)')"
		}

		# Define shortcut Properties
		$wshShell = New-Object -ComObject WScript.Shell
		$shortcut = $wshShell.Createshortcut($Path.FullName)

		$PSBoundParameters.GetEnumerator() |
		Where-Object -FilterScript {
			$_.key -notmatch '^Path$|^Admin$|Force'
		} |
		ForEach-Object -Process {
			$shortcut.$($_.key) = $_.value
		}

		try
		{
			# Create shortcut
			$shortcut.Save()
			# Set shortcut to Run Elevated
			If ($Admin)
			{
				$tempFileName = [IO.Path]::GetRandomFileName()
				$tempFile = [IO.FileInfo][IO.Path]::Combine($Path.Directory, $tempFileName)
				$writer = New-Object -TypeName System.IO.FileStream -ArgumentList $tempFile, ([System.IO.FileMode]::Create)
				$reader = $Path.OpenRead()
				While ($reader.Position -lt $reader.Length)
				{
					$Byte = $reader.ReadByte()
					If ($reader.Position -eq 22)
					{
						$Byte = 34
					}
					$writer.WriteByte($Byte)
				}
				$reader.Close()
				$writer.Close()
				$Path.Delete()
				[void](Rename-Item -Path $tempFile -NewName $Path.Name)
			}

			## Output genvejen
			if ($Passthru)
			{
				$shortcut
			}
		}
		catch
		{
			Write-Warning -Message $Error[0].Exception.Message
			throw "Unable to create $($Path.FullName)"
		}

	}

	end
	{

	}
}

function Add-JmaNstShortcut
{
<#
	.SYNOPSIS
		This script transform Dynamics Nav configuration Instances into shortcut-files consumed by the Microsoft Dynamics Nav Client

	.DESCRIPTION
		This script transform Dynamics Nav configuration Instance into shortcut-files consumed by the Microsoft Dynamics Nav Client.
		In JmaDsmNst module there's a function that creates these custom configuration data objects, so that you can create the input for this function.

	.PARAMETER ServerInstance
		Microsoft Dynamics Nav Server Instance.

	.PARAMETER DataBaseName
		SQL Server Database instance

	.PARAMETER nstServer
		The Window Server hosting the Microsoft Dynamics Nav Server instance

	.PARAMETER CompanyName
		Microsoft Dynamics Nav Companyname

	.PARAMETER clientServicesPort
		Microsoft Dynamics Nav Server instance port for client connections

	.PARAMETER DatabaseServer
		SQL Database name

	.PARAMETER Id
		Microsoft Dynamics Nav Tenant Id

	.PARAMETER serverInstanceIsMultiTenant
		Microsoft Dynamics Nav Server instance configuration setting

	.PARAMETER clientServicesCredentialType
		Microsoft Dynamics Nav Server instance configuration setting

	.PARAMETER Destination
		Shortcut file full Path

	.PARAMETER navTargetPath
		A description of the navTargetPath parameter.

	.NOTES
		ANP
		Version 1.0
		14.05.2016
#>

	[CmdletBinding()]
	[OutputType([System.IO.FileSystemInfo])]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$ServerInstance,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$DataBaseName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$nstServer,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$CompanyName,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$clientServicesPort,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$DatabaseServer,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('TenantId')]
		[string]
		$Id,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[bool]
		$serverInstanceIsMultiTenant,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]
		$clientServicesCredentialType,
		[Parameter(Mandatory = $true)]
		[ValidateScript({
			Test-Path -Path $_
		})]
		[string]
		$Destination,
		[Parameter(Mandatory = $true)]
		[string]
		$navTargetPath
	)

	begin
	{
		[xml]$clientSettings = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="ClientServicesCredentialType" value="NavUserPassword" />
  </appSettings>
</configuration>
"@
	}

	process
	{
		$databaseServerTrimmedName = $DatabaseServer.ToUpper() -replace '/|\\', '_'
		$companyTrimmedName = $CompanyName -replace '/|\\', ''

		Write-Progress -Activity 'Opretter genveje' -Status $databaseServerTrimmedName -CurrentOperation "$ServerInstance - $CompanyName"
		$arguments = "`"DynamicsNAV://$($nstServer.ToUpper()):$($clientServicesPort)/$($ServerInstance.ToUpper())/$($CompanyName)/?Tenant=$($Id)`" -language:da-DK"

		$relPath = "$databaseServerTrimmedName\$($DataBaseName.ToUpper())\$($ServerInstance.ToUpper())"

		$workingDirectory = Join-Path -Path $Destination -ChildPath $relPath

		if ($clientServicesCredentialType -eq 'NavUserPassword')
		{
			$arguments += " -settings:`"$($DataBaseName).config`""
		}

		$shortcutSettings = @{
			Path = "$($workingDirectory)\$($companyTrimmedName) ($($nstServer.ToUpper())).lnk"
			TargetPath = $navTargetPath
			Arguments = $arguments
		}

		Add-JmaShortcut @shortcutSettings

		if ($clientServicesCredentialType -eq 'NavUserPassword' -and (Test-Path -Path $workingDirectory))
		{
			$clientSettings.Save("$workingDirectory\$($DataBaseName.ToUpper()).config")
		}
	}
	end
	{
		Write-Progress -Activity 'Opretter genveje' -Completed
	}
}


Export-ModuleMember Get-JmaNstInstance,
					Add-JmaNstShortcut