Param(
	[String]$JobName,
	[String]$Id
)

####################
# Import Functions #
####################
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Import-Module "$PSScriptRoot\Helpers"

# Get the config from our config file
$config = (Get-Content "$PSScriptRoot\config\vsn.json") -Join "`n" | ConvertFrom-Json

# Logging
if($config.Debug_Log) {
	Start-Logging "$PSScriptRoot\log\debug.log"
}

#Add-PSSnapin Veeam.Backup.PowerShell
Import-Module Veeam.Backup.PowerShell

# Get the session
$session = Get-VBRBackupSession | ?{($_.OrigJobName -eq $JobName) -and ($Id -eq $_.Id.ToString())}
# Wait for the session to finish up
while ($session.IsCompleted -eq $false) {
	Write-LogMessage 'Info' 'Session not finished Sleeping...'
	Start-Sleep -m 200
	$session = Get-VBRBackupSession | ?{($_.OrigJobName -eq $JobName) -and ($Id -eq $_.Id.ToString())}
}

# Save same session info
[String]$Status = $session.Result
$JobName = $session.Name.ToString().Trim()
$JobType = $session.JobTypeString.Trim()
[Float]$JobSize = $session.BackupStats.DataSize
[Float]$TransfSize = $session.BackupStats.BackupSize
[Float]$ReadSize = $session.Progress.Readsize
[Float]$DataProcessedSize = $session.Progress.ProcessedUsedSize

# Report job/data size in B, KB, MB, GB, or TB depending on completed size.
## Job size
If([Float]$JobSize -lt 1KB) {
    [String]$JobSizeRound = [Float]$JobSize
    [String]$JobSizeRound += ' B'
}
ElseIf([Float]$JobSize -lt 1MB) {
    [Float]$JobSize = [Float]$JobSize / 1KB
    [String]$JobSizeRound = [math]::Round($JobSize,2)
    [String]$JobSizeRound += ' KB'
}
ElseIf([Float]$JobSize -lt 1GB) {
    [Float]$JobSize = [Float]$JobSize / 1MB
    [String]$JobSizeRound = [math]::Round($JobSize,2)
    [String]$JobSizeRound += ' MB'
}
ElseIf([Float]$JobSize -lt 1TB) {
    [Float]$JobSize = [Float]$JobSize / 1GB
    [String]$JobSizeRound = [math]::Round($JobSize,2)
    [String]$JobSizeRound += ' GB'
}
ElseIf([Float]$JobSize -ge 1TB) {
    [Float]$JobSize = [Float]$JobSize / 1TB
    [String]$JobSizeRound = [math]::Round($JobSize,2)
    [String]$JobSizeRound += ' TB'
}
### If no match then report in bytes
Else{
    [String]$JobSizeRound = [Float]$JobSize
    [String]$JobSizeRound += ' B'
}
## Transfer size
If([Float]$TransfSize -lt 1KB) {
    [String]$TransfSizeRound = [Float]$TransfSize
    [String]$TransfSizeRound += ' B'
}
ElseIf([Float]$TransfSize -lt 1MB) {
    [Float]$TransfSize = [Float]$TransfSize / 1KB
    [String]$TransfSizeRound = [math]::Round($TransfSize,2)
    [String]$TransfSizeRound += ' KB'
}
ElseIf([Float]$TransfSize -lt 1GB) {
    [Float]$TransfSize = [Float]$TransfSize / 1MB
    [String]$TransfSizeRound = [math]::Round($TransfSize,2)
    [String]$TransfSizeRound += ' MB'
}
ElseIf([Float]$TransfSize -lt 1TB) {
    [Float]$TransfSize = [Float]$TransfSize / 1GB
    [String]$TransfSizeRound = [math]::Round($TransfSize,2)
    [String]$TransfSizeRound += ' GB'
}
ElseIf([Float]$TransfSize -ge 1TB) {
    [Float]$TransfSize = [Float]$TransfSize / 1TB
    [String]$TransfSizeRound = [math]::Round($TransfSize,2)
    [String]$TransfSizeRound += ' TB'
}
### If no match then report in bytes
Else{
    [String]$TransfSizeRound = [Float]$TransfSize
    [String]$TransfSizeRound += ' B'
}
## Read size
If([Float]$ReadSize -lt 1KB) {
    [String]$ReadSizeRound = [Float]$ReadSize
    [String]$ReadSizeRound += ' B'
}
ElseIf([Float]$ReadSize -lt 1MB) {
    [Float]$ReadSize = [Float]$ReadSize / 1KB
    [String]$ReadSizeRound = [math]::Round($ReadSize,2)
    [String]$ReadSizeRound += ' KB'
}
ElseIf([Float]$ReadSize -lt 1GB) {
    [Float]$ReadSize = [Float]$ReadSize / 1MB
    [String]$ReadSizeRound = [math]::Round($ReadSize,2)
    [String]$ReadSizeRound += ' MB'
}
ElseIf([Float]$ReadSize -lt 1TB) {
    [Float]$ReadSize = [Float]$ReadSize / 1GB
    [String]$ReadSizeRound = [math]::Round($ReadSize,2)
    [String]$ReadSizeRound += ' GB'
}
ElseIf([Float]$ReadSize -ge 1TB) {
    [Float]$ReadSize = [Float]$ReadSize / 1TB
    [String]$ReadSizeRound = [math]::Round($ReadSize,2)
    [String]$ReadSizeRound += ' TB'
}
### If no match then report in bytes
Else{
    [String]$ReadSizeRound = [Float]$TransfSize
    [String]$ReadSizeRound += ' B'
}
## Data Processed
If([Float]$DataProcessedSize -lt 1KB) {
    [String]$DataProcessedSizeRound = [Float]$DataProcessedSize
    [String]$DataProcessedSizeRound += ' B'
}
ElseIf([Float]$DataProcessedSize -lt 1MB) {
    [Float]$DataProcessedSize = [Float]$DataProcessedSize / 1KB
    [String]$DataProcessedSizeRound = [math]::Round($DataProcessedSize,2)
    [String]$DataProcessedSizeRound += ' KB'
}
ElseIf([Float]$DataProcessedSize -lt 1GB) {
    [Float]$DataProcessedSize = [Float]$DataProcessedSize / 1MB
    [String]$DataProcessedSizeRound = [math]::Round($DataProcessedSize,2)
    [String]$DataProcessedSizeRound += ' MB'
}
ElseIf([Float]$DataProcessedSize -lt 1TB) {
    [Float]$DataProcessedSize = [Float]$DataProcessedSize / 1GB
    [String]$DataProcessedSizeRound = [math]::Round($DataProcessedSize,2)
    [String]$DataProcessedSizeRound += ' GB'
}
ElseIf([Float]$DataProcessedSize -ge 1TB) {
    [Float]$DataProcessedSize = [Float]$DataProcessedSize / 1TB
    [String]$DataProcessedSizeRound = [math]::Round($DataProcessedSize,2)
    [String]$DataProcessedSizeRound += ' TB'
}
### If no match then report in bytes
Else{
    [String]$ReadSizeRound = [Float]$TransfSize
    [String]$ReadSizeRound += ' B'
}
# Job duration
$Duration = $session.Info.EndTime - $session.Info.CreationTime
$TimeSpan = $Duration
$Duration = '{0:00}h {1:00}m {2:00}s' -f $TimeSpan.Hours, $TimeSpan.Minutes, $TimeSpan.Seconds

# Rate

if ($session.Progress.AvgSpeed > 0)
{
  $Rate = [math]::Round($session.Progress.AvgSpeed/1024/1024,0)
}else{
  $Rate = 0
}

# Compress ratio
if ($session.BackupStats.CompressRatio > 0)
{
  $CompressRatio = 100 / $session.BackupStats.CompressRatio
  $CompressRatio = [math]::Round($CompressRatio,2)
}else{
  $CompressRatio = 0
}

# Dedup ratio
if ($session.BackupStats.DedupRatio > 0)
{
  $DedupRatio = 100 / $session.BackupStats.DedupRatio
  $DedupRatio = [math]::Round($DedupRatio,2)
}else{
  $DedupRatio = 0
}

# Completion pourcentage
$CompletePourcent = ' ({0}%)' -f  $session.Info.CompletionPercentage

# Switch for card theme colour
Switch ([String]$Status) {
    None {$Colour = ''}
    Failed {$Colour = 'ff0000'}
    Warning {$Colour = 'ffe100'}
    Success {$Colour = '00ff00'}
    Default {$Colour = ''}
}
# Switch for status image
Switch ([String]$Status) {
    None {$StatusImg = $config.VeeamBRIcon}
    Failed {$StatusImg = $config.JobFailureImage}
    Warning {$StatusImg = $config.JobWarningImage}
    Success {$StatusImg = $config.JobSuccessImage}
    Default {$StatusImg = $config.VeeamBRIcon}
}

# Build the payload
$Card = ConvertTo-Json -Depth 4 @{
    Summary = 'Veeam B&R Report - ' + ($JobName)
    themeColor = $Colour
    sections = @(
        @{
            title = '**Veeam Backup & Replication**'
            activityImage = $StatusImg
            activityTitle = $JobName
            activitySubtitle = (Get-Date -UFormat “%A %d %B %Y %T”)
            facts = @(
                @{
                    name = "Job status:"
                    value = [String]$Status
                },
                @{
					name = "Backup size:"
					value = $JobSizeRound
				},
				@{
					name = "Rate Processing:"
					value = [String]$Rate += ' MB/s'
				},
				@{
					name = "Read data:"
					value = $ReadSizeRound
				},
				@{
					name = "Transferred data:"
					value = $TransfSizeRound
				},
				@{
					name = "Processed data:"
					value = $DataProcessedSizeRound += $CompletePourcent
				},
				@{
					name = "Dedupe ratio:"
					value = [String]$DedupRatio += ' %'
                },
				@{
					name = "Compress ratio:"
					value =	[String]$CompressRatio += ' %'
                },
				@{
					name = "Duration:"
					value = $Duration
				}
            )
        }
    )
}

# Send report to Teams
Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $Card -Uri $config.webhook
