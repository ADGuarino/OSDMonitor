# Load Global Data from MDT Webservice

 
$URL = # "http://'ServerHostName':9801/MDTMonitorData/Computers"
		
		function GetMDTData {
		  $Data = Invoke-RestMethod $URL
	
		  foreach($property in ($Data.content.properties) ) {
		    New-Object PSObject -Property @{
		    	Name = $($property.Name);
				PercentComplete = $($property.PercentComplete.'#text');
				Warnings = $($property.Warnings.'#text');
		    	Errors = $($property.Errors.'#text');
				DeploymentStatus = $(
			        Switch ($property.DeploymentStatus.'#text') {
				        1 { "Active/Running" }
				        2 { "Failed" }
				        3 { "Completed" }
						4 { "Unresponisve" }
			        Default { "Unknown" }
			        }
		    	);
                EndTime = $(
						if (($property.EndTime.'#text') -ne $null){
							$(Get-Date ($property.EndTime.'#text')).AddHours(-8)
						} 
					)
                StartTime = $(Get-Date ($property.StartTime.’#text’)).AddHours(-8);
				ID = $($property.ID.'#text')
				UniqueID = $($property.UniqueID.'#text')
				TotalSteps = $($property.TotalSteps.'#text')
				CurrentStep = $($property.StepName)
                TaskSequence = $(
                    Switch ($property.TotalSteps.'#text') {
                        114 { "All Models" }
                        35  { "In-Place Upgrade" }
                        6   { "Mandatory Apps" }
                    Default { "Reimage" }
                    }
                )
			    }
			}
		}

	
$Head = "<style>"
$Head = $Head + "BODY{text-align: center;background-color:white}"
$Head = $Head + "TABLE{margin-left: auto;margin-right: auto;border-width: 4px;border-style: solid;border-color: white;border-collapse: collapse;align: center}"
$Head = $Head + "TH{border-width: 5px;padding: 5px;border-style: solid;border-color: white;background-color:thistle}"
$Head = $Head + "TD{text-align: center;border-width: 5px;padding: 5px;border-style: solid;border-color: white}"
$Head = $Head + "H2{font-size: 28pt; align: center; font-family: 'Tw Cen MT'}"
$Head = $Head + "TD{font-size: 14pt}"
$Head = $Head + "</style>"

$Title = "MDT Deployment Status"

GetMDTData | Select Name, CurrentStep, TaskSequence, StartTime, EndTime, PercentComplete, DeploymentStatus | Sort -Property StartTime -Descending |
ConvertTo-Html  `
-Title $Title  `
-Head  $Head  `
-Body (Get-Date -Format G) `
-PreContent "<H2><u>OSD Status for Windows 10 Anniversary</u></H2>"  `
-PostContent "<P></P>"  `
-Property Name,CurrentStep,TaskSequence,StartTime,EndTime,PercentComplete,DeploymentStatus |
ForEach {
if($_ -like "*<td>Completed</td>*"){$_ -replace "<tr>", "<tr bgcolor=8cec6d>"}
elseif($_ -like "*<td>Failed</td>*"){$_ -replace "<tr>", "<tr bgcolor=f46054>"}
elseif($_ -like "*<td>Active/Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=ffd868>"}
else{$_}
} > C:\inetpub\1607\default.htm
#Invoke-Item C:\inetpub\1607\default.htm
