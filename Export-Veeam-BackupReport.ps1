function Load_Veeam_Module {

	Add-PSsnapin VeeamPSSnapin 
}

function Export-Veeam-HTML-Report {
	$HTML_Report_sytle = "<link rel=stylesheet type=text/css href=./mystyle.css charset=utf-8>"
	$HTML_Report_header = "<!DOCTYPE html><html><head><Title>Veeam Backup</Title>$HTML_Report_sytle</head><body><hr>"
	#$HTML_Report_menu = Export-Menu-HTML
	$today_modified = Get-date
	$Update_date = "Updated :" + $today_modified
	$Veeam_job_html = Export-Table-HTML($CSV_Veeam_Job_file)("Backup Job")
	$Veeam_ProtectedVM_html = Export-Table-HTML($CSV_Veeam_ProtectedVM_file)("Protected VM")
	$HTML_Report_footer = "</body></html>"
	
	$HTML_Report = $HTML_Report_header + $HTML_Report_menu + $Update_date + $Veeam_job_html + $Veeam_ProtectedVM_html + $HTML_Report_footer
	
	$HTML_Report | Out-File $HTML_Veeam_file
	Write-Host "File saved to $HTML_Veeam_file"
}

function Export-Table-HTML {
	param([string] $filepath, [string] $heading)
	
	$txt = Import-CSV -Path $filepath
	[string]$txt2html = $txt | ConvertTo-Html 
	
	$txt2html -match "<body>(?<content>.*)</body>" | out-null
	$txt2html_tbl = $matches['content']
	
	$txt2html_tbl = $txt2html_tbl -replace "<table>","<table id=myTable>"
	
	$heading = "<h2>$heading</h2>"
	$No_info = ""
	If (!$txt) {
		$No_info = "No Information"
	}
	$heading_table_html = $heading + $No_info + $txt2html_tbl
	
	Return $heading_table_html
}

function Remove_Veeam_CSV {
	$CSV_folder = ".\CSV"
	$CSV_Veeam_file = "$CSV_folder\Veeam_*.csv"
	
	if (Test-Path -Path $CSV_Veeam_file){
		Get-ChildItem $CSV_Veeam_file | Remove-Item
		Write-Host "Removed $CSV_Veeam_file"
	}
}

Load_Veeam_Module

$CSV_folder = ".\CSV"
$CSV_Veeam_Job_file = "$CSV_folder\Veeam_Job_Status.csv"
$CSV_Veeam_ProtectedVM_file = "$CSV_folder\Veeam_ProtectedVM_Status.csv"

$HTML_folder = ".\Report"
$HTML_Veeam_file = "$HTML_folder\Veeam.html"

Remove_Veeam_CSV

$Veeam_List = Import-CSV .\Veeam.csv

ForEach ($Veeam in $Veeam_List){
	Connect-VBRServer -Server $Veeam.vbrserver -User $Veeam.user -Password $Veeam.password
	
	$vbrserver = Get-VBRServersession
	if ((Get-Date).DayofWeek -eq 'Monday'){
		$vbrsessions = Get-VBRBackupSession | Where-object {$_.EndTime -ge (Get-Date).addhours(-72)}
	}
	else {
		$vbrsessions = Get-VBRBackupSession | Where-object {$_.EndTime -ge (Get-Date).addhours(-24)}
	}
	$vbrsessions | Select name, Result, CreationTime, EndTime, 
		@{N="Duration";E={$VMTime = ($_.EndTime - $_.CreationTime); $VMTime.ToString('hh') + " Hours " + $VMTime.ToString('mm') + " Minutes"}},
		@{N="BackupServer";E={$vbrserver.server}} | sort CreationTime | Export-CSV -Path $CSV_Veeam_Job_file -NoTypeInformation -Append
	Write-Host "File saved to $CSV_Veeam_Job_file"
		
	$backupedvms = foreach ($session in $vbrsessions) {$session.gettasksessions() | Select Name, Status, Jobname}
	$setarray = @("Name", "Status", "JobName") 
	$backupedvms | select-object -Property $setarray | Export-CSV -Path $CSV_Veeam_ProtectedVM_file -NoTypeInformation -Append
	Write-Host "File saved to $CSV_Veeam_ProtectedVM_file"
	
	Disconnect-VBRServer
}

Export-Veeam-HTML-Report
