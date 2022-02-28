<#
Created: 2022-02-26 15 h
Updated: 2022-02-26 18 h 15
#>

Function MainFK()
{
	Cls
	
	[string]$path = "C:/src/Send-MailMessageFK/"
	[string]$template = "$($path)modele.docx"
	[string]$attachment = "$($path)test.docx"
	[string]$body = "Bonjour,
		Ceci est un test.
		
		Merci. Bonne journée.
		
		Simon Charest, Analyste programmeur"
	[string]$from = "Simon Charest <simoncharest@gmail.com>"
	[string]$smtpServer = "infidem.biz"
	[string]$subject = "Ceci est un test"
	[string[]]$tos = @("Simon Charest <simoncharest@gmail.com>")
	[double]$seconds = 30
	[string]$findText = "{{PRENOM}}"
	[string]$replaceWith = "Simon"
	
	<#
	[object]$wordApp = New-Object –ComObject:"Word.Application" -Strict -Property:@{Visible = $True}
	[object]$document = $wordApp.Documents.Open($attachment)
	[object]$selection = $wordApp.Selection
	$selection.Find.Execute -FindText:$findtext -ReplaceWith:$replaceWith 
	$document.SaveAs($attachement)
    $document.Close()
	$wordApp.Quit()
	#>
	
	ForEach ($to In $tos)
	{
		Write-Host "Sending email to $($to)..." -ForegroundColor:"Yellow"
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		Send-MailMessage -Attachments:$attachment -Body:$body -From:$from -SmtpServer:$smtpServer -Subject:$subject -To:$to -UseSsl:$True -Port:25

		Write-Host "Waiting $($seconds) seconds..." -ForegroundColor:"Yellow"
		Start-Sleep -Seconds:$seconds
	}
	
	Write-Host "** DONE **" -ForegroundColor:"Green"
}

MainFK
	