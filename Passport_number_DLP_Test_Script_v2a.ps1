cd c:\DLP_Test
Remove-Item attachments -Force -Recurse
mkdir attachments

###Getting user input###
$policy="passport Details"
$record_array = @{}
$record_array.2 = 'Allow'
$record_array.9 = 'Warn'
$record_array.49 = 'Block'
$record_array.51 = 'BLOCK'

foreach($record_number in $record_array.keys)
{ 

###Generating passport random numbers###
$random_number_gen=1..$record_number| ForEach-Object {Get-Random -Minimum 111111111 -Maximum 999999999}
$passport_body=$random_number_gen |Out-String
$passport_attachment=$random_number_gen | Out-String

########################################
###Generating accompanying attachment###
########################################
$passport_attachment | Out-File .\attachments\$record_number'_records_passport_attachment.csv' 

##############################
###Generating docx document###
##############################
$FileName = "$record_number`_records_passport_attachment.docx"
$savepath="c:\DLP_Test\attachments\$record_number`_records_passport_attachment.docx"
$word=new-object -ComObject "Word.Application" 
$doc=$word.documents.Add() 
$passport_word=$word.Selection 
$passport_word.TypeParagraph() 
$passport_word.Style="Normal" 
$passport_word.TypeText("Data: $($passport_attachment)") 
$doc.SaveAs([ref]$savepath)     
$doc.Close()   
$word.quit() 

Invoke-Item $savepath

#############################
###Generating doc document###
#############################
$FileName = "$record_number`_records_passport_attachment.doc"
$savepath="c:\DLP_Test\attachments\$record_number`_records_passport_attachment.doc"
$word=new-object -ComObject "Word.Application" 
$doc=$word.documents.Add() 
$passport_word=$word.Selection 
$passport_word.TypeParagraph() 
$passport_word.Style="Normal" 
$passport_word.TypeText("Data: $($passport_attachment)") 
$doc.SaveAs([ref]$savepath)     
$doc.Close()   
$word.quit() 
Invoke-Item $savepath

#######################################
###Generating xlsX document from csv###
#######################################

### Set input and output path
$inputCSV = "c:\DLP_Test\attachments\$record_number`_records_passport_attachment.csv"
$outputXLSX = "c:\DLP_Test\attachments\$record_number`_records_passport_attachment.xlsx"

### Create a new Excel Workbook with one empty sheet
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

### Build the QueryTables.Add command
### QueryTables does the same as when clicking "Data » From Text" in Excel
$TxtConnector = ("TEXT;" + $inputCSV)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)

### Set the delimiter (, or ;) according to your regional settings
$query.TextFileOtherDelimiter = $Excel.Application.International(5)

### Set the format to delimited and text for every column
### A trick to create an array of 2s is used with the preceding comma
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

### Execute & delete the import query
$query.Refresh()
$query.Delete()

### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
$Workbook.SaveAs($outputXLSX,51)
$excel.Quit()

######################################
###Generating xls document from csv###
######################################

### Set input and output path
$inputCSV = "c:\DLP_Test\attachments\$record_number`_records_passport_attachment.csv"
$outputXLS = "c:\DLP_Test\attachments\$record_number`_records_passport_attachment.xls"

### Create a new Excel Workbook with one empty sheet
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

### Build the QueryTables.Add command
### QueryTables does the same as when clicking "Data » From Text" in Excel
$TxtConnector = ("TEXT;" + $inputCSV)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)

### Set the delimiter (, or ;) according to your regional settings
$query.TextFileOtherDelimiter = $Excel.Application.International(5)

### Set the format to delimited and text for every column
### A trick to create an array of 2s is used with the preceding comma
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

### Execute & delete the import query
$query.Refresh()
$query.Delete()

### Save & close the Workbook as XLS. Change the output extension for Excel 2003
$Workbook.SaveAs($outputXLS,-4143)
$excel.Quit()

###Quitting Word - RPC Bug in closing###
Stop-Process -Name "WINWORD"
Stop-Process -Name "WINWORD"
Stop-Process -Name "WINWORD"-Force
Stop-Process -Name "WINWORD"-Force
Stop-Process -Name "EXCEL"-Force


$passport_attachment | Out-File .\attachments\$record_number'_records_passport_attachment.txt'

Write-Host "All test data and files should be generated, moving on to emailing..."


###Emailing###
$to="xyz"
$from="xyz"
$bcc=$from
$subject="On Premises DLP Test - '$policy' - '$record_number' - '$($record_array.$record_number)'"
$attachment_subject="On Premises DLP Test ATTACHMENT - '$policy' - '$record_number' - '$($record_array.$record_number)'"
$smtpserver="xyz"
$smtp_port=25

#############################
###Emailing body test data###
#############################

###This line below can be commented out to just send attachments, or uncommented to send body test data emails also###
Send-MailMessage -From $from -To $to -Subject $subject -Body $passport_body -SmtpServer $smtpserver -Port $smtp_port
Write-Host "Body text emailed."
Clear-Variable passport_body
Clear-Variable passport_attachment

#################################################################
###Emailing no body, but with various attachments of same data###
#################################################################
Send-MailMessage -From $from -To $to -Subject $attachment_subject' csv' -Body 'Attachment-test-csv' -Attachments .\attachments\$record_number'_records_passport_attachment.csv' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' xls' -Body 'Attachment-test-xls' -Attachments .\attachments\$record_number'_records_passport_attachment.xls' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' xlsx' -Body 'Attachment-test-xlsx' -Attachments .\attachments\$record_number'_records_passport_attachment.xlsx' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' doc' -Body 'Attachment-test-doc' -Attachments .\attachments\$record_number'_records_passport_attachment.doc' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' docx' -Body 'Attachment-test-docx' -Attachments .\attachments\$record_number'_records_passport_attachment.docx' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' txt' -Body 'Attachment-test-txt' -Attachments .\attachments\$record_number'_records_passport_attachment.txt' -SmtpServer $smtpserver -Port $smtp_port
Write-host "Attachments emailed"
}


###Clearing the stage###
Clear-Variable record_number
Clear-Variable passport_body
Clear-Variable random_number_gen