cd c:\DLP_Test
Remove-Item attachments -Force -Recurse
mkdir attachments

###Getting user input###
$policy="Driving License"
$record_array = @{}
$record_array.2 = 'Allow'
$record_array.9 = 'Warn'
$record_array.49 = 'Block'
$record_array.51 = 'BLOCK'

foreach($record_number in $record_array.keys)
{ 
    
###Generating Driving License random numbers###
1..$record_number |ForEach-Object{    
    $DL_gen1=-join ((65..90) | Get-Random -Count 5 | % {[char]$_}) #first 5 chars of surname simulated
    $DL_gen2=Get-Random -minimum 1 -Maximum 9                      #decade single digit
    $DL_gen3=Get-Random -Minimum 11 -Maximum 12                    #month of birth double digit
    $DL_gen4=Get-Random -Minimum 11 -Maximum 31                    #date within month double digit
    $DL_gen5=Get-Random -minimum 1 -Maximum 9                      #last year digit
    $DL_gen6=-join ((65..90) | Get-Random -Count 2 | % {[char]$_}) #first 2 chars of first name simulated 
    $DL_gen7=Get-Random -Maximum 9                                 #random digit (normally number 9)
    $DL_gen8=-join ((65..90) | Get-Random -Count 2 | % {[char]$_}) #2 computer check characters
    #$DL_gen9=Get-Random -Minimum 10 -Maximum 99                    #2 characters representing number of licenses issued 
   

###Putting it all together###
    $DL_Number_join=-join ($DL_gen1 +$DL_gen2+ $DL_gen3 +$DL_gen4 + $DL_gen5 +$DL_gen6 +$DL_gen7 +$DL_gen8) | Out-String
    $DL_body +=$DL_number_join | Out-String
    $DL_attachment +=$DL_number_join | Out-String

}
########################################
###Generating accompanying attachment###
########################################
$DL_attachment | Out-File .\attachments\$record_number'_records_DL_attachment.csv' 

##############################
###Generating docx document###
##############################
$FileName = "$record_number`_records_DL_attachment.docx"
$savepath="c:\DLP_Test\attachments\$($record_number)`_records_DL_attachment.docx"
$word=new-object -ComObject "Word.Application" 
$doc=$word.documents.Add() 
$DL_word=$word.Selection 
$DL_word.TypeParagraph() 
$DL_word.Style="Normal" 
$DL_word.TypeText("Data: $($DL_attachment)") 
$doc.SaveAs([ref]$savepath)     
$doc.Close()   
$word.quit() 

Invoke-Item $savepath

#############################
###Generating doc document###
#############################
$FileName = "$($record_number)`_records_DL_attachment.doc"
$savepath="c:\DLP_Test\attachments\$($record_number)`_records_DL_attachment.doc"
$word=new-object -ComObject "Word.Application" 
$doc=$word.documents.Add() 
$DL_word=$word.Selection 
$DL_word.TypeParagraph() 
$DL_word.Style="Normal" 
$DL_word.TypeText("Data: $($DL_attachment)") 
$doc.SaveAs([ref]$savepath)     
$doc.Close()   
$word.quit() 
Invoke-Item $savepath

#######################################
###Generating xlsX document from csv###
#######################################

### Set input and output path
$inputCSV = "c:\DLP_Test\attachments\$($record_number)`_records_DL_attachment.csv"
$outputXLSX = "c:\DLP_Test\attachments\$($record_number)`_records_DL_attachment.xlsx"

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
$inputCSV = "c:\DLP_Test\attachments\$($record_number)`_records_DL_attachment.csv"
$outputXLS = "c:\DLP_Test\attachments\$($record_number)`_records_DL_attachment.xls"

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


$DL_attachment | Out-File .\attachments\$record_number'_records_DL_attachment.txt'

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
Send-MailMessage -From $from -To $to -Subject $subject -Body $DL_body -SmtpServer $smtpserver -Port $smtp_port
Write-Host "Body text emailed."
Clear-Variable DL_body
Clear-Variable DL_attachment

#################################################################
###Emailing no body, but with various attachments of same data###
#################################################################
Send-MailMessage -From $from -To $to -Subject $attachment_subject' csv' -Body 'Attachment-test-csv' -Attachments .\attachments\$record_number'_records_DL_attachment.csv' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' xls' -Body 'Attachment-test-xls' -Attachments .\attachments\$record_number'_records_DL_attachment.xls' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' xlsx' -Body 'Attachment-test-xlsx' -Attachments .\attachments\$record_number'_records_DL_attachment.xlsx' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' doc' -Body 'Attachment-test-doc' -Attachments .\attachments\$record_number'_records_DL_attachment.doc' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' docx' -Body 'Attachment-test-docx' -Attachments .\attachments\$record_number'_records_DL_attachment.docx' -SmtpServer $smtpserver -Port $smtp_port
Send-MailMessage -From $from -To $to -Subject $attachment_subject' txt' -Body 'Attachment-test-txt' -Attachments .\attachments\$record_number'_records_DL_attachment.txt' -SmtpServer $smtpserver -Port $smtp_port
Write-host "Attachments emailed"
}

###Clearing the stage###
#Clear-Variable record_number
#Clear-Variable DL_body

