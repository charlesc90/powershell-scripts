########################################################################################################
# Script: ConvertWord-ToPDF.ps1
# Author: Charles Cox
# Date: 26 December 2019
# Description: This script converts a word document to a PDF document
########################################################################################################

#WdSaveFormat Enumeration specifies the format to use when saving a document. PDF format is wdFormatPDF with a value of 17
$wdFormatPDF = 17
# define word COM object
$word = New-Object -ComObject Word.Application
$word.visible = $false
# name of the file to convert
$file = 'testword.docx'

Get-ChildItem -Path C:\Users\Charles.Cox\Documents\powershell-scripts-master\powershell-scripts-master\test\$file | ForEach-Object {
	
	$name = $file
	$destination = "C:\Users\Charles.Cox\Documents\powershell-scripts-master\powershell-scripts-master\test\pdf\"
	$doc = $word.documents.open($_.fullname)
	$doc.saveas([ref]"$destination\$name", [ref]$wdFormatPDF)
	$doc.close()

}
$word.Quit()
