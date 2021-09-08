#region Forms

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$decAutomation                   = New-Object system.Windows.Forms.Form
$decAutomation.ClientSize        = '400,870'
$decAutomation.text              = "DEC-Automation"
$decAutomation.TopMost           = $false

$cleanDEC                        = New-Object system.Windows.Forms.Button
$cleanDEC.text                   = "1. Clean DEC Directories"
$cleanDEC.width                  = 330
$cleanDEC.height                 = 30
$cleanDEC.location               = New-Object System.Drawing.Point(30,20)
$cleanDEC.Font                   = 'Microsoft Sans Serif,10'

$downloadDEC                     = New-Object system.Windows.Forms.Button
$downloadDEC.text                = "2. Download DEC Files"
$downloadDEC.width               = 330
$downloadDEC.height              = 30
$downloadDEC.location            = New-Object System.Drawing.Point(30,70)
$downloadDEC.Font                = 'Microsoft Sans Serif,10'

$extractDEC                      = New-Object system.Windows.Forms.Button
$extractDEC.text                 = "3. Extract DEC Files"
$extractDEC.width                = 330
$extractDEC.height               = 30
$extractDEC.location             = New-Object System.Drawing.Point(30,120)
$extractDEC.Font                 = 'Microsoft Sans Serif,10'

$moveDuplicates                  = New-Object system.Windows.Forms.Button
$moveDuplicates.text             = "4. Move Duplicates"
$moveDuplicates.width            = 330
$moveDuplicates.height           = 30
$moveDuplicates.location         = New-Object System.Drawing.Point(30,170)
$moveDuplicates.Font             = 'Microsoft Sans Serif,10'

$moveFiles                       = New-Object system.Windows.Forms.Button
$moveFiles.text                  = "6. Copy To Be Processed"
$moveFiles.width                 = 330
$moveFiles.height                = 30
$moveFiles.location              = New-Object System.Drawing.Point(30,220)
$moveFiles.Font                  = 'Microsoft Sans Serif,10'

$extractDownload                 = New-Object system.Windows.Forms.Button
$extractDownload.text            = "7. Extract Links and Download Files"
$extractDownload.width           = 330
$extractDownload.height          = 30
$extractDownload.location        = New-Object System.Drawing.Point(30,270)
$extractDownload.Font            = 'Microsoft Sans Serif,10'

$copyDownloads                   = New-Object system.Windows.Forms.Button
$copyDownloads.text              = "8. Copy Downloads and Files to be Processed"
$copyDownloads.width             = 330
$copyDownloads.height            = 30
$copyDownloads.location          = New-Object System.Drawing.Point(30,320)
$copyDownloads.Font              = 'Microsoft Sans Serif,10'

$sortFiles                       = New-Object system.Windows.Forms.Button
$sortFiles.text                  = "9. Sort Files"
$sortFiles.width                 = 330
$sortFiles.height                = 30
$sortFiles.location              = New-Object System.Drawing.Point(30,370)
$sortFiles.Font                  = 'Microsoft Sans Serif,10'

$sortPictures                    = New-Object system.Windows.Forms.Button
$sortPictures.text               = "10. Sort Pictures (Harris)"
$sortPictures.width              = 330
$sortPictures.height             = 30
$sortPictures.location           = New-Object System.Drawing.Point(30,420)
$sortPictures.Font               = 'Microsoft Sans Serif,10'

$sortPicturesJefferson           = New-Object system.Windows.Forms.Button
$sortPicturesJefferson.text      = "10. Sort Pictures (Jefferson)"
$sortPicturesJefferson.width     = 330
$sortPicturesJefferson.height    = 30
$sortPicturesJefferson.location  = New-Object System.Drawing.Point(30,470)
$sortPicturesJefferson.Font      = 'Microsoft Sans Serif,10'

$getEmbedded                     = New-Object system.Windows.Forms.Button
$getEmbedded.text                = "11. Get Embedded PDFs (DOCX)"
$getEmbedded.width               = 330
$getEmbedded.height              = 30
$getEmbedded.location            = New-Object System.Drawing.Point(30,520)
$getEmbedded.Font                = 'Microsoft Sans Serif,10'

$getEmbeddedDOC                  = New-Object system.Windows.Forms.Button
$getEmbeddedDOC.text             = "11. Get Embedded PDFs (DOC)"
$getEmbeddedDOC.width            = 330
$getEmbeddedDOC.height           = 30
$getEmbeddedDOC.location         = New-Object System.Drawing.Point(30,570)
$getEmbeddedDOC.Font             = 'Microsoft Sans Serif,10'

$sortEmbeddedHarris              = New-Object system.Windows.Forms.Button
$sortEmbeddedHarris.text         = "12. Sort PDFs (Harris)"
$sortEmbeddedHarris.width        = 330
$sortEmbeddedHarris.height       = 30
$sortEmbeddedHarris.location     = New-Object System.Drawing.Point(30,620)
$sortEmbeddedHarris.Font         = 'Microsoft Sans Serif,10'

$sortEmbeddedJefferson           = New-Object system.Windows.Forms.Button
$sortEmbeddedJefferson.text      = "12. Sort PDFs (Jefferson)"
$sortEmbeddedJefferson.width     = 330
$sortEmbeddedJefferson.height    = 30
$sortEmbeddedJefferson.location  = New-Object System.Drawing.Point(30,670)
$sortEmbeddedJefferson.Font      = 'Microsoft Sans Serif,10'

$prepareOracle                   = New-Object system.Windows.Forms.Button
$prepareOracle.text              = "13. Prepare for Oracle"
$prepareOracle.width             = 330
$prepareOracle.height            = 30
$prepareOracle.location          = New-Object System.Drawing.Point(30,720)
$prepareOracle.Font              = 'Microsoft Sans Serif,10'

$prepareOracle                   = New-Object system.Windows.Forms.Button
$prepareOracle.text              = "13. Prepare for Oracle Jefferson"
$prepareOracle.width             = 330
$prepareOracle.height            = 30
$prepareOracle.location          = New-Object System.Drawing.Point(30,770)
$prepareOracle.Font              = 'Microsoft Sans Serif,10'

$decAutomation.controls.AddRange(@($cleanDEC,$downloadDEC,$extractDEC,$moveDuplicates,$moveFiles,$extractDownload,$copyDownloads,$sortFiles,$sortPictures,$sortPicturesJefferson,$getEmbedded,$getEmbeddedDOC,$sortEmbeddedHarris,$sortEmbeddedJefferson,$prepareOracle,$prepareOracleJefferson))

$cleanDEC.Add_Click({ Clean-DECDirectories })
$downloadDEC.Add_Click({ Download-DECFiles })
$extractDEC.Add_Click({ Extract-DECFiles })
$moveDuplicates.Add_Click({ Move-Duplicates })
$moveFiles.Add_Click({ Move-ToBeProcessed })
$extractDownload.Add_Click({ ExtractDownload-Hyperlinks })
$copyDownloads.Add_Click({ Copy-Downloads })
$sortFiles.Add_Click({ Sort-Files })
$sortPictures.Add_Click({ Sort-Pictures })
$sortPicturesJefferson.Add_Click({ Sort-PicturesJefferson })
$getEmbedded.Add_Click({ Get-Embedded })
$getEmbeddedDOC.Add_Click({ Get-EmbeddedDOC })
$sortEmbeddedHarris.Add_Click({ Sort-EmbeddedHarris })
$sortEmbeddedJefferson.Add_Click({ Sort-EmbeddedJefferson })
$prepareOracle.Add_Click({ Prepare-Oracle })
$prepareOracle.Add_Click({ Prepare-OracleJefferson })
