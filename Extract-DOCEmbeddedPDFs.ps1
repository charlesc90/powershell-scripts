Function Get-EmbeddedDOC {
 
	# Set $docfiles to the path of the directory with the MS Word Documents
	# $docfiles = ''
	
	New-Item -ItemType Directory -Path "$($docfiles)\extracted"
	
	# Loop through all .doc files
    $files = Get-ChildItem -Path $docfiles -Filter "*.doc"
    Foreach ($file in $files) {
    
        # Define word COM object
	    $word = New-Object -ComObject word.application
	    $word.Visible = $False

        # Open each file
        $document = $word.documents.open($file.FullName, $false, $true)

        # Word documents that contain an embedded PDF will contain the regex 'oleobject' before the binary data
        $pdf = $document.WordOpenXML 
        If ( $pdf | Select-String -Pattern 'oleobject' -SimpleMatch) {
            
            # This function uses 7z to extract the word document to $docfiles
            Function Extract-Embedded([string]$Path, [string]$Destination) {

                # Give filepath of 7zip executable
                $7z_Application = "C:\Program Files\7-zip\7z.exe"
                $7z_Arguements = @(
                    # Extract files with fill paths
                    'x'
                    # Assume yes on all prompts
                    '-y'
                    # Set output directory
                    "`"-o$($Destination)`""
                    # Name of the archive to extract
                    "`"$($Path)`""
                )
                & $7z_Application $7z_Arguements
            }
            Extract-Embedded -Path $file.FullName -Destination "$($docfiles)\extracted"

            # Closes the current instance of MS Word
            $word.Quit()

            # Load embedded PDFs named 'CONTENTS' in the extracted archive
            $embeds = Get-ChildItem -LiteralPath "$($docfiles)\extracted" -Recurse -Filter "CONTENTS" 
            # Load embedded PDFs named '[1]Ole10Native' in the extracted archive
            $oles = Get-ChildItem -LiteralPath "$($docfiles)\extracted" -Recurse -Filter "[1]Ole10Native" 
            # Loops through embedded PDFs named 'CONTENTS'
            For ($i = 0; $i -lt $embeds.Count; $i++) {
            
                # Copy the item to $docfiles, rename it, and change its extention to .pdf
                Copy-Item -LiteralPath $embeds[$i].FullName -Destination "$($docfiles)\$($file.BaseName)_$i.pdf"
                # Write on the console
                Write-Host "contentsPDF extracted from:"$file.Name -ForegroundColor Yellow

            }
            # Loops through embedded PDFs named '[1]Ole10Native'
            For ($j = 0; $j -lt $oles.Count; $j ++) {

                Copy-Item -LiteralPath $oles[$j].FullName -Destination "$($docfiles)\$($file.BaseName)_$i.pdf"
                # Write on the console
                Write-Host "olePDF extracted from:"$file.Name -ForegroundColor Yellow
        
            }

        }
        # Clean $docfiles\extracted\* for the next iteration
        Remove-Item -Path "$($docfiles)\extracted\*" -Recurse

        $word.Quit() 
    }

}
Get-EmbeddedDOC
