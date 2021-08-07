# This is the directory of the files where you need to manuipulate filenames
$files = Get-ChildItem -Path 'C:\...'

# Loop through the files in this directory
For ($i = 0; $i -lt $files.Count; $i++) {

    #Replace a string in the filename of the original file with a new string
    Rename-Item -Path $files[$i].FullName -NewName $files[$i].FullName.Replace(".docx","_DEC.docx")
    
}
