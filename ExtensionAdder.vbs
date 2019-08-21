'This script scans folder files by identifying the file type and changing the extension of the file in case all files have no extensions
'Author Asem Saeedi Alawadhi 
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
objStartFolder = "Please ADD THE PATH"
 
 
Set objFolder = objFSO.GetFolder(objStartFolder)
 
 
Set colFiles = objFolder.Files
 
For Each objFile in colFiles
 
    'Wscript.Echo objFile.Name
    Dim filename
    filename = objFile.Name
    Dim filepath
    filepath = objStartFolder & filename
    'Wscript.Echo filepath
    Set objFile2 = objFSO.OpenTextFile(filepath, 1)
    Dim strCharacters       
    'Do Until objFile2.AtEndOfStream
       strCharacters= objFile2.Read(4)
       Wscript.Echo strCharacters
                'Loop
                objFile2.Close
    If strCharacters = "%PDF" Then
        Dim strNewFileName
        strNewFileName =  filename &".pdf"
                                objFile.Name = strNewFileName
    End If
    
Next