<% 

Sub WriteFile(sFilePathAndName,sFileContents)   

  Const ForWriting =2 

  Set oFS = Server.CreateObject("Scripting.FileSystemObject") 
  Set oFSFile = oFS.OpenTextFile(sFilePathAndName,ForWriting,True) 

  oFSFile.Write(sFileContents) 
  oFSFile.Close 

  Set oFSFile = Nothing 
  Set oFS = Nothing

End Sub 

Function ReadFile(sFilePathAndName) 

   dim sFileContents 

   Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

   If oFS.FileExists(sFilePathAndName) = True Then 
       
      Set oTextStream = oFS.OpenTextFile(sFilePathAndName,1) 
       
      sFileContents = oTextStream.ReadAll 
     
      oTextStream.Close 

      Set oTextStream = nothing 
   
   End if 
   
   Set oFS = nothing 

   ReadFile = sFileContents 

End Function 

Sub ReadFileLineByLine(sFilePathAndName) 

   Const ForReading = 1 
   Const ForWriting = 2 
   Const ForAppending = 8 
   Const TristateUseDefault = -2 
   Const TristateTrue = -1 
   Const TristateFalse = 0 

   Dim oFS 
   Dim oFile 
   Dim oStream 

   Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

   Set oFile = oFS.GetFile(sFilePathAndName) 
   
   Set oStream = oFile.OpenAsTextStream(ForReading, TristateUseDefault) 

   Do While Not oStream.AtEndOfStream 
     
      sRecord=oStream.ReadLine 

      Response.Write  sRecord 

   Loop 

   oStream.Close 

  End Sub 


Sub RemoveFolder(sPath,fRemoveSelf) 

  Dim oFS   
  Dim oFSFolder   
   
  Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

  If oFS.FolderExists(sPath)  <> True Then 
    Set oFS = Nothing 
    Exit Sub 
  End If 
   
  Set oFSFolder = oFS.GetFolder(sPath) 
   
  RemoveSubFolders oFSFolder 
   
  If fRemoveSelf = True Then 

     If oFS.FolderExists(sPath) = True Then 
        oFSFolder.Delete True 
     Else 
        Set oFSFolder = Nothing 
        Set oFS = Nothing 
        Exit Sub 
     End If 

  End If 
   
   Set oFSFolder = Nothing 
   Set oFS = Nothing 

End Sub 


Sub RemoveFolderIfEmpty(sPath) 
  Dim oFS   
  Dim oFSFolder   
   
  Set oFS = Server.CreateObject("Scripting.FileSystemObject") 

  If oFS.FolderExists(sPath)  <> True Then 
    Set oFS = Nothing 
    Exit Sub 
  End If 
   
  Set oFSFolder = oFS.GetFolder(sPath) 
   
  If oFSFolder.files.count = 0 Then 
     RemoveFolder sPath,True
  else
	if oFSFolder.files.count = 1 Then
		'Si el archivo es default.asp, se considera un folder vacio
		For each item in oFSFolder.files
			if item.name = "default.asp" then
				RemoveFolder sPath,True
			end if
		next
	end if
  End If 
   
   Set oFSFolder = Nothing 
   Set oFS = Nothing 

End Sub 

Sub RemoveSubFolders(oFSFolder) 

   Dim oFSFile 
   Dim oFSSubFolder   
   
   For Each oFSFile In oFSFolder.Files 
         oFSFile.Delete True 
   Next 

   For Each oFSSubFolder In oFSFolder.SubFolders 
         RemoveSubFolders oFSSubFolder 
         oFSSubFolder.Delete True 
   Next 
     
   Set oFSFile = Nothing 
   
End Sub 


Sub RemoveFile(sFilePathAndName) 

  Set oFS = Server.CreateObject("Scripting.FileSystemObject") 
   
  If oFS.FileExists(sFilePathAndName) = True Then 
     oFS.DeleteFile sFilePathAndName, True 
  end if 

  Set oFS = Nothing 
   
End Sub 

function UploadFile(objUpload, fieldName, strFileName, completePath, replaceFile)
	Dim files_originalFileName, files_subFSO, files_counter
	
	if IsNull(fieldName) or fieldName="" then
		' -- limpiar variables y terminar
		set files_originalFileName = nothing
		set files_subFSO = nothing
		
		UploadFile = ""
		Exit function
		
	else
		if strFileName = "" Then
			' Grab the file name
			strFileName = objUpload.Fields(fieldName).FileName	
		end if
		
		if strFileName = "" Then
			' -- limpiar variables y terminar
			set files_originalFileName = nothing
			set files_subFSO = nothing
			
			UploadFile = ""
			Exit function
			
		else
            strFileName = Replace(strFileName, " ", "_")
			files_originalFileName = strFileName
			' Compile path to save file to
			strPath = completePath & strFileName
			
			if NOT replaceFile Then
				'Check if file exists
				Set files_subFSO = server.CreateObject ("Scripting.FileSystemObject")
				files_counter = 1
			
				while(files_subFSO.FileExists(strPath))
					strFileName = files_subFSO.GetBaseName(files_originalFileName) & "(" & files_counter & ")." & files_subFSO.GetExtensionName(files_originalFileName)
					strPath = completePath & strFileName
					files_counter = files_counter+1
				wend
			end if
			' Save the binary data to the file system
			objUpload(fieldName).SaveAs strPath
			
			UploadFile = strFileName
		end if
	end if
	
	set files_originalFileName = nothing
	set files_subFSO = nothing
	set files_counter = nothing
End function

%>