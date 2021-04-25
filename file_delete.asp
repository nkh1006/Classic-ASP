' // 폴더 삭제
Function folderDelete(path)
	DIM FSO
	SET FSO = CreateObject("Scripting.FileSystemObject")
	IF (FSO.FolderExists(path)) THEN FSO.DeleteFolder(path)
	SET FSO = NOTHING
END Function

' // 파일 삭제
Function FileDelete(path, filename)
	Dim fso, strfile
	SET fso = Server.CreateObject("scripting.FileSystemObject")
	strfile = path & filename
	IF fso.FileExists(strfile) THEN fso.DeleteFile(strfile)
	SET fso = NOTHING
End Function