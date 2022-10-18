Option Explicit

Dim baseFolder, linkFile1, linkFile2, targetPath, objShell1, objShell2, desktopPath, progPath, oFSO
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")	
	oFSO.CreateFolder desktopPath & "\My program"
   
	Set objShell2 = Wscript.CreateObject("Wscript.Shell")
	progPath = objShell2.SpecialFolders("Programs")
	 

    With WScript.CreateObject("Scripting.FileSystemObject")
        baseFolder = .BuildPath( .GetParentFolderName( WScript.ScriptFullName ), "\")
        linkFile1   = .BuildPath( desktopPath, "My program.lnk" )
	linkFile2   = .BuildPath( progPath, "My program.lnk" )
	targetPath = .BuildPath( baseFolder, "My program.pyw" )
    End With 

    With WScript.CreateObject("WScript.Shell").CreateShortcut( linkFile1 )

        .TargetPath = targetPath
        .WorkingDirectory = baseFolder
		.IconLocation= baseFolder & "icon.ico"
        .Save
    End With
	
	With WScript.CreateObject("WScript.Shell").CreateShortcut( linkFile2 )

        .TargetPath = targetPath
        .WorkingDirectory = baseFolder
		.IconLocation= baseFolder & "icon.ico"
        .Save
    End With
	
	
