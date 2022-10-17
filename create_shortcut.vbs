Option Explicit

Dim baseFolder, linkFile, targetPath, objShell, desktopPath

   Set objShell = Wscript.CreateObject("Wscript.Shell")
   desktopPath = objShell.SpecialFolders("Desktop") 

    With WScript.CreateObject("Scripting.FileSystemObject")
        baseFolder = .BuildPath( .GetParentFolderName( WScript.ScriptFullName ), "\")
        linkFile   = .BuildPath( desktopPath, "Link name.lnk" )
	targetPath = .BuildPath( baseFolder, "Program name.exe" )
    End With 

    With WScript.CreateObject("WScript.Shell").CreateShortcut( linkFile )

        .TargetPath = targetPath
        .WorkingDirectory = baseFolder
		.IconLocation= baseFolder & "icon.ico"
        .Save
    End With