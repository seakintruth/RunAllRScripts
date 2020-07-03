Option Explicit
Const strRscriptExe = "%userprofile%\Documents\R-3.5.1\bin\x64\Rscript.exe" 

Dim objShell
Set objShell = CreateObject("Wscript.Shell")

Dim strPath
strPath = Wscript.ScriptFullName

Dim ObjFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objFile
Set objFile = objFSO.GetFile(strPath)

'Folder of this script...
Dim strFolder
strFolder = objFSO.GetParentFolderName(objFile) 

Dim strCommand
strCommand = objShell.ExpandEnvironmentStrings("""" & strRscriptExe & """ """ & strFolder & "\RunAll.R""")

on error resume next
objShell.Run(strCommand), 0, true 

if err.number = 70 then
    msgbox err.description & " while attempting to run:" & vbcrlf & strCommand ,0,"ERROR: RunAllWithRscript.vbs"
end if
