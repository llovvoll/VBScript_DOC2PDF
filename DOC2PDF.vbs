Dim fso, fld, Path
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Path = fso.GetParentFolderName(WScript.ScriptFullName) ' Get the folder path where the script is located
Set fld = fso.GetFolder(Path) ' Get the folder object using the path

Dim Sum, IsChooseDelete, ThisTime
Sum = 0
Dim LogFile
Set LogFile = fso.OpenTextFile("log.txt", 8, true)

Dim List
Set List = fso.OpenTextFile("ConvertFileList.txt", 2, true)

Call LogOut("Starting to scan files")
Call TreatSubFolder(fld) ' Recursively scan all files and subfolders in the current folder

Sub LogOut(msg)
    ThisTime = Now
    LogFile.WriteLine(Year(ThisTime) & "-" & Month(ThisTime) & "-" & Day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld)
    Dim File
    Dim ts
    For Each File In fld.Files ' Loop through all files in the folder
        If UCase(fso.GetExtensionName(File)) = "DOC" Or UCase(fso.GetExtensionName(File)) = "DOCX" Then
            List.WriteLine(File.Path)
            Sum = Sum + 1
        End If
    Next
    Dim subfld
    For Each subfld In fld.SubFolders ' Recursively scan subfolders
        TreatSubFolder subfld
    Next
End Sub
List.Close

Call LogOut("File scan completed, found " & Sum & " Word documents")

If MsgBox("File scan completed, found " & Sum & " Word documents. The list is saved at" & vbCrLf & fso.GetFolder(Path).Path & "\ConvertFileList.txt" & vbCrLf & "You can edit this file to add or remove files before conversion." & vbCrLf & vbCrLf & "Do you want to convert these documents to PDF format?", vbYesNo + vbInformation, "File Scan Complete") = vbYes Then
    If MsgBox("Do you want to delete the original DOC files after conversion?", vbYesNo + vbInformation, "Confirm File Deletion?") = vbYes Then
        IsChooseDelete = MsgBox("Please confirm again: Do you really want to delete the original DOC files after conversion?", vbYesNo + vbExclamation, "Final Confirmation for File Deletion")
    End If
Else
    MsgBox("Conversion canceled")
    WScript.Quit
End If
MsgBox "Before starting the conversion, close all open Word documents to avoid file access errors.", vbOKOnly + vbExclamation, "Warning"

' Create Word application object (also supports WPS)
Const wdFormatPDF = 17
On Error Resume Next
Set WordApp = CreateObject("Word.Application")
' Try to connect to WPS if Word is not available
If WordApp Is Nothing Then
    Set WordApp = CreateObject("WPS.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("KWPS.Application")
        If WordApp Is Nothing Then
            MsgBox "This script requires Microsoft Office Word 2010 or later, or WPS. Please install Word or WPS before using this script.", vbCritical + vbOKOnly, "Cannot Convert Files"
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

WordApp.Visible = false ' Run in the background

Sum = 0
Dim FilePath, FileLine
Set List = fso.OpenTextFile("ConvertFileList.txt", 1, true)
Do While List.AtEndOfLine <> True
    FileLine = List.ReadLine
    If FileLine <> "" And Mid(FileLine, 1, 2) <> "~$" Then
        Sum = Sum + 1 ' Count the number of files in the list
    End If
Loop
List.Close
MsgBox "Conversion is starting. If Word windows pop up during the process," & vbCrLf & "DO NOT CLOSE THEM! Just minimize them." & vbCrLf & "DO NOT CLOSE THEM! Just minimize them." & vbCrLf & "DO NOT CLOSE THEM! Just minimize them." & vbCrLf & "I said it three times because it's important!", vbOKOnly + vbExclamation, "Warning"

Dim Finished
Finished = 0
Set List = fso.OpenTextFile("ConvertFileList.txt", 1, true)
Do While List.AtEndOfLine <> True
    FilePath = List.ReadLine
    If Mid(FilePath, 1, 2) <> "~$" Then ' Ignore temporary Word files
        Set objDoc = WordApp.Documents.Open(FilePath)
        ' WordApp.Visible = false ' (Commented out due to issues with macro-heavy documents)
        If WordApp.Visible = true Then
            WordApp.ActiveDocument.ActiveWindow.WindowState = 2 ' Minimize window
        End If
        objDoc.SaveAs Left(FilePath, InstrRev(FilePath, ".")) & "pdf", wdFormatPDF ' Save as PDF
        LogOut("Converted file: " & FilePath & " (" & Finished & "/" & Sum & ")")
        WordApp.ActiveDocument.Close
        Finished = Finished + 1
    End If
    If IsChooseDelete = vbYes Then
        fso.DeleteFile FilePath
        LogOut("Deleted file: " & FilePath)
    End If
Loop
' Cleanup
List.Close
LogOut("Conversion complete")
LogFile.Close
' Uncomment these lines if you want to automatically delete ConvertFileList.txt and log.txt after completion
' fso.DeleteFile "ConvertFileList.txt"
' fso.DeleteFile "log.txt"

Dim Msg
Msg = "Successfully converted " & Finished & " files"
If IsChooseDelete = vbYes Then
    Msg = Msg & " and deleted the original files."
End If
MsgBox Msg & vbCrLf & "Log file saved at: " & fso.GetFolder(Path).Path & "\log.txt"
Set fso = Nothing
WordApp.Quit
WScript.Quit
