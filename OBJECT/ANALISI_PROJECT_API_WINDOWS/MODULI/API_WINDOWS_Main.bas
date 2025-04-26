Attribute VB_Name = "API_WINDOWS_Main"


Option Compare Database
Option Explicit

Const GB As Currency = 107374.1824@
Const MB As Currency = 104.8576@
Const KB As Currency = 0.1024@

Public Sub Main()
    Dim oFile As New clsFile
    oFile.OpenFile "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\Project_PROGETTI_MDB.sublime-project"

    oFile.SeekFile64bit 6 * GB, so_Begin

End Sub
