Attribute VB_Name = "Module1"
Option Explicit

Public Sub LoadDocPrint(ByVal sFile As String)
On Error GoTo errLP
Dim sData As String
Dim oFile As New clsorCGSA
    ofile.GetDataFile2(sfile, sData)
    
End Sub
