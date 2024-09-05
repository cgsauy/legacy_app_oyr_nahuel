Attribute VB_Name = "modFiles"
Option Explicit

Public clsGeneral As New clsorCGSA

Const mFileName = "Copiar_Archivos.ini"

Public Function fnc_LeoArchivo() As String

Dim mData As String
Dim mFile As String, mRet As Long

    mFile = App.Path & "\" & mFileName
    
    mRet = clsGeneral.GetDataFile2(mFile, mData)
    If mRet <> 0 Then mData = ""
    
    fnc_LeoArchivo = mData
    
End Function


Public Function fnc_GrabarArchivo(mTXTFile As String)
Dim mFile As String

    mFile = App.Path & "\" & mFileName
    Open mFile For Output As #1
    Print #1, mTXTFile
    Close #1
    
End Function

