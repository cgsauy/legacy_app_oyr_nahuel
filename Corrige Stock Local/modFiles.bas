Attribute VB_Name = "modFiles"
Option Explicit

Public Function fileAddArticulo(idArticulo As Long)
On Error GoTo errAdd
Dim mTextFile As String
Dim bErr As Boolean
    
    mTextFile = CreateTextFromFile(prmFileName, bErr)
    If bErr Then Exit Function
    
    If Trim(mTextFile) <> "" Then
        mTextFile = mTextFile & CStr(idArticulo)
    Else
        mTextFile = CStr(idArticulo)
    End If
    GrabarArchivo (mTextFile)

errAdd:
End Function

Public Function CreateTextFromFile(ByVal sFileName As String, retError As Boolean) As String
    
    On Error GoTo ErrorHandler
    Err.Clear
    retError = False
    
    Dim bExiste As Boolean
    bExiste = (Dir(sFileName) <> "")
    If Not bExiste Then Exit Function
    
    Dim oFileSys As Object, oFileObj As Object
    Dim sData As String
    
    CreateTextFromFile = vbNullString
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Set oFileObj = oFileSys.opentextfile(sFileName, 1, False, 0)
    sData = oFileObj.Readall()
    oFileObj.Close
    CreateTextFromFile = sData

ErrorHandler:
        Set oFileObj = Nothing
        Set oFileSys = Nothing
        If Err.Number > 0 Then
            CreateTextFromFile = vbNullString
            retError = True
        End If
End Function

Private Function GrabarArchivo(myText As String)
On Error GoTo errGrabar

Dim iFile As Integer
    
    iFile = FreeFile
    Open prmFileName For Output As iFile
    
    Print #iFile, myText

    Close iFile

errGrabar:
End Function
