VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clslPrintConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tPrintDoc
    TipoDoc As Integer
    Impresora As String
    Bandeja As Integer
    EsXDef As Boolean
    Papel As Integer
End Type
Private arrPrinterDoc() As tPrintDoc

Dim objPrint As clsCnfgPrintDocument

Private Sub Class_Initialize()
    Set objPrint = New clsCnfgPrintDocument
    Set objPrint.Connect = cBase
    ReDim arrPrinterDoc(0)
End Sub

Private Sub Class_Terminate()
    Set objPrint = Nothing
    Erase arrPrinterDoc
End Sub

Public Function ShowPrinterSetup(ByVal sTipoDocs As String, ByVal lTerminal As Long) As Boolean
On Error GoTo errSPS
    ShowPrinterSetup = True
    If sTipoDocs = "" Or lTerminal <= 0 Then Exit Function
    With objPrint
        .Terminal = lTerminal
        .CnfgTipoDocumento = sTipoDocs
        .ShowConfig
    End With
    Exit Function
errSPS:
    ShowPrinterSetup = False
    MsgBox "Error al mostrar el formulario de configuración de impresión." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "Atención"
End Function

Public Function LoadPrinterConfig(ByVal sTipoDocs As String, ByVal lTerminal As Long) As Boolean
'Parámetros sTipoDocs = "1,2" --> Ctdo y Credito.
'                  lTerminal = Código de Terminal.
On Error GoTo errGPC
Dim vDocs() As String, vPrint() As String
Dim sPrint As String
Dim iCont As Integer

    LoadPrinterConfig = True
    If sTipoDocs = "" Or lTerminal <= 0 Then Exit Function
    
    vDocs = Split(sTipoDocs, ",")
    
    With objPrint
        .Terminal = lTerminal
        For iCont = 0 To UBound(vDocs)
            
            sPrint = .getDocumentoImpresora(vDocs(iCont))
            
            ReDim Preserve arrPrinterDoc(iCont)
            
            arrPrinterDoc(iCont).TipoDoc = vDocs(iCont)
            If sPrint <> "" Then
                vPrint = Split(sPrint, "|")
                arrPrinterDoc(iCont).Impresora = Trim(vPrint(0))
                arrPrinterDoc(iCont).Papel = vPrint(2)
                arrPrinterDoc(iCont).Bandeja = Trim(vPrint(1))
                arrPrinterDoc(iCont).EsXDef = (Val(vPrint(3)) = 1)
            Else
                arrPrinterDoc(iCont).Impresora = ""
                arrPrinterDoc(iCont).Papel = 1
                arrPrinterDoc(iCont).Bandeja = 0
                arrPrinterDoc(iCont).EsXDef = False
            End If
            If arrPrinterDoc(iCont).Papel = 0 Then arrPrinterDoc(iCont).Papel = 1
        
        Next
    End With
    Exit Function
    
errGPC:
    MsgBox "Error al leer la configuración de impresión para su terminal." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "Atención"
    LoadPrinterConfig = False
End Function

Public Function GetPrinterDoc(ByVal iTipoDoc As Integer, ByRef sPrint As String, ByRef iBandeja As Integer, ByRef bEsXDef As Boolean, Optional ByRef iPapel As Integer) As Boolean
Dim iCont As Integer
On Error GoTo errGPD
    GetPrinterDoc = False
    sPrint = ""
    bEsXDef = False
    iBandeja = 1        'pongo una x def.
    For iCont = 0 To UBound(arrPrinterDoc)
        With arrPrinterDoc(iCont)
            If .TipoDoc = iTipoDoc Then
                sPrint = .Impresora
                iPapel = .Papel
                iBandeja = .Bandeja
                bEsXDef = .EsXDef
                Exit For
            End If
        End With
    Next
    GetPrinterDoc = (sPrint <> "")
    Exit Function
errGPD:
    MsgBox "Error al leer la configuración de impresión para el documento de tipo & " & iTipoDoc & "." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "Atención"
    GetPrinterDoc = False
End Function
