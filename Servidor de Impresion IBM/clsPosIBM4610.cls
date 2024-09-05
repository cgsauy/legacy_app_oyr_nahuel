VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPosIBM4610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PrinterIndex
    Journal = 1
    Customer = 2
End Enum

Public Enum VaciarMemoria
    logo = 1
    mensaje = 2
    flash = 5
End Enum

Public Function AlmacenarLogo(ByVal nroLogo As Byte, ByVal width As Byte, ByVal heigh As Byte, ByVal logo As String) As String
    AlmacenarLogo = Chr$(&H1D) & Chr$(&H2A) & GetTxtHexadecimal(nroLogo) & GetTxtHexadecimal(width) & GetTxtHexadecimal(heigh) & logo
End Function

Public Function ImprimirLogoAlmacenado(ByVal nroLogo As Byte, Optional densidad As Byte = 0) As String
    ImprimirLogoAlmacenado = Chr$(&H1D) & Chr$(&H2F) & GetTxtHexadecimal(densidad) & GetTxtHexadecimal(nroLogo)
End Function

Public Function ImprimirLogo(ByVal width As Byte, ByVal heigh As Byte, ByVal logo As String, Optional densidad As Byte = 0) As String
    ImprimirLogo = Chr$(&H1B) & Chr$(&H2A) & GetTxtHexadecimal(densidad) & GetTxtHexadecimal(width) & GetTxtHexadecimal(heigh) & logo
End Function

Public Function Key_Continue()
    Key_Continue = Chr$(&H1B) & Chr$(&H2E)
End Function

Public Function LimpiarMemoria(ByVal que As VaciarMemoria)
    LimpiarMemoria = Chr$(&H1B) & Chr$(&H23) & GetTxtHexadecimal(que)
End Function

Public Function Status() As String
    Status = Chr$(&H1B) & Chr$(&H76)
End Function

Public Function key_SetHighQualityOn() As String
    key_SetHighQualityOn = Chr$(&H1B) & Chr$(&H2F) & Chr$(&H1)
End Function

Public Function key_SetHighQualityOFF() As String
    key_SetHighQualityOFF = Chr$(&H1B) & Chr$(&H2F) & Chr$(&H0)
End Function

Public Function key_AlignLeft() As String
    key_AlignLeft = Chr$(&H1B) & Chr$(&H61) & Chr$(&H0)
End Function
Public Function key_AlignCenter() As String
    key_AlignCenter = Chr$(&H1B) & Chr$(&H61) & Chr$(&H1)
End Function
Public Function key_AlignRight() As String
    key_AlignRight = Chr$(&H1B) & Chr$(&H61) & Chr$(&H2)
End Function
Public Function key_AlignColumnRight() As String
    key_AlignColumnRight = Chr$(&H1B) & Chr$(&H61) & Chr$(&H4)
End Function

Public Function Key_SetTAB() As String
    Key_SetTAB = Chr$(&H1B) & Chr$(&H44)
End Function

Public Function Key_SetTABEND() As String
    Key_SetTABEND = Chr$(&H0) & Chr$(&H0)
End Function

Public Function key_TAB() As String
    key_TAB = Chr$(&H9)
End Function

Public Function key_LF() As String
    key_LF = Chr$(&HD)
End Function

Public Function key_CutPaper() As String
    key_CutPaper = Chr(&HC)
End Function

Public Function key_SetFontA() As String
    key_SetFontA = key_InitFont & Chr$(&H0)
End Function

Public Function key_SetFontB() As String
    key_SetFontB = key_InitFont & Chr$(&H1)
End Function

Private Function key_InitFont() As String
    key_InitFont = Chr$(&H1B) & Chr$(&H21)
End Function

Public Function key_SetFontADoubleHigh() As String
    key_SetFontADoubleHigh = key_InitFont & Chr$(&H10)
End Function
Public Function key_SetFontBDoubleHigh() As String
    key_SetFontBDoubleHigh = key_InitFont & Chr$(&H11)
End Function

Public Function key_SetFontADoubleWide() As String
    key_SetFontADoubleWide = key_InitFont & Chr$(&H20)
End Function
Public Function key_SetFontBDoubleWide() As String
    key_SetFontBDoubleWide = key_InitFont & Chr$(&H21)
End Function

Public Function key_SetFontADoubleWideDoubleHigh() As String
    key_SetFontADoubleWideDoubleHigh = key_InitFont & Chr$(&H30)
End Function
Public Function key_SetFontBDoubleWideDoubleHigh() As String
    key_SetFontBDoubleWideDoubleHigh = key_InitFont & Chr$(&H31)
End Function

Private Function key_InitFontSize() As String
    key_InitFontSize = Chr$(&H1B) & Chr$(&H20)
End Function

Public Function key_SetFontSize12() As String
    key_SetFontSize12 = key_InitFontSize & Chr$(&H5)
End Function
Public Function key_SetFontSize15() As String
    key_SetFontSize15 = key_InitFontSize & Chr$(&H3)
End Function
Public Function key_SetFontSize17() As String
    key_SetFontSize17 = key_InitFontSize & Chr$(&H2)
End Function

Public Function key_InitBARCODE() As String
    key_InitBARCODE = Chr$(&H1D) & Chr$(&H6B)
End Function

Public Function key_EndBARCODE() As String
    key_EndBARCODE = Chr$(&H0)
End Function

Public Function key_FONTCB_UPCA() As String
    key_FONTCB_UPCA = Chr$(&H0)
End Function
Public Function key_FONTCB_UPCE() As String
    key_FONTCB_UPCE = Chr$(&H1)
End Function
Public Function key_FONTCB_JAN13() As String
    key_FONTCB_JAN13 = Chr$(&H2)
End Function
Public Function key_FONTCB_JAN8() As String
    key_FONTCB_JAN8 = Chr$(&H3)
End Function
Public Function key_FONTCB_CODE39() As String
    key_FONTCB_CODE39 = Chr$(&H4)
End Function
Public Function key_FONTCB_ITF() As String
    key_FONTCB_ITF = Chr$(&H5)
End Function

Public Function key_FONTCB_CODABAR() As String
    key_FONTCB_CODABAR = Chr$(&H6)
End Function
Public Function key_FONTCB_CODE128() As String
    key_FONTCB_CODE128 = Chr$(&H7)
End Function
Public Function key_FONTCB_CODE93() As String
    key_FONTCB_CODE93 = Chr$(&H8)
End Function

Public Function key_SetearHeighCodeBar(ByVal valor As Integer) As String
    key_SetearHeighCodeBar = Chr$(&H1D) & Chr$(&H68) & GetTxtHexadecimal(valor)
End Function

Public Function key_SetearWidthCodeBar(ByVal valor As Byte) As String
    key_SetearWidthCodeBar = Chr$(&H1D) & Chr$(&H77) & GetTxtHexadecimal(valor)
End Function

Public Function Key_TextInvertInit() As String
    Key_TextInvertInit = Chr$(&H1B) & Chr$(&H48) & Chr$(&H1)
End Function

Public Function Key_TextInvertEnd() As String
    Key_TextInvertEnd = Chr$(&H1B) & Chr$(&H48) & Chr$(&H0)
End Function

Public Function Key_TextEnphasizedON() As String
    Key_TextEnphasizedON = Chr$(&H1B) & Chr$(&H47) & Chr$(&H1)
End Function
Public Function Key_TextEnphasizedOFF() As String
    Key_TextEnphasizedOFF = Chr$(&H1B) & Chr$(&H47) & Chr$(&H0)
End Function

Public Function SetPrinter(ByVal bandeja As PrinterIndex) As String
    SetPrinter = Chr$(&H1B) & Chr$(&H63) & Chr$(&H30)
    Select Case bandeja
        Case PrinterIndex.Journal
            SetPrinter = SetPrinter & Chr$(&H1)
        Case Else
            SetPrinter = SetPrinter & Chr$(&H2)
    End Select
End Function

Public Function GetTxtHexadecimal(ByVal numero As Integer) As String
Dim sHex As String
'Dim ibytes As Byte
    sHex = Hex(numero)
    'ibytes = Len(sHex) \ 2
    Do While sHex <> ""
        GetTxtHexadecimal = Chr$(Val("&H" & VBA.Right(sHex, 2))) & GetTxtHexadecimal
        If Len(sHex) < 3 Then
            sHex = ""
        Else
            sHex = Left(sHex, Len(sHex) - 2)
        End If
    Loop
    
End Function


