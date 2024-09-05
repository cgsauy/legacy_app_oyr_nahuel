VERSION 5.00
Begin VB.Form frmImpresoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Impresoras"
   ClientHeight    =   1440
   ClientLeft      =   3300
   ClientTop       =   2940
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpresoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4950
   Begin VB.CommandButton bGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   4020
      TabIndex        =   4
      Top             =   1020
      Width           =   855
   End
   Begin VB.ComboBox cPrinter 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2955
   End
   Begin VB.OptionButton oConformes 
      Caption         =   "Conformes"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   1335
   End
   Begin VB.OptionButton oCaja 
      Caption         =   "Caja"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione la Nueva impresora."
      Height          =   255
      Left            =   1980
      TabIndex        =   3
      Top             =   180
      Width           =   2415
   End
End
Attribute VB_Name = "frmImpresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargoImpresoras()

    'Cargo las impresoras definidas en el sistema
    Dim x As Printer
    For Each x In Printers
        cPrinter.AddItem Trim(x.DeviceName)
    Next
    
    '----------------------------------------------------
End Sub

Private Sub bGrabar_Click()

Dim aNombre As String, bCaja As Boolean

    If Not oCaja.Value And Not oConformes.Value Then
        MsgBox "Seleccione el tipo de impresora a cambiar.", vbExclamation, "Faltan Datos"
        Exit Sub
    End If
    
    If cPrinter.ListIndex = -1 Then
        MsgBox "Seleccione el nombre de la nueva impresora.", vbExclamation, "Faltan Datos"
        Exit Sub
    End If
    
    If oCaja.Value = True Then
        aNombre = "CAJA": bCaja = True
    End If
    If oConformes.Value = True Then
        aNombre = "CONFORMES": bCaja = False
    End If
    
    If MsgBox("Confirma cambiar la impresora de " & aNombre & " por la impresora " & Trim(cPrinter.Text), vbQuestion + vbYesNo, "Cambiar Impresora") = vbNo Then Exit Sub
    
    
    If AccionGrabar(Trim(cPrinter.Text), bCaja) Then
        MsgBox "La impresora se cambio con éxito." & vbCrLf & _
                    "Recurde cerrar los sistemas que utilicen las impresoras para que esto funcione." & vbCrLf & _
                    "ej: Caja, Mostrador, Factura Contado, etc.", vbInformation, "CAMBIO OK "
    Else
        MsgBox "Las impresoras no se cambiaron", vbExclamation, "No Hubo Cambio"
    End If
    
End Sub

Private Sub Form_Load()

    CargoImpresoras
    
    
    Screen.MousePointer = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set cBase = Nothing
    Set eBase = Nothing
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Function AccionGrabar(sImp As String, bCaja As Boolean) As Boolean

    AccionGrabar = True
    cons = "Select * from Local Where LocCodigo = " & paLocalColonia
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        rsAux.Edit
        
        If bCaja Then
            rsAux!LocICoNombre = Trim(sImp)
            rsAux!LocICrNombre = Trim(sImp)
            rsAux!LocINDNombre = Trim(sImp)
            rsAux!LocINCNombre = Trim(sImp)
            rsAux!LocINENombre = Trim(sImp)
            rsAux!LocIRENombre = Trim(sImp)
        Else
            rsAux!LocIRMNombre = Trim(sImp)
            rsAux!LocICNNombre = Trim(sImp)
        End If
        
        rsAux.Update
    Else
        MsgBox "No hay datos para el local COLONIA.", vbExclamation, "No hay datos"
        AccionGrabar = False
    End If
    rsAux.Close
    
End Function
