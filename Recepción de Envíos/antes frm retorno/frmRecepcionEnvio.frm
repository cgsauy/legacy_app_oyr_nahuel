VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmRecepcionEnvio 
   BackColor       =   &H00CEBAB3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Envíos"
   ClientHeight    =   5205
   ClientLeft      =   4020
   ClientTop       =   2265
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecepcionEnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5415
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEBAB3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5175
      TabIndex        =   19
      Top             =   840
      Width           =   5175
      Begin VB.OptionButton opGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEBAB3&
         Caption         =   "I&ngresar todo el resto como entregado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   21
         Top             =   0
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton opGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEBAB3&
         Caption         =   "&Individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   20
         Top             =   300
         Width           =   3975
      End
   End
   Begin VB.CheckBox chSendMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEBAB3&
      Caption         =   "Enviar mensaje"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1920
      TabIndex        =   12
      Top             =   3150
      Width           =   1815
   End
   Begin VB.TextBox tMotivo 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmRecepcionEnvio.frx":0442
      Top             =   3390
      Width           =   4815
   End
   Begin VB.OptionButton opEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEBAB3&
      Caption         =   "En&tregó"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   2190
      Width           =   1575
   End
   Begin VB.OptionButton opEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEBAB3&
      Caption         =   "&Nueva Fecha"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   2550
      Width           =   1455
   End
   Begin VB.OptionButton opEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEBAB3&
      Caption         =   "&A Confirmar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   2910
      Width           =   1575
   End
   Begin VB.TextBox tEnvio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1470
      Width           =   855
   End
   Begin VB.TextBox tFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   2670
      Width           =   1095
   End
   Begin VB.ComboBox cHora 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2670
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0448
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0874
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "save"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Envío"
            Key             =   "envio"
            Object.ToolTipText     =   "Ir a formulario de envíos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dividir"
            Key             =   "dividir"
            Object.ToolTipText     =   "Dividir un envío en dos"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   8
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin prjHiperLink.orHiperLink hlVaCon 
      Height          =   255
      Left            =   2040
      Top             =   1440
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   450
      BackColor       =   13548211
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   16711680
      Caption         =   "6"
      MouseIcon       =   "frmRecepcionEnvio.frx":097E
      MousePointer    =   99
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjHiperLink.orHiperLink hlDesvVaCon 
      Height          =   255
      Left            =   3120
      Top             =   1440
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
      BackColor       =   13548211
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   16711680
      Caption         =   "Desvincular Va Con"
      MouseIcon       =   "frmRecepcionEnvio.frx":0C98
      MousePointer    =   99
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Motivo:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   3150
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Envío:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1470
      Width           =   495
   End
   Begin VB.Label lDireccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hora"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   2670
      Width           =   375
   End
   Begin VB.Label lVaCon 
      BackStyle       =   0  'Transparent
      Caption         =   "Va Con"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1470
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label lHelp 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label lCamion 
      BackColor       =   &H007280FA&
      BackStyle       =   0  'Transparent
      Caption         =   "Camión: Martín"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape shfac 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      FillColor       =   &H00DCFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   5220
   End
   Begin VB.Menu MnuVaCon 
      Caption         =   "Va Con"
      Visible         =   0   'False
      Begin VB.Menu MnuVaConItem 
         Caption         =   "Item"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmRecepcionEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10/1/2007      Deshabilito entrega parcial si no se entregó ningún artículo.
Option Explicit

Dim lDirEnvio As Long, lAgeEnvio As Long
Dim bCamionTieneMerc As Boolean

Private Type tDatosFlete
    Agenda As Double
    AgendaAbierta As Double
    AgendaCierre As Date
    HorarioRango As Integer
    HoraEnvio As String
End Type

Private rDatosFlete As tDatosFlete

Private Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Private Sub act_DividirEnvio()
On Error GoTo errDE
    'Dim oFrm As New frmDividoEnvio
    Dim oFrm As New frmMercaAReclamar
    oFrm.prmInvocacion = 1
    oFrm.prmEnvio = Val(tEnvio.Tag)
    
    oFrm.Show vbModal
    Set oFrm = Nothing
    Exit Sub
errDE:
    objGral.OcurrioError "Error al acceder al formulario para dividir el envío.", Err.Description, "Dividir Envíos"
End Sub

Private Sub db_IndependizarVaCon()
On Error GoTo errIVC
Dim rsEnvio As rdoResultset
Dim lOld As Long, iQ As Integer

    Screen.MousePointer = 11
    Cons = "Select * From Envio  Where EnvCodigo = " & Val(tEnvio.Text)
    Set rsEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Cuento Q de envíos en todo el va con.
    Cons = "Select count(*) from Envio Where Abs(EnvVaCon) = " & Abs(rsEnvio("EnvVaCon"))
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    iQ = RsAux(0)
    RsAux.Close
    
    'Si está casado con más de uno le tengo que decir que tiene que desvincular los otros ya que si no me quedán todos
    'desvinculados.
    If iQ > 2 And rsEnvio("EnvVaCon") < 0 Then
        RsAux.Close
        rsEnvio.Close
        Screen.MousePointer = 0
        MsgBox "Este envío posee más de un envío en el va con y este es el que une a todos. No puede desvincular este envío ingrese el código de otro de los envíos del va con.", vbCritical, "Atención"
        Exit Sub
    End If
    
    FechaDelServidor
    
    lOld = rsEnvio!EnvVaCon
    rsEnvio.Edit
    rsEnvio!EnvVaCon = Null
    rsEnvio!EnvFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsEnvio.Update
    rsEnvio.Close

'Quito al otro de este va con
    If iQ = 2 Then
        cBase.Execute "Update Envio Set EnvVaCon = Null, EnvFModificacion = '" & Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss") & "' Where abs(EnvVaCon) = " & Abs(lOld)
    End If
    
    Screen.MousePointer = 0
    Exit Sub
errIVC:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al intentar independizar.", Err.Description, "Eliminar va con"
End Sub


Private Function fnc_ValidoNoEntregado() As Boolean
On Error GoTo errVNE
Dim rs As rdoResultset
Dim rsE As rdoResultset

    fnc_ValidoNoEntregado = False
    
    Cons = "Select VTeDocumento, EnvReclamoCobro, EnvDocumento, EnvCodigo, EnvFormaPago from Envio " & _
                " Left Outer Join VentaTelefonica On VTeDocumento = EnvDocumento" & _
            " Where ((EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Null) " & _
            "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Not Null)))"
    
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsE.EOF
        
        'OPCION NUEVA
        '   DOCUMENTOPENDIENTE
        
        Cons = "Select * From DocumentoPendiente Where DPeTipo = 1 And DPeIDTipo = " & rsE("EnvCodigo")
        Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rs.EOF Then
            rs.Close
            'Opción VIEJA
            
            'Veo si es Venta Telefónica
            If Not IsNull(rsE("VTeDocumento")) Then
                
                'Es una venta telefónica.
                If Not IsNull(rsE("EnvReclamoCobro")) Then
                    
                    Cons = "Select * From Envio" _
                        & " Where EnvCodigo <> " & rsE("EnvCodigo") _
                        & " And EnvDocumento = " & rsE("EnvDocumento") _
                        & " And EnvEstado IN (3, 4)"  ' EstadoEnvio.Impreso  & ", " & EstadoEnvio.Entregado
                    
                    Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
                    If Not rs.EOF Then
                        MsgBox "Este envío es el que cobra una venta telefónica y tiene pendiente el cobro." _
                            & " Verifique pues ya  hay envíos entregados o impresos para esta venta." & Chr(13) & "Proceda por el formato anterior.", vbExclamation, "ATENCIÓN VENTAS TELEFONICAS"
                        rs.Close
                        rsE.Close
                        Exit Function
                    Else
                        rs.Close
                    End If
                    
                    'Ahora veo si hay + de un envío en la venta telefónica.
                    Cons = "Select * From Envio Where EnvCodigo <> " & rsE("EnvCodigo") & " And EnvDocumento = " & rsE("EnvDocumento")
                    Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not rs.EOF Then
                        'No es el único.
                        MsgBox "Este envío es el que cobra la venta telefónica." & vbCrLf & " Primero debe darle un destino a los otros envíos.", vbInformation, "Ventas telefónicas."
                    Else
                        MsgBox "Este envío es una venta telefónica, debe proceder por el formato anterior.", vbCritical, "FORMATO ANTERIOR"
                    End If
                    rs.Close
                    rsE.Close
                    Exit Function
                Else
                    'Es seguro que hay otro con el valor de cobranza.
                    If rsE!EnvFormaPago = 2 Then      'TipoPagoEnvio.PagaDomicilio
                        
                        MsgBox "El envío " & rsE("EnvCodigo") & " es parte de una venta telefónica, para anular el documento que cobro el flete se debe acceder por el envío.", vbInformation, "Ventas telefónicas"
                    End If
                    
                End If
            Else
                'NO es vta telefónica.
                If rsE!EnvFormaPago = 2 Then      'TipoPagoEnvio.PagaDomicilio
                    'Esta factura la hizo el deposito.
                    Cons = "Select * From Documento Where DocCodigo = " & RsAux!EnvDocumentoFactura
                    Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly)
                    If Not rs.EOF Then
                        rsE.Close
                        rs.Close
                        MsgBox "Este envío tiene un documento pendiente, debe proceder por el formato anterior.", vbCritical, "FORMATO ANTERIOR"
                        Exit Function
                    End If
                    rs.Close
                End If
            End If
        Else
            rs.Close
        End If
        rsE.MoveNext
    Loop
    rsE.Close
    fnc_ValidoNoEntregado = True
    Exit Function

errVNE:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al validar.", Err.Description

End Function

Private Sub db_StockFisicoEnLocal(lArt As Long, iEstado As Integer, cCant As Currency, iTipoDoc As Integer, lDoc As Long, ByVal iTipoLocal As Byte, ByVal iLocal As Integer)
    'stockMovFisicoEnLocal @iUser smallint, @iArticulo int, @iCantidad smallint, @iEstado smallint, @iTipoLocal tinyint, @iLocal int, @iTipoDocumento smallint = Null,  @iDocumento Int = Null, @iTerminal smallint = Null
    Cons = "EXEC stockMovFisicoEnLocal " & paCodigoDeUsuario & ", " & lArt & ", " & cCant & ", " & iEstado & ", " & iTipoLocal & ", " & iLocal & ", " & iTipoDoc & ", " & lDoc & ", " & paCodigoDeTerminal
    cBase.Execute Cons
End Sub

Private Function db_RestoARenglonEntrega(ByVal lArt As Long, ByVal cCant As Currency, ByVal bEsEntregado As Boolean) As Boolean
Dim RsRE As rdoResultset

    db_RestoARenglonEntrega = True
    Cons = "Select * From RenglonEntrega Where ReECodImpresion = " & Val(tCodigo.Tag) _
        & " And ReEArticulo = " & lArt & " And ReECantidadTotal > 0"
    Set RsRE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Si no hay que de error.
    If RsRE!ReECantidadEntregada > cCant Then
        RsRE.Edit
        RsRE!ReECantidadTotal = RsRE!ReECantidadTotal - cCant
        RsRE!ReECantidadEntregada = RsRE!ReECantidadEntregada - cCant
        RsRE.Update
    ElseIf RsRE!ReECantidadEntregada = cCant Then
        If RsRE!ReECantidadTotal = RsRE!ReECantidadEntregada Then
            RsRE.Delete
        Else
            RsRE.Edit
            RsRE!ReECantidadTotal = RsRE!ReECantidadTotal - cCant
            RsRE!ReECantidadEntregada = RsRE!ReECantidadEntregada - cCant
            RsRE.Update
        End If
    Else
        If bEsEntregado Then
'ocasiono error ya que no tengo esa mercadería en el registro.
            RsRE.Close
            RsRE.Edit
        Else
            'Si no tiene mercadería dejo.
            If RsRE!ReECantidadEntregada = 0 Then
                db_RestoARenglonEntrega = False
                'Si la cantidad total = a la de este envío --> no tengo más envíos con este artículo si no resto la total.
                If RsRE!ReECantidadTotal = cCant Then
                    RsRE.Delete
                Else
                    RsRE.Edit
                    RsRE!ReECantidadTotal = RsRE!ReECantidadTotal - cCant
                    RsRE.Update
                End If
            Else
                'No puede ocurrir, pero x las dudas.
                RsRE.Close
                RsRE.Edit
            End If
        End If
    End If
    RsRE.Close
End Function

Private Sub db_SetEntregado(ByVal lEnvio As Long, ByVal iTipo As Integer, sErr As String)
Dim iTD As Integer
Dim lDoc As Long
    sErr = ""
    
     If iTipo <> 2 Then 'TipoEnvio.Service
        Cons = "Select DocTipo, DocCodigo From Envio, Documento" _
            & " Where EnvCodigo = " & lEnvio _
            & " And EnvDocumento = DocCodigo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        iTD = RsAux!DocTipo
        lDoc = RsAux!DocCodigo
    Else
        Cons = "Select EnvDocumento From Envio Where EnvCodigo = " & lEnvio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        lDoc = RsAux!EnvDocumento
    End If
    RsAux.Close
    
    Cons = "Select * From RenglonEnvio Where REvEnvio = " & lEnvio _
        & " And REvArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        
        sErr = "Error al restar en la tabla RenglonEntrega. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
        db_RestoARenglonEntrega RsAux!REvArticulo, RsAux!REvAEntregar, True
        
        sErr = "Error al actualizar StockLocal. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
        'Le quito la mercadería entregada al camión
        db_StockFisicoEnLocal RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar * -1, iTD, lDoc, 1, lCamion.Tag
        
        'Doy la baja al movimiento de estados.
        Cons = "EXEC StockMovEstadoStockTotal " & paCodigoDeUsuario & ", " & RsAux!REvArticulo & ", " & RsAux!REvAEntregar * -1 & ", " & TipoMovimientoEstado.AEntregar & ", " & iTD & ", " & lDoc & ", " & paCodigoDeSucursal
        cBase.Execute Cons
        
        
        sErr = "Error al actualizar Renglón Envío. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
        RsAux.Edit
        RsAux!REvAEntregar = 0
        RsAux.Update
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Pongo los artículos de servicio como entregados.--------------------------------------
    sErr = "Error al actualizar los artículos de servicio para el envío: " & lEnvio
    Cons = "Update RenglonEnvio Set REvAEntregar = 0" _
        & " Where REvEnvio = " & lEnvio _
        & " And REvArticulo IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
    cBase.Execute (Cons)
    '-----------------------------------------------------------------------------------------------------
    sErr = ""

End Sub

Private Sub db_SetDevuelve(ByVal lEnvio As Long, ByVal iTipo As Integer, sErr As String)
Dim iTD As Integer
Dim lDoc As Long
    
    sErr = ""
    
     If iTipo <> 2 Then 'TipoEnvio.Service
        Cons = "Select DocTipo, DocCodigo From Envio, Documento" _
            & " Where EnvCodigo = " & lEnvio _
            & " And EnvDocumento = DocCodigo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        iTD = RsAux!DocTipo
        lDoc = RsAux!DocCodigo
    Else
        Cons = "Select EnvDocumento From Envio Where EnvCodigo = " & lEnvio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        lDoc = RsAux!EnvDocumento
    End If
    RsAux.Close
    
    Cons = "Select * From RenglonEnvio Where REvEnvio = " & lEnvio _
        & " And REvArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        
        sErr = "Error al restar en la tabla RenglonEntrega. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
        If db_RestoARenglonEntrega(RsAux!REvArticulo, RsAux!REvAEntregar, False) Then
            If bCamionTieneMerc Then
                'Si entro es xq no tenía mercadería para este artículo.
                sErr = "Error al actualizar StockLocal. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
                db_StockFisicoEnLocal RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar * -1, iTD, lDoc, 1, lCamion.Tag
            'LE doy al local
                db_StockFisicoEnLocal RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar, iTD, lDoc, 2, paCodigoDeSucursal
            End If
        End If
        
'No va más pero lo dejo hasta que lo viejo no exista más.
        RsAux.Edit
        RsAux!REvCodImpresion = Null
        RsAux.Update
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    sErr = ""

End Sub

Private Sub db_SaveEntregado()
Dim rsE As rdoResultset
Dim sError As String

    On Error GoTo errBT
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errRB
    Cons = "Select * From Envio " & _
        " Where ((EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Null) " & _
        "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Not Null)))"

    Cons = Cons & " And EnvEstado = 3" 'IMPRESO
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsE.EOF
        db_SetEntregado rsE!EnvCodigo, rsE!EnvTipo, sError
    
        '....................................EDIT ENVIO
        rsE.Edit
        rsE!EnvEstado = 4   'EstadoEnvio.Entregado
        rsE!EnvFechaEntregado = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        rsE!EnvFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        rsE.Update
        '....................................EDIT ENVIO
        rsE.MoveNext
    Loop
    rsE.Close

    cBase.CommitTrans

    On Error Resume Next
    tEnvio.Text = ""
    tEnvio.SetFocus
    Exit Sub
    
errBT:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Exit Sub
    
Relajo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al pasar todo como entregado.", sError & vbCr & Err.Description
    Exit Sub
    
errRB:
    Resume Relajo
    Exit Sub
    
End Sub

Private Sub loc_EnvioConDocumentoPendiente(ByVal iEnvio As Long)
On Error GoTo errEDP
Dim rsD As rdoResultset
Dim sQy As String
    sQy = "Select DocSerie, DocNumero From DocumentoPendiente, Documento " & _
         " Where DPeTipo = 1 and DPeIDTipo IN (" & _
                     "Select EnvCodigo From Envio " & _
                    " Where ((EnvCodigo = " & iEnvio & " And EnvVaCon Is Null) " & _
                    "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Not Null)))" & _
                    " And EnvEstado = 3)" & _
         " And DPeDocumento = DocCodigo"
    
    Set rsD = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    sQy = ""
    Do While Not rsD.EOF
        sQy = sQy & IIf(sQy = "", "", vbCrLf) & rsD("DocSerie") & " " & rsD("DocNumero")
        rsD.MoveNext
    Loop
    rsD.Close
    If sQy <> "" Then
        MsgBox "Atención el envío seleccionado posee los siguientes documentos pendientes asociados a él y debe reclamarselos al camionero si los tiene: " & vbCrLf & sQy, vbInformation, "Facturas del envío"
    End If
    Exit Sub
errEDP:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar si el documento posee documentos pendientes.", Err.Description, "Documentos pendientes"
End Sub

Private Sub db_SaveAImprimirAConfirmar(ByVal bEsAImprimir As Boolean)
Dim rsE As rdoResultset
Dim arrCli() As Long
Dim arrDocE() As Long
Dim sError As String
Dim iIndex As Integer, iQ As Integer

    'Aca tengo que controlar el documento pendiente
    'Si es con el formato anterior doy msg y me voy.
    If Not fnc_ValidoNoEntregado() Then Exit Sub
    
    
    'Doy aviso si el envío tiene una factura pendiente.
    loc_EnvioConDocumentoPendiente Val(tEnvio.Tag)
    
    ReDim arrDocE(0)
    ReDim arrCli(0)
    iIndex = 0
    On Error GoTo errBT
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errRB
    
    Cons = "Select * From Envio " & _
            " Where ((EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Null) " & _
            "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Not Null)))" & _
            " And EnvEstado = 3"
    
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsE.EOF
        'Mercadería entregada.
        db_SetDevuelve rsE!EnvCodigo, rsE!EnvTipo, sError
        
        If Not bEsAImprimir And paTComEnvConf <> 0 Then
            If arrDocE(0) > 0 Then
                iIndex = iIndex + 1
                ReDim Preserve arrDocE(iIndex)
                ReDim Preserve arrCli(iIndex)
            End If
            arrDocE(iIndex) = rsE!EnvDocumento
            arrCli(iIndex) = rsE!EnvCliente
        End If
        
        '....................................EDIT ENVIO
        rsE.Edit
        rsE!EnvEstado = IIf(bEsAImprimir, 0, 1)
    'Saco el ID de imrpesión
        rsE!EnvCodImpresion = Null
        If IsDate(tFecha.Text) Then
            rsE!EnvFechaPrometida = Format(tFecha.Text, "mm/dd/yyyy")
            
            If Trim(cHora.Text) <> vbNullString Then
                If cHora.ListIndex > -1 Then
                    'Busco en codigotexto el valor.
                    If cHora.ItemData(cHora.ListIndex) > 0 Then
                        Cons = "Select * from CodigoTexto Where Codigo = " & cHora.ItemData(cHora.ListIndex)
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If Len(Trim(RsAux!Clase)) < 4 Then
                            rsE!EnvRangoHora = "0" & Trim(RsAux!Clase) & "-" & Trim(RsAux!Puntaje)
                        Else
                            rsE!EnvRangoHora = Trim(RsAux!Clase) & "-" & Trim(RsAux!Puntaje)
                        End If
                        RsAux.Close
                    Else
                        rsE!EnvRangoHora = cHora.Text
                    End If
                Else
                    rsE!EnvRangoHora = Trim(cHora.Text)
                End If
            Else
                rsE!EnvRangoHora = Null
            End If
        End If
        
        rsE!EnvFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        rsE.Update
        '....................................EDIT ENVIO
        rsE.MoveNext
    Loop
    rsE.Close
    
    'Grabo Comentario asociado al documento.
    If Not bEsAImprimir And paTComEnvConf <> 0 Then
        Cons = "Select * From Comentario Where ComCodigo = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        For iQ = 0 To iIndex
            RsAux.AddNew
            RsAux!ComCliente = arrCli(iQ)
            RsAux!ComDocumento = arrDocE(iQ)
            RsAux!ComFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
            RsAux!ComTipo = paTComEnvConf
            RsAux!ComUsuario = paCodigoDeUsuario
            RsAux!ComComentario = tMotivo.Text
            RsAux.Update
        Next
        RsAux.Close
    End If
    
    cBase.CommitTrans

    On Error Resume Next
    If Not bEsAImprimir And chSendMsg.Value Then
        Dim objConnect As New clsConexion
        objConnect.EnviaMensaje paUIDEnvConf, "Envío(s) a Confirmar", tMotivo.Text, DateAdd("s", 30, Now), 751, paCodigoDeUsuario
        Set objConnect = Nothing
    End If

    tEnvio.Text = ""
    tEnvio.SetFocus
    Exit Sub
    
errBT:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Exit Sub
    
Relajo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al pasar todo como entregado.", sError & vbCr & Err.Description
    Exit Sub
    
errRB:
    Resume Relajo
    Exit Sub
    
End Sub

Private Sub act_Envio()
On Error GoTo errE
    Dim objEnvio As New clsEnvio
    objEnvio.InvocoEnvio tEnvio.Tag, App.Path & "\Reportes"
    Set objEnvio = Nothing
    
    Me.Refresh
    'Por si lo modifico o le quito el va con vuelvo a cargar.
    tEnvio.Tag = ""
    s_FindEnvio
    If Val(tEnvio.Tag) = 0 Then s_SetCtrlIndividual
    tooMenu.Buttons("envio").Enabled = (Val(tEnvio.Tag) > 0)
    
    Exit Sub
errE:
    objGral.OcurrioError "Error al acceder al envío.", Err.Description, "Envío"
End Sub

Private Sub loc_SaveTodoEntregado()
On Error GoTo errSTE
    Screen.MousePointer = 11
    Cons = "EXEC repartoDarEntregadoEnvio " & Val(tCodigo.Tag) & ", " & paTipoArticuloServicio & ", " & paCodigoDeUsuario & ", " & paCodigoDeSucursal & ", " & paCodigoDeTerminal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = -1 Then
        MsgBox "Error al dar los envíos como entregados, refresque el código de impresión y reintente.", vbExclamation, "Atención"
    End If
    RsAux.Close
    tCodigo.Text = ""
    tCodigo.SetFocus
    Screen.MousePointer = 0
    Exit Sub
errSTE:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al dar todo por entregado.", Err.Description, "Grabar todo entregado"
End Sub
Private Sub act_Save()

    If opGrabar(0).Value Then
        If MsgBox("¿Confirmar pasar el resto de los envíos como entregados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        
        'válido que no hayan fichas asignadas sin cumplir para algún envío.
        Cons = "Select * From Devolucion " _
            & " Where DevEnvio IN (Select EnvCodigo From Envio Where EnvCodImpresion = " & Val(tCodigo.Tag) & ")" _
            & " And DevLocal Is Null And DevAnulada Is Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            MsgBox "Existen envíos asignados a Fichas de Alta de Stock." & vbCrLf & "Debe cumplir las mismas para poder cumplir los envíos.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        loc_SaveTodoEntregado
    Else
    
        If opEstado(2).Value Then
            Cons = "Select * From Devolucion " _
                & " Where DevEnvio = " & Val(tEnvio.Text) & " And DevLocal Is Null And DevAnulada Is Null"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                RsAux.Close
                MsgBox "Existen envíos asignados a Fichas de Alta de Stock." & vbCrLf & "Debe cumplir las mismas para poder cumplir los envíos.", vbExclamation, "ATENCIÓN"
                Screen.MousePointer = 0
                Exit Sub
            End If
            RsAux.Close
            
            If fnc_MercaderiaParcial Then
                MsgBox "El camión tiene asignada mercadería parcialmente, si oucrre un error al grabar debe darle toda la mercadería al camión.", vbExclamation, "ATENCIÓN"
            End If
            
            If MsgBox("¿Confirma dar por entregado al envío?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
            
            db_SaveEntregado
            
        Else
            'Para nueva fecha es obligatoria.
        
            If tFecha.Text <> "" Then
                If Not IsDate(tFecha.Text) Then
                    MsgBox "Ingrese la nueva fecha para el envío.", vbExclamation, "Atención"
                    tFecha.SetFocus
                    Exit Sub
                ElseIf CDate(tFecha.Text) < Date Then
                    MsgBox "La fecha es menor a hoy.", vbExclamation, "Atención"
                    tFecha.SetFocus
                    Exit Sub
                End If
                If Not f_EsDiaAbierto Then
                    If MsgBox("El día no está abierto." & vbCr & vbCr & "¿Confirma guardar el envío con esa fecha?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then Exit Sub
                End If
                If Not ValidoRangoHorario Then Exit Sub
            Else
                If opEstado(1).Value Then
                    MsgBox "Debe indicar la nueva fecha de envío.", vbExclamation, "Atención"
                    tFecha.SetFocus
                    Exit Sub
                End If
            End If
            
            
            If fnc_MercaderiaParcial Then
                MsgBox "El camión tiene asignada mercadería parcialmente, debe darle toda la mercadería al camión o pasar los envíos entregados primero.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
                                
            If opEstado(0).Value Then
            
                If Trim(tMotivo.Text) = "" Then
                    MsgBox "Ingrese el motivo por el cual pone el envío a confirmar.", vbExclamation, "Atención"
                    tMotivo.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Confirma poner el envío a Confirmar?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
                db_SaveAImprimirAConfirmar False
            Else
                If MsgBox("¿Confirma poner el envío a imprimir?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
                db_SaveAImprimirAConfirmar True
            End If
        End If
    End If
    
End Sub

Private Function fnc_MercaderiaParcial() As Boolean
    fnc_MercaderiaParcial = False
    'Si no tiene toda la mercadería el camión no lo dejo hacer
    'ya que no se si el camión tenía mercadería o no.
    Cons = "Select IsNull(Sum(ReECantidadTotal), 0) as QTotal, IsNull(Sum(ReECantidadEntregada), 0) as QCamion " & _
                "From RenglonEntrega Where ReECodImpresion = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!QTotal <> RsAux!QCamion And RsAux!QCamion > 0 Then
        RsAux.Close
        fnc_MercaderiaParcial = True
        Exit Function
    End If
    RsAux.Close
End Function

Private Sub loc_CleanMenuVaCon()
On Error GoTo errCMVC
Dim iQ As Integer
    
    MnuVaConItem(0).Caption = ""
    MnuVaConItem(0).Tag = ""
    For iQ = MnuVaConItem.LBound + 1 To MnuVaConItem.UBound
        Unload MnuVaConItem(iQ)
    Next
Exit Sub
errCMVC:
    objGral.OcurrioError "Error al limpiar el menú va con.", Err.Description
End Sub

Private Sub s_FindEnvio()
On Error GoTo errFE
Dim lAux As Long
Dim sVaCon As String

    Screen.MousePointer = 11
    lDirEnvio = 0
    lAgeEnvio = 0
    loc_CleanMenuVaCon
    
    Cons = "Select EnvCodigo, EnvDireccion, EnvTipoFlete, EnvAgencia,  IsNull(EnvVaCon, 0) as VaCon " & _
                " From Envio " & _
                " Where ((EnvCodigo = " & Val(tEnvio.Text) & " And EnvVaCon Is Null) " & _
                        "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & Val(tEnvio.Text) & " And EnvVaCon Is Not Null)))" & _
                " And EnvCodImpresion = " & Val(tCodigo.Text)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un envío pendiente con ese código que pertenezca al código de impresión.", vbExclamation, "Atención"
        Exit Sub
    Else
        tEnvio.Tag = Val(tEnvio.Text)
        Do While Not RsAux.EOF
            
            If sVaCon <> "" Then sVaCon = sVaCon & ","
            sVaCon = sVaCon & RsAux("EnvCodigo")
            
            If RsAux("EnvCodigo") = Val(tEnvio.Tag) Then
                lDireccion.Caption = objGral.ArmoDireccionEnTexto(cBase, RsAux("EnvDireccion"), , True, , True)
                lDirEnvio = RsAux("EnvDireccion")
                tFecha.Tag = RsAux!EnvTipoFlete
                If Not IsNull(RsAux!EnvAgencia) Then lAgeEnvio = RsAux!EnvAgencia
            Else
                If RsAux("VaCon") <> 0 And RsAux("EnvCodigo") <> Val(tEnvio.Tag) Then
                    If Val(MnuVaConItem(0).Tag) > 0 Then Load MnuVaConItem(MnuVaConItem.UBound + 1)
                    With MnuVaConItem(MnuVaConItem.UBound)
                        .Visible = True
                        .Enabled = True
                        .Caption = Trim(RsAux("EnvCodigo"))
                        .Tag = .Caption
                    End With
                End If
            End If
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    
    If MnuVaConItem(0).Tag <> "" Then
        With hlVaCon
            .Enabled = True
            .ForeColor = &HC00000
        End With
        hlDesvVaCon.Visible = True
        lVaCon.Enabled = True
    End If
    Screen.MousePointer = 0
    Exit Sub
errFE:
    Screen.MousePointer = 0
    tEnvio.Tag = ""
    objGral.OcurrioError "Error al buscar el envío.", Err.Description
End Sub
Private Sub s_GetDatosReparto()
On Error GoTo errGDR
Dim QTotal As Integer, QCamion As Integer
    
    'Busco los datos de la tabla repartoimpresión.
    Screen.MousePointer = 11
    bCamionTieneMerc = False
    
    Cons = "Select IsNull(Sum(ReECantidadTotal), 0) as QTotal, IsNull(Sum(ReECantidadEntregada), 0) as QCamion,  CamCodigo, CamNombre " & _
            " From RenglonEntrega, Camion Where ReECodImpresion = " & tCodigo.Text & _
            " And ReECamion = CamCodigo" & _
            " Group by CamCodigo, CamNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lCamion.Caption = "Camión: " & Trim(RsAux!CamNombre)
        lCamion.Tag = RsAux!CamCodigo
        QTotal = RsAux!QTotal
        QCamion = RsAux!QCamion
        bCamionTieneMerc = (QCamion > 0)
        RsAux.Close
        tCodigo.Tag = Trim(tCodigo.Text)
    Else
        RsAux.Close
        'Tengo que ver si tengo envíos que tengan sólo artículos de servicio.
        Cons = "Select EnvCodigo, CamCodigo, CamNombre From Envio, Camion " & _
                    "Where EnvCodImpresion = " & Val(tCodigo.Text) & _
                    " And EnvEstado = 3 And EnvTipo = 2 And EnvCamion = CamCodigo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsAux.EOF Then
            
            With lCamion
                .Caption = "Camión: " & Trim(RsAux!CamNombre)
                .Tag = RsAux!CamCodigo
            End With
            tCodigo.Tag = Trim(tCodigo.Text)
            bCamionTieneMerc = True
            With opGrabar(0)
                .Enabled = True: .Value = True
            End With
            opGrabar(1).Enabled = True
            With tooMenu
                .Buttons("save").Enabled = True
            End With
            RsAux.Close
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No se encontraron datos para el código ingresado.", vbExclamation, "Buscar"
        Exit Sub
    End If
    
    If QTotal > QCamion Then
        MsgBox "Al camión no se le entregó " & IIf(QCamion > 0, "la totalidad de ", "") & "la mercadería, no se podrá dar todo como entregado.", vbExclamation, "Atención"
    Else
        If QTotal = 0 Then MsgBox "No hay mercadería a entregar.", vbExclamation, "Atención"
    End If

    'Pasar todo como ingresado.
    opGrabar(0).Enabled = (QTotal = QCamion And QCamion > 0)
    opGrabar(0).Value = (QTotal = QCamion And QCamion > 0)
    
    opGrabar(1).Enabled = QTotal > 0
    opGrabar(1).Value = (Not opGrabar(0).Value And QTotal > 0)

    With tooMenu
        .Buttons("save").Enabled = opGrabar(1).Enabled And opGrabar(0).Value
    End With
    Screen.MousePointer = 0
    Exit Sub
    
errGDR:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el código de impresión.", Err.Description
End Sub

Private Sub s_SetCtrlEstado()
    
    With tFecha
        .Enabled = (opEstado(0).Value Or opEstado(1).Value) And Val(tEnvio.Tag) > 0
        If Not .Enabled Then .Text = ""
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    
    With cHora
        .Enabled = tFecha.Enabled
        .BackColor = tFecha.BackColor
        If Not .Enabled Then .Text = ""
    End With
    
    With tMotivo
        .Enabled = opEstado(0).Value And Val(tEnvio.Tag) > 0
        If Not .Enabled Then .Text = ""
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With chSendMsg
        .Enabled = opEstado(0).Value And Val(tEnvio.Tag) > 0
        If Not .Enabled Then .Value = 0 Else .Value = 1
    End With
    
End Sub

Private Sub s_SetCtrlIndividual()
Dim lColor As Long
    
    tFecha.Tag = 0          'Guardo el ID de Tipo de Flete
    
    If opGrabar(1).Enabled And opGrabar(1).Value Then lColor = vbWindowBackground Else lColor = vbButtonFace
        
    With tEnvio
        .Enabled = opGrabar(1).Value
        .BackColor = lColor
        If Not .Enabled Then .Text = ""
    End With
    
    opEstado(0).Enabled = opGrabar(1).Value
    opEstado(1).Enabled = opGrabar(1).Value
    opEstado(2).Enabled = opGrabar(1).Value And bCamionTieneMerc
    
    opEstado(0).Value = False
    opEstado(1).Value = False
    opEstado(2).Value = False
    
    lVaCon.Enabled = False
    With hlVaCon
        .Enabled = False
        .ForeColor = vbButtonFace
    End With
    hlDesvVaCon.Visible = False
    lDireccion.Caption = ""
    tooMenu.Buttons("save").Enabled = opGrabar(0).Value Or (opGrabar(1).Enabled And opGrabar(1).Value And Val(tEnvio.Tag) > 0)
    tooMenu.Buttons("envio").Enabled = (opGrabar(1).Enabled And opGrabar(1).Value And Val(tEnvio.Tag) > 0)
    
End Sub

Private Sub s_CtrlClean()
    bCamionTieneMerc = False
    With tooMenu
        .Buttons("save").Enabled = False
    End With
    
    opGrabar(0).Value = False
    opGrabar(1).Value = False
    opGrabar(0).Enabled = False
    opGrabar(1).Enabled = False
    
    lCamion.Caption = ""
    lHelp.Caption = ""
    s_SetCtrlIndividual
    s_SetCtrlEstado
    
    tEnvio.Text = ""
    tMotivo.Text = ""
    
End Sub

Private Sub cHora_GotFocus()
On Error Resume Next
    With cHora
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text) Else .SelLength = 0
    End With
End Sub

Private Sub cHora_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Val(tEnvio.Tag) > 0 Then
        If Trim(cHora.Text) <> "" Then If Not ValidoRangoHorario Then Exit Sub
        If opEstado(0).Value Then
            If tMotivo.Enabled Then tMotivo.SetFocus
        Else
            act_Save
        End If
    End If
    
End Sub

Private Sub Form_Load()
    s_CtrlClean
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set objGral = Nothing
    cBase.Close
    Set cBase = Nothing
    End
End Sub

Private Sub hlDesvVaCon_Click()
On Error GoTo errDV
    If MsgBox("¿Confirma quitar del va con al envío seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar va con") = vbYes Then
        db_IndependizarVaCon
    End If
Exit Sub
errDV:
    objGral.OcurrioError "Error al intentar desvincular el envío.", Err.Description
End Sub

Private Sub hlVaCon_Click()
    PopupMenu MnuVaCon, , hlVaCon.Left, hlVaCon.Top + hlVaCon.Height
End Sub

Private Sub Label1_Click()
    With tCodigo
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
On Error Resume Next
    cHora.SetFocus
End Sub

Private Sub Label3_Click()
    With tEnvio
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label4_Click()
    With tMotivo
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub lVaCon_Click()
    hlVaCon_Click
End Sub

Private Sub opEstado_Click(Index As Integer)
    s_SetCtrlEstado
    tooMenu.Buttons("save").Enabled = (Val(tEnvio.Tag) > 0)
End Sub

Private Sub opEstado_GotFocus(Index As Integer)
    Select Case Index
        Case 2: lHelp.Caption = "El envío se entrego normalmente."
        Case 1: lHelp.Caption = "El envío se imprimirá nuevamente en la fecha que indique, se mantiene el camionero."
        Case 0: lHelp.Caption = "El envío quedará sin fecha de entrega y en estado a confirmar."
    End Select
End Sub

Private Sub opEstado_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Select Case Index
            Case 0: ' A Confirmar
                tFecha.SetFocus
            Case 2: ' Entrego
                If Val(tEnvio.Tag) > 0 Then act_Save
            Case 1: ' A imprimir (nueva fecha)
                tFecha.SetFocus
        End Select
    End If
End Sub

Private Sub opEstado_LostFocus(Index As Integer)
    lHelp.Caption = ""
End Sub

Private Sub opGrabar_Click(Index As Integer)
    s_SetCtrlIndividual
End Sub

Private Sub opGrabar_GotFocus(Index As Integer)
    Select Case Index
        Case 0: lHelp.Caption = "Se pasan todos los envíos del código de impresión como entregados."
        Case 1: lHelp.Caption = "Permite seleccionar un envío del código de impresión."
    End Select
End Sub

Private Sub opGrabar_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            act_Save
        Else
            On Error Resume Next
            tEnvio.SetFocus
        End If
    End If
End Sub

Private Sub opGrabar_LostFocus(Index As Integer)
    lHelp.Caption = ""
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then tCodigo.Tag = "": s_CtrlClean
    lHelp.Caption = "Ingrese el código de impresión a recepcionar."
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    lHelp.Caption = "Ingrese el código de impresión a recepcionar."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tCodigo.Tag) > 0 Then
            If opGrabar(0).Enabled Then
                opGrabar(0).SetFocus
            ElseIf opGrabar(1).Enabled Then
                opGrabar(1).SetFocus
            End If
        Else
            s_GetDatosReparto
        End If
    End If
End Sub

Private Sub tCodigo_LostFocus()
    lHelp.Caption = ""
End Sub

Private Sub tEnvio_Change()
    If tEnvio.Tag <> "" Then tEnvio.Tag = "": s_SetCtrlIndividual: s_SetCtrlEstado
End Sub

Private Sub tEnvio_GotFocus()
    With tEnvio
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    lHelp.Caption = "Ingrese el código de envío."
End Sub

Private Sub tEnvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tEnvio.Tag) > 0 Then
            If opEstado(2).Enabled Then
                opEstado(2).SetFocus
            ElseIf opEstado(1).Enabled Then
                opEstado(1).SetFocus
            End If
        Else
            If Not IsNumeric(tEnvio.Text) Then
                MsgBox "No es un código válido.", vbExclamation, "Atención"
            Else
                'Busco el envío.
                s_FindEnvio
                s_SetCtrlEstado
                tooMenu.Buttons("envio").Enabled = (Val(tEnvio.Tag) > 0)
            End If
        End If
    End If
End Sub

Private Sub tEnvio_LostFocus()
    lHelp.Caption = ""
End Sub

Private Sub tFecha_Change()
    cHora.Clear
End Sub

Private Sub tFecha_GotFocus()
    With tFecha
        If .Text = "" Then
            s_GetDatosTipoFlete
            s_SetFirstDay
            s_SetHoraEntrega
        End If
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    lHelp.Caption = "Ingrese la fecha en que se volverá a enviar."
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            If Not f_EsDiaAbierto Then
                MsgBox "El día ingresado no está abierto.", vbExclamation, "Atención"
            Else
                s_SetHoraEntrega
                cHora.SetFocus
            End If
        End If
    End If
End Sub

Private Sub tMotivo_GotFocus()
    lHelp.Caption = "Ingrese el motivo por el cual el envío queda a confirmar, se graba un comentario para el cliente."
End Sub

Private Sub tMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tMotivo.Text) <> "" Then act_Save
End Sub

Private Sub tMotivo_LostFocus()
    lHelp.Caption = ""
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "save": act_Save
        Case "envio": act_Envio
        Case "dividir": act_DividirEnvio
    End Select
End Sub

Private Sub s_GetDatosTipoFlete()
Dim RsF As rdoResultset
Dim lZona As Long

    With rDatosFlete
        .Agenda = 0
        .AgendaAbierta = 0
        .AgendaCierre = Date
        .HoraEnvio = ""
        .HorarioRango = 0
    End With
    
    'Ya lo cargue o no hay tipo de flete
    If Val(tFecha.Tag) = 0 Or rDatosFlete.Agenda > 0 Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select IsNull(TFlAgenda, 0) as Agenda, IsNull(TFlAgendaHabilitada, 0) as AgendaH, IsNull(TFLFechaAgeHab, GetDate()) as FAgenda, TFLHoraEnvio, IsNull(THoRangoHS, 0) as RangoHS " & _
                " From TipoFlete " & _
                        "Left Outer Join TipoHorario On TFlRangoHs = THoID" & _
                " Where TFLCodigo = " & Val(tFecha.Tag)
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsF.EOF Then
        With rDatosFlete
            .Agenda = RsF("Agenda")
            .AgendaAbierta = RsF("AgendaH")
            .AgendaCierre = RsF("FAgenda")
            If Not IsNull(RsF("TFLHoraEnvio")) Then .HoraEnvio = Trim(RsF!TFLHoraEnvio)
            .HorarioRango = RsF("RangoHS")
        End With
    End If
    RsF.Close
    
    'Si no es de Agencia --> busco para la zona.
    If lAgeEnvio > 0 Then
        'Tengo que buscar la zona de la agencia.
        Cons = "Select IsNull(CZoZona, 0) From Agencia, Direccion, CalleZona" _
                & " Where AgeCodigo = " & lAgeEnvio _
                & " And AgeDireccion = DirCodigo And DirCalle = CZoCalle " _
                & " And CZoDesde <= DirPuerta And CZoHasta >= DirPuerta"

        Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsF.EOF Then
            lZona = RsF(0)
        End If
        RsF.Close
    End If

    'Si no tengo zona de agencia busco para la dirección del envío.
    If lZona = 0 Then lZona = db_FindZona(lDirEnvio)
        
    Cons = "Select * From FleteAgendaZona " & _
            " Where FAZZona = " & lZona & " And FAZTipoFlete = " & Val(tFecha.Tag)
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsF.EOF Then
        With rDatosFlete
            If Not IsNull(RsF!FAZAgenda) Then
                .Agenda = RsF!FAZAgenda
                If Not IsNull(RsF!FAZAgendaHabilitada) Then .AgendaAbierta = RsF!FAZAgendaHabilitada Else .AgendaAbierta = .Agenda
                If Not IsNull(RsF!FAZFechaAgeHab) Then .AgendaCierre = RsF("FAZFechaAgeHab")
            End If
            If Not IsNull(RsF!FAZRangoHS) Then .HorarioRango = RsF!FAZRangoHS
            If Not IsNull(RsF!FAZHoraEnvio) Then .HoraEnvio = Trim(RsF!FAZHoraEnvio)
        End With
    End If
    RsF.Close
    Screen.MousePointer = 0
    
End Sub

Private Function db_FindZona(lCodDireccion As Long) As Long
On Error GoTo errFZ
Dim lZonP As Long
Dim lIDComp As Long

    Cons = "Select IsNull(CZoZona,0) as CZoZona, IsNull(DirComplejo,0) as DirComplejo From Direccion " _
            & " Left Outer Join CalleZona On DirCalle = CZoCalle And CZoDesde <= DirPuerta And CZoHasta >= DirPuerta" _
        & " Where DirCodigo = " & lCodDireccion
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        lZonP = 0
        lIDComp = 0
    Else
        lIDComp = RsAux!DirComplejo
        lZonP = RsAux!CZoZona
    End If
    RsAux.Close
    
    If lIDComp > 0 Then
        'Si tengo complejo --> busco la zona para el mismo.
        Cons = "Select CZoZona From Complejo, CalleZona" _
            & " Where ComCodigo = " & lIDComp _
            & " And CZoCalle = ComCalle And CZoDesde <= ComNumero And CZoHasta >= ComNumero"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            lZonP = RsAux!CZoZona
        End If
        RsAux.Close
    End If
    db_FindZona = lZonP
    Exit Function

errFZ:
    objGral.OcurrioError "Error al buscar el código de la zona.", Err.Description
End Function

Private Sub s_SetFirstDay()
Dim sMat As String
Dim iSuma As Integer
Dim dAux As Date
    
    If rDatosFlete.AgendaCierre < Date Then dAux = Date Else dAux = rDatosFlete.AgendaCierre
    
    If DateDiff("d", rDatosFlete.AgendaCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(rDatosFlete.Agenda)
    Else
        sMat = superp_MatrizSuperposicion(rDatosFlete.AgendaAbierta)
    End If
    
    If sMat <> "" Then
        iSuma = BuscoProximoDia(dAux, sMat)
        If iSuma <> -1 Then tFecha.Text = Format(DateAdd("d", iSuma, dAux), "dd/mm/yyyy")
    End If
    If tFecha.Text = "" Then
        MsgBox "No hay agenda abierta para el tipo de flete del envío.", vbExclamation, "Atención"
        tFecha.Text = Date
    End If
    
End Sub

Private Function BuscoProximoDia(dFecha As Date, strMat As String)
Dim rsHora As rdoResultset
Dim intDia As Integer, intSuma As Integer
    
    'Por las dudas que no cumpla en la semana paso la agenda normal.
    
    On Error GoTo errBDER
    
    BuscoProximoDia = -1
    
    'Consulto en base a la matriz devuelta.
    Cons = "Select * From HorarioFlete Where HFlIndice IN (" & strMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsHora.EOF Then
        
        'Busco el valor que coincida con el dia de hoy y ahí busco para arriba.
        intSuma = 0
        Do While intSuma < 7
            rsHora.MoveFirst
            intDia = Weekday(dFecha + intSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = intDia Then
                    BuscoProximoDia = intSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            intSuma = intSuma + 1
        Loop
        rsHora.Close
    End If

Encontre:
    rsHora.Close
    Exit Function
    
errBDER:
    objGral.OcurrioError "Error al buscar el primer día disponible para el tipo de flete.", Trim(Err.Description)
End Function

Private Sub s_SetHoraEntrega()
On Error GoTo errCHEPD
Dim sMat As String
    Screen.MousePointer = 11
    cHora.Clear
    If DateDiff("d", rDatosFlete.AgendaCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(rDatosFlete.Agenda)
    Else
        sMat = superp_MatrizSuperposicion(rDatosFlete.AgendaAbierta)
    End If
    If rDatosFlete.HoraEnvio <> "" Then
        loc_SetHoraEnvio rDatosFlete.HoraEnvio, sMat
    Else
        If sMat <> "" Then
            Cons = "Select HFlCodigo, HFlNombre From HorarioFlete Where HFlIndice IN (" & sMat & ")" _
                & " And HFlDiaSemana = " & Weekday(CDate(tFecha.Text)) & " Order by HFlInicio"
            CargoCombo Cons, cHora
        End If
    End If
    If cHora.ListCount > 0 Then cHora.ListIndex = 0
    Screen.MousePointer = 0
    Exit Sub
errCHEPD:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar los horarios para el día de semana.", Trim(Err.Description)
End Sub
Private Sub loc_SetHoraEnvio(ByVal sHora As String, ByVal sMat As String)
Dim arrHoraE() As String, arrID() As String
Dim iQ As Integer
On Error Resume Next
Dim rsHF As rdoResultset
Dim sIn As String

    arrHoraE = Split(sHora, ",")
    
    Cons = "Select HEnIndice From HorarioFlete, HoraEnvio Where HFlIndice IN (" & sMat & ")" _
            & " And HFlDiaSemana = " & Weekday(CDate(tFecha.Text)) _
            & " And HEnCodigo = HFlCodigo  Order by HFlInicio"
    Set rsHF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsHF.EOF
        If sIn <> "" Then sIn = sIn & ","
        sIn = sIn & Trim(rsHF("HEnIndice"))
        rsHF.MoveNext
    Loop
    rsHF.Close
    If sIn <> "" Then sIn = "," & sIn & ","
    
    For iQ = 0 To UBound(arrHoraE)
        arrID = Split(arrHoraE(iQ), ":")
        If InStr(1, sIn, "," & arrID(0) & ",") > 0 Then cHora.AddItem arrID(1)
    Next
End Sub

Private Function f_EsDiaAbierto() As Boolean
Dim sMat As String
Dim dAux As Date

    f_EsDiaAbierto = False
    If Val(tFecha.Tag) = 0 Then Exit Function
    
    If DateDiff("d", rDatosFlete.AgendaCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(rDatosFlete.Agenda)
    Else
        sMat = superp_MatrizSuperposicion(rDatosFlete.AgendaAbierta)
    End If
    dAux = CDate(tFecha.Text)
    If sMat <> "" Then f_EsDiaAbierto = (BuscoProximoDia(dAux, sMat) = 0)
    
End Function

Private Function ValidoRangoHorario() As Boolean

    ValidoRangoHorario = True
    If cHora.ListIndex > -1 Then Exit Function
    
    If InStr(1, cHora.Text, "-") > 0 Then
        Select Case Len(cHora.Text)
            Case 9
                If CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1)) > CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    cHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                End If
                
            Case 5
                If InStr(1, cHora.Text, "-") = 1 Then
                    If CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) < paPrimeraHoraEnvio Then
                        MsgBox "El horario ingresado es menor a la primera hora de entrega.", vbExclamation, "ATENCIÓN"
                        ValidoRangoHorario = False
                        Exit Function
                    Else
                        If paPrimeraHoraEnvio < 1000 Then
                            cHora.Text = "0" & paPrimeraHoraEnvio & cHora.Text
                        Else
                            cHora.Text = paPrimeraHoraEnvio & cHora.Text
                        End If
                        Exit Function
                    End If
                Else
                    If InStr(1, cHora.Text, "-") = 5 Then
                        If CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1)) > paUltimaHoraEnvio Then
                            MsgBox "El horario ingresado es mayor que la última hora de envio.", vbExclamation, "ATENCIÓN"
                            ValidoRangoHorario = False
                            Exit Function
                        Else
                            cHora.Text = cHora.Text & paUltimaHoraEnvio
                        End If
                    Else
                        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                        cHora.SetFocus
                        ValidoRangoHorario = False
                        Exit Function
                    End If
                End If
            
            Case 8
                If CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1)) > CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    cHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                End If
                
                If InStr(1, cHora.Text, "-") = 4 Then
                    cHora.Text = "0" & cHora.Text
                End If
            
            Case Else
                    MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                    cHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                    
        End Select
    Else
        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
        cHora.SetFocus
        ValidoRangoHorario = False
        Exit Function
    End If
    
    'Ahora válido el rango de horas.
    If Val(tFecha.Tag) > 0 And rDatosFlete.HorarioRango > 0 Then
        
        Dim lhora As Long
        
        lhora = (CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) - CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1))) / 100
        If lhora < rDatosFlete.HorarioRango Then
            If MsgBox("El rango ingresado es menor al posible para el flete seleccionado." & vbCr & vbCr & _
                        "El flete tiene un rango de " & rDatosFlete.HorarioRango & " hora(s) y se asigno un rango de " & lhora & " hora(s)" & vbCr & vbCr & _
                        "¿Confirma mantener el rango ingresado?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible error en horario") = vbNo Then
                cHora.SetFocus
                ValidoRangoHorario = False
            End If
        End If
    End If
    
End Function


