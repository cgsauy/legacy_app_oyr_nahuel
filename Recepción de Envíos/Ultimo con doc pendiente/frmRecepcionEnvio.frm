VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmRecepcionEnvio 
   BackColor       =   &H00B3DEF5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Envíos"
   ClientHeight    =   6390
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
   ScaleHeight     =   6390
   ScaleWidth      =   5415
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0554
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
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
         NumButtons      =   3
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
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F1FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   5175
      Begin prjHiperLink.orHiperLink hlVaCon 
         Height          =   255
         Left            =   1920
         Top             =   120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   450
         BackColor       =   13234687
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
         MouseIcon       =   "frmRecepcionEnvio.frx":086E
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
      Begin VB.ComboBox cHora 
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox tEnvio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   8
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "&A Confirmar"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "&Nueva Fecha"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "En&tregó"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "Entregó &Parcial"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox tMotivo 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmRecepcionEnvio.frx":0B88
         Top             =   2040
         Width           =   4815
      End
      Begin VB.CheckBox chSendMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "Enviar mensaje"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsEntParcial 
         Height          =   975
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1720
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15794175
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483639
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin prjHiperLink.orHiperLink hlDesvVaCon 
         Height          =   255
         Left            =   3000
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   13234687
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
         MouseIcon       =   "frmRecepcionEnvio.frx":0B8E
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   450
         Width           =   705
      End
      Begin VB.Label lVaCon 
         BackStyle       =   0  'Transparent
         Caption         =   "Va Con"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hora"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lHelp 
         BackColor       =   &H00008000&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   4200
         Width           =   4935
      End
      Begin VB.Label lDireccion 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   450
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Envío:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.OptionButton opGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F1FF&
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
   Begin VB.OptionButton opGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F1FF&
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   5175
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
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   1800
      TabIndex        =   17
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

Private Sub db_StockCamion(lArt As Long, iEstado As Integer, cCant As Currency, iTipoDoc As Integer, lDoc As Long)
Dim RsLocal As rdoResultset

    Cons = "Select * From StockLocal " _
        & " Where StlArticulo = " & lArt & " And StlTipolocal = 1" _
        & " And StlLocal = " & Val(lCamion.Tag) & " And StlEstado = " & iEstado
        
    Set RsLocal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsLocal!StLCantidad - cCant = 0 Then
        RsLocal.Delete
    Else
        RsLocal.Edit
        RsLocal!StLCantidad = RsLocal!StLCantidad - cCant
        RsLocal.Update
    End If
    RsLocal.Close

    MarcoMovimientoStockFisico paCodigoDeUsuario, 1, lCamion.Tag, lArt, cCant, iEstado, -1, iTipoDoc, lDoc
    
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
            RsRE.Close
            RsRE.Edit
        Else
            'Si no tiene mercadería dejo.
            If RsRE!ReECantidadEntregada = 0 Then
                db_RestoARenglonEntrega = False
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
        db_StockCamion RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar, iTD, lDoc
        
        'Doy la baja al movimiento de estados.
        MarcoMovimientoStockEstado paCodigoDeUsuario, RsAux!REvArticulo, RsAux!REvAEntregar, TipoMovimientoEstado.AEntregar, -1, iTD, lDoc, paCodigoDeSucursal
        MarcoMovimientoStockTotal RsAux!REvArticulo, TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.AEntregar, RsAux!REvAEntregar, -1
        
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

Private Sub db_SetEntregaParcial(ByVal lEnvio As Long, ByVal iTipo As Integer, sErr As String)
Dim iTD As Integer
Dim lDoc As Long
Dim iQ As Integer
Dim rsR As rdoResultset
Dim bArtServ As Boolean
    
    sErr = ""
    Cons = "Select DocTipo, DocCodigo From Envio, Documento" _
        & " Where EnvCodigo = " & lEnvio & " And EnvDocumento = DocCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    iTD = RsAux!DocTipo
    lDoc = RsAux!DocCodigo
    RsAux.Close
    
    'Recorro la lista.
    With vsEntParcial
        For iQ = 0 To .Rows - 1
        'Los artículos que no entrego se los doy al documento.
            bArtServ = False
            Cons = "Select ArtTipo From Articulo Where ArtID = " & .Cell(flexcpData, iQ, 0)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            bArtServ = (RsAux(0) = paTipoArticuloServicio)
            RsAux.Close
                   
            Cons = "Select * From RenglonEnvio Where REvEnvio = " & lEnvio _
                & " And REvArticulo = " & .Cell(flexcpData, iQ, 0)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
            If Not bArtServ Then
                'De RenglonEntrega RESTO TODO
                sErr = "Error al restar en la tabla RenglonEntrega. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
                db_RestoARenglonEntrega RsAux!REvArticulo, RsAux!REvAEntregar, True
    
                'LE QUITO TODA LA MERCADERIA AL CAMION
                sErr = "Error al actualizar StockLocal. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
                db_StockCamion RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar, iTD, lDoc
            End If
            
            If .Cell(flexcpValue, iQ, 0) <> .Cell(flexcpData, iQ, 1) Then
                
                If Not bArtServ Then
                    'Le tengo que dar el resto al local.
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, RsAux!REvArticulo, RsAux!REvAEntregar - .Cell(flexcpValue, iQ, 0), paEstadoArticuloEntrega, 1
                    MarcoMovimientoStockFisico paCodigoDeUsuario, TipoLocal.Deposito, paCodigoDeSucursal, RsAux!REvArticulo, RsAux!REvAEntregar - .Cell(flexcpValue, iQ, 0), paEstadoArticuloEntrega, 1, iTD, lDoc
                
                    'Quito del stock total todos los a Entregar y la diferencia de lo que no entrego lo pongo a RETIRAR
                    'Doy de alta a RETIRAR la resta a Entregar la hago abajo.
                    MarcoMovimientoStockTotal RsAux!REvArticulo, TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, RsAux!REvAEntregar - .Cell(flexcpValue, iQ, 0), 1
                    
                    'Doy la baja al movimiento de estados.
                    MarcoMovimientoStockEstado paCodigoDeUsuario, RsAux!REvArticulo, RsAux!REvAEntregar - .Cell(flexcpValue, iQ, 0), TipoMovimientoEstado.ARetirar, 1, iTD, lDoc, paCodigoDeSucursal
                End If
                
                'Pongo en la tabla Renglon los que quedarón para RETIRAR.
                Cons = "Select * From Renglon Where RenDocumento = " & lDoc _
                    & " And RenArticulo = " & RsAux!REvArticulo
            
                Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
                rsR.Edit
                rsR!RenARetirar = rsR!RenARetirar + (.Cell(flexcpData, iQ, 1) - .Cell(flexcpValue, iQ, 0))
                rsR.Update
                rsR.Close
            
                Cons = "Update Documento set DocFModificacion = '" & Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss") & "'" _
                    & "Where DocCodigo = " & lDoc
                cBase.Execute (Cons)
                
            End If
            
            If Not bArtServ Then
                'Movimientos común resto los A ENVIAR
                MarcoMovimientoStockTotal RsAux!REvArticulo, TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.AEntregar, RsAux!REvAEntregar, -1
                MarcoMovimientoStockEstado paCodigoDeUsuario, RsAux!REvArticulo, RsAux!REvAEntregar, TipoMovimientoEstado.AEntregar, -1, iTD, lDoc, paCodigoDeSucursal
            End If
            
            sErr = "Error al actualizar Renglón Envío. (Artículo ID = " & RsAux!REvArticulo & ", Envío= " & lEnvio & ")"
            If .Cell(flexcpValue, iQ, 0) = 0 Then
                'Si da cero borro renglon.
                RsAux.Delete
            Else
                RsAux.Edit
                'Al envío lo dejo con la cantidad que realmente entregó.
                If .Cell(flexcpValue, iQ, 0) <> .Cell(flexcpData, iQ, 1) Then RsAux!RevCantidad = RsAux!RevCantidad - (.Cell(flexcpData, iQ, 1) - .Cell(flexcpValue, iQ, 0))
                RsAux!REvAEntregar = 0          'LE RESTO TODO AL ENVÍO
                RsAux.Update
            End If
            RsAux.Close
        Next
    End With
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
                db_StockCamion RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar, iTD, lDoc
            
                'Doy la baja al movimiento de estados.
                MarcoMovimientoStockFisicoEnLocal 2, paCodigoDeSucursal, RsAux!REvArticulo, RsAux!REvAEntregar, paEstadoArticuloEntrega, 1
                MarcoMovimientoStockFisico paCodigoDeUsuario, 2, paCodigoDeSucursal, RsAux!REvArticulo, RsAux!REvAEntregar, paEstadoArticuloEntrega, 1, iTD, lDoc
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

Private Sub db_SaveEntregaParcial()
Dim rsE As rdoResultset
Dim sError As String

    On Error GoTo errBT
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errRB
    
    Cons = "Select * From Envio Where EnvCodigo = " & Val(tEnvio.Tag) & " And EnvEstado = 3"
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsE.EOF
        db_SetEntregaParcial rsE!EnvCodigo, rsE!EnvTipo, sError
    
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

Private Sub db_SaveEntregado(ByVal bEsTodo As Boolean)
Dim rsE As rdoResultset
Dim sError As String

    On Error GoTo errBT
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errRB
    If bEsTodo Then
        'Aca toma todos los envíos (incluye el Va Con)
        Cons = "Select * From Envio Where EnvCodImpresion = " & Val(tCodigo.Tag)
    Else
        Cons = "Select * From Envio " & _
            " Where ((EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Null) " & _
            "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & Val(tEnvio.Tag) & " And EnvVaCon Is Not Null)))"
    End If
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
    If bEsTodo Then
        tCodigo.Text = ""
        tCodigo.SetFocus
    Else
        tEnvio.Text = ""
        tEnvio.SetFocus
    End If
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

Private Sub db_SaveAImprimirAConfirmar(ByVal bEsAImprimir As Boolean)
Dim rsE As rdoResultset
Dim arrCli() As Long
Dim arrDocE() As Long
Dim sError As String
Dim iIndex As Integer, iQ As Integer

    'Aca tengo que controlar el documento pendiente
    'Si es con el formato anterior doy msg y me voy.
    If Not fnc_ValidoNoEntregado() Then Exit Sub
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
        If bEsAImprimir Then
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
    Dim objEnvio As New clsEnvio
    objEnvio.InvocoEnvio tEnvio.Tag, App.Path & "\Reportes"
    Set objEnvio = Nothing
    'Por si lo modifico o le quito el va con vuelvo a cargar.
    tEnvio.Tag = "": s_SetCtrlIndividual
    s_FindEnvio
    tooMenu.Buttons("envio").Enabled = (Val(tEnvio.Tag) > 0)
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
        Screen.MousePointer = 11
        db_SaveEntregado True
        Screen.MousePointer = 0
    Else
    
        If opEstado(2).Value Or opEstado(3).Value Then
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
            
            If opEstado(2).Value Then
                db_SaveEntregado False
            Else
                'Entrega Parcial.
                Dim iQ As Integer
                For iQ = 0 To vsEntParcial.Rows - 1
                    If vsEntParcial.Cell(flexcpValue, iQ, 0) > 0 Then
                        iQ = -1: Exit For
                    End If
                Next
                If iQ = -1 Then
                    db_SaveEntregaParcial
                Else
                    MsgBox "Lo correcto es eliminar el envío, pongalo primero a confirmar y luego acceda al formulario de envío.", vbCritical, "ATENCIÓN"
                End If
            End If
        
        ElseIf opEstado(1).Value Then
            
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
            
            If fnc_MercaderiaParcial Then
                MsgBox "El camión tiene asignada mercadería parcialmente, debe darle toda la mercadería al camión o pasar los envíos entregados primero.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If MsgBox("¿Confirma poner el envío a imprimir?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
                   
            db_SaveAImprimirAConfirmar True
                                
        ElseIf opEstado(0).Value Then
            
            If Trim(tMotivo.Text) = "" Then
                MsgBox "Ingrese el motivo por el cual pone el envío a confirmar.", vbExclamation, "Atención"
                tMotivo.SetFocus
                Exit Sub
            End If
            
            If fnc_MercaderiaParcial Then
                MsgBox "El camión tiene asignada mercadería parcialmente, debe darle toda la mercadería al camión o pasar los envíos entregados primero.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            
            If MsgBox("¿Confirma poner el envío a Confirmar?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
                   
            db_SaveAImprimirAConfirmar False
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
        
    Cons = "Select Sum(REvAEntregar) as QArt, ArtID, ArtCodigo, rTrim(ArtNombre) as ArtNombre " & _
            " From RenglonEnvio, Articulo " & _
            " Where REvEnvio IN (" & sVaCon & ")" & _
            " And RevArticulo = ArtID And RevAEntregar > 0" & _
            " Group by ArtID, ArtCodigo, ArtNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Cargo la lista por si selecciona la opción EntregaParcial.
    Do While Not RsAux.EOF
        With vsEntParcial
            .AddItem RsAux!QArt
            .Cell(flexcpText, .Rows - 1, 1) = "(" & Format(RsAux!ArtCodigo, "000,000") & ") " & Trim(RsAux!ArtNombre)
            lAux = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = lAux
            lAux = RsAux!QArt
            .Cell(flexcpData, .Rows - 1, 1) = lAux
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If MnuVaConItem(0).Tag <> "" Then
        With hlVaCon
            .Enabled = True
            .ForeColor = &HC00000
        End With
        hlDesvVaCon.Visible = True
        lVaCon.Enabled = True
        opEstado(3).Enabled = False
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

Private Sub s_SetCtrlEntregaParcial()
Dim iQ As Integer
    With vsEntParcial
        If .Rows > 0 Then
            For iQ = 0 To .Rows - 1
                If opEstado(3).Value = False Then
                    .Cell(flexcpText, iQ, 0) = .Cell(flexcpData, iQ, 1)
                End If
            Next
        End If
    End With
    
End Sub

Private Sub s_SetCtrlEstado()
    
    With tFecha
        .Enabled = opEstado(1).Value And Val(tEnvio.Tag) > 0
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
    
    opEstado(3).Enabled = (Not hlVaCon.Enabled And Val(tEnvio.Tag) > 0 And bCamionTieneMerc)
    
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
    opEstado(3).Enabled = opEstado(2).Enabled
    
    opEstado(0).Value = False
    opEstado(1).Value = False
    opEstado(2).Value = False
    opEstado(3).Value = False
    
    lVaCon.Enabled = False
    With hlVaCon
        .Enabled = False
        .ForeColor = vbButtonFace
    End With
    hlDesvVaCon.Visible = False
    
    With vsEntParcial
        .Rows = 0
        .ColWidth(0) = 500
    End With
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
        act_Save
    End If
End Sub

Private Sub Form_Load()

    s_CtrlClean
    With vsEntParcial
        .Rows = 0: .ColWidth(0) = 500
    End With
    
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
    PopupMenu MnuVaCon, , hlVaCon.Left + Frame1.Left, hlVaCon.Top + hlVaCon.Height + Frame1.Top
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
    s_SetCtrlEntregaParcial
    tooMenu.Buttons("save").Enabled = (Val(tEnvio.Tag) > 0)
End Sub

Private Sub opEstado_GotFocus(Index As Integer)
    Select Case Index
        Case 2: lHelp.Caption = "El envío se entrego normalmente."
        Case 1: lHelp.Caption = "El envío se imprimirá nuevamente en la fecha que indique, se mantiene el camionero."
        Case 0: lHelp.Caption = "El envío quedará sin fecha y puesto a confirmar."
        Case 3: lHelp.Caption = "Seleccione en la lista un artículo y reste los que no fueron entregados."
    End Select
End Sub

Private Sub opEstado_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Select Case Index
            Case 0: ' A Confirmar
                tMotivo.SetFocus
            Case 2: ' Entrego
                If Val(tEnvio.Tag) > 0 Then act_Save
            Case 1: ' A imprimir (nueva fecha)
                If opEstado(1).Value Then
                    tFecha.SetFocus
                Else
                    opEstado(3).SetFocus
                End If
            Case 3: ' Entrega Parcial
                vsEntParcial.SetFocus
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
    lHelp.Caption = "Ingrese el motivo por el cual el envío queda a confirmar."
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
    End Select
End Sub

Private Sub vsEntParcial_GotFocus()
    If opEstado(3).Value Then
        lHelp.Caption = "Deje en c/artículo la cantidad que se entregó al cliente. (+,- suma o resta)"
    Else
        lHelp.Caption = "Lista de Artículos del envío."
    End If
End Sub

Private Sub vsEntParcial_KeyDown(KeyCode As Integer, Shift As Integer)
    If opEstado(3).Value And vsEntParcial.Rows > 0 Then
        Select Case KeyCode
            Case vbKeyAdd
                If vsEntParcial.Cell(flexcpValue, vsEntParcial.Row, 0) < vsEntParcial.Cell(flexcpData, vsEntParcial.Row, 1) Then
                    vsEntParcial.Cell(flexcpText, vsEntParcial.Row, 0) = vsEntParcial.Cell(flexcpValue, vsEntParcial.Row, 0) + 1
                End If
            Case vbKeySubtract
                If vsEntParcial.Cell(flexcpValue, vsEntParcial.Row, 0) > 0 Then
                    vsEntParcial.Cell(flexcpText, vsEntParcial.Row, 0) = vsEntParcial.Cell(flexcpValue, vsEntParcial.Row, 0) - 1
                End If
        End Select
    End If
End Sub

Private Sub vsEntParcial_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If opEstado(3).Value And vsEntParcial.Rows > 0 Then act_Save
    End If
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


