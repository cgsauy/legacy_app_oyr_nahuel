VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form VerArribo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arribo de Mercadería"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VerArribo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   3500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin AACombo99.AACombo cFolder 
      Height          =   315
      Left            =   960
      TabIndex        =   12
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      ListIndex       =   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
   End
   Begin VB.Frame fDatos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   5655
      Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
         Height          =   1455
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2566
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
      Begin VB.TextBox tDias 
         Height          =   285
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox tPuerto 
         Height          =   285
         Left            =   840
         MaxLength       =   11
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "&Días Libres:"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lFArribo 
         Caption         =   "Arribo:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   615
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4380
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3069
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label labProveedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   540
      Width           =   2895
   End
   Begin VB.Label lZF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Carpeta:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pro&veedor:"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   540
      Width           =   855
   End
   Begin VB.Label lArribo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Local"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lEmbarque 
      Caption         =   "&Local"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lDesde 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   5655
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "VerArribo.frx":0DC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "VerArribo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsSC As rdoResultset
Private FEmbarque As Date
Private sNuevo As Boolean, sModificar As Boolean
Private NroFila As Byte

Private Sub cFolder_Click()
On Error GoTo ErrFC
    RelojA
    LimpioCampos
    If cFolder.ListIndex <> -1 Then
        CargoDatosEmbarque
    Else
        Botones False, False, False, False, False, Toolbar1, Me
    End If
    RelojD
    Exit Sub
ErrFC:
    msgError.MuestroError "Ocurrio un error al cargar la información.", Trim(Err.Description)
    RelojD
End Sub

Private Sub cFolder_Change()
    LimpioCampos
End Sub

Private Sub cFolder_GotFocus()
On Error GoTo ErrCF
    cFolder.SelStart = 0
    cFolder.SelLength = Len(cFolder.Text)
    If cFolder.ListCount = 0 Then CargoComboFolder
    Exit Sub
ErrCF:
    msgError.MuestroError "Ocurrio un error inesperado.", Trim(Err.Description)
    RelojD
End Sub

Private Sub cFolder_KeyDown(KeyCode As Integer, Shift As Integer)
    Exit Sub        'Anulado porque vuelve un lugar atras el combo.
    If KeyCode = vbKeyReturn Then
        If cFolder.ListIndex = -1 Then Exit Sub
        If MnuNuevo.Enabled Then
            If MsgBox("Confirma realizar un nuevo arribo de mercadería.", vbQuestion + vbOKCancel, "NUEVO") = vbOK Then
                AccionNuevo
            End If
        Else
            If MnuModificar.Enabled Then
                If MsgBox("Confirma modificar el arribo de mercadería.", vbQuestion + vbOKCancel, "MODIFICAR") = vbOK Then
                    AccionModificar
                End If
            End If
        End If
    End If
End Sub

Private Sub cFolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Seleccione el embarque para registrar el arribo."
    
End Sub


Private Sub Form_Activate()
On Error Resume Next
    Me.Refresh
    RelojD
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    NroFila = 0
    LimpioCampos
    
    'Obtengo la fecha del servidor.
    FechaDelServidor
    
    'Cargo los parámetros de importaciones.
    CargoParametrosImportaciones
    
    LimpioEtiquetas
    
    DeshabilitoIngreso

    Exit Sub
ErrLoad:
    msgError.MuestroError "Ocurrio un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(4).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CierroConexion
    Set msgError = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco cFolder
End Sub
Private Sub Label7_Click()
    Foco tDias
End Sub

Private Sub lFArribo_Click()
    Foco tPuerto
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuVolver_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub AccionEliminar()
    AccionEliminarEmbarque
End Sub

Private Sub AccionEliminarEmbarque()
Dim RsArt As rdoResultset, RsST As rdoResultset

    If MsgBox("Confirma eliminar el arribo de mercadería.", vbQuestion + vbOKCancel, "ATENCION") = vbCancel Then Exit Sub
    
    On Error GoTo ErrEli
    RelojA
    FechaDelServidor
    
    If InStr(lZF.Tag, "Z") > 0 Then
        'Si es Zona Franca, verifico que no posea subcarpetas abiertas.
        Cons = "Select * From SubCarpeta Where SubEmbarque = " & CLng(cFolder.ItemData(cFolder.ListIndex))
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then
            RsAux.Close
            MsgBox "El embarque seleccionado posee subcarpetas abiertas, no podrá eliminar la fecha de arribo.", vbExclamation, "ATENCIÓN"
            RelojD
            Exit Sub
        End If
    End If
    
    cBase.BeginTrans        'Comienzo Transaccion------------------------------------------------------------------!!!!!!
    
    On Error GoTo ErrResumir
    Cons = "Select * From Embarque " _
            & " Where EmbID = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
            & " And EmbCodigo = '" & UCase(Mid(cFolder.Text, InStr(cFolder.Text, ".") + 1, 1)) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
    If RsAux.EOF Then
        RsAux.Close
        cBase.CommitTrans   'Cierro Transacción.---------------------
        
        RelojD
        MsgBox "Otra terminal pudo eliminar el embarque, verifique.", vbExclamation, "ATENCION"
        Exit Sub
    Else
        If RsAux!EmbLocal = paLocalZF Or RsAux!EmbLocal = paLocalPuerto Then
            If IsNull(RsAux!EmbFLocal) Then
                'Doy de baja al stock los artículos.---
                Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
                    & " And AFoCodigo = " & RsAux!EmbID
                Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                Do While Not RsArt.EOF
                    'Veo si en el local hay los artículos que le estoy dando de baja.
                    Cons = "Select * From StockLocal " _
                        & " Where StLArticulo = " & RsArt!AFoArticulo & " And StLTipoLocal = " & TipoLocal.Deposito _
                        & " And StLLocal = " & RsAux!EmbLocal & " And StLEstado = " & paEstadoArticuloEntrega
                    Set RsST = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsST.EOF Then
                        Cons = "IDArtículo = " & RsArt!AFoArticulo
                            RsST.Close
                            RsArt.Close
                            RsAux.Close
                            cBase.RollbackTrans
                            MsgBox "No hay stock para el artículo en el local, verifique si no se hizo algún traslado o algo que le pueda haber quitado al local." & Chr(13) & Cons, vbExclamation, "ATENCIÓN"
                            Screen.MousePointer = 0
                            Exit Sub
                    Else
                        If RsST!StLCantidad < RsArt!AFoCantidad Then
                            Cons = "IDArtículo = " & RsArt!AFoArticulo
                            RsST.Close
                            RsArt.Close
                            RsAux.Close
                            cBase.RollbackTrans
                            MsgBox "No hay tanto stock para el artículo en el local, verifique si no se hizo algún traslado o algo que le pueda haber quitado mercadería al local." & Chr(13) & Cons, vbExclamation, "ATENCIÓN"
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                    End If
                    RsST.Close
                    'Agregó el artículo al stock.---------
                    MarcoMovimientoStockFisico UsuarioLogueado, TipoLocal.Deposito, RsAux!EmbLocal, RsArt!AFoArticulo, RsArt!AFoCantidad, paEstadoArticuloEntrega, -1, TipoDocumento.CompraCarpeta, RsAux!EmbID
                    MarcoMovimientoStockTotal RsArt!AFoArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsArt!AFoCantidad, -1
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, RsAux!EmbLocal, RsArt!AFoArticulo, RsArt!AFoCantidad, paEstadoArticuloEntrega, -1
                    RsArt.MoveNext
                Loop
                RsArt.Close
                'Edito el embarque.-----------------------
                RsAux.Edit
                RsAux!EmbFArribo = Null
                RsAux!EmbDiasLibres = Null
                RsAux!EmbFModificacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux.Update
                RsAux.Close
                cBase.CommitTrans                                   'Cierro Transacción.(Faltaba)---------------------
            Else
                RsAux.Close
                cBase.CommitTrans   'Cierro Transacción.---------------------
                
                RelojD
                MsgBox "La mercadería de este embarque fue dada de alta en los depósitos, no podrá eliminar la fecha de arribo del mismo.", vbInformation, "ATENCIÓN"
            End If
        Else
            RsAux.Close
            cBase.CommitTrans               'Cierro Transacción.---------------------
            
            RelojD
            MsgBox "Al embarque le fue modificado el local de destino por un local de depósito, verifique.", vbInformation, "ATENCIÓN"
        End If
    End If
    
    cFolder_Click
    RelojD
    Exit Sub
    
ErrEli:
    msgError.MuestroError "Ha ocurrido un error al eliminar el arribo.", Trim(Err.Description)
    RelojD
    Exit Sub
ErrResumir:
    Resume ErrTransaccion
ErrTransaccion:
    cBase.RollbackTrans
    msgError.MuestroError "Ha ocurrido un error al intentar eliminar el arribo.", Trim(Err.Description)
    RelojD
End Sub
Private Sub AccionGrabar()
    
    If Not ValidoCamposE Then Exit Sub
    If MsgBox("Confirma grabar el arribo de la mercadería", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
        AccionGrabarEmbarque
    End If
    
End Sub

Private Sub AccionGrabarEmbarque()
Dim CodLocal As Long
    RelojA
    FechaDelServidor
    'Comienzo la transacción.---------------------
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrET

    Cons = "Select * From Embarque " _
        & " Where EmbID = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
            & " And EmbCodigo = '" & UCase(Mid(cFolder.Text, InStr(cFolder.Text, ".") + 1, 1)) & "'"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
    If RsAux.EOF Then
        RsAux.Close
        cBase.RollbackTrans
        RelojD
        MsgBox "Otra terminal pudo eliminar el embarque, verifique.", vbExclamation, "ATENCION"
        Exit Sub
    Else
        If RsAux!EmbFModificacion <> FEmbarque Then
            RsAux.Close
            cBase.RollbackTrans
            RelojD
            MsgBox "El embarque fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        CodLocal = RsAux!EmbLocal
        RsAux.Edit
        RsAux!EmbFArribo = Format(tPuerto.Text, sqlFormatoF)
        If IsNumeric(tDias.Text) Then RsAux!EmbDiasLibres = Val(tDias.Text)
        RsAux!EmbFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update
        RsAux.Close
    End If
    
    For I = 1 To vsArticulo.Rows - 1
        Cons = "Select * From ArticuloFolder " _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
            & " And AFoArticulo =" & CLng(vsArticulo.Cell(flexcpText, I, 0))
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
        If RsAux.EOF Then
            RsAux.Close
            'Cierro transacción.------------------------------------------------
            cBase.RollbackTrans
            MsgBox "Otra terminal pudo eliminar los artículos del embarque, verifique.", vbExclamation, "ATENCION"
            RelojD
            Exit Sub
        Else
            If sNuevo Then
                 'Agregó el artículo al stock.---------
                 gFechaServidor = CDate(tPuerto.Text) & " " & Time
                MarcoMovimientoStockFisico UsuarioLogueado, TipoLocal.Deposito, CodLocal, CLng(vsArticulo.Cell(flexcpText, I, 0)), CCur(vsArticulo.Cell(flexcpText, I, 3)), paEstadoArticuloEntrega, 1, TipoDocumento.CompraCarpeta, RsAux!AFoCodigo
                MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(vsArticulo.Cell(flexcpText, I, 3)), 1
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, CodLocal, vsArticulo.Cell(flexcpText, I, 0), CCur(vsArticulo.Cell(flexcpText, I, 3)), paEstadoArticuloEntrega, 1
            Else
                If RsAux!AFoCantidad < CLng(vsArticulo.Cell(flexcpText, I, 3)) Then
                    'Agregó el artículo al stock.---------
                    MarcoMovimientoStockFisico UsuarioLogueado, TipoLocal.Deposito, CodLocal, CLng(vsArticulo.Cell(flexcpText, I, 0)), CCur(vsArticulo.Cell(flexcpText, I, 3)) - RsAux!AFoCantidad, paEstadoArticuloEntrega, 1, TipoDocumento.CompraCarpeta, RsAux!AFoCodigo
                    MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(vsArticulo.Cell(flexcpText, I, 3)) - RsAux!AFoCantidad, 1
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, CodLocal, vsArticulo.Cell(flexcpText, I, 0), CCur(vsArticulo.Cell(flexcpText, I, 3)) - RsAux!AFoCantidad, paEstadoArticuloEntrega, 1
                ElseIf RsAux!AFoCantidad > CLng(vsArticulo.Cell(flexcpText, I, 3)) Then
                    'Quitó artículo del stock.---------
                    MarcoMovimientoStockFisico UsuarioLogueado, TipoLocal.Deposito, CodLocal, CLng(vsArticulo.Cell(flexcpText, I, 0)), RsAux!AFoCantidad - CCur(vsArticulo.Cell(flexcpText, I, 3)), paEstadoArticuloEntrega, -1, TipoDocumento.CompraCarpeta, RsAux!AFoCodigo
                    MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsAux!AFoCantidad - CCur(vsArticulo.Cell(flexcpText, I, 3)), -1
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, CodLocal, vsArticulo.Cell(flexcpText, I, 0), RsAux!AFoCantidad - CCur(vsArticulo.Cell(flexcpText, I, 3)), paEstadoArticuloEntrega, -1
                End If
            End If
            RsAux.Edit
            RsAux!AFoCantidad = CLng(vsArticulo.Cell(flexcpText, I, 3))
            RsAux.Update
            RsAux.Close
        End If
    Next I
    
    cBase.CommitTrans
    'Finalizó Transacción---------------------------------
    On Error GoTo ErrAjusto
    AccionCancelar
    RelojD
    cFolder_Click
    Exit Sub

ErrBT:
    msgError.MuestroError "No se pudo iniciar la transacción."
    RelojD
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    msgError.MuestroError "No se pudo almacenar la ficha, reintente.", Trim(Err.Description)
    RelojD
    Exit Sub
ErrAjusto:
    msgError.MuestroError "Ocurrio un error luego de almacenar la información.", Trim(Err.Description)
    RelojD
End Sub
Private Sub AccionCancelar()
On Error GoTo ErrAC
Dim strAux As String
    RelojA
    strAux = cFolder.Text
    DeshabilitoIngreso
    LimpioCampos
    CargoComboFolder
    Call Botones(False, False, False, False, False, Toolbar1, Me)
    sModificar = False
    sNuevo = False
    cFolder.Text = strAux
    If cFolder.ListIndex = -1 Then cFolder.Text = ""
    Foco cFolder
    RelojD
    Exit Sub
ErrAC:
    msgError.MuestroError "Ocurrio un error al intentar restaurar el formulario."
    RelojD
End Sub

Private Sub DeshabilitoIngreso()

    tPuerto.Enabled = False
    tDias.Enabled = False
    tPuerto.BackColor = inactivo
    tDias.BackColor = inactivo
    vsArticulo.BackColor = inactivo
    cFolder.Enabled = True
    cFolder.BackColor = obligatorio

End Sub
Private Sub HabilitoIngreso()
    
    cFolder.Enabled = False
    cFolder.BackColor = inactivo
    
    tPuerto.Enabled = True
    tPuerto.BackColor = obligatorio
    
    vsArticulo.BackColor = vbWhite
    
    If InStr(lZF.Tag, "P") > 0 Then
        tDias.Enabled = True
        tDias.BackColor = vbWhite
    End If
    
End Sub

Private Sub tDias_GotFocus()

    tDias.SelStart = 0
    tDias.SelLength = Len(tDias.Text)
        
End Sub

Private Sub tDias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tDias_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(4).Text = "Ingrese la cantidad de días libres en el puerto."
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo"
            AccionNuevo
            
        Case "modificar"
            AccionModificar
        
        Case "eliminar"
            AccionEliminar
            
        Case "grabar"
            AccionGrabar
        
        Case "cancelar"
            AccionCancelar
        
        Case "salir"
            Unload Me
            
    End Select

End Sub

Sub CargoComboFolder()
    
    RelojA
    cFolder.Clear
    
    Cons = "Select * From Embarque, Carpeta" _
        & " Where EmbFEmbarque < '" & Format(gFechaServidor, sqlFormatoF) & "'" _
        & " And EmbCarpeta = CarID" _
        & " And EmbCosteado = 0 And EmbFLocal IS Null And EmbFEmbarque IS Not Null"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
        cFolder.AddItem RsAux!CarCodigo & "." & RsAux!EmbCodigo
        cFolder.ItemData(cFolder.NewIndex) = RsAux!EmbID
        RsAux.MoveNext
    Loop
    RsAux.Close
    RelojD
    
End Sub
Sub AccionModificar()
On Error GoTo ErrAM

    sModificar = True
    RelojA
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    CargoArticulos
    HabilitoIngreso
    RelojA
    Cons = "Select * From Embarque " _
        & " Where EmbID = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
        & " And EmbCodigo = '" & UCase(Mid(cFolder.Text, InStr(cFolder.Text, ".") + 1, InStr(cFolder.Text, ".") + 1)) & "'"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
    If RsAux.EOF Then
        MsgBox "El embarque pudo ser eliminado por otra terminal, verifique.", vbExclamation, "ATENCION"
        RsAux.Close: AccionCancelar
    Else
        If RsAux!EmbFModificacion <> FEmbarque Then
            RsAux.Close
            MsgBox "El embarque fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
            AccionCancelar
        Else
            If Not IsNull(RsAux!EmbFArribo) Then tPuerto.Text = Format(RsAux!EmbFArribo, FormatoFP)
            If Not IsNull(RsAux!EmbDiasLibres) Then tDias.Text = RsAux!EmbDiasLibres
            RsAux.Close
        End If
    End If
    If vsArticulo.Rows > 1 Then vsArticulo.Select 1, 3, 1, 3: vsArticulo.SetFocus
    RelojD
    Exit Sub
ErrAM:
    msgError.MuestroError "Ocurrio un error al intentar actualizar el formulario para modificar.", Trim(Err.Description)
    RelojD
End Sub

Private Sub AccionNuevo()
On Error GoTo ErrAN
    RelojA
    sNuevo = True
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    CargoArticulos
    HabilitoIngreso
    tPuerto.SetFocus
    tPuerto.Text = Format(gFechaServidor, FormatoFP)
    vsArticulo.Select 1, 3, 1, 3: vsArticulo.SetFocus
    RelojD
    Exit Sub
ErrAN:
    msgError.MuestroError "Ocurrio un error al dar acción nuevo.", Trim(Err.Description)
    AccionCancelar
    RelojD
End Sub
Sub LimpioCampos()
    LimpioEtiquetas
    labProveedor.Caption = ""
    tPuerto.Text = ""
    tDias.Text = ""
    LimpioGrilla
End Sub

Private Sub CargoArticulos()
On Error GoTo ErrCA
    
    RelojA
    LimpioGrilla
    Cons = "Select AFoArticulo, AFoCantidad, ArtNombre From ArticuloFolder, Articulo " _
        & " Where AFoTipo = " & Folder.cFEmbarque _
        & " And AFoCodigo = " & cFolder.ItemData(cFolder.ListIndex) _
        & " And AFoArticulo = ArtID" _
        & " Order By AFoArticulo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)

    Do While Not RsAux.EOF
        vsArticulo.AddItem ""
        vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 0) = RsAux!AFoArticulo
        vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 2) = Trim(RsAux!ArtNombre)
        If sModificar Then
            vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 3) = RsAux!AFoCantidad
        Else
            If lZF.Tag = "P" Then vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 3) = RsAux!AFoCantidad
        End If
        Cons = "Select Sum(AFoCantidad) From ArticuloFolder, SubCarpeta" _
            & " Where AFoTipo = " & Folder.cFSubCarpeta _
            & " And AFoArticulo = " & RsAux!AFoArticulo _
            & " And SubEmbarque = " & cFolder.ItemData(cFolder.ListIndex) _
            & " And SubID = AFoCodigo"
        Set RsSC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not IsNull(RsSC(0)) Then
            vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 1) = RsSC(0)
        Else
            vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 1) = 0
        End If
        RsSC.Close
        RsAux.MoveNext
    Loop
    RsAux.Close
    RelojD
    Exit Sub
ErrCA:
    msgError.MuestroError "Ocurrio un error al cargar la lista de artículos.", Trim(Err.Description)
    RelojD
End Sub

Private Sub CargoDatosEmbarque()
On Error GoTo errCargar

    RelojA
        
    Cons = "Select Embarque.*, PExNombre, Destino.CiuNombre 'Destino', Origen.CiuNombre 'Origen' " _
        & " From Embarque, Carpeta, Ciudad Destino, Ciudad Origen, ProveedorExterior" _
        & " Where EmbID = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
        & " And EmbCodigo = '" & UCase(Mid(cFolder.Text, InStr(cFolder.Text, ".") + 1, InStr(cFolder.Text, ".") + 1)) & "'" _
        & " And EmbCiudadDestino = Destino.CiuCodigo And EmbCiudadOrigen = Origen.CiuCodigo" _
        & " And CarID = EmbCarpeta And CarProveedor = PExCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)

    If Not RsAux.EOF Then
        FEmbarque = RsAux!EmbFModificacion
        If Not IsNull(RsAux!PExNombre) Then labProveedor.Caption = Trim(RsAux!PExNombre)
        
        'FECHAS DE ARRIBO--------------------------------------------------------------------
        If RsAux!EmbLocal = paLocalZF Then
            If IsNull(RsAux!EmbFArribo) Then
                lZF.Caption = "Mercadería va a Zona Franca"
                lZF.Tag = "Z"
                Call Botones(True, False, False, False, False, Toolbar1, Me)
            Else
                lZF.Caption = "Mercadería en Zona Franca"
                lZF.Tag = "EZ"
                If TieneSubCarpetas Then
                    Call Botones(False, True, False, False, False, Toolbar1, Me)
                Else
                    Call Botones(False, True, True, False, False, Toolbar1, Me)
                End If
            End If
        Else
            If RsAux!EmbLocal = paLocalPuerto Then
                If IsNull(RsAux!EmbFArribo) Then
                    lZF.Caption = "Mercadería va a Puerto"
                    lZF.Tag = "P"
                    Call Botones(True, False, False, False, False, Toolbar1, Me)
                Else
                    lZF.Caption = "Mercadería en Puerto"
                    lZF.Tag = "EP"
                    Call Botones(False, True, True, False, False, Toolbar1, Me)
                End If
            Else
                lZF.Caption = "Mercadería va directo a Local "
                Call Botones(False, False, False, False, False, Toolbar1, Me)
            End If
        End If
        
        If Not IsNull(RsAux!Origen) Then
            lDesde.Caption = "Desde " & Trim(RsAux!Origen)
            If Not IsNull(RsAux!Destino) Then
                lDesde.Caption = lDesde.Caption & " Hacia " & Trim(RsAux!Destino)
            Else
                lDesde.Caption = lDesde.Caption & " Hacia ......................."
            End If
        Else
            If Not IsNull(RsAux!Destino) Then
                lDesde.Caption = "Desde ......... Hacia " & Trim(RsAux!Destino)
            Else
                lDesde.Caption = "Desde ................. Hacia ......................."
            End If
        End If
        
        lEmbarque.Caption = "Embarcó el " & Format(RsAux!EmbFEmbarque, "d-mmm-yy")
        If Not IsNull(RsAux!EmbFAPrometido) Then
            lArribo.Caption = "Arribo prometido para el " & Format(RsAux!EmbFAPrometido, "d-mmm-yy")
        Else
            lArribo.Caption = "Arribo prometido para el ..........."
        End If
    End If
    
    RsAux.Close
    RelojD
    Exit Sub

errCargar:
    msgError.MuestroError "Ha ocurrido un error al cargar los datos del embarque.", Trim(Err.Description)
    RelojD
    
End Sub
Function ValidoCamposE()
On Error GoTo ErrVCE

    ValidoCamposE = True
    If Not IsDate(tPuerto.Text) Then
        MsgBox "La fecha ingresada no es correctas. Verifique", vbExclamation, "ATENCION"
        ValidoCamposE = False: tPuerto.SetFocus: Exit Function
    End If
    If CDate(tPuerto.Text) > gFechaServidor Then
        MsgBox "La fecha ingresada no es correctas. Verifique", vbExclamation, "ATENCION"
        ValidoCamposE = False: tPuerto.SetFocus: Exit Function
    End If
    If tDias.Enabled Then
        If Not IsNumeric(tDias.Text) Or Trim(tDias.Text) = "" Then
            MsgBox "Se debe ingresar la cantidad de días libres en el puerto.", vbExclamation, "ATENCION"
            tDias.SetFocus: ValidoCamposE = False: Exit Function
        End If
    End If
        
    Cons = "Select * From Embarque " _
        & " Where EmbID = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
        & " And EmbCodigo = '" & UCase(Mid(cFolder.Text, InStr(cFolder.Text, ".") + 1, InStr(cFolder.Text, ".") + 1)) & "'"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
    If RsAux.EOF Then
        MsgBox "El embarque pudo ser eliminado por otra terminal, verifique.", vbExclamation, "ATENCION"
        ValidoCamposE = False: RsAux.Close
        Exit Function
    Else
        If RsAux!EmbFModificacion <> FEmbarque Then
            MsgBox "El embarque fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
            ValidoCamposE = False: RsAux.Close: Exit Function
        End If
        If Not IsNull(RsAux!EmbFEmbarque) Then
            If Trim(tPuerto.Text) <> "" Then
                If RsAux!EmbFEmbarque > CDate(tPuerto.Text) Then
                    MsgBox "La mercadería embarcó con fecha superior al arribo, verifique.", vbExclamation, "ATENCION"
                    ValidoCamposE = False: RsAux.Close
                    Exit Function
                End If
            End If
        Else
            MsgBox "El embarque no tiene fecha de embarcado, verifique.", vbExclamation, "ATENCION"
            ValidoCamposE = False: RsAux.Close
            Exit Function
        End If
    End If
    RsAux.Close
    

    For I = 1 To vsArticulo.Rows - 1
        'Cantidad menores que cero.----------------------------
        If Val(vsArticulo.Cell(flexcpText, I, 3)) <= 0 Then
            MsgBox "Existen artículos cuyas cantidades no son correctas.", vbExclamation, "ATENCION"
            ValidoCamposE = False: Exit Function
        End If
        'Verifico cantidad en subcarpetas.-----------------------
        If Val(vsArticulo.Cell(flexcpText, I, 3)) < Val(vsArticulo.Cell(flexcpText, I, 1)) Then
            MsgBox "El artículo " & vsArticulo.Cell(flexcpText, I, 2) & " posee subcarpetas abiertas y la cantidad ingresada es menor a las que posee en las mismas.", vbExclamation, "ATENCION"
            ValidoCamposE = False: Exit Function
        End If
    Next I
    
    For I = 1 To vsArticulo.Rows - 1
        Cons = "Select * From ArticuloFolder " _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = " & cFolder.ItemData(cFolder.ListIndex) _
            & " And AFoArticulo =" & vsArticulo.Cell(flexcpText, I, 0)
    
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
        
        If Not RsAux.EOF Then
            If CLng(vsArticulo.Cell(flexcpText, I, 3)) <> RsAux!AFoCantidad Then
                If MsgBox("La cantidad arribada para el artículo " & vsArticulo.Cell(flexcpText, I, 2) & " no coincide con las del embarque." & Chr(13) _
                    & "Cantidad original: " & RsAux!AFoCantidad & " , Ingresada: " & vsArticulo.Cell(flexcpText, I, 3) & Chr(13) & Chr(13) _
                    & "Desea modificar el embarque y continuar?", vbExclamation + vbYesNo, "ATENCION") = vbNo Then
                    
                    ValidoCamposE = False
                    RsAux.Close
                    Exit Function
                End If
            End If
        Else
            MsgBox "El articulo " & vsArticulo.Cell(flexcpText, I, 2) & " fue eliminado del embarque, verifique.", vbExclamation, "ATENCION"
            ValidoCamposE = False: RsAux.Close: Exit Function
        End If
        RsAux.Close
    Next I

    Exit Function
ErrVCE:
    msgError.MuestroError "Ocurrio un error al validar los datos.", Trim(Err.Description)
    ValidoCamposE = False
    RelojD
End Function

Private Sub tPuerto_GotFocus()
    tPuerto.SelStart = 0
    tPuerto.SelLength = Len(tPuerto.Text)
End Sub

Private Sub tPuerto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If tDias.Enabled Then
            If tDias.Text = "" And IsDate(tPuerto.Text) Then
                RelojA
                'Saco la cantidad de dias de la agencia.
                Cons = "Select * From Embarque, AgenciaTransporte " _
                    & " Where EmbID = " & CLng(cFolder.ItemData(cFolder.ListIndex)) _
                    & " And EmbCodigo = '" & UCase(Mid(cFolder.Text, InStr(cFolder.Text, ".") + 1, 1)) & "'" _
                    & " And EmbAgencia = ATrCodigo"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                If Not RsAux.EOF Then
                    If Not IsNull(RsAux!ATrDemora) Then tDias.Text = RsAux!ATrDemora
                End If
                RsAux.Close
                RelojD
            End If
            tDias.SetFocus
        Else
            AccionGrabar
        End If
    End If
    
End Sub

Private Sub tPuerto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Fecha de arribo de mercadería al puerto."
    
End Sub

Private Sub LimpioGrilla()

    With vsArticulo
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Editable = True
        .Rows = 1
        .Cols = 3
        .FormatString = "IDArticulo|ArtEnSubCarp|Articulo|>Cantidad"
        .ColWidth(2) = 3500
        .ColWidth(3) = 900
        .ColHidden(0) = True
        .ColHidden(1) = True
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With

End Sub

Private Sub LimpioEtiquetas()
    lDesde.Caption = "Arribo desde......"
    lArribo.Caption = "Arribo prometido para el.........."
    lEmbarque.Caption = "Embarcó el..... "
    lZF.Caption = "Mercadería en......"
    lZF.Tag = ""
End Sub

Private Sub vsArticulo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 3 Then Cancel = True
    If vsArticulo.BackColor = inactivo Then Cancel = True
End Sub

Private Sub vsArticulo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not IsNumeric(vsArticulo.EditText) Then
        Cancel = True
        MsgBox "Ingrese un numéro mayor que cero.", vbInformation, "ATENCIÓN"
        Exit Sub
    End If
End Sub

Private Function TieneSubCarpetas() As Boolean
On Error GoTo ErrTSC
    Cons = "Select * From SubCarpeta" _
        & " Where SubEmbarque = " & cFolder.ItemData(cFolder.ListIndex)
    Set RsSC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsSC.EOF Then TieneSubCarpetas = False Else TieneSubCarpetas = True
    RsSC.Close
    Exit Function
ErrTSC:
    msgError.MuestroError "Ocurrio un error al buscar subcarpetas para el embarque.", Trim(Err.Description)
    TieneSubCarpetas = True
End Function
