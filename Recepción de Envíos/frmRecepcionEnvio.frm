VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmRecepcionEnvio 
   BackColor       =   &H00CEBAB3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepci�n de Env�os"
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
      Caption         =   "En&treg�"
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
            Caption         =   "Env�o"
            Key             =   "envio"
            Object.ToolTipText     =   "Ir a formulario de env�os"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dividir"
            Key             =   "dividir"
            Object.ToolTipText     =   "Dividir un env�o en dos"
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
      Caption         =   "&Env�o:"
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
      Caption         =   "Direcci�n:"
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
      Caption         =   "Cami�n: Mart�n"
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
      Caption         =   "&C�digo:"
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
'10/1/2007      Deshabilito entrega parcial si no se entreg� ning�n art�culo.
Option Explicit

Dim lDirEnvio As Long, lAgeEnvio As Long
Dim bCamionTieneMerc As Boolean

Private Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Private Sub act_DividirEnvio()
On Error GoTo errDE
    'Dim oFrm As New frmDividoEnvio
    Dim oFrm As New frmMercaAReclamar
    oFrm.prmInvocacion = 0
    oFrm.prmEnvio = Val(tEnvio.Tag)
   oFrm.Show vbModal
    Set oFrm = Nothing
    Exit Sub
errDE:
    objGral.OcurrioError "Error al acceder al formulario para dividir el env�o.", Err.Description, "Dividir Env�os"
End Sub

Private Sub db_IndependizarVaCon()
On Error GoTo errIVC
Dim rsEnvio As rdoResultset
Dim lOld As Long, iQ As Integer

    Screen.MousePointer = 11
    Cons = "Select * From Envio  Where EnvCodigo = " & Val(tEnvio.Text)
    Set rsEnvio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Cuento Q de env�os en todo el va con.
    Cons = "Select count(*) from Envio Where Abs(EnvVaCon) = " & Abs(rsEnvio("EnvVaCon"))
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    iQ = RsAux(0)
    RsAux.Close
    
    'Si est� casado con m�s de uno le tengo que decir que tiene que desvincular los otros ya que si no me qued�n todos
    'desvinculados.
    If iQ > 2 And rsEnvio("EnvVaCon") < 0 Then
        RsAux.Close
        rsEnvio.Close
        Screen.MousePointer = 0
        MsgBox "Este env�o posee m�s de un env�o en el va con y este es el que une a todos. No puede desvincular este env�o ingrese el c�digo de otro de los env�os del va con.", vbCritical, "Atenci�n"
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
'ocasiono error ya que no tengo esa mercader�a en el registro.
            RsRE.Close
            RsRE.Edit
        Else
            'Si no tiene mercader�a dejo.
            If RsRE!ReECantidadEntregada = 0 Then
                db_RestoARenglonEntrega = False
                'Si la cantidad total = a la de este env�o --> no tengo m�s env�os con este art�culo si no resto la total.
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
        
        sErr = "Error al restar en la tabla RenglonEntrega. (Art�culo ID = " & RsAux!REvArticulo & ", Env�o= " & lEnvio & ")"
        db_RestoARenglonEntrega RsAux!REvArticulo, RsAux!REvAEntregar, True
        
        sErr = "Error al actualizar StockLocal. (Art�culo ID = " & RsAux!REvArticulo & ", Env�o= " & lEnvio & ")"
        'Le quito la mercader�a entregada al cami�n
        db_StockFisicoEnLocal RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar * -1, iTD, lDoc, 1, lCamion.Tag
        
        'Doy la baja al movimiento de estados.
        Cons = "EXEC StockMovEstadoStockTotal " & paCodigoDeUsuario & ", " & RsAux!REvArticulo & ", " & RsAux!REvAEntregar * -1 & ", " & TipoMovimientoEstado.AEntregar & ", " & iTD & ", " & lDoc & ", " & paCodigoDeSucursal
        cBase.Execute Cons
        
        
        sErr = "Error al actualizar Rengl�n Env�o. (Art�culo ID = " & RsAux!REvArticulo & ", Env�o= " & lEnvio & ")"
        RsAux.Edit
        RsAux!REvAEntregar = 0
        RsAux.Update
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Pongo los art�culos de servicio como entregados.--------------------------------------
    sErr = "Error al actualizar los art�culos de servicio para el env�o: " & lEnvio
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
        
        sErr = "Error al restar en la tabla RenglonEntrega. (Art�culo ID = " & RsAux!REvArticulo & ", Env�o= " & lEnvio & ")"
        If db_RestoARenglonEntrega(RsAux!REvArticulo, RsAux!REvAEntregar, False) Then
            If bCamionTieneMerc Then
                'Si entro es xq no ten�a mercader�a para este art�culo.
                sErr = "Error al actualizar StockLocal. (Art�culo ID = " & RsAux!REvArticulo & ", Env�o= " & lEnvio & ")"
                db_StockFisicoEnLocal RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar * -1, iTD, lDoc, 1, lCamion.Tag
            'LE doy al local
                db_StockFisicoEnLocal RsAux!REvArticulo, paEstadoArticuloEntrega, RsAux!REvAEntregar, iTD, lDoc, 2, paCodigoDeSucursal
            End If
        End If
        
'No va m�s pero lo dejo hasta que lo viejo no exista m�s.
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
    objGral.OcurrioError "Error al iniciar la transacci�n.", Err.Description
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
        MsgBox "Atenci�n el env�o seleccionado posee los siguientes documentos pendientes asociados a �l y debe reclamarselos al camionero si los tiene: " & vbCrLf & sQy, vbInformation, "Facturas del env�o"
    End If
    Exit Sub
errEDP:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar si el documento posee documentos pendientes.", Err.Description, "Documentos pendientes"
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
    objGral.OcurrioError "Error al acceder al env�o.", Err.Description, "Env�o"
End Sub

Private Sub loc_SaveTodoEntregado()
On Error GoTo errSTE
    Screen.MousePointer = 11
    Cons = "EXEC repartoDarEntregadoEnvio " & Val(tCodigo.Tag) & ", " & paTipoArticuloServicio & ", " & paCodigoDeUsuario & ", " & paCodigoDeSucursal & ", " & paCodigoDeTerminal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = -1 Then
        MsgBox "Error al dar los env�os como entregados, refresque el c�digo de impresi�n y reintente.", vbExclamation, "Atenci�n"
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
        If MsgBox("�Confirmar pasar el resto de los env�os como entregados?", vbQuestion + vbYesNo, "ATENCI�N") = vbNo Then Exit Sub
        
        'v�lido que no hayan fichas asignadas sin cumplir para alg�n env�o.
        Cons = "Select * From Devolucion " _
            & " Where DevEnvio IN (Select EnvCodigo From Envio Where EnvCodImpresion = " & Val(tCodigo.Tag) & ")" _
            & " And DevLocal Is Null And DevAnulada Is Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            MsgBox "Existen env�os asignados a Fichas de Alta de Stock." & vbCrLf & "Debe cumplir las mismas para poder cumplir los env�os.", vbExclamation, "ATENCI�N"
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
                MsgBox "Existen env�os asignados a Fichas de Alta de Stock." & vbCrLf & "Debe cumplir las mismas para poder cumplir los env�os.", vbExclamation, "ATENCI�N"
                Screen.MousePointer = 0
                Exit Sub
            End If
            RsAux.Close
            
            If fnc_MercaderiaParcial Then
                MsgBox "El cami�n tiene asignada mercader�a parcialmente, si oucrre un error al grabar debe darle toda la mercader�a al cami�n.", vbExclamation, "ATENCI�N"
            End If
            
            If MsgBox("�Confirma dar por entregado al env�o?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
            
            db_SaveEntregado
        End If
    End If
    
End Sub

Private Function fnc_MercaderiaParcial() As Boolean
    fnc_MercaderiaParcial = False
    'Si no tiene toda la mercader�a el cami�n no lo dejo hacer
    'ya que no se si el cami�n ten�a mercader�a o no.
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
    objGral.OcurrioError "Error al limpiar el men� va con.", Err.Description
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
        MsgBox "No existe un env�o pendiente con ese c�digo que pertenezca al c�digo de impresi�n.", vbExclamation, "Atenci�n"
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
    objGral.OcurrioError "Error al buscar el env�o.", Err.Description
End Sub
Private Sub s_GetDatosReparto()
On Error GoTo errGDR
Dim QTotal As Integer, QCamion As Integer
    
    'Busco los datos de la tabla repartoimpresi�n.
    Screen.MousePointer = 11
    bCamionTieneMerc = False
    
    Cons = "Select IsNull(Sum(ReECantidadTotal), 0) as QTotal, IsNull(Sum(ReECantidadEntregada), 0) as QCamion,  CamCodigo, CamNombre " & _
            " From RenglonEntrega, Camion Where ReECodImpresion = " & tCodigo.Text & _
            " And ReECamion = CamCodigo" & _
            " Group by CamCodigo, CamNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lCamion.Caption = "Cami�n: " & Trim(RsAux!CamNombre)
        lCamion.Tag = RsAux!CamCodigo
        QTotal = RsAux!QTotal
        QCamion = RsAux!QCamion
        bCamionTieneMerc = (QCamion > 0)
        RsAux.Close
        tCodigo.Tag = Trim(tCodigo.Text)
    Else
        RsAux.Close
        'Tengo que ver si tengo env�os que tengan s�lo art�culos de servicio.
        Cons = "Select EnvCodigo, CamCodigo, CamNombre From Envio, Camion " & _
                    "Where EnvCodImpresion = " & Val(tCodigo.Text) & _
                    " And EnvEstado = 3 And EnvTipo = 2 And EnvCamion = CamCodigo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsAux.EOF Then
            
            With lCamion
                .Caption = "Cami�n: " & Trim(RsAux!CamNombre)
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
        MsgBox "No se encontraron datos para el c�digo ingresado.", vbExclamation, "Buscar"
        Exit Sub
    End If
    
    If QTotal > QCamion Then
        MsgBox "Al cami�n no se le entreg� " & IIf(QCamion > 0, "la totalidad de ", "") & "la mercader�a, no se podr� dar todo como entregado.", vbExclamation, "Atenci�n"
    Else
        If QTotal = 0 Then MsgBox "No hay mercader�a a entregar.", vbExclamation, "Atenci�n"
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
    objGral.OcurrioError "Error al buscar el c�digo de impresi�n.", Err.Description
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
    If MsgBox("�Confirma quitar del va con al env�o seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar va con") = vbYes Then
        db_IndependizarVaCon
    End If
Exit Sub
errDV:
    objGral.OcurrioError "Error al intentar desvincular el env�o.", Err.Description
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
        Case 2: lHelp.Caption = "El env�o se entrego normalmente."
        Case 1: lHelp.Caption = "El env�o se imprimir� nuevamente en la fecha que indique, se mantiene el camionero."
        Case 0: lHelp.Caption = "El env�o quedar� sin fecha de entrega y en estado a confirmar."
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
        Case 0: lHelp.Caption = "Se pasan todos los env�os del c�digo de impresi�n como entregados."
        Case 1: lHelp.Caption = "Permite seleccionar un env�o del c�digo de impresi�n."
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
    lHelp.Caption = "Ingrese el c�digo de impresi�n a recepcionar."
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    lHelp.Caption = "Ingrese el c�digo de impresi�n a recepcionar."
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
    lHelp.Caption = "Ingrese el c�digo de env�o."
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
                MsgBox "No es un c�digo v�lido.", vbExclamation, "Atenci�n"
            Else
                'Busco el env�o.
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


Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "save": act_Save
        Case "envio": act_Envio
        Case "dividir": act_DividirEnvio
    End Select
End Sub





