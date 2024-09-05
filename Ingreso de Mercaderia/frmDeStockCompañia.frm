VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDeStockCompañia 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Stock"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeStockCompañia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lLista 
      Height          =   4095
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tipo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Número"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fecha"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cantidad"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "A Retirar"
         Object.Width           =   970
      EndProperty
   End
   Begin VB.Label lProveedor 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmDeStockCompañia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gLista As vsFlexGrid                      'Lista de Artículos a cargar
Dim gProveedor As Long                   'Codigo del Proveedor
Dim gNombreProveedor As String      'Nombre del Proveedor
Dim aTexto As String
Dim itmx As ListItem

Public Property Get pProveedor() As Long
    pProveedor = gProveedor
End Property
Public Property Let pProveedor(Codigo As Long)
    gProveedor = Codigo
End Property
Public Property Get pProveedorNombre() As String
    pProveedorNombre = gNombreProveedor
End Property
Public Property Let pProveedorNombre(Texto As String)
    gNombreProveedor = Texto
End Property

Public Property Get pLista() As vsFlexGrid
End Property
Public Property Let pLista(nLista As vsFlexGrid)
    Set gLista = nLista
End Property

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    Screen.MousePointer = 11
    
    'SetearLView lvValores.Grilla Or lvValores.FullRow, lLista
    
    CargoLocalCompania
    lProveedor.Caption = "Proveedor:  " & Trim(gNombreProveedor)
    CargoStock gProveedor
    
End Sub

Private Sub CargoStock(Proveedor As Long)

    On Error GoTo errPago
    cons = "Select * from RemitoCompra, RemitoCompraRenglon, Articulo" _
           & " Where RCoProveedor = " & Proveedor _
           & " And RCoLocal = " & paLocalCompañia _
           & " And RCoCodigo = RCRRemito" _
           & " And RCREnCompania > 0 " _
           & " And RCRArticulo = ArtId"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not rsAux.EOF
        
        Set itmx = lLista.ListItems.Add(, , RetornoNombreDocumento(rsAux!RCoTipo))
        itmx.Tag = rsAux!ArtID & ": " & rsAux!RCRRemito
        
        aTexto = ""
        If Not IsNull(rsAux!RCoSerie) Then aTexto = Trim(rsAux!RCoSerie) & " "
        aTexto = aTexto & rsAux!RCoNumero
        itmx.SubItems(1) = aTexto
        
        itmx.SubItems(2) = Format(rsAux!RCoFecha, "dd/mm/yyyy")
        itmx.SubItems(3) = rsAux!RCREnCompania
        
        itmx.SubItems(4) = Trim(rsAux!ArtNombre)
        itmx.SubItems(5) = "0"
                
        rsAux.MoveNext
    Loop
    rsAux.Close
    Exit Sub
    
errPago:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos en la compañía.", Err.Description
End Sub

Private Sub CargoLocalCompania()

    On Error GoTo errCom
    cons = "Select * from Local Where LocCodigo = " & paLocalCompañia
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not rsAux.EOF Then Me.Caption = Me.Caption & " (En local " & Trim(rsAux!LocNombre) & ")"
    rsAux.Close
    Exit Sub
    
errCom:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del local compañía."
End Sub

Private Sub lLista_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyAdd               'Agrego un articulo a la cantidad ARetirar
            Set itmx = lLista.SelectedItem
            If CCur(itmx.SubItems(5)) = CCur(itmx.SubItems(3)) Then Exit Sub
            itmx.SubItems(5) = CCur(itmx.SubItems(5)) + 1
            
        Case vbKeySubtract        'Quito un articulo a la cantidad ARetirar
            Set itmx = lLista.SelectedItem
            If CCur(itmx.SubItems(5)) = 0 Then Exit Sub
            itmx.SubItems(5) = CCur(itmx.SubItems(5)) - 1
            
        Case vbKeySpace           'Agrego Todos los Articulos
            Set itmx = lLista.SelectedItem
            itmx.SubItems(5) = itmx.SubItems(3)
            
        Case vbKeyDelete           'Quito Todos los Articulos
            Set itmx = lLista.SelectedItem
            itmx.SubItems(5) = "0"
        
        Case vbKeyEscape: Unload Me
        
        Case vbKeyReturn: AccionSalir
    End Select
    
End Sub

Private Sub AccionSalir()

    If MsgBox("Aplica los cambios al formulario de ingreso de mercadría.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Dim aValor As Long
    
    With gLista
        .Rows = 1
        For Each itmx In lLista.ListItems
            If CCur(itmx.SubItems(5)) > 0 Then
                .AddItem Trim(itmx.SubItems(4))
                aValor = ArticuloDelTag(itmx.Tag): .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = itmx.SubItems(5)
                .Cell(flexcpText, .Rows - 1, 2) = RemitoDelTag(itmx.Tag)
            End If
        Next
    End With
    
    Unload Me
    Exit Sub

errAgregar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al agregar el artículo."
End Sub

Private Function ArticuloDelTag(Texto As String) As String
    ArticuloDelTag = Trim(Mid(Texto, 1, InStr(Texto, ":") - 1))
End Function
Private Function RemitoDelTag(Texto As String) As String
    RemitoDelTag = Trim(Mid(Texto, InStr(Texto, ":") + 1, Len(Texto)))
End Function

