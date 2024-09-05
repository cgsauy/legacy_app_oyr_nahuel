VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArbol 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView trTipos 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5106
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ilImage 
      Left            =   5280
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArbol.frx":0000
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArbol.frx":0352
            Key             =   "tipo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArbol.frx":07A4
            Key             =   "open"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmArbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub loc_FillTree(Optional sKeySel As String = "")
On Error GoTo errFT
    Screen.MousePointer = 11
    
    Dim iIdx As Integer
    Dim sQy As String, sPadre As String, sKey As String
    
'    Botones True, False, False, False, False, Toolbar1, Me
    trTipos.Visible = False
    trTipos.Nodes.Clear
    Dim ndSelect As Node
   
'    Do While trTipos.Nodes.Count > 0
 '       trTipos.Nodes.Remove 1
  '  Loop
    
    sQy = "Select TipCodigo, TipNombre, TipHijoDe From Tipo Order by TipHijoDe, TipNombre"
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        sPadre = ""
        If Not IsNull(RsAux("TipHijoDe")) Then If RsAux("TipHijoDe") > 0 Then sPadre = "T" & RsAux("TipHijoDe")
        With trTipos
        
            sKey = "T" & RsAux("TipCodigo")
            If sPadre <> "" Then
                iIdx = .Nodes(sPadre).Index
                .Nodes.Add iIdx, tvwChild, sKey, Trim(RsAux("TipNombre"))
                .Nodes(iIdx).Image = ilImage.ListImages("close").Index
            Else
                .Nodes.Add , , sKey, Trim(RsAux("TipNombre"))
            End If
            .Nodes(sKey).Image = ilImage.ListImages("tipo").Index
            .Nodes(sKey).ExpandedImage = ilImage.ListImages("open").Index
            If Not (.Nodes(sKey).Parent Is Nothing) Then .Nodes(sKey).Parent.Expanded = True
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    
    With trTipos
        If .Nodes.Count > 0 Then
'            Botones True, True, True, False, False, Toolbar1, Me
            If sKeySel = "" Then
                .Nodes.Item(1).Selected = True
            Else
                For Each ndSelect In .Nodes
                    If ndSelect.Key = sKeySel Then
                        ndSelect.Selected = True
                        ndSelect.EnsureVisible
                        Exit For
                    End If
                Next
            End If
        End If
        .Visible = True
    End With
    
    Screen.MousePointer = 0
    Exit Sub
errFT:
    Screen.MousePointer = 11
    trTipos.Visible = True
    clsGeneral.OcurrioError "Error al cargar el árbol.", Err.Description, "Tipo de artículos."
End Sub

Private Sub Form_DblClick()
    
    loc_FillTree
End Sub

Private Sub Form_Load()
trTipos.ImageList = ilImage
    loc_FillTree
End Sub
