VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migrar Contenedores"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bMigrar 
      Caption         =   "&Migrar"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar pbProgreso 
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Total de Embarques:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bMigrar_Click()
    If MsgBox("¿Confirma migrar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    bMigrar.Enabled = False
    pbProgreso.Value = 0
    Migro
End Sub

Private Sub Form_Load()
    pbProgreso.Value = 0
    Cons = "Select count(*) From Embarque Where EmbContenedor Is Not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Label1.Caption = Label1.Caption & " " & RsAux(0)
    pbProgreso.Max = RsAux(0) + 1
    RsAux.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
End Sub

Private Sub Migro()
On Error GoTo errM
Dim rsE As rdoResultset
    Screen.MousePointer = 11
    Cons = "Select * From Embarque Where EmbContenedor Is Not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        'Veo si ya lo inserte
        Cons = "Select * From EmbarqueContenedor Where ECoEmbarque = " & RsAux!EmbID
        Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rsE.EOF Then
            rsE.AddNew
            rsE!ECoEmbarque = RsAux!EmbID
            rsE!ECoContenedor = RsAux!EmbContenedor
            If IsNull(RsAux!EmbCantContenedor) Then
                rsE!ECoCantidad = 1
            Else
                rsE!ECoCantidad = RsAux!EmbCantContenedor
            End If
            rsE.Update
        End If
        rsE.Close
        pbProgreso.Value = pbProgreso.Value + 1
        Me.Refresh
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    MsgBox "La migración se completo.", vbInformation, "ATENCIÓN"
    pbProgreso.Value = 0
    Exit Sub
errM:
    Screen.MousePointer = 0
    MsgBox "Ocurrió el siguiente error: " & Err.Description, vbCritical, "ATENCIÓN"
End Sub
