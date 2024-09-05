VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmDetCobro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobro "
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3165
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tQ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox tDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton bOK 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsDocumento 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2143
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
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
   Begin VB.TextBox tInstalacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Q:"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "En &otros documentos:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "En &instalación:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDetCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Lista As Integer

Public bSeteo As Boolean

Public prmDocInstalacion As Long
Public prmCliente As Long
Public prmInstalacion As Long
Public prmIDArticuloInstalacion As Long

Public prmQNecesitoCobrar As Integer
Public prmQInstalacion As Integer
Public prmEnOtrosDocumentos As String

Public Property Get QEnLista() As Integer
    QEnLista = m_Lista
End Property

Private Function f_GetTotalLista() As Integer
Dim iCont As Integer
    f_GetTotalLista = 0
    For iCont = 1 To vsDocumento.Rows - 1
        f_GetTotalLista = f_GetTotalLista + Val(vsDocumento.Cell(flexcpValue, iCont, 0))
    Next
End Function

Private Function f_GetQUsados(ByVal sDato As String, ByVal lDoc As Long) As Integer
Dim sDoc() As String
Dim iCont As Integer
    
    f_GetQUsados = 0
    sDoc = Split(sDato, "|")
    If UBound(sDoc) = 1 Then
        sDoc = Split(sDoc(1), ";")
        For iCont = 0 To UBound(sDoc)
            If Val(Mid(sDoc(iCont), InStr(1, sDoc(iCont), ":") + 1)) = lDoc Then
                f_GetQUsados = Val(Mid(sDoc(iCont), 1, InStr(1, sDoc(iCont), ":", vbTextCompare) - 1)) + f_GetQUsados
            End If
        Next
    End If
    
End Function

Private Sub s_ArmoGrillaDocumentos()
On Error GoTo errAGD
Dim sDoc() As String
Dim iCont As Integer
Dim rsD As rdoResultset
        
    sDoc = Split(prmEnOtrosDocumentos, ";")
        
    For iCont = 0 To UBound(sDoc)
        'Tengo cantidad en otros documentos.
        With vsDocumento
            .AddItem Val(Mid(sDoc(iCont), 1, InStr(1, sDoc(iCont), ":", vbTextCompare) - 1))
            .Cell(flexcpData, .Rows - 1, 0) = Val(Mid(sDoc(iCont), InStr(1, sDoc(iCont), ":", vbTextCompare) + 1))
            'Busco los datos del documento.
            Cons = "Select DocSerie, DocNumero, DocTipo From Documento Where DocCodigo = " & .Cell(flexcpData, .Rows - 1, 0)
            Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsD.EOF Then
                .Cell(flexcpText, .Rows - 1, 1) = IIf(rsD!DocTipo = 1, "Ctdo ", "Créd ") & Trim(rsD!DocSerie) & " " & Trim(rsD!DocNumero)
            End If
            rsD.Close
        End With
    Next
    Exit Sub
errAGD:
    clsGeneral.OcurrioError "Error al armar la grilla con los documentos.", Err.Description
End Sub

Private Sub bCancel_Click()
    Unload Me
End Sub

Private Sub bOK_Click()
Dim iQLista As Integer
On Error GoTo errBOK
    If Not IsNumeric(tInstalacion.Text) And Trim(tInstalacion.Text) <> "" Then
        MsgBox "Formato incorrecto.", vbExclamation, "Atención"
        tInstalacion.SetFocus
        Exit Sub
    ElseIf Val(tInstalacion.Text) < 0 Then
        MsgBox "Valor negativo.", vbCritical, "Atención"
        tInstalacion.SetFocus
        Exit Sub
    End If
    
    If f_GetTotalLista + Val(tInstalacion.Text) > prmQNecesitoCobrar Then
        MsgBox "Se excedió de la cantidad necesaria.", vbExclamation, "Atención"
        Exit Sub
    End If
   
   If MsgBox("¿Confirma los valores ingresados?", vbQuestion + vbYesNo, "Confirmar") = vbNo Then tInstalacion.SetFocus: Exit Sub
    
    prmQInstalacion = Val(tInstalacion.Text)
    m_Lista = f_GetTotalLista
    
    'Armo el string
    Dim iCont As Integer
    prmEnOtrosDocumentos = ""
    For iCont = 1 To vsDocumento.Rows - 1
        If prmEnOtrosDocumentos <> "" Then prmEnOtrosDocumentos = prmEnOtrosDocumentos & ";"
        prmEnOtrosDocumentos = prmEnOtrosDocumentos & vsDocumento.Cell(flexcpValue, iCont, 0) & ":" & vsDocumento.Cell(flexcpData, iCont, 0)
    Next
    bSeteo = True
    Unload Me
    Exit Sub
errBOK:
    clsGeneral.OcurrioError "Error al validar y setear los valores.", Err.Description
End Sub

Private Sub Form_Load()
    
    bSeteo = False
    With vsDocumento
        .Rows = 1
        .Cols = 2
        .ExtendLastCol = True
        .FormatString = "Cantidad|Documento"
    End With
    
    If prmQInstalacion > 0 Then tInstalacion.Text = prmQInstalacion
    If prmEnOtrosDocumentos <> "" Then s_ArmoGrillaDocumentos
    
    Me.Caption = Me.Caption & "(necesarios " & prmQNecesitoCobrar & ")"
End Sub

Private Sub Label1_Click()
    Foco tInstalacion
End Sub

Private Sub tDoc_Change()
    tDoc.Tag = ""
End Sub

Private Sub tDoc_GotFocus()
    With tDoc
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDoc_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
Dim sDoc As String, sSerie As String, sNro As String
Dim rsD As rdoResultset
Dim iCont As Integer

    If KeyAscii = vbKeyReturn Then
        If Val(tDoc.Tag) > 0 Then
            tQ.SetFocus
        ElseIf Trim(tDoc.Text) = "" Then
            bOK.SetFocus
        Else
            sDoc = Trim(tDoc.Text)
            If InStr(sDoc, "-") <> 0 Then
                sSerie = Mid(sDoc, 1, InStr(sDoc, "-") - 1)
                sNro = Val(Mid(sDoc, InStr(sDoc, "-") + 1))
            Else
                sSerie = Mid(sDoc, 1, 1)
                sNro = Val(Mid(sDoc, 2))
            End If
            tDoc.Text = UCase(sSerie) & "-" & sNro
            
            '................................................................
            'Busco el documento.
            Cons = "Select * From Documento, Renglon " _
                & "Where DocCodigo <> " & prmDocInstalacion & " And DocSerie = '" & sSerie & "' And DocNumero = " & sNro _
                & " And DocTipo In (1,2)" _
                & " And RenArticulo = " & prmIDArticuloInstalacion & " And DocAnulado = 0" _
                & " And DocCodigo = RenDocumento"
            Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsD.EOF Then
                MsgBox "No se encontró un documento con esos datos y que además contenga al artículo instalación y no este anulado.", vbInformation, "Atención"
            Else
                If rsD!DocCliente <> prmCliente Then
                    If MsgBox("El cliente del documento no es el mismo del documento de los artículos a instalar." & vbCrLf & "¿Esta seguro que este es el documento para cobrar esta instalación?", vbQuestion + vbYesNo, "Distintos Clientes") = vbNo Then
                        rsD.Close
                        Exit Sub
                    End If
                End If
                tQ.Tag = rsD!RenCantidad
                With tDoc
                    .Text = IIf(rsD!DocTipo = 1, "Ctdo ", "Créd ") & UCase(sSerie) & "-" & sNro
                    .Tag = rsD!DocCodigo
                End With
                tQ.SetFocus
            End If
            rsD.Close
            
            If Val(tDoc.Tag) = 0 Then Exit Sub
            
            'Válido que no este insertada en la lista.
            For iCont = 1 To vsDocumento.Rows - 1
                If vsDocumento.Cell(flexcpData, iCont, 0) = Val(tDoc.Tag) Then
                    MsgBox "El documento está insertado en la lista.", vbInformation, "Atención"
                    tQ.Tag = 0
                    Exit Sub
                End If
            Next iCont
            If Val(tQ.Tag) > 0 Then
                'Busco si ya lo inserto en otro artículo de la lista del formulario de instalaciones.
                With frmInstall.vsArticulo
                    For iCont = 0 To .Rows - 1
                        If .Row <> iCont Then
                            If .Cell(flexcpText, iCont, 7) <> "" Then
                                tQ.Tag = Val(tQ.Tag) - f_GetQUsados("|" & .Cell(flexcpText, iCont, 7), Val(tDoc.Tag))
                            End If
                        End If
                    Next
                End With
            End If
            
            If Val(tQ.Tag) > 0 Then
                'Busco si este documento no esta en alguna instalación.
                'En el cobro tengo que tener el valor de la instalacion + alguno de documento --> va el |
                Cons = "Select * From RenglonInstalacion, Instalacion " _
                    & " Where InsID <> " & prmInstalacion & " And InsAnulada Is null " _
                    & " And RInCobro Like '%|%:" & tDoc.Tag & "%' And InsID = RInInstalacion"
                Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                Do While Not rsD.EOF
                    tQ.Tag = Val(tQ.Tag) - f_GetQUsados(rsD!RInCobro, Val(tDoc.Tag))
                    'Tomo las cantidades.
                    rsD.MoveNext
                Loop
                rsD.Close
            End If
            If Val(tQ.Tag) > 0 Then
                If Val(tQ.Tag) > prmQNecesitoCobrar - Val(tInstalacion.Text) Then
                    tQ.Text = prmQNecesitoCobrar - Val(tInstalacion.Text)
                Else
                    tQ.Text = Val(tQ.Tag)
                End If
                tQ.SetFocus
            Else
                MsgBox "No quedan artículos disponibles para asignar a esta instalación.", vbInformation, "Atención"
            End If
            '................................................................
        End If
    End If
    Exit Sub
errKP:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
End Sub

Private Sub tInstalacion_GotFocus()
    With tInstalacion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tInstalacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tDoc.SetFocus
End Sub

Private Sub tQ_GotFocus()
    Foco tQ
End Sub

Private Sub tQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tQ.Tag) > 0 Then
            If IsNumeric(tQ.Tag) Then
                If Val(tQ.Tag) >= Val(tQ.Text) And Val(tQ.Text) > 0 Then
                    With vsDocumento
                        .AddItem tQ.Text
                        .Cell(flexcpData, .Rows - 1, 0) = Val(tDoc.Tag)
                        .Cell(flexcpText, .Rows - 1, 1) = tDoc.Text
                    End With
                    tDoc.Text = ""
                    tQ.Text = "": tQ.Tag = ""
                Else
                    MsgBox "Cantidad incorrecta.", vbExclamation, "Atención"
                End If
            Else
                bOK.SetFocus
            End If
        Else
            bOK.SetFocus
        End If
    End If
End Sub

Private Sub vsDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyDelete Then
        If vsDocumento.Row >= vsDocumento.FixedRows Then vsDocumento.RemoveItem vsDocumento.Row
    End If
End Sub

