VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmWizGasto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Depósitos"
   ClientHeight    =   5400
   ClientLeft      =   3945
   ClientTop       =   2625
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizGasto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame f2 
      Height          =   75
      Left            =   -60
      TabIndex        =   22
      Top             =   4860
      Width           =   6255
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton bPrevious 
      Caption         =   "Anterior"
      Height          =   315
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton bNext 
      Caption         =   "Siguiente"
      Height          =   315
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox tDIImporte 
      Appearance      =   0  'Flat
      Height          =   305
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3120
      Width           =   1275
   End
   Begin VB.TextBox tADImporte 
      Appearance      =   0  'Flat
      Height          =   305
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox tDISubRubro 
      Appearance      =   0  'Flat
      Height          =   305
      Left            =   1080
      TabIndex        =   9
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox tDIRubro 
      Appearance      =   0  'Flat
      Height          =   305
      Left            =   1080
      TabIndex        =   7
      Top             =   3600
      Width           =   3495
   End
   Begin AACombo99.AACombo cADTipoMov 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Width           =   2835
      _ExtentX        =   5001
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
   End
   Begin AACombo99.AACombo cADDispDesde 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   2835
      _ExtentX        =   5001
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
   End
   Begin AACombo99.AACombo cADDispA 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   1980
      Width           =   2835
      _ExtentX        =   5001
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
   End
   Begin AACombo99.AACombo cDIDispA 
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   4380
      Width           =   2835
      _ExtentX        =   5001
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
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada de Caja de"
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   3180
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Por un Valor de"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Para los Cheques Diferidos se hace un Gasto del tipo ""Entrada de Caja"" con el siguiente formato:"
      Height          =   495
      Left            =   60
      TabIndex        =   17
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "A la Disponibilidad"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Subrubro:"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   4020
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde Rubro:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "A la Disponibilidad"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde "
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Movimiento"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Para los Cheques del Día se hace una Transferencia con el siguiente formato:"
      Height          =   375
      Left            =   60
      TabIndex        =   16
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label lSucursal 
      BackStyle       =   0  'Transparent
      Caption         =   "Depósitos en "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   5895
   End
End
Attribute VB_Name = "frmWizGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmOK  As Boolean

Private Sub bCancel_Click()
    prmOK = False
    Unload Me
End Sub

Private Sub bNext_Click()
    
    If Not ValidoFicha Then Exit Sub
    
    If Val(bNext.Tag) > UBound(dGastos) Then
        prmOK = True
        ActualizoParametrosDepositos
        Unload Me
    Else
        ArmoPantalla
    End If
    
End Sub

Private Sub bPrevious_Click()
    If Not ValidoFicha Then Exit Sub
    
    bNext.Tag = Val(bNext.Tag) - 2
    ArmoPantalla
End Sub

Private Sub cADDispA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cADDispA.ListIndex <> -1 Then Foco tDIRubro
End Sub

Private Sub cADDispDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cADDispDesde.ListIndex <> -1 Then Foco cADDispA
End Sub

Private Sub cADTipoMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cADTipoMov.ListIndex <> -1 Then Foco cADDispDesde
End Sub

Private Sub cDIDispA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cDIDispA.ListIndex <> -1 Then bNext.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next

    Me.BackColor = RGB(255, 248, 220)
    f2.BackColor = Me.BackColor
    
    bCancel.BackColor = Me.BackColor
    bPrevious.BackColor = bCancel.BackColor
    bNext.BackColor = bCancel.BackColor
    
    prmOK = False
    
    bNext.Tag = 0
    ArmoPantalla
    
End Sub

Private Sub InicializoForm(idMoneda As Long)

On Error Resume Next

    If cADTipoMov.ListCount = 0 Then
        Cons = "Select TMDCodigo, TMDNombre From TipoMovDisponibilidad " & _
                    " Where TMDTransferencia = 1 Order by TMDNombre"
        CargoCombo Cons, cADTipoMov
    End If
    
    Cons = "Select DisId, DisNombre From Disponibilidad " & _
                " Where DisSucursal is Not Null" & _
                " And DisMoneda = " & idMoneda & _
                " Order by DisNombre"

    CargoCombo Cons, cADDispA
    CargoCombo Cons, cDIDispA
    
    Cons = "Select DisId, DisNombre From Disponibilidad " & _
                " Where DisSucursal is Null" & _
                " And DisMoneda = " & idMoneda & _
                " Order by DisNombre"
     CargoCombo Cons, cADDispDesde
     
End Sub

Private Function LimpioFicha()

Dim bkColor As Long

    bkColor = RGB(255, 250, 240)
        
    cADTipoMov.Text = "": cADTipoMov.BackColor = bkColor
    tADImporte.Text = "": tADImporte.BackColor = bkColor
    cADDispDesde.Text = "": cADDispDesde.BackColor = bkColor
    cADDispA.Text = "": cADDispA.BackColor = bkColor
    
    tDIImporte.Text = "": tDIImporte.BackColor = bkColor
    tDIRubro.Text = "": tDIRubro.BackColor = bkColor
    tDISubRubro.Text = "": tDISubRubro.BackColor = bkColor
    cDIDispA.Text = "": cDIDispA.BackColor = bkColor
    
    cADTipoMov.Enabled = True
    tADImporte.Enabled = True
    cADDispDesde.Enabled = True
    cADDispA.Enabled = True
    
    tDIImporte.Enabled = True
    tDIRubro.Enabled = True
    tDISubRubro.Enabled = True
    cDIDispA.Enabled = True
    
End Function

Private Sub ArmoPantalla()

    On Error Resume Next
    LimpioFicha
    
    Dim mIdx As Integer: mIdx = Val(bNext.Tag)
    
    If UBound(dGastos) = mIdx Then bNext.Caption = "Finalizar" Else bNext.Caption = "Siguiente"
    bPrevious.Enabled = Not (mIdx = 0)
    
    bNext.Tag = mIdx + 1
    
    With dGastos(mIdx)
        InicializoForm .idMoneda
        
        lSucursal.Caption = " Depósitos en " & .SucursalNombre
        
        BuscoCodigoEnCombo cADTipoMov, .IdTipoTransferencia
        tADImporte.Text = Format(.ImporteAlDia, "#,##0.00")
        BuscoCodigoEnCombo cADDispDesde, .IdDisponibilidadSalida
        BuscoCodigoEnCombo cADDispA, .IdDisponibilidadEntrada
        
        tDIImporte.Text = Format(.ImporteDiferido, "#,##0.00")
        tDIRubro.Text = Trim(.NameRSalida): tDIRubro.Tag = .IdRubroSalida
        tDISubRubro.Text = Trim(.NameSRSalida): tDISubRubro.Tag = .IdSubrubroSalida
        BuscoCodigoEnCombo cDIDispA, .IdDisponibilidadEntrada
        
        If .ImporteAlDia = 0 Then
            cADTipoMov.Enabled = False: cADTipoMov.BackColor = Colores.Gris
            tADImporte.Enabled = False: tADImporte.BackColor = Colores.Gris
            cADDispDesde.Enabled = False: cADDispDesde.BackColor = Colores.Gris
            cADDispA.Enabled = False: cADDispA.BackColor = Colores.Gris
        End If
        
        If .ImporteDiferido = 0 Then
            tDIImporte.Enabled = False: tDIImporte.BackColor = Colores.Gris
            tDIRubro.Enabled = False: tDIRubro.BackColor = Colores.Gris
            tDISubRubro.Enabled = False: tDISubRubro.BackColor = Colores.Gris
            cDIDispA.Enabled = False: cDIDispA.BackColor = Colores.Gris
        End If
        
        If .ImporteAlDia <> 0 Then
            If cADTipoMov.ListIndex = -1 Then Foco cADTipoMov
        Else
            If .ImporteDiferido <> 0 Then If Val(tDISubRubro.Tag) = 0 Then Foco tDISubRubro
        End If
        
    End With
    
End Sub

Private Function ValidoFicha() As Boolean
On Error GoTo errValido

    ValidoFicha = False
    
    Dim mIdx As Integer
    mIdx = Val(bNext.Tag) - 1
    
    With dGastos(mIdx)
        If .ImporteAlDia <> 0 Then
            If cADTipoMov.ListIndex = -1 Then Foco cADTipoMov: Exit Function
            If cADDispDesde.ListIndex = -1 Then Foco cADDispDesde: Exit Function
            If cADDispA.ListIndex = -1 Then Foco cADDispA: Exit Function
            
            .IdTipoTransferencia = cADTipoMov.ItemData(cADTipoMov.ListIndex)
            .IdDisponibilidadSalida = cADDispDesde.ItemData(cADDispDesde.ListIndex)
        
            .IdDisponibilidadEntrada = cADDispA.ItemData(cADDispA.ListIndex)
        End If
        
        If .ImporteDiferido <> 0 Then
            If Val(tDIRubro.Tag) = 0 Then Foco tDIRubro: Exit Function
            If Val(tDISubRubro.Tag) = 0 Then Foco tDISubRubro: Exit Function
            If cDIDispA.ListIndex = -1 Then Foco cDIDispA: Exit Function
            
            .NameRSalida = Trim(tDIRubro.Text)
            .IdRubroSalida = Val(tDIRubro.Tag)
            
            .NameSRSalida = Trim(tDISubRubro.Text)
            .IdSubrubroSalida = Val(tDISubRubro.Tag)
            .IdDisponibilidadEntrada = cDIDispA.ItemData(cDIDispA.ListIndex)
        End If
        
        If .ImporteAlDia <> 0 And .ImporteDiferido <> 0 Then
            If cADDispA.ItemData(cADDispA.ListIndex) <> cDIDispA.ItemData(cDIDispA.ListIndex) Then
                MsgBox "La disponibilidad destino de la Transferencia y del Gasto debe ser la misma.", vbExclamation, "Disponibilidad A"
                Foco cADDispA: Exit Function
            End If
        End If
    End With
    
    ValidoFicha = True
    Exit Function

errValido:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
End Function

Private Sub tDIRubro_Change()
    
    If Val(tDIRubro.Tag) <> 0 Then
        tDIRubro.Tag = 0
        If Val(tDISubRubro.Tag) <> 0 Then tDISubRubro.Text = ""
    End If
    
End Sub
Private Sub tDIRubro_GotFocus()
    tDIRubro.SelStart = 0: tDIRubro.SelLength = Len(tDIRubro.Text)
End Sub

Private Sub tDIRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If KeyAscii = vbKeyReturn Then
        If Val(tDIRubro.Tag) <> 0 Then Foco tDISubRubro: Exit Sub
        If Trim(tDIRubro.Text) = "" Then Foco tDISubRubro: Exit Sub
    
        ing_BuscoRubro tDIRubro
        
        Exit Sub
    End If
    Exit Sub

errBS:
    clsGeneral.OcurrioError "Error al buscar el rubro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tDISubRubro_Change()
    tDISubRubro.Tag = 0
End Sub

Private Sub tDISubRubro_GotFocus()
    tDISubRubro.SelStart = 0: tDISubRubro.SelLength = Len(tDISubRubro.Text)
End Sub

Private Sub tDISubRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If KeyAscii = vbKeyReturn Then
        If Val(tDISubRubro.Tag) <> 0 Then Foco cDIDispA: Exit Sub
        
        If Trim(tDISubRubro.Text) <> "" And Val(tDISubRubro.Tag) = 0 Then
            ing_BuscoSubrubro tDIRubro, tDISubRubro
            Exit Sub
        End If
    End If
    Exit Sub

errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ActualizoParametrosDepositos()
On Error GoTo errAct
Dim bChange As Boolean
    
    Screen.MousePointer = 11
    
    Dim I As Integer, mParams As String
    For I = LBound(dGastos) To UBound(dGastos)
        With dGastos(I)
            mParams = Trim(.zPrmsDepositos)
            bChange = False
            bChange = get_PrmsDepositos(.zPrmsDepositos, TipoTranferencia) <> .IdTipoTransferencia
            If Not bChange Then bChange = get_PrmsDepositos(.zPrmsDepositos, SRGasto) <> .IdSubrubroSalida
            If Not bChange Then bChange = (get_PrmsDepositos(.zPrmsDepositos, DispDeposito) <> .IdDisponibilidadEntrada And get_PrmsDepositos(.zPrmsDepositos, DispDepositoFletes) <> .IdDisponibilidadEntrada)
            
            If bChange Then
                
                mParams = put_PrmsDepositos(mParams, enuPrmsDepositos.DispDepositoFletes, get_PrmsDepositos(.zPrmsDepositos, DispDepositoFletes))
                mParams = put_PrmsDepositos(mParams, enuPrmsDepositos.DispDeposito, .IdDisponibilidadEntrada)
                mParams = put_PrmsDepositos(mParams, enuPrmsDepositos.SRGasto, .IdSubrubroSalida)
                mParams = put_PrmsDepositos(mParams, enuPrmsDepositos.TipoTranferencia, .IdTipoTransferencia)
                
'            End If
'            If Trim(mParams) <> Trim(.zPrmsDepositos) Then

                Cons = "Select * from SucursalDeBanco Where SBaCodigo = " & .SucursalID
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    RsAux.Edit
                    RsAux!SBaPrmsDepositos = mParams
                    RsAux.Update
                End If
                RsAux.Close
            End If
            
        End With
    Next
    Screen.MousePointer = 0
    Exit Function
    
errAct:
    clsGeneral.OcurrioError "Error al actualizar valores por defecto del depósitos.", Err.Description
    Screen.MousePointer = 0
End Function
