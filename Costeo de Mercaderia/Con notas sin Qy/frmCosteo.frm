VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmCosteo 
   Caption         =   "Costeo de Mercaderia"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
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
   ScaleHeight     =   3000
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5535
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   3240
         ScaleHeight     =   405
         ScaleWidth      =   2175
         TabIndex        =   9
         Top             =   180
         Width           =   2175
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   720
            Picture         =   "frmCosteo.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1080
            Picture         =   "frmCosteo.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Quitar filtros."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   1800
            Picture         =   "frmCosteo.frx":04C8
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmCosteo.frx":05CA
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin VB.TextBox tMes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "8/1999"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Mes A Costear:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin ComctlLib.ProgressBar bProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fin del Costeo:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lCosteoF 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio del Costeo:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lCosteoI 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/12/1999 00:00:00"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lAccion 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "frmCosteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum TipoCV
    Compra = 1
    Comercio = 2
    Importacion = 3
End Enum

Dim RsAux As rdoResultset

Private Sub AccionCostear()

Dim aIDCosteo As Long
    
    aIDCosteo = 1
    
    Screen.MousePointer = 11
    
    lAccion.Caption = "Procesando compras del mes...": lAccion.Refresh
    CargoTablaCMCompra CDate(tMes.Text)
    
    lAccion.Caption = "Procesando ventas del mes...": lAccion.Refresh
    CargoTablaCMVenta CDate(tMes.Text)
    
    lAccion.Caption = "Costeando Mercadería...": lAccion.Refresh
    lCosteoI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoI.Refresh
    CargoTablaCMCosteo aIDCosteo
    lCosteoF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoF.Refresh
    lAccion.Caption = "Costeo Finalizado OK.": lAccion.Refresh
    Screen.MousePointer = 0

End Sub

Private Sub CargoTablaCMCompra(Mes As Date)

    '1) Cargar las compras del mes (Credito y Contado)
    '2) Cargar las importes del mes (costeadas en el mes)
Dim aCosto As Currency
Dim QyCom As rdoQuery
    
    Cons = "Insert Into CMCompra (ComFecha, ComArticulo, ComTipo, ComCodigo, ComCantidad, ComCosto, ComQOriginal) Values (?,?,?,?,?,?,?)"
    Set QyCom = cBase.CreateQuery("", Cons)
    
    '1) Compras del Mes (contado y credito)------------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from Compra, CompraRenglon" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraCredito & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        QyCom.rdoParameters(0) = RsAux!ComFecha
        QyCom.rdoParameters(1) = RsAux!CReArticulo
        QyCom.rdoParameters(2) = TipoCV.Compra
        QyCom.rdoParameters(3) = RsAux!ComCodigo
        
        QyCom.rdoParameters(4) = RsAux!CReCantidad
        QyCom.rdoParameters(6) = RsAux!CReCantidad
        
        If RsAux!ComMoneda <> paMonedaPesos Then
            aCosto = RsAux!CRePrecioU * RsAux!ComTC
        Else
            aCosto = RsAux!CRePrecioU
        End If
        QyCom.rdoParameters(5) = aCosto
        
        QyCom.Execute
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '2) Importaciones Costeadas en el Mes-------------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from CosteoCarpeta, CosteoArticulo" _
           & " Where CCaFCosteo Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
           & " And CCaID = CArIDCosteo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        QyCom.rdoParameters(0) = RsAux!CCaFCosteo
        QyCom.rdoParameters(1) = RsAux!CArIdArticulo
        QyCom.rdoParameters(2) = TipoCV.Importacion
        QyCom.rdoParameters(3) = RsAux!CCaID
        
        QyCom.rdoParameters(4) = RsAux!CArCantidad
        QyCom.rdoParameters(5) = RsAux!CArCostoP
        QyCom.Execute
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    QyCom.Close
    
End Sub

Private Sub CargoTablaCMCosteo(IDCosteo As Long)

Dim QyCos As rdoQuery
Dim RsVen As rdoResultset, RsCom As rdoResultset

Dim aFVenta As Date, aArticulo As Long
Dim aQVenta As Long, aQCompra As Long, aQCosteo As Long
Dim aQVentaOriginal As Long
Dim bSalir As Boolean

    Cons = "Insert Into CMCosteo (CosID, CosArticulo, CosTipoCompra, CosCompra, CosCantidad, CosCosto, CosVenta) Values (?,?,?,?,?,?,?)"
    Set QyCos = cBase.CreateQuery("", Cons)
    QyCos.rdoParameters(0) = IDCosteo

    
    '-------------------------------------------------------------------------------------------------------------------
    Cons = "Select Count(*) from CMVenta"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) <> 0 Then bProgress.Max = RsAux(0)
    RsAux.Close
    bProgress.Value = 0
    '-------------------------------------------------------------------------------------------------------------------

    Cons = "Select * from CMVenta Order by VenFecha, VenArticulo, VenTipo, VenCodigo"
    Set RsVen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
       
    Do While Not RsVen.EOF
        bProgress.Value = bProgress.Value + 1
        aFVenta = RsVen!VenFecha
        aArticulo = RsVen!VenArticulo
        aQVenta = RsVen!VenCantidad
        aQVentaOriginal = aQVenta
        
        QyCos.rdoParameters(1) = aArticulo
        
        Do While aQVenta <> 0
            'Voy a la maxima fecha de Compra <= a la fecha de venta ------------------------------------
            Cons = "Select * from CMCompra " _
                   & " Where ComFecha <= '" & Format(aFVenta, sqlFormatoF) & " 23:59:59'" _
                   & " And ComArticulo = " & aArticulo _
                   & " And ComCantidad > 0 " _
                   & " Order by ComFecha DESC"
            Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsCom.EOF Then               'Hay una FC <= FV
                If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                    aQCompra = RsCom!ComCantidad
                    If aQVenta > aQCompra Then
                        aQVenta = aQVenta - aQCompra
                        aQCosteo = aQCompra
                    Else
                        aQCosteo = aQVenta
                        aQVenta = 0
                    End If
                    
                    QyCos.rdoParameters(2) = RsCom!ComTipo
                    QyCos.rdoParameters(3) = RsCom!ComCodigo
                    QyCos.rdoParameters(4) = aQCosteo
                    QyCos.rdoParameters(5) = RsCom!ComCosto
                    QyCos.rdoParameters(6) = RsVen!VenPrecio
                    QyCos.Execute
                    
                    RsCom.Edit
                    RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                    RsCom.Update
                
                Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                              'La cantidad debe ser siempre menor a la original, sino voy al inmediato anterior
                     bSalir = False
                    'Como viene DESC hago move next hasta encontrar uno
                    Do While Not bSalir
                        If RsCom!ComCantidad < RsCom!ComQOriginal Then
                            bSalir = True
                        
                            aQCompra = RsCom!ComQOriginal - RsCom!ComCantidad
                            If Abs(aQVenta) > aQCompra Then
                                aQVenta = (Abs(aQVenta) - aQCompra) * -1
                                aQCosteo = -aQCompra * -1
                            Else
                                aQCosteo = aQVenta
                                aQVenta = 0
                            End If
                            
                            QyCos.rdoParameters(2) = RsCom!ComTipo
                            QyCos.rdoParameters(3) = RsCom!ComCodigo
                            QyCos.rdoParameters(4) = aQCosteo
                            QyCos.rdoParameters(5) = RsCom!ComCosto
                            QyCos.rdoParameters(6) = RsVen!VenPrecio
                            QyCos.Execute
                        
                            RsCom.Edit
                            RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                            RsCom.Update
                            
                        Else
                            RsCom.MoveNext
                            If RsCom.EOF Then bSalir = True
                        End If
                    Loop
                    
                End If
                RsCom.Close

            Else                                        'NO Hay una FC <= FV
                RsCom.Close
                'Voy a la minima fecha de Compra >= a la fecha de venta------------------------------------
                Cons = "Select * from CMCompra " _
                       & " Where ComFecha >= '" & Format(aFVenta, sqlFormatoF) & " 23:59:59'" _
                       & " And ComArticulo = " & aArticulo _
                       & " And ComCantidad > 0 " _
                       & " Order by ComFecha"
                Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsCom.EOF Then               'Hay una FC >= FV
                
                    If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                        aQCompra = RsCom!ComCantidad
                        If aQVenta > aQCompra Then
                            aQVenta = aQVenta - aQCompra
                            aQCosteo = aQCompra
                        Else
                            aQCosteo = aQVenta
                            aQVenta = 0
                        End If
                        
                        QyCos.rdoParameters(2) = RsCom!ComTipo
                        QyCos.rdoParameters(3) = RsCom!ComCodigo
                        QyCos.rdoParameters(4) = aQCosteo
                        QyCos.rdoParameters(5) = RsCom!ComCosto
                        QyCos.rdoParameters(6) = RsVen!VenPrecio
                        QyCos.Execute
                        
                        RsCom.Edit
                        RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                        RsCom.Update
                    
                    Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                              'La cantidad debe ser siempre menor a la original, sino voy al inmediato siguiente
                        bSalir = False
                        Do While Not bSalir
                            If RsCom!ComCantidad < RsCom!ComQOriginal Then
                                bSalir = True
                            
                                aQCompra = RsCom!ComQOriginal - RsCom!ComCantidad
                                If Abs(aQVenta) > aQCompra Then
                                    aQVenta = (Abs(aQVenta) - aQCompra) * -1
                                    aQCosteo = -aQCompra * -1
                                Else
                                    aQCosteo = aQVenta
                                    aQVenta = 0
                                End If
                                
                                QyCos.rdoParameters(2) = RsCom!ComTipo
                                QyCos.rdoParameters(3) = RsCom!ComCodigo
                                QyCos.rdoParameters(4) = aQCosteo
                                QyCos.rdoParameters(5) = RsCom!ComCosto
                                QyCos.rdoParameters(6) = RsVen!VenPrecio
                                QyCos.Execute
                            
                                RsCom.Edit
                                RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                                RsCom.Update
                                
                            Else
                                RsCom.MoveNext
                                If RsCom.EOF Then bSalir = True
                            End If
                        Loop
                    
                    RsCom.Close
                    End If
                    
                Else
                    RsCom.Close
                    'Si no hay datos queda remanente, Primero updateo con lo que queda remanente en la venta
                    If aQVenta <> aQVentaOriginal Then
                        Cons = "Update CMVenta Set VenCantidad = " & aQVenta _
                                & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                                & " And VenArticulo = " & RsVen!VenArticulo _
                                & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
                        cBase.Execute Cons
                    End If
                    Exit Do
                End If
            End If
        Loop
        
        'Si la venta quedó en cero elimino el registro de la venta
        If aQVenta = 0 Then
            Cons = " Delete CMVenta " _
                    & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                    & " And VenArticulo = " & RsVen!VenArticulo _
                    & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
            cBase.Execute Cons
        End If
        RsVen.MoveNext
    Loop
    
    RsVen.Close
    
    
    'Hay que borrar las compras en que las cantidades son iguales a 0
    Cons = "Delete CMCompra Where ComCantidad = 0"
    cBase.Execute Cons
    '---------------------------------------------------------------------------
    bProgress.Value = 0
    
End Sub

Private Sub bConsultar_Click()
    AccionCostear
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    lAccion.Caption = ""
    lCosteoI.Caption = "": lCosteoF.Caption = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set msgError = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub CargoTablaCMVenta(Mes As Date)
'Parámetro: Recibe el primer día del mes a costear.
On Error GoTo ErrCTCMV
Dim QyVta As rdoQuery
Dim strDocumentos As String
Dim Monedas As String
Dim aCosto As Currency

    Screen.MousePointer = 11
    
    Monedas = RetornoTCMonedas(PrimerDia(Mes))
    'Creo query para insertar los datos
    Cons = "Insert Into CMVenta (VenFecha, VenArticulo, VenTipo, VenCodigo, VenCantidad,VenPrecio) Values (?,?,?,?,?,?)"
    Set QyVta = cBase.CreateQuery("", Cons)
    
    'Primer Paso Copio las Ventas---------------------------------------
    'Traigo los documentos Ctdo y Cred, Nota Esp, Nota de Cred. y  Nota de Dev. que no estén anulados
    strDocumentos = TipoDocumento.Contado & ", " & TipoDocumento.Credito _
        & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial
    
    Cons = "Select DocFecha, DocMoneda, DocTipo, Renglon.* From Documento, Renglon" _
        & " Where DocTipo IN (" & strDocumentos & ")" _
        & " And DocFecha BetWeen '" & Format(PrimerDia(Mes) & " 00:00:00", sqlFormatoFH) & "'" _
        & " And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
        & " And DocAnulado = 0 And DocCodigo = RenDocumento"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)

    Do While Not RsAux.EOF
        QyVta.rdoParameters(0) = Format(RsAux!DocFecha, sqlFormatoFH)
        QyVta.rdoParameters(1) = RsAux!RenArticulo
        QyVta.rdoParameters(2) = TipoCV.Comercio
        QyVta.rdoParameters(3) = RsAux!RenDocumento
        
        Select Case RsAux!DocTipo
            Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                            QyVta.rdoParameters(4) = RsAux!RenCantidad * -1
                            
            Case Else: QyVta.rdoParameters(4) = RsAux!RenCantidad
        End Select
        
        'Es el precio neto
        QyVta.rdoParameters(5) = (RsAux!RenPrecio - RsAux!RenIva) * ValorTC(RsAux!DocMoneda, Monedas)
        
        QyVta.Execute
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Segundo paso Cargo Notas de Compras.
    Cons = "Select * from Compra, CompraRenglon" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(Mes, sqlFormatoFH) & "' And '" & Format(UltimoDia(Mes) & " 23:59:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    Do While Not RsAux.EOF
        QyVta.rdoParameters(0) = Format(RsAux!ComFecha, sqlFormatoFH)
        QyVta.rdoParameters(1) = RsAux!CReArticulo
        QyVta.rdoParameters(2) = TipoCV.Compra
        QyVta.rdoParameters(3) = RsAux!CReCompra
        QyVta.rdoParameters(4) = RsAux!CReCantidad
        If RsAux!ComMoneda <> paMonedaPesos Then
            aCosto = RsAux!CRePrecioU * RsAux!ComTC
        Else
            aCosto = RsAux!CRePrecioU
        End If
        QyVta.rdoParameters(5) = aCosto
        QyVta.Execute
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    QyVta.Close 'Cierro query
    Screen.MousePointer = 0
    Exit Sub
    
ErrCTCMV:
    clsGeneral.OcurrioError "Ocurrio un error al cargar la tabla de Ventas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValorTC(PosMoneda As Integer, ByVal strMonedas As String) As Currency
Dim Cont As Integer
    
    Cont = 1: ValorTC = 1
    Do While strMonedas <> ""
        If PosMoneda = Cont Then
            ValorTC = Mid(strMonedas, 1, InStr(1, strMonedas, ":") - 1)
            Exit Function
        Else
            strMonedas = Mid(strMonedas, InStr(1, strMonedas, ":") + 1, Len(strMonedas))
            Cont = Cont + 1
        End If
    Loop
End Function

Private Function RetornoTCMonedas(Fecha As Date) As String

Dim aTC As Currency, Contador As Integer

    'Armo vector con las TC de las monedas que existen.
    RetornoTCMonedas = ""
    Cons = "Select * From Moneda"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Contador = 1
    Do While Not RsAux.EOF
        If Contador = RsAux!MonCodigo Then
            If RsAux!MonCodigo = paMonedaPesos Then
                aTC = 1
            Else
                aTC = TasadeCambio(RsAux!MonCodigo, paMonedaPesos, Fecha)
            End If
        Else
            aTC = 1
        End If
        If RetornoTCMonedas = "" Then RetornoTCMonedas = aTC Else RetornoTCMonedas = RetornoTCMonedas & ":" & aTC
        Contador = Contador + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
End Function

