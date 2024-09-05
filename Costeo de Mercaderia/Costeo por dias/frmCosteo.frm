VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmCosteo 
   Caption         =   "Costeo de Mercaderia"
   ClientHeight    =   5175
   ClientLeft      =   2925
   ClientTop       =   2295
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCosteo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6480
   Begin VB.Frame Frame2 
      Caption         =   "Información del Último Costeo"
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6255
      Begin VB.Label lCUsuario 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label lCFecha 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label lCMes 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diciembre 2000"
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Mes Costeado:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox tMesF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3300
         TabIndex        =   30
         Text            =   "8/1999"
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   5160
         ScaleHeight     =   405
         ScaleWidth      =   975
         TabIndex        =   3
         Top             =   180
         Width           =   975
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   600
            Picture         =   "frmCosteo.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmCosteo.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin VB.TextBox tMesI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "8/1999"
         Top             =   240
         Width           =   1575
      End
      Begin ComctlLib.ProgressBar bProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Generación de Ventas"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Generación de Compras"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin :"
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lVentaF 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   26
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lVentaI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/12/1999 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lCompraF 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lCompraI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/12/1999 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   1800
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
         TabIndex        =   18
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label lCosteoI 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/12/1999 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio del Costeo:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2760
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
         TabIndex        =   15
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin del Costeo:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "&Mes A Costear:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   4920
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3228
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
    
    If Not ValidoCampos Then Exit Sub
    If MsgBox("Confirma realizar el costeo de mercadería desde el " & tMesI.Text & " al " & tMesF.Text, vbQuestion + vbYesNo, "Costear Mercadería") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errCostear
    
    FechaDelServidor
    
    '-------------------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from CMCabezal"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!CabMesCosteo = Format(tMesI.Text, sqlFormatoF)
    RsAux!CabFecha = Format(gFechaServidor, sqlFormatoFH)
    RsAux!CabUsuario = paCodigoDeUsuario
    RsAux.Update: RsAux.Close
    
    Cons = "Select Max(CabID) from CMCabezal"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aIDCosteo = RsAux(0)
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    lAccion.Caption = "Procesando compras del mes...": lAccion.Refresh
    lCompraI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCompraI.Refresh
    CargoTablaCMCompra CDate(tMesI.Text), CDate(tMesF.Text)
    lCompraF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCompraF.Refresh
    
    lAccion.Caption = "Procesando ventas del mes...": lAccion.Refresh
    lVentaI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lVentaI.Refresh
    CargoTablaCMVenta CDate(tMesI.Text), CDate(tMesF.Text)
    lVentaF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lVentaF.Refresh
    
    lAccion.Caption = "Costeando Mercadería...": lAccion.Refresh
    lCosteoI.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoI.Refresh
    CargoTablaCMCosteo aIDCosteo
    lCosteoF.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss"): lCosteoF.Refresh
    lAccion.Caption = "Costeo Finalizado OK.": lAccion.Refresh
    
    MsgBox "El costeo para el mes de " & Trim(tMesI.Text) & " se ha finalizado con éxito.", vbExclamation, "Costeo Finalizado"
    CargoInformacionCosteo
    
    Screen.MousePointer = 0
    Exit Sub

errCostear:
    clsGeneral.OcurrioError "Ocurrió un error al costear la mercadería.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoTablaCMCompra(MesI As Date, MesF As Date)

    '1) Cargar las compras del mes (Credito y Contado)
    '2) Cargar las importes del mes (con fecha de arribo del costeo en el mes)
Dim aCosto As Currency
Dim QyCom As rdoQuery
    
    Cons = "Insert Into CMCompra (ComFecha, ComArticulo, ComTipo, ComCodigo, ComCantidad, ComCosto, ComQOriginal) Values (?,?,?,?,?,?,?)"
    Set QyCom = cBase.CreateQuery("", Cons)
    
    '1) Compras del Mes (contado y credito)------------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from Compra, CompraRenglon" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(MesI, sqlFormatoFH) & "' And '" & Format(MesF & " 23:59:59", sqlFormatoFH) & "'" _
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
           & " Where CCaFArribo Between '" & Format(MesI, sqlFormatoFH) & "' And '" & Format(MesF & " 23:59:59", sqlFormatoFH) & "'" _
           & " And CCaID = CArIDCosteo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        QyCom.rdoParameters(0) = RsAux!CCaFArribo
        QyCom.rdoParameters(1) = RsAux!CArIdArticulo
        QyCom.rdoParameters(2) = TipoCV.Importacion
        QyCom.rdoParameters(3) = RsAux!CCaID
        
        QyCom.rdoParameters(4) = RsAux!CArCantidad
        QyCom.rdoParameters(6) = RsAux!CArCantidad
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
Dim bSalir As Boolean, bBorroVenta As Boolean

    'Preparo las Queries para costear-----------------------------------------------------------------------------------------------------------------------------
    Cons = "Insert Into CMCosteo (CosID, CosArticulo, CosTipoVenta, CosIDVenta, CosCantidad, CosCosto, CosVenta) Values (?,?,?,?,?,?,?)"
    Set QyCos = cBase.CreateQuery("", Cons)
    QyCos.rdoParameters(0) = IDCosteo

    Dim QyCMenores As rdoQuery, QyCMayores As rdoQuery
    
    Cons = "Select * from CMCompra " _
            & " Where ComFecha <= ?" _
            & " And ComArticulo = ? " _
            & " And ComCantidad > 0 " _
            & " Order by ComFecha DESC"
    Set QyCMenores = cBase.CreateQuery("", Cons)
    
    Cons = "Select * from CMCompra " _
            & " Where ComFecha >= ?" _
            & " And ComArticulo = ? " _
            & " And ComCantidad > 0 "
    Set QyCMayores = cBase.CreateQuery("", Cons)
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------------------------------------
    Cons = "Select Count(*) from CMVenta"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) <> 0 Then bProgress.Max = RsAux(0)
    RsAux.Close
    bProgress.Value = 0
    '-------------------------------------------------------------------------------------------------------------------

    Cons = "Select * from CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " Order by VenFecha, VenArticulo, VenTipo, VenCodigo"
    Set RsVen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
       
    Do While Not RsVen.EOF
        bProgress.Value = bProgress.Value + 1
        aFVenta = RsVen!VenFecha
        aArticulo = RsVen!VenArticulo
        aQVenta = RsVen!VenCantidad
        aQVentaOriginal = aQVenta
        
        QyCos.rdoParameters(1) = aArticulo
        bBorroVenta = True
        Do While aQVenta <> 0
        
            'Si el artículo es del tipo Servicio lo costeo contra costo 0
            If RsVen!ArtTipo = paTipoArticuloServicio Then
                aQCosteo = aQVenta
                aQVenta = 0
                
                QyCos.rdoParameters(2) = RsVen!VenTipo
                QyCos.rdoParameters(3) = RsVen!VenCodigo
                QyCos.rdoParameters(4) = aQCosteo
                QyCos.rdoParameters(5) = 0
                QyCos.rdoParameters(6) = RsVen!VenPrecio
                QyCos.Execute
            
            Else
        
                'Voy a la maxima fecha de Compra <= a la fecha de venta ------------------------------------
                QyCMenores.rdoParameters(0) = Format(aFVenta, sqlFormatoF) & " 23:59:59"
                QyCMenores.rdoParameters(1) = aArticulo
                Set RsCom = QyCMenores.OpenResultset(rdOpenDynamic, rdConcurValues)
                
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
                        
                        QyCos.rdoParameters(2) = RsVen!VenTipo
                        QyCos.rdoParameters(3) = RsVen!VenCodigo
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
                                
                                QyCos.rdoParameters(2) = RsVen!VenTipo
                                QyCos.rdoParameters(3) = RsVen!VenCodigo
                                QyCos.rdoParameters(4) = aQCosteo
                                QyCos.rdoParameters(5) = RsCom!ComCosto
                                QyCos.rdoParameters(6) = RsVen!VenPrecio
                                QyCos.Execute
                            
                                RsCom.Edit
                                RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                                RsCom.Update
                                
                            Else
                                RsCom.MoveNext
                                If RsCom.EOF Then bSalir = True: bBorroVenta = False: aQVenta = 0
                            End If
                        Loop
                        
                    End If
                    RsCom.Close
    
                Else                                        'NO Hay una FC <= FV
                    RsCom.Close
                    'Voy a la minima fecha de Compra >= a la fecha de venta------------------------------------
                    QyCMayores.rdoParameters(0) = Format(aFVenta, sqlFormatoF) & " 23:59:59"
                    QyCMayores.rdoParameters(1) = aArticulo
                    Set RsCom = QyCMayores.OpenResultset(rdOpenDynamic, rdConcurValues)
                    
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
                            
                            QyCos.rdoParameters(2) = RsVen!VenTipo
                            QyCos.rdoParameters(3) = RsVen!VenCodigo
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
                                    
                                    QyCos.rdoParameters(2) = RsVen!VenTipo
                                    QyCos.rdoParameters(3) = RsVen!VenCodigo
                                    QyCos.rdoParameters(4) = aQCosteo
                                    QyCos.rdoParameters(5) = RsCom!ComCosto
                                    QyCos.rdoParameters(6) = RsVen!VenPrecio
                                    QyCos.Execute
                                
                                    RsCom.Edit
                                    RsCom!ComCantidad = RsCom!ComCantidad - aQCosteo
                                    RsCom.Update
                                    
                                Else
                                    RsCom.MoveNext
                                    If RsCom.EOF Then bSalir = True: bBorroVenta = False: aQVenta = 0
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
            End If
        Loop
        
        'Si la venta quedó en cero elimino el registro de la venta
        If aQVenta = 0 And bBorroVenta Then
            Cons = " Delete CMVenta " _
                    & " Where VenFecha = '" & Format(RsVen!VenFecha, sqlFormatoFH) & "'" _
                    & " And VenArticulo = " & RsVen!VenArticulo _
                    & " And VenTipo = " & RsVen!VenTipo & " And VenCodigo = " & RsVen!VenCodigo
            cBase.Execute Cons
        End If
        RsVen.MoveNext
    Loop
    
    RsVen.Close
    QyCMenores.Close: QyCMayores.Close
    
    'Hay que borrar las compras en que las cantidades son iguales a 0
    Cons = "Delete CMCompra Where ComCantidad = 0"
    cBase.Execute Cons
    '---------------------------------------------------------------------------
    bProgress.Value = 0
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    LimpioDatos
    AccionCostear
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    LimpioDatos
    
    CargoInformacionCosteo
        
End Sub

Private Sub LimpioDatos()
    lAccion.Caption = ""
    lCosteoI.Caption = "": lCosteoF.Caption = ""
    lCompraI.Caption = "": lCompraF.Caption = ""
    lVentaI.Caption = "": lVentaF.Caption = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set msgError = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub CargoTablaCMVenta(MesI As Date, MesF As Date)
'Parámetro: Recibe el primer día del mes a costear.
On Error GoTo ErrCTCMV
Dim QyVta As rdoQuery
Dim strDocumentos As String
Dim Monedas As String
Dim aCosto As Currency
Dim aMoneda As Long, aTC As Currency

    Screen.MousePointer = 11
    
    Monedas = RetornoTCMonedas(PrimerDia(MesI))
    aMoneda = 0
    'Creo query para insertar los datos
    Cons = "Insert Into CMVenta (VenFecha, VenArticulo, VenTipo, VenCodigo, VenCantidad,VenPrecio) Values (?,?,?,?,?,?)"
    Set QyVta = cBase.CreateQuery("", Cons)
    
    'Primer Paso Copio las Ventas---------------------------------------
    'Traigo los documentos Ctdo y Cred, Nota Esp, Nota de Cred. y  Nota de Dev. que no estén anulados
    strDocumentos = TipoDocumento.Contado & ", " & TipoDocumento.Credito _
        & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial
    
    Cons = "Select DocFecha, DocMoneda, DocTipo, Renglon.* From Documento, Renglon" _
        & " Where DocTipo IN (" & strDocumentos & ")" _
        & " And DocFecha BetWeen '" & Format(MesI & " 00:00:00", sqlFormatoFH) & "'" _
        & " And '" & Format(MesF & " 23:59:59", sqlFormatoFH) & "'" _
        & " And DocAnulado = 0 And DocCodigo = RenDocumento"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)

    Do While Not RsAux.EOF
        QyVta.rdoParameters(0) = Format(RsAux!DocFecha, sqlFormatoF)
        QyVta.rdoParameters(1) = RsAux!RenArticulo
        QyVta.rdoParameters(2) = TipoCV.Comercio
        QyVta.rdoParameters(3) = RsAux!RenDocumento
        
        Select Case RsAux!DocTipo
            Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                            QyVta.rdoParameters(4) = RsAux!RenCantidad * -1
                            
            Case Else: QyVta.rdoParameters(4) = RsAux!RenCantidad
        End Select
        
        'Es el precio neto-----------------------------------------------------------------------------
        If aMoneda <> RsAux!DocMoneda Then
            aMoneda = RsAux!DocMoneda
            aTC = ValorTC(RsAux!DocMoneda, Monedas)
        End If
        
        QyVta.rdoParameters(5) = (RsAux!RenPrecio - RsAux!RenIva) * aTC
        '-------------------------------------------------------------------------------------------------
        QyVta.Execute
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Segundo paso Cargo Notas de Compras.
    Cons = "Select * from Compra, CompraRenglon" _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha Between '" & Format(MesI, sqlFormatoFH) & "' And '" & Format(MesF & " 23:59:59", sqlFormatoFH) & "'" _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    Do While Not RsAux.EOF
        QyVta.rdoParameters(0) = Format(RsAux!ComFecha, sqlFormatoF)
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

Private Sub Label1_Click()
    Foco tMesI
End Sub

Private Sub CargoInformacionCosteo()

    On Error GoTo errInfo
    Cons = "Select * from CMCabezal Where CabId In (Select Max(CabID) from CMCabezal)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    lCMes.Caption = ""
    lCFecha.Caption = "": lCUsuario.Caption = ""
    
    If Not RsAux.EOF Then
        lCMes.Caption = Format(RsAux!CabMesCosteo, "Mmmm yyyy")
        lCFecha.Caption = Format(RsAux!CabFecha, "dd/mm/yyyy hh:mm")
        lCUsuario.Caption = miConexion.UsuarioLogueado(Nombre:=True)
    End If
    
    'If IsDate(lCMes.Caption) Then tMes.Text = Format(DateAdd("m", 1, CDate(lCMes.Caption)), "Mmmm yyyy") Else tMes.Text = ""
    
    If IsDate(lCMes.Caption) Then
        tMesI.Text = Format(DateAdd("m", 1, CDate(lCMes.Caption)), "dd/mm/yyyy")
        tMesF.Text = Format(DateAdd("m", 1, CDate(lCMes.Caption)) - 1, "dd/mm/yyyy")
        
    Else
        tMesI.Text = "": tMesF.Text = ""
    End If
    RsAux.Close
    Exit Sub
    
errInfo:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la información del último costeo.", Err.Description
End Sub

Private Sub tMesI_GotFocus()
    With tMesI: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Function ValidoCampos() As Boolean

    On Error GoTo errValido
    ValidoCampos = False
    If Not IsDate(tMesI.Text) Then
        MsgBox "El mes ingresado para realizar el costeo no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tMesI: Exit Function
    End If
    If Not IsDate(tMesF.Text) Then
        MsgBox "El mes ingresado para realizar el costeo no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tMesF: Exit Function
    End If
    If CDate(tMesF.Text) < CDate(tMesI.Text) Then
        MsgBox "El rango de fechas para realizar el costeo no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tMesI: Exit Function
    End If
    
    If Trim(lCMes.Caption) <> "" Then
        If CDate(lCMes.Caption) >= CDate(tMesI.Text) Then
            MsgBox "El mes ingresado para realizar el costeo debe ser mayor al último mes costeado.", vbExclamation, "ATENCIÓN"
            Foco tMesI: Exit Function
        End If
        
        If Abs(DateDiff("m", CDate(lCMes.Caption), CDate(tMesI.Text))) <> 1 Then
            If MsgBox("El mes ingresado para realizar el costeo no es el siguiente al último mes costeado." & Chr(vbKeyReturn) _
                    & "Está seguro de costear el mes ingresado.", vbYesNo + vbDefaultButton2 + vbExclamation, "ATENCIÓN") = vbNo Then
                Foco tMesI: Exit Function
            End If
        End If
    End If
    
    ValidoCampos = True
    Exit Function

errValido:
    clsGeneral.OcurrioError "Ocurrió un error al validar datos.", Err.Description
End Function

Private Sub tMesF_LostFocus()
    If IsDate(tMesF.Text) Then tMesF.Text = Format(tMesF.Text, "dd/mm/yyyy")
End Sub

Private Sub tMesI_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tMesF
End Sub

Private Sub tMesI_LostFocus()
    If IsDate(tMesI.Text) Then tMesI.Text = Format(tMesI.Text, "dd/mm/yyyy")
End Sub
