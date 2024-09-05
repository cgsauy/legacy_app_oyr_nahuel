Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public paClienteEmpresa As Long

Public Sub Main()
Dim aTexto As String
Dim arrCom() As String
Dim iQ As Byte, lCli As Long

'Ingreso del command
    ' id                           = Solo id de Cliente
    'N + Id de Cliente      = Nuevo con id de un cliente
    'P + Id de Producto
    'A + Id de Artículo

'Si combina van separados con ;

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("comercio")
        CargoParametros
        
        aTexto = Command()
        
        lCli = 0
        If aTexto <> "" Then
            arrCom = Split(aTexto, ";")
            For iQ = 0 To UBound(arrCom)
                Select Case LCase(Mid(arrCom(iQ), 1, 1))
                    Case "n"
                        frmProducto.prmNuevo = True
                        lCli = Mid(arrCom(iQ), 2)
                        
                    Case "a"
                        frmProducto.m_IDArticulo = Mid(arrCom(iQ), 2)
                    
                    Case "p"
                        frmProducto.m_IDProducto = Mid(arrCom(iQ), 2)
                    
                    Case 0 To 9
                        lCli = arrCom(iQ)
                End Select
            Next
        End If
        frmProducto.prmCliente = lCli
        frmProducto.Show
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & _
                Err.Number & " - " & Err.Description
    End
End Sub

Private Sub CargoParametros()

    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub
