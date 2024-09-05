Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA 'clsLibGeneral
Public miConexion As New clsConexion
Public txtConexion As String

Public paPlantillasIDMail As Long
Public prmPathApp As String

Public Sub Main()

On Error GoTo ErrMain
Dim aTexto As String, aPos As Integer

    If App.PrevInstance Then ActivatePrevInstance 'End
    
    Screen.MousePointer = 11
    txtConexion = miConexion.TextoConexion("comercio")
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    InicioConexionBD txtConexion
    CargoParametrosLocales
    
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        'Solo = IdCliente       /M = idEMail
        
        If IsNumeric(aTexto) Then
            frmEMails.prmIDCliente = Val(aTexto)
        Else
            aPos = InStr(aTexto, "/M")
            If aPos <> 0 Then frmEMails.prmIDEMail = Mid(aTexto, aPos + 2)
        End If
        
    End If
    
    frmEMails.Show
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
    End
End Sub


Public Sub CargoParametrosLocales()

    On Error GoTo errParametro
 
    cons = "Select * from Parametro Where ParNombre like 'plantilla%' or ParNombre like 'Path%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            
            Case "plantillasidmail": paPlantillasIDMail = rsAux!ParValor
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close

    Exit Sub
errParametro:
    clsGeneral.OcurrioError "Error al cargar los parámetros. (" & App.Path & ")", Err.Description
End Sub
