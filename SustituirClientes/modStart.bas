Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String

Public paLocalidad As Long

Public Sub Main()
    On Error GoTo errMain
    
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        txtConexion = miConexion.TextoConexion("comercio")
        If Not InicioConexionBD(txtConexion) Then End
        
'        CargoParametrosLocales
                       
        frmSustituir.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Sin Acceso"
            Screen.MousePointer = 0
        End If
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()

    On Error GoTo errParametro
 
    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            'Case "monedaempleo": paMonedaEmpleo = rsAux!ParValor
            
            
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
errParametro:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


Public Function TelefonoATexto(Cliente As Long, Optional TieneLlamoDe As Boolean = False) As String

Dim rsTel As rdoResultset
Dim aTelefonos As String

    On Error GoTo errTelefono
    TieneLlamoDe = False
    
    Cons = "Select * from Telefono, TipoTelefono" _
        & " Where TelCliente = " & Cliente _
        & " And TelTipo = TTeCodigo"
    Set rsTel = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsTel.EOF Then
        aTelefonos = ""
        Do While Not rsTel.EOF
            aTelefonos = aTelefonos & Trim(rsTel!TTeNombre) & ": " & Trim(rsTel!TelNumero)
            If Not IsNull(rsTel!telInterno) Then aTelefonos = aTelefonos & "(" & Trim(rsTel!telInterno) & ")"
            aTelefonos = aTelefonos & ", "
            
            'If RsTel!TelTipo = paTipoTelefonoLlamoDe Then TieneLlamoDe = True
            rsTel.MoveNext
        Loop
        aTelefonos = Mid(aTelefonos, 1, Len(aTelefonos) - 2)
    Else
        aTelefonos = "S/D"
    End If
    rsTel.Close
    
    TelefonoATexto = aTelefonos

errTelefono:
End Function

Public Function sql_Update(Tbla As String, Campo As String, Malo As Long, Bueno As Long, Optional Error As Boolean = False, Optional sqlAnd As String = "") As String
On Error GoTo ErrUpdateDAO
    Error = False
    Dim Cant As Integer ', C As Integer
'
'    Cons = "SELECT " & Campo & " FROM " & Tbla & " WHERE " & Campo & "=" & Malo
'    If Trim(sqlAnd) <> "" Then Cons = Cons & " AND " & sqlAnd
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    Do Until RsAux.EOF
'        Cant = Cant + 1
'        RsAux.Edit
'        RsAux(Campo) = Bueno
'        RsAux.Update: RsAux.MoveNext
'    Loop
'    sql_Update = "UpD." & Cant & " " & Tbla & ". " & Campo & "=" & Malo
'FinUpdateDAO: RsAux.Close


    Dim sQy As String
    sQy = "EXEC prg_UpdateCampoEnTabla '" & Tbla & "', '" & Campo & "', '" & Bueno & "', '" & Malo & "', " & _
        IIf(sqlAnd <> "", "'" & sqlAnd & "'", "'1=1'")
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(1)) Then Cant = RsAux(1)
    RsAux.Close
    sql_Update = "UpD." & Cant & " " & Tbla & ". " & Campo & "=" & Malo
    
'EXECUTE @RC = [CGSA].[dbo].[prg_UpdateCampoEnTabla]
'   @Tabla
'  ,@Campo
'  ,@ValorCampoNew
'  ,@ValorCampoOld
'  ,@CondAnd


Fin2UpdateDAO: Exit Function

ErrUpdateDAO:
'    C = C + 1
    Error = True
'    sql_Update = "Error " & Err & " UpDateando " & Tbla & " dónde " & Campo & "=" & Malo
'    If C = 1 Then Resume FinUpdateDAO Else Resume Fin2UpdateDAO
Debug.Print "ACA"
End Function

Public Function sql_Update_OLD(Tbla As String, Campo As String, Malo As Long, Bueno As Long, Optional Error As Boolean = False, Optional sqlAnd As String = "") As String
On Error GoTo ErrUpdateDAO
    Error = False
    Dim Cant As Integer, C As Integer

    Cons = "SELECT " & Campo & " FROM " & Tbla & " WHERE " & Campo & "=" & Malo
    If Trim(sqlAnd) <> "" Then Cons = Cons & " AND " & sqlAnd

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do Until RsAux.EOF
        Cant = Cant + 1
        RsAux.Edit
        RsAux(Campo) = Bueno
        RsAux.Update: RsAux.MoveNext
    Loop
    sql_Update_OLD = "UpD." & Cant & " " & Tbla & ". " & Campo & "=" & Malo
FinUpdateDAO: RsAux.Close


Fin2UpdateDAO: Exit Function
ErrUpdateDAO:
    C = C + 1
    Error = True
    sql_Update_OLD = "Error " & Err & " UpDateando " & Tbla & " dónde " & Campo & "=" & Malo
    If C = 1 Then Resume FinUpdateDAO Else Resume Fin2UpdateDAO
End Function


Public Function sql_Delete(Tbla As String, Campo As String, Malo As Long, Optional sqlAnd As String = "") As String
On Error GoTo Errsql_Delete
Dim C As Integer, Cant As Integer

    Cons = "SELECT * FROM " & Tbla & " WHERE " & Campo & "=" & Malo
    If Trim(sqlAnd) <> "" Then Cons = Cons & " AND " & sqlAnd
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do Until RsAux.EOF
        RsAux.Delete: Cant = Cant + 1: RsAux.MoveNext
    Loop
    
    sql_Delete = "Borr." & Cant & " " & Tbla & ". " & Campo & "=" & Malo

Finsql_Delete: RsAux.Close
Fin2sql_Delete: Exit Function

Errsql_Delete:
    C = C + 1
    sql_Delete = "Error " & Err & " Borrando Reg. en " & Tbla & " dónde " & Campo & "=" & Malo
    If C = 1 Then Resume Finsql_Delete Else Resume Fin2sql_Delete
End Function

Function ProxTipoDisponible(nCliente As Long, miTipoT As Long) As Integer
    On Error GoTo errFunc
    ProxTipoDisponible = 0
    
    Dim rs1 As rdoResultset
    Cons = "Select Top 1 Codigo from CodigoTexto " & _
               " Where Tipo = 8 " & _
               " And Codigo Not In (Select TelTipo FROM Telefono WHERE TelCliente=" & nCliente & ")" & _
               " And Clase > (Select isnull(Clase, 0) From CodigoTexto Where Tipo = 8 And Codigo = " & miTipoT & ")" & _
               " Order by Clase"
        
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then ProxTipoDisponible = rs1!Codigo
    rs1.Close

errFunc:
End Function


