Attribute VB_Name = "modPersistencia"
Option Explicit
'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno

Public Cons As String
Public Function Usuario_Buscar(ByRef oUID As clsUsuario) As Boolean
On Error GoTo errUB
Dim rsAux As rdoResultset
Dim sQuery As String

Screen.MousePointer = 11
    If oUID.Codigo > 0 Then
        sQuery = " Where UsuCodigo = " & oUID.Codigo
    ElseIf oUID.Digito > 0 Then
        sQuery = " Where UsuDigito = " & oUID.Digito
    ElseIf oUID.Identificacion <> "" Then
        sQuery = " Where UsuIdentificacion = '" & oUID.Identificacion & "'"
    Else
        Screen.MousePointer = 0
        Exit Function
    End If
    
    sQuery = "Select UsuCodigo, UsuDigito, UsuIDentificacion, UsuHabilitado " & _
                " From Usuario " & sQuery
    Set rsAux = cBase.OpenResultset(sQuery, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        With oUID
            .Codigo = rsAux("UsuCodigo")
            .Digito = rsAux("UsuDigito")
            .Identificacion = Trim(rsAux("UsuIdentificacion"))
            .Habilitado = rsAux("UsuHabilitado")
        End With
        Usuario_Buscar = True
    End If
    rsAux.Close

Screen.MousePointer = 0
Exit Function
errUB:
    Screen.MousePointer = 0
    objG.OcurrioError "Error al buscar el usuario.", Err.Description, "Buscar Usuario"
End Function

Public Function Documento_BuscoDocPorTexto(adTexto As String, retIDDoc As Long, retIDTipoD) As Boolean
On Error GoTo errDoc

    Documento_BuscoDocPorTexto = False
    
    Dim mDSerie As String, mDNumero As Long
    Dim adQ As Integer, adCodigo As Long, adTipoD As Integer
    Dim sQy As String, rsAux As rdoResultset
        
    If InStr(adTexto, "-") <> 0 Then
        mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
        mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
    Else
        mDSerie = Mid(adTexto, 1, 1)
        mDNumero = Val(Mid(adTexto, 2))
    End If
    
    adTexto = UCase(mDSerie) & "-" & mDNumero
        
    Screen.MousePointer = 11
    adQ = 0: adTexto = ""
    
    'Cargo combo con tipos de docuemento--------------------------------------
    sQy = "Select DocCodigo, DocTipo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
               " From Documento " & _
               " Where DocSerie = '" & mDSerie & "'" & _
               " And DocNumero = " & mDNumero & _
               " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.NotaCredito & ", " & _
                                               TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
        
    Set rsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        adCodigo = rsAux!DocCodigo
        adTipoD = rsAux!DocTipo
        adQ = 1
        rsAux.MoveNext: If Not rsAux.EOF Then adQ = 2
    End If
    rsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                If miLDocs.ActivarAyuda(cBase, sQy, 4100, 2) <> 0 Then
                    adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                    adTipoD = miLDocs.RetornoDatoSeleccionado(1)
                End If
                Set miLDocs = Nothing
        End Select
        
        If adCodigo > 0 Then
            'lDoc.Tag = adCodigo: lDoc.Caption = adTexto
            Documento_BuscoDocPorTexto = True
            retIDDoc = adCodigo
            retIDTipoD = adTipoD
        Else
            'lDoc.Caption = " No Existe !!"
            Documento_BuscoDocPorTexto = False
        End If
        
        Screen.MousePointer = 0
    'End If
    
    Exit Function
errDoc:
    objG.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Function

