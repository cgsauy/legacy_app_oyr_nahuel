Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
        
        CargoParametrosImportaciones
        
        frmMaCEmpresa.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmMaCEmpresa.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmMaCEmpresa.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        frmMaCEmpresa.Show
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title
    End
End Sub

'--------------------------------------------------------------------------------------------------------
'   PROCEDMIENTO ProcesoCalle: Para la calle XXX realiza el proceso de búsqueda.
'       Si hay mas de una llama al Help y si se selecciona ya la pone en el control.
'       .Text = nombre de la calle
'       .Tag  = id de calle
'
'   PARÁMETROS:
'       Nombre: Nombre de la calle a buscar.
'       Localidad: Codigo de localidad al que pertenece la calle
'       aControl: Control de texto de la calle
'--------------------------------------------------------------------------------------------------------
Public Function ProcesoCalle(Nombre As String, Localidad As Integer, aControl As Control) As Boolean

Dim aNomCalle As String  'Nombre auxiliar de la calle
Dim aCantidad As Long: aCantidad = 0
Dim RsCal As rdoResultset

    On Error GoTo errProceso
    Screen.MousePointer = 11
    ProcesoCalle = False
    aNomCalle = Trim(aControl.Text)
    
    If aNomCalle <> "" Then
        
        Cons = "Select Count(*) from Calle Where CalLocalidad = " & Localidad & " And CalNombre like '" & aNomCalle & "%'"
        Set RsCal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsCal.EOF Then If Not IsNull(RsCal(0)) Then aCantidad = RsCal(0)
        RsCal.Close
        
        Select Case aCantidad
            Case 0          'No hay calles
                Screen.MousePointer = 0
                MsgBox "No existen calles que coincidan con el texto ingresado.", vbExclamation, "ATENCIÓN"
            
            Case 1     'Cargo los datos de la calle al control
                Cons = "Select * from Calle  Where CalLocalidad = " & Localidad & " And CalNombre like '" & aNomCalle & "%'"
                Set RsCal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aControl.Text = Trim(RsCal!CalNombre)
                aControl.Tag = RsCal!CalCodigo
                RsCal.Close
                ProcesoCalle = True
                
            Case Is > 1     'Hay mas de una calle --- Lista de ayuda
                Dim aLista As New clsListadeAyuda
                Cons = "Select CalCodigo, CalNombre from Calle Where CalLocalidad = " & Localidad & " And CalNombre like '" & aNomCalle & "%'"
                aLista.ActivoListaAyuda Cons, False, cBase.Connect
                
                If aLista.ValorSeleccionado > 0 Then
                    aControl.Text = Trim(aLista.ItemSeleccionado)
                    aControl.Tag = aLista.ValorSeleccionado
                    ProcesoCalle = True
                End If
        End Select
    End If
    
    Screen.MousePointer = 0
    Exit Function
    
errProceso:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al buscar la calle ingresada.", Err.Description
End Function


Public Function DireccionATexto(Codigo As Long, Optional Confirmada As Boolean = True)

Dim aTexto As String
Dim RsAux2 As rdoResultset
    
    On Error Resume Next
    Cons = "Select Direccion.*, LocNombre, DepNombre, CalNombre From Direccion, Calle, Localidad, Departamento" _
            & " Where DirCodigo = " & Codigo _
            & " And DirCalle = CalCodigo And CalLocalidad = LocCodigo" _
            & " And LocDepartamento = DepCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Trim(RsAux!DepNombre) <> Trim(RsAux!LocNombre) Then
        aTexto = Trim(RsAux!DepNombre) & ", " & Trim(RsAux!LocNombre)
    Else
        aTexto = Trim(RsAux!DepNombre)
    End If
    
    'Cargo el Complejo Habitacional-----------------------------------------
    If Not IsNull(RsAux!DirComplejo) Then
        Cons = "Select ComNombre from Complejo Where ComCodigo = " & RsAux!DirComplejo
        Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        aTexto = aTexto & " (" & Trim(RsAux2!ComNombre) & ")"
        RsAux2.Close
    End If
    
    aTexto = aTexto & Chr(vbKeyReturn) & Chr(10)
    
    aTexto = aTexto & Trim(RsAux!CalNombre) & " "
    
    If Trim(RsAux!DirPuerta) = 0 Then aTexto = aTexto & "S/N" Else: aTexto = aTexto & Trim(RsAux!DirPuerta)
    
    If Not IsNull(RsAux!DirLetra) Then aTexto = aTexto & Trim(RsAux!DirLetra)
    If Not IsNull(RsAux!DirApartamento) Then aTexto = aTexto & "/" & Trim(RsAux!DirApartamento)
    If RsAux!DirBis Then aTexto = aTexto & " Bis"
    
    'Campo 1 de la Direccion------------------------------------------------------------------------------------------
    If Not IsNull(RsAux!DirCampo1) Then
        Cons = "Select CDiAbreviacion from CamposDireccion Where CDiCodigo = " & RsAux!DirCampo1
        Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " " & Trim(RsAux2!CDiAbreviacion)
            If Not IsNull(RsAux!DirSenda) Then aTexto = aTexto & " " & Trim(RsAux!DirSenda)
        End If
        RsAux2.Close
    End If
    '-----------------------------------------------------------------------------------------------------------------------
    'Campo 2 de la Direccion------------------------------------------------------------------------------------------
    If Not IsNull(RsAux!DirCampo2) Then
        Cons = "Select CDiAbreviacion from CamposDireccion Where CDiCodigo = " & RsAux!DirCampo2
        Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " " & Trim(RsAux2!CDiAbreviacion)
            If Not IsNull(RsAux!DirBloque) Then aTexto = aTexto & " " & Trim(RsAux!DirBloque)
        End If
        RsAux2.Close
    End If
    '-----------------------------------------------------------------------------------------------------------------------
    
    If Not IsNull(RsAux!DirEntre1) Or Not IsNull(RsAux!DirEntre2) Then
        aTexto = aTexto & Chr(vbKeyReturn) & Chr(10)
        If Not IsNull(RsAux!DirEntre1) And Not IsNull(RsAux!DirEntre2) Then
            
            Cons = "Select * from Calle where CalCodigo = " & RsAux!DirEntre1
            Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = aTexto & "Entre " & Trim(RsAux2!CalNombre)
            RsAux2.Close
            
            Cons = "Select * from Calle where CalCodigo = " & RsAux!DirEntre2
            Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = aTexto & " y " & Trim(RsAux2!CalNombre)
            RsAux2.Close
            
        Else
            If Not IsNull(RsAux!DirEntre1) Then
                Cons = "Select * from Calle where CalCodigo = " & RsAux!DirEntre1
                Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aTexto = aTexto & "Esquina " & Trim(RsAux2!CalNombre)
                RsAux2.Close
            Else
                Cons = "Select * from Calle where CalCodigo = " & RsAux!DirEntre2
                Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aTexto = aTexto & "Esquina " & Trim(RsAux2!CalNombre)
                RsAux2.Close
            End If
        End If
    End If
    
    If Not IsNull(RsAux!DirAmpliacion) Then aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) & Trim(RsAux!DirAmpliacion)
    If Confirmada Then
        If RsAux!DirConfirmada Then
            aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) & Chr(vbKeyReturn) & Chr(10) & "(Confirmada"
        Else
            aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) & Chr(vbKeyReturn) & Chr(10) & "(No Confirmada"
        End If
        
        If Not IsNull(RsAux!DirVive) Then aTexto = aTexto & ", Vive desde " & Format(RsAux!DirVive, "Mmm-yyyy")
        aTexto = aTexto & ")"
    End If
        
    RsAux.Close
    DireccionATexto = aTexto
    
End Function

