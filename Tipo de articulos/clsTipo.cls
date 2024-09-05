VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Codigo As Long
Public Especie As Long
Public HijoDe As Long
Public Nombre As String
Public Abreviacion As String
Public LocalReparacion As Integer
Public BusquedaWeb As String
Public ArrayCaracteristicas As String
Public RecargoMS As Long
Public RecargoRM As Long

Public Function DeleteTipo(ByVal iTipo As Long) As Boolean
Dim sQy As String
        
    On Error GoTo errValidar
    Screen.MousePointer = 11
    'Verifico que no existan artículos con ese tipo.
    Cons = "Select * from Articulo Where ArtTipo = " & iTipo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "Existen dependencias al tipo seleccionado, no podrá eliminarlo.", vbCritical, "ATENCION"
        RsAux.Close
        Exit Function
    End If
    RsAux.Close
    
    'Si el tipo tiene hijos doy aviso
    Cons = "Select TipCodigo From Tipo Where TipHijoDe = " & iTipo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Screen.MousePointer = 0
        RsAux.Close
        If MsgBox("El tipo tiene hijos asignados, estos quedaran sin padre." & vbCrLf & vbCrLf & "¿Confirma continuar?", vbQuestion + vbYesNo, "Validación") = vbNo Then
            Exit Function
        End If
    Else
        RsAux.Close
    End If

    
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB

    Cons = "UPDATE Tipo SET TipHijoDe = Null WHERE TipHijoDe = " & iTipo
    cBase.Execute Cons

    Cons = "DELETE Tipo WHERE TipCodigo = " & iTipo
    cBase.Execute Cons
    
    cBase.CommitTrans
    DeleteTipo = True
    Exit Function
    Screen.MousePointer = 0

errValidar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al validar la eliminación.", Err.Description, "Eliminar Tipo"
    Exit Function
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description, "Eliminar Tipo"
    Exit Function
errSave:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar grabar los datos.", Err.Description, "Eliminar Tipo"
    Exit Function
errRB:
    Resume errSave
    Exit Function
End Function

Public Function LoadTipo(ByVal iCodigo As Long) As Boolean
On Error GoTo errLoad
    Screen.MousePointer = 11
    Cons = "SELECT * FROM Tipo WHERE TipCodigo = " & iCodigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Codigo = iCodigo
        Nombre = Trim(RsAux!TipNombre)
        If Not IsNull(RsAux!TipAbreviacion) Then Abreviacion = RsAux!TipAbreviacion
        If Not IsNull(RsAux!TipLocalRep) Then Me.LocalReparacion = RsAux!TipLocalRep
        If Not IsNull(RsAux!TipBusqWeb) Then Me.BusquedaWeb = Trim(RsAux!TipBusqWeb)
        If Not IsNull(RsAux!TipArrayCaract) Then Me.ArrayCaracteristicas = Trim(RsAux!TipArrayCaract)
        If Not IsNull(RsAux!TipRecargoMS) Then Me.RecargoMS = RsAux!TipRecargoMS
        If Not IsNull(RsAux!TipRecargoRM) Then Me.RecargoRM = RsAux!TipRecargoRM
        If Not IsNull(RsAux("TipHijoDe")) Then Me.HijoDe = RsAux("TipHijoDe")
        If Not IsNull(RsAux("TipEspecie")) Then Me.Especie = RsAux("TipEspecie")
    End If
    RsAux.Close
    LoadTipo = (Codigo > 0)
    Screen.MousePointer = 0
    Exit Function
errLoad:
    clsGeneral.OcurrioError "Error al cargar la información del tipo.", Err.Description, "Buscar tipo"
End Function

Public Function SaveTipo() As Boolean
On Error GoTo errST
Dim rsT As rdoResultset

    'Verifico duplicación de nombre
    Cons = "SELECT * FROM Tipo WHERE TipCodigo <> " & Codigo & " AND TipNombre = '" & Nombre & "'"
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsT.EOF Then
        MsgBox "Ya existe un tipo ingresado con ese nombre, no puede duplicar el nombre.", vbExclamation, "Validación"
        rsT.Close
        Exit Function
    End If
    rsT.Close

    Cons = "SELECT * FROM Tipo WHERE TipCodigo = " & Codigo
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Codigo <= 0 Then                  'Nuevo----------
        rsT.AddNew
    Else
        rsT.Edit
    End If
    If Me.HijoDe > 0 Then rsT("TipHijoDe") = Me.HijoDe Else rsT("TipHijoDe") = Null
    rsT!TipNombre = Nombre
    If Abreviacion <> "" Then rsT!TipAbreviacion = Abreviacion Else rsT!TipAbreviacion = Null
    If Trim(Me.BusquedaWeb) <> "" Then rsT!TipBusqWeb = Me.BusquedaWeb Else rsT!TipBusqWeb = Null
    If Me.ArrayCaracteristicas <> "" Then
        rsT!TipArrayCaract = Me.ArrayCaracteristicas
    Else
        rsT!TipArrayCaract = Null
    End If
    rsT!TipEspecie = Me.Especie
    
    If Me.LocalReparacion > 0 Then rsT!TipLocalRep = Me.LocalReparacion Else rsT("TipLocalRep") = Null
    If Me.RecargoMS > 0 Then rsT!TipRecargoMS = RecargoMS Else rsT!TipRecargoMS = Null
    If RecargoRM > 0 Then rsT!TipRecargoRM = RecargoRM Else rsT!TipRecargoRM = Null
    rsT.Update
    rsT.Close
    
    If Codigo <= 0 Then
        Cons = "SELECT Max(TipCodigo) From Tipo Where TipNombre = '" & Nombre & "'"
        Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsT.EOF Then Codigo = rsT(0)
        rsT.Close
    End If

    SaveTipo = True
    Exit Function
    
errST:
    clsGeneral.OcurrioError "Error al grabar el tipo de artículo.", Err.Description, "Grabar"

End Function
