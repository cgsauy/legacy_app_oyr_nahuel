Attribute VB_Name = "modMainEntrega"
Option Explicit

Public prmTCAlEntregar As String
Public prmNombreLocal As String

Public Function fnc_ControlComentariosAlEntregar(mIdCliente As Long)
On Error GoTo errCOntrol
Dim rsCom As rdoResultset
Dim bHay As Boolean
    
    If Trim(prmTCAlEntregar) = "" Then Exit Function
    
    Cons = "Select Top 1 * From Comentario" & _
                " Where ComCliente = " & mIdCliente & _
                " And ComTipo IN (" & prmTCAlEntregar & ")"
                
    Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsCom.EOF Then bHay = True
    rsCom.Close
    
    If Not bHay Then Screen.MousePointer = 0: Exit Function
    
    Dim aObj As New clsCliente
    aObj.Comentarios idCliente:=mIdCliente
    Set aObj = Nothing
    
    DoEvents
    Screen.MousePointer = 0
    Exit Function
    
errCOntrol:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0

End Function
