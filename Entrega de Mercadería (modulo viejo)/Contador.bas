Attribute VB_Name = "Contador"
Option Explicit

'DOCUMENTOS

Global Const DocCredito = "Crédito"
Global Const DocContado = "Contado"
Global Const DocRemito = "Remito"
Global Const DocNDevolucion = "Nota Devolución"
Global Const DocNCredito = "Nota Crédito"
Global Const DocRecibo = "Recibo de Pago"
Global Const DocCreditoDomicilio = "Crédito Domicilio"
Global Const DocContadoDomicilio = "Contado Domicilio"
Global Const DocServicioDomicilio = "Servicio a Domicilio"
Global Const DocNEspecial = "Nota Especial"


Public paDContado As String
Public paDCredito As String
Public paDNDevolucion As String
Public paDNCredito As String
Public paDRecibo As String
Public paDNEspecial As String
Public paDNDebito As String

'--------------------------------------------------------------------------------
'   Devuelve un número para el tipo de documento pasado como parametro.
'   Los numeros son leidos desde la tabla CONTADOR, que maneja los correlativos
'   para todos los documentos.
'
'   EN LA TABLA QUEDA EL ULTIMO NUMERO USADO !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'   RETORNA: UN STRING compuesto por
'                                   - 1 char   (Serie del Documento)
'                                   - 1 a 9 char (Numero del Documento)
'
'   ACTUALIZA EL CONTADOR A EL RETORNADO (ULTIMO CONTADOR USADO).
'--------------------------------------------------------------------------------
Public Function NumeroDocumento(Documento As String, Optional ret_Serie As String, Optional ret_Numero As Long)

    Dim Auxiliar As String    'Auxiliar para retornar el NRO DOC (Serie + Nro)
    Dim RsDoc As rdoResultset
    
    Cons = "Select * from Contador Where ConDocumento = '" & Trim(Documento) & "'"
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurLock)
    
    If RsDoc.EOF Then
        RsDoc.AddNew
        RsDoc!ConDocumento = Trim(Documento)
        RsDoc!ConValor = 1
        RsDoc!ConSerie = "A"
        Auxiliar = RsDoc!ConSerie & RsDoc!ConValor
        ret_Serie = Trim(RsDoc!ConSerie)
        ret_Numero = RsDoc!ConValor

        RsDoc.Update
    
    Else
        If RsDoc!ConValor = 999999 Then
            RsDoc.Edit
            RsDoc!ConValor = 1
            RsDoc!ConSerie = Chr(Asc(RsDoc!ConSerie) + 1)
            
            Auxiliar = RsDoc!ConSerie & RsDoc!ConValor
            ret_Serie = Trim(RsDoc!ConSerie)
            ret_Numero = RsDoc!ConValor
            RsDoc.Update
        Else
            Auxiliar = Trim(RsDoc!ConSerie) & RsDoc!ConValor + 1
            ret_Serie = Trim(RsDoc!ConSerie)
            ret_Numero = RsDoc!ConValor + 1
            
            RsDoc.Edit
            RsDoc!ConValor = RsDoc!ConValor + 1
            RsDoc.Update
        End If
    End If
    RsDoc.Close
    
    
    NumeroDocumento = Auxiliar
    
End Function

