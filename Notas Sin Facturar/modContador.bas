Attribute VB_Name = "modContador"
Option Explicit

'DOCUMENTOS

Global Const DocCredito = "Cr�dito"
Global Const DocContado = "Contado"
Global Const DocRemito = "Remito"
Global Const DocNDevolucion = "Nota Devoluci�n"
Global Const DocNCredito = "Nota Cr�dito"
Global Const DocRecibo = "Recibo de Pago"
Global Const DocCreditoDomicilio = "Cr�dito Domicilio"
Global Const DocContadoDomicilio = "Contado Domicilio"
Global Const DocServicioDomicilio = "Servicio a Domicilio"
Global Const DocNEspecial = "Nota Especial"
Global Const DocNDebito = "Nota de D�bito"

Public paDContado As String
Public paDCredito As String
Public paDNDevolucion As String
Public paDNCredito As String
Public paDRecibo As String
Public paDNEspecial As String
Public paDNDebito As String
Public paDRemito As String

'--------------------------------------------------------------------------------
'   Devuelve un n�mero para el tipo de documento pasado como parametro.
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
Public Function NumeroDocumento(Documento As String)

    Dim Auxiliar As String    'Auxiliar para retornar el NRO DOC (Serie + Nro)
    Dim RsDoc As rdoResultset
    
    cons = "Select * from Contador Where ConDocumento = '" & Trim(Documento) & "'"
    Set RsDoc = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurLock)
    
    If RsDoc.EOF Then
        RsDoc.AddNew
        RsDoc!ConDocumento = Trim(Documento)
        RsDoc!ConValor = 1
        RsDoc!ConSerie = "A"
        Auxiliar = RsDoc!ConSerie & RsDoc!ConValor
        RsDoc.Update
    
    Else
        If RsDoc!ConValor = 999999 Then
            RsDoc.Edit
            RsDoc!ConValor = 1
            RsDoc!ConSerie = Chr(Asc(RsDoc!ConSerie) + 1)
            Auxiliar = RsDoc!ConSerie & RsDoc!ConValor
            RsDoc.Update
        Else
            Auxiliar = Trim(RsDoc!ConSerie) & RsDoc!ConValor + 1
            RsDoc.Edit
            RsDoc!ConValor = RsDoc!ConValor + 1
            RsDoc.Update
        End If
    End If
    RsDoc.Close
    
    NumeroDocumento = Auxiliar
    
End Function

Public Function CodigoDeBarras(TipoDoc As Integer, CodigoDoc As Long)

    If Len(CodigoDoc) < 6 Then
        CodigoDeBarras = TipoDoc & "D" & Format(CodigoDoc, "000000")
    Else
        CodigoDeBarras = TipoDoc & "D" & CodigoDoc
    End If
    CodigoDeBarras = "*" & CodigoDeBarras & "*"
    
End Function
