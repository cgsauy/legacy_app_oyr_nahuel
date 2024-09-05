Attribute VB_Name = "modValidEmail"
Option Explicit
' list of top level Domains, obtained from www.IANA.com.
' Can and will change, used in host name syntax checking
Private Const TOP_DOMAINS = " COM ORG NET EDU GOV MIL INT BIZ AF AL DZ AS " & _
        "AD AO AI AQ AG AR AM AW AC AU AT AZ BS BH BD BB BY BZ BT BJ " & _
        "BE BM BO BA BW BV BR IO BN BG BF BI KH CM CA CV KY CF TD CL " & _
        "CN CX CC CO KM CD CG CK CR CI HR CU CY CZ DK DJ DM DO TP EC " & _
        "EG SV GQ ER EE ET FK FO FJ FI FR GF PF TF GA GM GE DE GH GI " & _
        "GR GL GD GP GU GT GG GN GW GY HT HM VA HN HK HU IS IN ID IR " & _
        "IQ IE IM IL IT JM JP JE JO KZ KE KI KP KR KW KG LA LV LB LS " & _
        "LR LY LI LT LU MO MK MG MW MY MV ML MT MH MQ MR MU YT MX FM " & _
        "MD MC MN MS MA MZ MM NA NR NP NL AN NC NZ NI NE NG NU NF MP " & _
        "NO OM PK PW PA PG PY PE PH PN PL PT PR QA RE RO RU RW KN LC " & _
        "VC WS SM ST SA SN SC SL SG SK SI SB SO ZA GS ES LK SH PM SD " & _
        "SR SJ SZ SE CH SY TW TJ TZ TH TG TK TO TT TN TR TM TC TV UG " & _
        "UA AE GB US UM UY UZ VU VE VN VG VI WF EH YE YU ZR ZM ZW UK "
        
Private Function IsDottedQuad(ByVal HostString As String) As Boolean

    ' verify that a string is 'xxx.xxx.xxx.xxx' format
    
    Dim sSplit()        As String
    Dim iCtr            As Integer

    ' split at the "."
     or_Split sSplit, HostString, "."
    
    ' should be 4 elements
    If UBound(sSplit) <> 3 Then Exit Function

    ' check each element
    For iCtr = 0 To 3
        ' should be numeric
        If Not IsNumeric(sSplit(iCtr)) Then Exit Function

        ' range check
        If iCtr = 0 Then
            If Val(sSplit(iCtr)) > 239 Then Exit Function
        Else
            If Val(sSplit(iCtr)) > 255 Then Exit Function
        End If
    Next
    
    IsDottedQuad = True

End Function

Public Function IsValidIPHost(ByVal HostString As String) As Boolean

    ' validate a host string
    Dim nCharacter As Integer
    Dim sBuffer As String
    
    For nCharacter = 1 To Len(HostString)
        sBuffer = Mid$(HostString, nCharacter, 1)
        If sBuffer = "@" Or sBuffer = " " Then
            IsValidIPHost = False
            Exit Function
        End If
    Next nCharacter
    
    Dim sHost               As String
    Dim sSplit()            As String

    sHost = UCase$(Trim$(HostString))

    ' if it's a dotted quad it's OK
    If IsDottedQuad(sHost) Then
        IsValidIPHost = True
        Exit Function
    End If

    or_Split sSplit, sHost, "."

    ' it's dotted quad, top level domain?
    If UBound(sSplit) > 0 And InStr(TOP_DOMAINS, " " & sSplit(UBound(sSplit)) & " ") > 0 Then
        IsValidIPHost = True
        Exit Function
    End If

End Function

Private Sub or_Split(arrObj As Variant, ByVal sIn As String, Optional sDelim As String = " ", _
    Optional nLimit As Long = -1, Optional bCompare As VbCompareMethod = vbBinaryCompare)

    Dim nC As Long, nPos As Long, nDelimLen As Long
'    Dim sOut() As String
    
    If sDelim <> "" Then
        nDelimLen = Len(sDelim)
        nPos = InStr(1, sIn, sDelim, bCompare)
        Do While nPos
            ReDim Preserve arrObj(nC)
            arrObj(nC) = Left(sIn, nPos - 1)
            sIn = Mid(sIn, nPos + nDelimLen)
            nC = nC + 1
            If nLimit <> -1 And nC >= nLimit Then Exit Do
            nPos = InStr(1, sIn, sDelim, bCompare)
        Loop
    End If

    ReDim Preserve arrObj(nC)
    arrObj(nC) = sIn

End Sub


Public Function IsEMailAddress(ByVal sEmail As String, Optional ByRef sReason As String) As Boolean
        
    Dim sPreffix As String
    Dim sSuffix As String
    Dim sMiddle As String
    Dim nCharacter As Integer
    Dim sBuffer As String

    sEmail = Trim(sEmail)

    If Len(sEmail) < 8 Then
        IsEMailAddress = False
        sReason = "Muy corta"
        Exit Function
    End If

    If InStr(sEmail, "@") = 0 Then
        IsEMailAddress = False
        sReason = "Falta @"
        Exit Function
    End If


    If InStr(InStr(sEmail, "@") + 1, sEmail, "@") <> 0 Then
        IsEMailAddress = False
        sReason = "Demasiadas @"
        Exit Function
    End If

    If InStr(sEmail, ".") = 0 Then
        IsEMailAddress = False
        sReason = "Falta el punto"
        Exit Function
    End If

    If InStr(sEmail, "@") = 1 Or InStr(sEmail, "@") = Len(sEmail) Or _
        InStr(sEmail, ".") = 1 Or InStr(sEmail, ".") = Len(sEmail) Then
        IsEMailAddress = False
        sReason = "Formato no válido"
        Exit Function
    End If

    For nCharacter = 1 To Len(sEmail)
        sBuffer = Mid$(sEmail, nCharacter, 1)
        If Not (LCase(sBuffer) Like "[a-z]" Or sBuffer = "@" Or sBuffer = "." Or sBuffer = "-" Or sBuffer = "_" Or IsNumeric(sBuffer)) Then
            IsEMailAddress = False: sReason = "Caracter no válido."
            Exit Function
        End If
    Next nCharacter
    
    nCharacter = 0
    
    On Error Resume Next
    
    sBuffer = Right(sEmail, 4)
    If InStr(sBuffer, ".") = 0 Then GoTo TooLong:
    If Left(sBuffer, 1) = "." Then sBuffer = Right(sBuffer, 3)
    If Left(Right(sBuffer, 3), 1) = "." Then sBuffer = Right(sBuffer, 2)
    If Left(Right(sBuffer, 2), 1) = "." Then sBuffer = Right(sBuffer, 1)
    
    
    If Len(sBuffer) < 2 Then
        IsEMailAddress = False
        sReason = "Sufijo muy corto"
        Exit Function
    End If
    
TooLong:
    
    If Len(sBuffer) > 3 Then
        IsEMailAddress = False
        sReason = "Sufijo muy largo"
        Exit Function
    End If
    
    sReason = Empty
    IsEMailAddress = True

End Function

