Attribute VB_Name = "modEncryption"
Private m_strPassword As String
Private mintPassWordIndex As Integer
Private mabytePassword() As Byte

Function EncryptoString(strIn As String) As String
  ' Comments  : Encrypts/Decrypts the passed string with XOR encryption
  ' Parameters: strIn - string to encrypt/decrypt
  ' Returns   : encrypted/unencrypted string
  ' Source    : Total VB SourceBook 6
  '
Dim intCounter As Long

  On Error GoTo PROC_ERR
  Dim aux As String
  aux = ""
  For I = 1 To Len(strIn)
    aux = aux & Chr(Asc(Mid(strIn, I, 1)))
  Next
  strIn = aux
  
  
  
  
  
  ' reset the password to the byte array
  'm_strPassword = "Password"
  m_strPassword = "0123456789"
  ReDim mabytePassword(LenB(m_strPassword)) As Byte
    
  mabytePassword = m_strPassword
  
  ' Create a byte array to store the string to encrypt
  ReDim abytIn(LenB(strIn)) As Byte
  
  ' assign the string to the byte array
  abytIn = strIn
  
  ' Reset the password index
  mintPassWordIndex = 0

  ' For each byte in the string
  For intCounter = 0 To LenB(strIn) - 1
    ' Encrypt the byte
    EncryptByte abytIn(intCounter)
  Next
  
  ' Return the encrypted string
  EncryptoString = abytIn
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "EncryptString"
  Resume PROC_EXIT
  
End Function

Private Function EncryptByte(bytIn As Byte) As Byte
  ' Comments  : This function Encrypts one byte, and modifies the password.
  '                     Modifying the password as we encrypt makes the encyption slightly harder to break.
  ' Parameters: byteIn - The byte to encrypt. The encrypted byte is returned in this parameter
  ' Returns   : The encrypted byte
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  ' Exclusive or the byte with the current password byte
  bytIn = bytIn Xor mabytePassword(mintPassWordIndex)
  ' Exclusive or the byte with the first character of the password
  ' multiplied by the current index into the password. And the result with
  ' 256 to avoid possible overflow errors
  bytIn = (bytIn Xor CInt(mabytePassword(mintPassWordIndex)) * _
    mintPassWordIndex) And &HFF
  
  ' Assign the encrypted byte to the function return value
  EncryptByte = bytIn
  
  ' Modify the password.
    
  If mintPassWordIndex < UBound(mabytePassword) Then
    ' set the current byte in the password to the current byte plus the
    ' next byte.
    mabytePassword(mintPassWordIndex) = _
      (CInt(mabytePassword(mintPassWordIndex)) + _
      mabytePassword(mintPassWordIndex + 1)) And &HFF
    
    ' Increment the password index
    mintPassWordIndex = mintPassWordIndex + 1
  Else
    ' If the password length has been exceeded, wrap around to the
    ' beginning set the current byte in the password to the current byte
    ' plus the first byte. And the result with 256 to avoid possible
    ' overflow errors
    mabytePassword(mintPassWordIndex) = _
      (CInt(mabytePassword(mintPassWordIndex)) + mabytePassword(1)) _
      And &HFF

    ' Reset the password index
    mintPassWordIndex = 1
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "EncryptByte"
  Resume PROC_EXIT

End Function




