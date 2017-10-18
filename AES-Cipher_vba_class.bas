'AES encryption / decryption class using CryptAPI

Option Explicit

'Constant definition for CryptAPI
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const ALG_TYPE_BLOCK As Long = 1536
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_SID_AES_128 As Long = 14
Private Const ALG_SID_AES_192 As Long = 15
Private Const ALG_SID_AES_256 As Long = 16
Private Const PROV_RSA_AES As Long = 24
Private Const CALG_AES_128 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_128
Private Const CALG_AES_192 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_192
Private Const CALG_AES_256 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_256
Private Const KP_IV As Long = 1
Private Const KP_PADDING As Long = 3
Private Const KP_MODE As Long = 4
Private Const PKCS5_PADDING As Long = 1
Private Const CRYPT_MODE_CBC As Long = 1
Private Const PLAINTEXTKEYBLOB As Long = 8
Private Const CUR_BLOB_VERSION As Long = 2

'Constant definition for Windows API
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Const FORMAT_MESSAGE_FROM_STRING As Long = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = 8192
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = 255
Private Const LANG_NEUTRAL As Long = &H0
Private Const SUBLANG_DEFAULT As Long = &H1


'CryptoAPI definition
#If VBA7 And Win64 Then

Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptImportKey Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As LongPtr) As Long
Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function CryptSetKeyParam Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare PtrSafe Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

#Else

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" ( _
    ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

#End If


'WindowsAPI definition
#If VBA7 And Win64 Then

Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" ( _
    ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByRef lpBuffer As LongPtr, ByVal nSize As Long, _
    ByRef Arguments As Any) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" ( _
    ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long

#Else

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" ( _
    ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByRef lpBuffer As Long, ByVal nSize As Long, _
    ByRef Arguments As Any) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" ( _
    ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

#End If


'BLOBHEADER user-defined type
Private Type BLOBHEADER
    bType As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
End Type

'User defined type of key data for import
'
'PUBLICKEYSTRUC BLOB header followed by key size and key data,
'As for the key data, the array size changes according to the key size,
'Dynamically allocate memory in logic and make it undefined here
Private Type keyBlob
    hdr As BLOBHEADER
    keySize As Long
'    keyData() As Byte
End Type

'Key Length Constant Definition
Public Enum AESKeyBits
    AES_KEY128 = 128
    AES_KEY192 = 192
    AES_KEY256 = 256
End Enum

'Error code definition
Private Const ERR_CRYPT_API = vbObjectError + 513   'CryptAPI error
Private Const ERR_KEY_LENGTH = vbObjectError + 514  'Key length error
Private Const ERR_IV_LENGTH = vbObjectError + 515   'IV length error

'AES/CBC/PKCS5 Padding decoding processing
'
'argument:
'   [in]         key: key byte string
'   [in]          iv: IV byte sequence
'   [in,out]    data: [in]encrypted byte string /[out] decryppted byte string
'   [in]     keyBits: key bit length (default 128bit)
'
'return value:
'   None
Public Sub decrypt(ByRef key() As Byte, ByRef iv() As Byte, ByRef data() As Byte, Optional ByVal keyBits As AESKeyBits = AES_KEY128)
#If VBA7 And Win64 Then
    Dim hProv As LongPtr   'CSP handler
    Dim hKey As LongPtr    'cryptographic key handler
#Else
    Dim hProv As Long   'CSP handler
    Dim hKey As Long    'cryptographic key handler
#End If
    Dim algid As Long   'encryption algorithm

    On Error GoTo ErrorHandler

    'Set encryption algorithm ID from key length of AES
    Select Case keyBits
        Case AES_KEY128
            algid = CALG_AES_128
        Case AES_KEY192
            algid = CALG_AES_192
        Case AES_KEY256
            algid = CALG_AES_256
    End Select

    Dim keyLength As Long   'key byte length
    keyLength = keyBits / 8 'bit->byte conversion

    'Check key length
    If UBound(key) + 1 <> keyLength Then
        Err.Raise ERR_KEY_LENGTH, "decrypt()", "The key length is invalid: " & UBound(key) + 1 & "byte"
    End If

    'Check IV length
    If UBound(iv) + 1 <> 16 Then
        Err.Raise ERR_IV_LENGTH, "decrypt()", "IV length is invalid: " & UBound(iv) + 1 & "byte"
    End If

    'CSP(Cryptographic Service Provider) handle
    If Not CBool(CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptAcquireContext()", Err.LastDllError
    End If

    Dim blob As keyBlob 'Key data (user defined type)
    Dim keyData() As Byte   'key data (byte sequence)
    
    'Create key data
    'KeyBlob Forcibly create a byte string combining key data into a user-defined type

    blob.hdr.bType = PLAINTEXTKEYBLOB
    blob.hdr.bVersion = CUR_BLOB_VERSION
    blob.hdr.reserved = 0
    blob.hdr.aiKeyAlg = algid
    blob.keySize = keyLength
    ReDim keyData(LenB(blob) + blob.keySize - 1)
    Call CopyMemory(keyData(0), blob, LenB(blob))
    Call CopyMemory(keyData(LenB(blob)), key(0), keyLength)

    'Import key
    If Not CBool(CryptImportKey(hProv, keyData(0), UBound(keyData) + 1, 0, 0, hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptImportKey()", Err.LastDllError
    End If

    'Setting of padding method (PKCS#5)
    If Not CBool(CryptSetKeyParam(hKey, KP_PADDING, PKCS5_PADDING, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_PADDING", Err.LastDllError
    End If

    'IV (Initialization Vector) setting
    If Not CBool(CryptSetKeyParam(hKey, KP_IV, iv(0), 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_IV", Err.LastDllError
    End If

    'Setting encryption mode (ciphertext block chain mode)
    If Not CBool(CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_CBC, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_MODE", Err.LastDllError
    End If

    'Encrypted byte sequence length
    Dim dwDataLen As Long
    dwDataLen = UBound(data) + 1

    'CryptDecrypt is a specification that returns the decrypted byte sequence to the encrypted byte sequence of the argument
    'Copy the encrypted byte string of the method argument to the local variable and use it
    Dim pbData() As Byte
    ReDim pbData(dwDataLen - 1)
    Call CopyMemory(pbData(0), data(0), UBound(data) + 1)

    'Decryption processing
    If Not CBool(CryptDecrypt(hKey, 0, True, 0, pbData(0), dwDataLen)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptDecrypt()", Err.LastDllError
    End If

    ReDim Preserve pbData(dwDataLen - 1)
    data = pbData

    'Release of encryption key handler
    If Not CBool(CryptDestroyKey(hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptDestroyKey()", Err.LastDllError
    End If

    'Release of CSP handler
    If Not CBool(CryptReleaseContext(hProv, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptReleaseContext()", Err.LastDllError
    End If

    Exit Sub

ErrorHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    errMessage = ""

    If Err.Number <> 0 Then
        If Err.Number = ERR_CRYPT_API Then
            errNumber = Err.Description
            errSource = Err.Source
            errMessage = GetErrorText(Err.Description)
        Else
            errNumber = Err.Number
            errSource = Err.Source
            errMessage = Err.Description
        End If
    End If

    Err.Clear

    If Not hKey <> 0 Then
        'Release of encryption key handler
        Call CryptDestroyKey(hKey)
    End If

    If Not hProv <> 0 Then
        'Release of CSP handler
        Call CryptReleaseContext(hProv, 0)
    End If

    On Error GoTo 0
    If errMessage <> "" Then
        Err.Raise Number:=errNumber, Source:=errSource, Description:=errMessage
    End If
End Sub

'AES/CBC/PKCS5Padding encryption processing
'
'argument:
'   [in]         key: key byte string
'   [in]          iv: IV byte sequence
'   [in,out]    data: [in] plain text byte string / [out] encrypted byte string
'   [in]     keyBits: Key bit length (default 128bit)
'
'Return value:
'   None
Public Sub encrypt(ByRef key() As Byte, ByRef iv() As Byte, ByRef data() As Byte, Optional ByVal keyBits As AESKeyBits = AES_KEY128)
#If VBA7 And Win64 Then
    Dim hProv As LongPtr   'CSP handler
    Dim hKey As LongPtr    'cryptographic key handler
#Else
    Dim hProv As Long   'CSP handler
    Dim hKey As Long    'cryptographic key handler
#End If
    Dim algid As Long   'encryption algorithm

    On Error GoTo ErrorHandler

    'Set encryption algorithm ID from key length of AES
    Select Case keyBits
        Case AES_KEY128
            algid = CALG_AES_128
        Case AES_KEY192
            algid = CALG_AES_192
        Case AES_KEY256
            algid = CALG_AES_256
    End Select

    Dim keyLength As Long   'key length
    keyLength = keyBits / 8 'bit->byte conversion

    'Check key length
    If UBound(key) + 1 <> keyLength Then
        Err.Raise ERR_KEY_LENGTH, "decrypt()", "The key length is invalid: " & UBound(key) + 1 & "byte"
    End If

    'Check IV length
    If UBound(iv) + 1 <> 16 Then
        Err.Raise ERR_IV_LENGTH, "decrypt()", "IV length is invalid: " & UBound(iv) + 1 & "byte"
    End If

    'Get the CSP (Cryptographic Service Provider) handle
    If Not CBool(CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptAcquireContext()", Err.LastDllError
    End If

    Dim blob As keyBlob 'Key data (user defined type)
    Dim keyData() As Byte  'key data (byte sequence)

    'Create key data
    'KeyBlob Forcibly create a byte string combining key data into a user-defined type
    blob.hdr.bType = PLAINTEXTKEYBLOB
    blob.hdr.bVersion = CUR_BLOB_VERSION
    blob.hdr.reserved = 0
    blob.hdr.aiKeyAlg = algid
    blob.keySize = keyLength
    ReDim keyData(LenB(blob) + blob.keySize - 1)
    Call CopyMemory(keyData(0), blob, LenB(blob))
    Call CopyMemory(keyData(LenB(blob)), key(0), keyLength)

    'import key
    If Not CBool(CryptImportKey(hProv, keyData(0), UBound(keyData) + 1, 0, 0, hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptImportKey()", Err.LastDllError
    End If

    'Setting of padding method (PKCS#5)
    If Not CBool(CryptSetKeyParam(hKey, KP_PADDING, PKCS5_PADDING, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_PADDING", Err.LastDllError
    End If

    'IV(Initialization Vector) setting
    If Not CBool(CryptSetKeyParam(hKey, KP_IV, iv(0), 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_IV", Err.LastDllError
    End If

    'Setting encryption mode (ciphertext block chain mode)
    If Not CBool(CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_CBC, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_MODE", Err.LastDllError
    End If

    'Plain text byte sequence length
    Dim dwPlainDataLen As Long
    dwPlainDataLen = UBound(data) + 1

    'Encrypted byte sequence length
    Dim dwCryptDataLen As Long
    dwCryptDataLen = dwPlainDataLen

    'CryptEncrypt is a specification that returns an encrypted byte sequence to the plaintext byte sequence of the argument
    'Copy the plaintext byte string of the method argument to the local variable and use it
    Dim pbData() As Byte
    ReDim pbData(dwPlainDataLen - 1)
    Call CopyMemory(pbData(0), data(0), dwPlainDataLen)

    'Encryption processing
    'Inquire beforehand the byte sequence length after encryption and extend the buffer
    If Not CBool(CryptEncrypt(hKey, 0, True, 0, ByVal 0&, dwCryptDataLen, dwPlainDataLen)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptEncrypt()", Err.LastDllError
    End If
    If dwCryptDataLen > dwPlainDataLen Then
        ReDim Preserve pbData(dwCryptDataLen - 1)
    End If
    If Not CBool(CryptEncrypt(hKey, 0, True, 0, pbData(0), dwPlainDataLen, dwCryptDataLen)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptEncrypt()", Err.LastDllError
    End If

    data = LeftB(pbData, dwCryptDataLen)

    'Release of encryption key handler
    If Not CBool(CryptDestroyKey(hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptDestroyKey()", Err.LastDllError
    End If

    'Release of CSP handler
    If Not CBool(CryptReleaseContext(hProv, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptReleaseContext()", Err.LastDllError
    End If

    Exit Sub

ErrorHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    errMessage = ""

    If Err.Number <> 0 Then
        If Err.Number = ERR_CRYPT_API Then
            errNumber = Err.Description
            errSource = Err.Source
            errMessage = GetErrorText(Err.Description)
        Else
            errNumber = Err.Number
            errSource = Err.Source
            errMessage = Err.Description
        End If
    End If

    Err.Clear

    If Not hKey <> 0 Then
        'Release of encryption key handler
        Call CryptDestroyKey(hKey)
    End If

    If Not hProv <> 0 Then
        'Release of CSP handler
        Call CryptReleaseContext(hProv, 0)
    End If

    On Error GoTo 0
    If errMessage <> "" Then
        Err.Raise Number:=errNumber, Source:=errSource, Description:=errMessage
    End If
End Sub

'Method implementation implementation of the MAKELANGID macro
Private Function MAKELANGID(ByVal p As Long, ByVal s As Long) As Long
    MAKELANGID = (CLng(CInt(s)) * 1024) Or CLng(CInt(p))
End Function

'Get error message from error code
Private Function GetErrorText(ByVal ErrorCode As Long) As String
#If VBA7 And Win64 Then
    Dim lpBuffer As LongPtr
#Else
    Dim lpBuffer As Long
#End If
    Dim messageLength As Long

    messageLength = FormatMessage( _
        FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        0, ErrorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
        lpBuffer, 0, 0)

    If messageLength = 0 Then
        GetErrorText = ""
    Else
        GetErrorText = Space$(messageLength)
        Call lstrcpy(ByVal StrPtr(GetErrorText), ByVal lpBuffer)
        Call LocalFree(lpBuffer)
    End If
End Function
