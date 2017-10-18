'**************************************************
'
' VBA code to create encrypted PDF with qpdf.exe
'
' http://pdf-file.nnn2.com/?p=867
'
'**************************************************

Option Explicit

Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare PtrSafe Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare PtrSafe Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessID As Long) As Long
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SYNCHRONIZE = 1048576
Const PROCESS_QUERY_INFORMATION = &H400

'Qpdf-6.00 qpdf.exe
'Const CON_QPDF_PATH = "c:\qpdf\bin\qpdf.exe"

'   Temp File No
Private gFileCnt    As Long

'   Debug Mode [ True=On | False=Off ]
Private gDebugMode  As Boolean

'Main Call
'Sub encryptPDF()
Public Sub EncryptPDF( _
    ByVal userPassword As String, _
    ByVal inFile As String, _
    ByVal outFile As String)

    gDebugMode = False

    Dim qpdfPara_Password(1)    As String
    Dim qpdfPara_keyLength      As Long
    Dim qpdfPara_Flags(7)       As String
    Dim qpdfPara_InPdfPath      As String
    Dim qpdfPara_InPdfPassword  As String
    Dim qpdfPara_OutPdfPath     As String
    Dim qpdfPara_OrverWrite     As Boolean
    Dim strErr                  As String

    Dim i                       As Long

    'Initialize
    For i = 0 To UBound(qpdfPara_Password)
        qpdfPara_Password(i) = ""
    Next i
    For i = 0 To UBound(qpdfPara_Flags)
        qpdfPara_Flags(i) = ""
    Next i

    'user - Password: Specify the user password
    qpdfPara_Password(0) = userPassword
    'owner - Password: Specify the owner password
    qpdfPara_Password(1) = OWNER_PASSWORD
    'key-length: Specify 40, 128, or 256 as the key length
    qpdfPara_keyLength = 128

    'flags: This depends on the above key-length value

    'Let's make unused places comments
    Select Case qpdfPara_keyLength
    Case 40
    'When 40 (= qpdfPara_keyLength) [Y: permit] [N: disallow]
        '0.--print=[yn] : Print
        qpdfPara_Flags(0) = "N"
        '1.--modify=[yn] : Document change
        qpdfPara_Flags(1) = "N"
        '2.--extract=[yn] : Extract text / graphics
        qpdfPara_Flags(2) = "N"
        '3.--annotate=[yn] : comments and form filling and signing
        qpdfPara_Flags(3) = "Y"
    '
    Case 128
    'When 128 (= qpdfPara_keyLength) [Y: permit] [N: disallow]
        '0.--accessibility=[yn] : Access to blind people
        qpdfPara_Flags(0) = "Y"
        '1.--extract=[yn] : text / graphics extraction
        qpdfPara_Flags(1) = "N"
        '2.--print=print-opt : Control print access
            'full: 1. Full printing possible
            'If you mean high-resolution printing
            'low: 2. Allow only low resolution printing
            'none: 3. Do not allow printing
        qpdfPara_Flags(2) = "3"
        '3.--modify=modify-opt : Control change access
        'all: 1. Allow full document change
        'annotate: 2. Allow creation of comments and manipulation of forms
        'form: 3. Allow form field input and signature
        'assembly: 4. Allow assembly of documents only
        'none: 5. Do not allow changes
        qpdfPara_Flags(3) = "5"
        '4.--cleartext-metadata : Prevent encryption of metadata
            '[Y: Prevent] [N: not prevent]
        qpdfPara_Flags(4) = "Y"
        '5.--use-aes=[yn] : Whether to use AES encryption [Y: Use] [N: Not used]
        qpdfPara_Flags(5) = "Y"
        '6.--force-V4 : V = 4 Forcing the use of encryption handlers [Y: Use] [N: Not used]
        qpdfPara_Flags(6) = "Y"
    '
    Case 256
    'When '256 (= qpdfPara_keyLength) [Y: permit] [N: disallow]
        '0.--accessibility=[yn] : Access to blind people
        qpdfPara_Flags(0) = "Y"
        '1.--extract=[yn] : text / graphics extraction
        qpdfPara_Flags(1) = "N"
        '2.--print=print-opt : Control print access
            'full: 1. Full printing possible
            'If you mean high-resolution printing
            'low: 2. Allow only low resolution printing
            'none: 3. Do not allow printing
        qpdfPara_Flags(2) = "1"
        '3.--modify=modify-opt : Control change access
        'all: 1. Allow full document change
        'annotate: 2. Allow creation of comments and manipulation of forms
        'form: 3. Allow form field input and signature
        'assembly: 4. Allow assembly of documents only
        'none: 5. Do not allow changes
        qpdfPara_Flags(3) = "5"
        '4.--cleartext-metadata : Prevent encryption of metadata [Y: Prevent] [N: not prevent]
        qpdfPara_Flags(4) = "Y"
        '5.--use-aes=y : Use of AES encryption is always on for 256-bit keys
        qpdfPara_Flags(5) = "Y"  'fixed
        '6.--force-V4 : Not available with 256 bits
        qpdfPara_Flags(6) = "N"  'fixed
        '7.--force-R5 : deprecated R5 encryption [Y: Use] [N: Not used]
        qpdfPara_Flags(7) = "N"

    End Select

    If gDebugMode Then Debug.Print "Start:" & Now
'    Dim j   As Long
'    For j = 0 To ****

    'Full path of input PDF
    'qpdfPara_InPdfPath = ThisWorkbook.Path & _
        "\" & "TEST.pdf"
    qpdfPara_InPdfPath = inFile
    'User password (when opening document) of input PDF
    qpdfPara_InPdfPassword = ""
    'Full path of output PDF
    'qpdfPara_OutPdfPath = ThisWorkbook.Path & _
        "\" & "2017_Tony.pdf"
    qpdfPara_OutPdfPath = outFile

    'Overwrite on output?
    qpdfPara_OrverWrite = True

    strErr = ""
    Call qpdfSetEncryption( _
        qpdfPara_Password(), qpdfPara_keyLength, _
        qpdfPara_Flags(), qpdfPara_InPdfPath, _
        qpdfPara_InPdfPassword, qpdfPara_OutPdfPath, _
        qpdfPara_OrverWrite, strErr)

    If strErr <> "" Then
        MsgBox strErr, vbCritical, "execution error"
        Exit Sub
    End If

'    Next j
    If gDebugMode Then Debug.Print "End  :" & Now

    If gDebugMode Then MsgBox "End"
End Sub

'**************************************************
'
' Set Encryption Options for the PDF file
'
' Function: Perform security setting with qpdf.exe
' Create  : 2016/06/27
' Update  : 2016/06/27
' Vertion : 1.0.0
'
' Argument: See description of argument below URL
' http://pdf-file.nnn2.com/?p=867
'
' Returns: Nothing
'
' Remarks: When strErr <> "", treat it as error.
' URL: http://pdf-file.nnn2.com/?p=867
' Others: We do not claim copyright.
' I'm happy if you can comment on the above URL.
'
'**************************************************

Public Sub qpdfSetEncryption( _
    ByRef qpdfPara_Password() As String, _
    ByVal qpdfPara_keyLength As Long, _
    ByRef qpdfPara_Flags() As String, _
    ByVal qpdfPara_InPdfPath As String, _
    ByVal qpdfPara_InPdfPassword As String, _
    ByVal qpdfPara_OutPdfPath As String, _
    ByVal qpdfPara_OrverWrite As Boolean, _
    ByRef strErr As String)

On Error GoTo Err_qpdfSetEncryption:

    Dim strQpdfPath  As String
    Dim strTempFilePath  As String
    Dim strCmd          As String
    Dim i               As Long
    Dim objFileSystem   As Object
    Dim CON_QPDF_PATH As String

    'Initialization
    CON_QPDF_PATH = ThisWorkbook.Path & "\bin\qpdf.exe"

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    strErr = ""

    'Error checking
    If qpdfPara_Password(0) = "" And _
        qpdfPara_Password(1) = "" Then
        strErr = "No password has been set"
        Exit Sub
    End If
    If Not (qpdfPara_keyLength = 40 Or _
        qpdfPara_keyLength = 128 Or _
        qpdfPara_keyLength = 256) Then
        strErr = "keyLength is not 40, 128, 256"
        Exit Sub
    End If
    If objFileSystem.FileExists(qpdfPara_InPdfPath) = False Then
        strErr = qpdfPara_InPdfPath & vbCrLf & _
            "Source PDF file does not exist"
        Exit Sub
    End If
    If objFileSystem.FileExists(qpdfPara_OutPdfPath) = True Then
        If qpdfPara_OrverWrite = False Then
            strErr = qpdfPara_OutPdfPath & vbCrLf & _
                "Output PDF file exists"
            Exit Sub
        End If
    End If
    If objFileSystem.FileExists(CON_QPDF_PATH) = False Then
        strErr = CON_QPDF_PATH & vbCrLf & _
            "This file does not exist"
        Exit Sub
    End If

    'Edit command line
    strCmd = CON_QPDF_PATH & " --encrypt "
    If qpdfPara_Password(0) <> "" Then
        strCmd = strCmd & qpdfPara_Password(0) & " "
    Else
        strCmd = strCmd & """"" "
    End If
    If qpdfPara_Password(1) <> "" Then
        strCmd = strCmd & qpdfPara_Password(1) & " "
    Else
        strCmd = strCmd & """"" "
    End If
    strCmd = strCmd & qpdfPara_keyLength & " "

    Select Case qpdfPara_keyLength
    Case 40
    'When it is' 40 (= qpdfPara_keyLength) [Y: permit] [N: disallow]
        '0.--print=[yn] : Print strCmd
        Select Case qpdfPara_Flags(0)
        Case "Y": strCmd = strCmd & "--print=y "
        Case "N": strCmd = strCmd & "--print=n "
        Case Else
            strErr = "Use for --print=[yn] " & _
                    "qpdfPara_Flags(0) is not set"
            Exit Sub
        End Select
        '1.--modify=[yn] : Document change
        Select Case qpdfPara_Flags(1)
        Case "Y": strCmd = strCmd & "--modify=y "
        Case "N": strCmd = strCmd & "--modify=n "
        Case Else
            strErr = "Use for --modify=[yn] " & _
                "qpdfPara_Flags(1) is not set"
            Exit Sub
        End Select
        '2.--extract=[yn] : Extract text / graphics
        Select Case qpdfPara_Flags(2)
        Case "Y": strCmd = strCmd & "--extract=y "
        Case "N": strCmd = strCmd & "--extract=n "
        Case Else
            strErr = "Use for --extract=[yn] " & _
                "qpdfPara_Flags(2) is not set"
            Exit Sub
        End Select
        '3.--annotate=[yn] : comments and form filling and signing
        Select Case qpdfPara_Flags(3)
        Case "Y": strCmd = strCmd & "--annotate=y "
        Case "N": strCmd = strCmd & "--annotate=n "
        Case Else
            strErr = "Use for --annotate=[yn] " & _
                "qpdfPara_Flags(3) is not set"
            Exit Sub
        End Select
    '
    Case 128
    'When 128 (= qpdfPara_keyLength) [Y: permit] [N: disallow]
        '0.--accessibility=[yn] : Access to blind people
        Select Case qpdfPara_Flags(0)
        Case "Y": strCmd = strCmd & "--accessibility=y "
        Case "N": strCmd = strCmd & "--accessibility=n "
        Case Else
            strErr = "Use for --accessibility=[yn] " & _
                "qpdfPara_Flags(0) is not set"
            Exit Sub
        End Select
        '1.--extract=[yn] : text / graphics extraction
        Select Case qpdfPara_Flags(1)
        Case "Y": strCmd = strCmd & "--extract=y "
        Case "N": strCmd = strCmd & "--extract=n "
        Case Else
            strErr = "Use for --extract=[yn] " & _
                "qpdfPara_Flags(1) is not set"
            Exit Sub
        End Select
        '2.--print=print-opt : Control print access
            'full    : 1.Full printing possible
            '  If you mean high-resolution printing
            'low     : 2.Allow only low resolution printing
            'none    : 3.Do not allow printing
        Select Case qpdfPara_Flags(2)
        Case "1": strCmd = strCmd & "--print=" & "full "
        Case "2": strCmd = strCmd & "--print=" & "low "
        Case "3": strCmd = strCmd & "--print=" & "none "
        Case Else
            strErr = "Use for --print=print-opt " & _
                "qpdfPara_Flags(2) is not set"
            Exit Sub
        End Select
        '3.--modify=modify-opt : Control change access
            'all     : 1.Allow full document change
            'annotate: 2.Allow creation of comments and manipulation of forms
            'form    : 3.Allow form field input and signature
            'assembly: 4.Allow assembly of documents only
            'none    : 5.Do not allow changes
        Select Case qpdfPara_Flags(3)
        Case "1": strCmd = strCmd & "--modify=" & "all "
        Case "2": strCmd = strCmd & "--modify=" & "annotate "
        Case "3": strCmd = strCmd & "--modify=" & "form "
        Case "4": strCmd = strCmd & "--modify=" & "assembly "
        Case "5": strCmd = strCmd & "--modify=" & "none "
        Case Else
            strErr = "Use for --modify=modify-opt " & _
                "qpdfPara_Flags(3) is not set"
            Exit Sub
        End Select
        '4.--cleartext-metadata : Prevent encryption of metadata
            '[Y: Prevent] [N: not prevent]
        Select Case qpdfPara_Flags(4)
        Case "Y": strCmd = strCmd & "--cleartext-metadata "
        Case "N":
        Case Else
            strErr = "Use for --cleartext-metadata " & _
                "qpdfPara_Flags(4) is not set"
            Exit Sub
        End Select
        '5.--use-aes=[yn] : Indicates whether to use AES encryption
            '[Y: Use] [N: Not used]
        Select Case qpdfPara_Flags(5)
        Case "Y": strCmd = strCmd & "--use-aes=y "
        Case "N": strCmd = strCmd & "--use-aes=n "
        Case Else
            strErr = "Use for --use-aes=[yn] " & _
                "qpdfPara_Flags(5) is not set"
            Exit Sub
        End Select
        '6.--force-V4 : V = 4 Forcing the use of encryption handlers
            '[Y: Use] [N: Not used]
        Select Case qpdfPara_Flags(6)
        Case "Y": strCmd = strCmd & "--force-V4 "
        Case "N":
        Case Else
            strErr = "Use for --force-V4 " & _
                "qpdfPara_Flags(6) is not set"
            Exit Sub
        End Select
    '
    Case 256
    'When '256 (= qpdfPara_keyLength) [Y: permit] [N: disallow]
        '0.--accessibility=[yn] : Access to blind people
        Select Case qpdfPara_Flags(0)
        Case "Y": strCmd = strCmd & "--accessibility=y "
        Case "N": strCmd = strCmd & "--accessibility=n "
        Case Else
            strErr = "Use for --accessibility=[yn] " & _
                "qpdfPara_Flags(0) is not set"
            Exit Sub
        End Select
        '1.--extract=[yn] : text / graphics extraction
        Select Case qpdfPara_Flags(1)
        Case "Y": strCmd = strCmd & "--extract=y "
        Case "N": strCmd = strCmd & "--extract=n "
        Case Else
            strErr = "Use for --extract=[yn] " & _
                "qpdfPara_Flags(1) is not set"
            Exit Sub
        End Select
        '2.--print=print-opt : Control print access
            'full    : 1. Full printing possible
            'If you mean high-resolution printing
            'low     : 2. Allow only low resolution printing
            'none    : 3. Do not allow printing
        Select Case qpdfPara_Flags(2)
        Case "1": strCmd = strCmd & "--print=" & "full "
        Case "2": strCmd = strCmd & "--print=" & "low "
        Case "3": strCmd = strCmd & "--print=" & "none "
        Case Else
            strErr = "Use for --print=print-opt " & _
                "qpdfPara_Flags(2) is not set"
            Exit Sub
        End Select
        '3.--modify=modify-opt : Control change access
            'all     : 1. Allow full document change
            'annotate: 2. Allow creation of comments and manipulation of forms
            'form    : 3. Allow form field input and signature
            'assembly: 4. Allow assembly of documents only
            'none    : 5. Do not allow changes
        Select Case qpdfPara_Flags(3)
        Case "1": strCmd = strCmd & "--modify=" & "all "
        Case "2": strCmd = strCmd & "--modify=" & "annotate "
        Case "3": strCmd = strCmd & "--modify=" & "form "
        Case "4": strCmd = strCmd & "--modify=" & "assembly "
        Case "5": strCmd = strCmd & "--modify=" & "none "
        Case Else
            strErr = "Use for --modify=modify-opt " & _
                "qpdfPara_Flags(3) is not set"
            Exit Sub
        End Select
        '4.--cleartext-metadata : Prevent encryption of metadata
            '[Y: Prevent] [N: not prevent]
        Select Case qpdfPara_Flags(4)
        Case "Y": strCmd = strCmd & "--cleartext-metadata "
        Case "N":
        Case Else
            strErr = "Use for --cleartext-metadata " & _
                "qpdfPara_Flags(4) is not set"
            Exit Sub
        End Select
        '5.--use-aes=y : Use of AES encryption is always on for 256-bit keys
        'qpdfPara_Flags(5) is fixed with'Y'ignored
        strCmd = strCmd & "--use-aes=y "
        '6.--force-V4 : Not available with 256 bits
        'qpdfPara_Flags (6) is "N" fixed and ignored
        '7.--force-R5 : deprecated R = 5 encryption
            '[Y: Use] [N: Not used]
        Select Case qpdfPara_Flags(7)
        Case "Y": strCmd = strCmd & "--force-R5 "
        Case "N":
        Case Else
            strErr = "Use for --force-R5 " & _
                "qpdfPara_Flags(7) is not set"
            Exit Sub
        End Select

    End Select
    strCmd = strCmd & "-- "

    If qpdfPara_InPdfPassword <> "" Then
        strCmd = strCmd & "--password=" & _
            qpdfPara_InPdfPassword & " "
    End If

    'temporary file
    gFileCnt = gFileCnt + 1
    'strTempFilePath = Application.ActiveWorkbook.Path &
    strTempFilePath = ThisWorkbook.Path & _
        "\" & Format(Now(), "yyyymmdd-hhmmss-") & gFileCnt & ".txt"

    'Note: put a single quotation "around the file path
    strCmd = strCmd & _
        """" & qpdfPara_InPdfPath & _
        """ """ & qpdfPara_OutPdfPath & _
        """ > """ & strTempFilePath & """ 2>&1"

    'Command line execution
    strCmd = "cmd /s /c " & strCmd
    Call RunCommandLine(strCmd, strErr)

'    Dim wsh As Object
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    Dim waitOnReturn As Boolean: waitOnReturn = True
'    Dim windowStyle As Integer: windowStyle = 1
'
'    wsh.Run strCmd, windowStyle, waitOnReturn
    
    If gDebugMode Then Debug.Print strCmd

On Error GoTo Skip:
    'Read standard output text
    Dim strInput        As String
    Dim lFileNo         As Long
    lFileNo = FreeFile
    Open strTempFilePath For Input As #lFileNo
    Do Until EOF(lFileNo)
        Line Input #lFileNo, strInput
        strErr = strErr & vbCrLf & Trim(strInput)
    Loop
    Close #lFileNo

    'Delete temporary file
    If Trim$(strErr) = "" Then Kill strTempFilePath
Skip:
    Set objFileSystem = Nothing
    Exit Sub
Err_qpdfSetEncryption:
    strErr = "(qpdfSetEncryption) Runtime Error :" & _
        Err.Number & vbCrLf & Err.Description
End Sub

'Wait for shell function to finish

Sub RunCommandLine(ByRef strCmd As String, _
                   ByRef strErr As String)
On Error GoTo Err_RunCommandLine:

    Dim hProcess        As Long
    Dim lpdwExitCode    As Long
    Dim dwProcessID     As Long
    Dim retVal          As Long
    Dim lCnt            As Long
    Const CON_SLEEP = 20
    Const CON_LOOP_CNT = 250
    lCnt = 0
    dwProcessID = Shell(strCmd, vbHide)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, _
        True, dwProcessID)
    Do
        Sleep CON_SLEEP
        DoEvents
        retVal = GetExitCodeProcess(hProcess, lpdwExitCode)
        lCnt = lCnt + 1
        If lCnt > CON_LOOP_CNT Then
            If gDebugMode Then Debug.Print vbCrLf & strCmd
            strErr = "Shell Error : Time Over " & _
                CON_SLEEP * CON_LOOP_CNT & "ms"
            Exit Sub
        End If
        'Loop until the application executed with shell function terminates
    Loop While lpdwExitCode <> 0
    Exit Sub
Err_RunCommandLine:
    strErr = "(RunCommandLine) Runtime Error :" & _
        Err.Number & vbCrLf & Err.Description
End Sub
