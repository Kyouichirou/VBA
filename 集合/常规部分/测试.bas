Attribute VB_Name = "测试"

Sub GetMD5()
    Dim sFilePath       As String           '// ファイルパス
    Dim wsh             As WshShell         '// WshShellクラス
    Dim ex              As WshExec          '// WshExecクラス
    Dim ar()            As String           '// ファイルパス配列
    Dim i               As Integer          '// ループカウンタ
    Dim sCmd            As String           '// コマンド文字列

    '// WshShellインスタンス作成
    Set wsh = New WshShell

'    ReDim ar(4)
'    ar(0) = "C:\test\テキスト ドキュメント.txt"
'    ar(1) = "C:\test\テキスト ドキュメント.txt"
'    ar(2) = "C:\test\テキスト ドキュメント.txt"
'    ar(3) = "C:\test\テキスト ドキュメント.txt"
'    ar(4) = "C:\test\テキスト ドキュメント.txt"

    '// コマンド文字列の設定（ドライブとディレクトリの指定は必須）
    sCmd = "C: & "
    sCmd = sCmd & "cd C:\test\ "

    '// 配列ループ
'    For i = 0 To UBound(ar)
        '// certutilコマンドを連結
        sCmd = sCmd & " & certutil -hashfile """ & Range("e6").Value & """ MD5|findstr -v "":"">>MD5.txt "
'    Next

    '// コマンドを実行
    Set ex = wsh.Exec("cmd.exe /c " & sCmd)
End Sub
Private Sub TestMD5()
    Debug.Print FileToMD5Hex("C:\test.txt")
    Debug.Print FileToSHA1Hex("C:\test.txt")
End Sub

Public Function FileToMD5Hex(sFilename As String) As String
    Dim enc
    Dim Bytes
    Dim outstr As String
    Dim pos As Integer
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFilename)
    Bytes = enc.ComputeHash_2((Bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(Bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(Bytes, pos, 1))), 2))
    Next
    FileToMD5Hex = outstr
    Set enc = Nothing
End Function

Public Function FileToSHA1Hex(sFilename As String) As String
    Dim enc
    Dim Bytes
    Dim outstr As String
    Dim pos As Integer
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFilename)
    Bytes = enc.ComputeHash_2((Bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(Bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(Bytes, pos, 1))), 2))
    Next
    FileToSHA1Hex = outstr 'Returns a 40 byte/character hex string
    Set enc = Nothing
End Function

Private Function GetFileBytes(ByVal Path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    lngFileNum = FreeFile
    If LenB(Dir(Path)) Then ''// Does file exist?
        Open Path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function

Option Explicit

Private Sub TestFileHashes()
    'run this to obtain file hashes in a choice of algorithms
    'select any one algorithm call below
    'Limited to unrestricted files less than 200MB and not zero
    'Set a reference to mscorlib 4.0 64-bit, and Scripting Runtime

    Dim sFPath As String, b64 As Boolean, bOK As Boolean
    Dim sH As String, sSecret As String, nSize As Long, reply

    'USER SETTINGS
    '======================================================
    '======================================================
    'set output format here
    b64 = True     'true for output base-64, false for hex
    '======================================================
    'set chosen file here
    'either set path to target file in hard-typed line
    'or choose a file using the file dialog procedure
    'sFPath = "C:\Users\Your Folder\Documents\test.txt" 'eg.
    sFPath = SelectFile2("SELECT A FILE TO HASH...") 'uses file dialog

    'check the file
    If sFPath = "" Then 'exit sub for no file selection
        MsgBox "No selection made - closing"
        Exit Sub
    End If
    bOK = GetFileSize(sFPath, nSize)
'    If nSize = 0 Or nSize > 200000000 Then 'exit sub for zero size
'        MsgBox "File has zero contents or greater than 200MB - closing"
'        Exit Sub
'    End If
    '======================================================
    'set secret key here if using HMAC class of algorithms
    sSecret = "Set secret key for FileToSHA512Salt selection"
    '======================================================
    'choose algorithm
    'enable any one line to obtain that hash result
    sH = FileToMD5(sFPath, b64)
    'sH = FileToSHA1(sFPath, b64)
    'sH = FileToSHA256(sFPath, b64)
    'sH = FileToSHA384(sFPath, b64)
    'sH = FileToSHA512Salt(sFPath, sSecret, b64)
'    sH = FileToSHA512(sFPath, b64)
    '======================================================
    '======================================================

    'Results Output - open the immediate window as required
    Debug.Print sFPath & vbNewLine & sH & vbNewLine & Len(sH) & " characters in length"
    MsgBox sFPath & vbNewLine & sH & vbNewLine & Len(sH) & " characters in length"
    'reply = InputBox("The selected text can be copied with Ctrl-C", "Output is in the box...", sH)

    'decomment this block to place the hash in first cell of sheet1
'    With ThisWorkbook.Worksheets("Sheet1").Cells(1, 1)
'        .Font.Name = "Consolas"
'        .Select: Selection.NumberFormat = "@" 'make cell text
'        .Value = sH
'    End With
End Sub

Public Function FileToMD5(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an MD5 hash
    'Set a reference to mscorlib 4.0 64-bit

    Dim enc, Bytes, outstr As String, pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFullPath)
    Bytes = enc.ComputeHash_2((Bytes))

    If bB64 = True Then
       FileToMD5 = ConvToBase64String(Bytes)
    Else
       FileToMD5 = ConvToHexString(Bytes)
    End If

    Set enc = Nothing

End Function

Public Function FileToSHA1(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA1 hash
    'Set a reference to mscorlib 4.0 64-bit

    Dim enc, Bytes, outstr As String, pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFullPath) 'returned as a byte array
    Bytes = enc.ComputeHash_2((Bytes))

    If bB64 = True Then
       FileToSHA1 = ConvToBase64String(Bytes)
    Else
       FileToSHA1 = ConvToHexString(Bytes)
    End If

    Set enc = Nothing

End Function
Sub xkk()
Debug.Print FileToSHA512Salt("abc", "123")
End Sub
Function FileToSHA512Salt(ByVal sPath As String, ByVal sSecretKey As String, Optional ByVal bB64 As Boolean = False) As String
    Dim asc As Object, enc As Object
    Dim SecretKey() As Byte
    Dim Bytes() As Byte
    Dim i As Integer, k As Integer
    Dim Result As String

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")
    Bytes = asc.GetBytes_4(sPath)
    SecretKey = asc.GetBytes_4(sSecretKey)
    enc.key = SecretKey
    Bytes = enc.ComputeHash_2((Bytes))
    If bB64 = True Then
       FileToSHA512Salt = ConvToBase64String(Bytes)
    Else
       FileToSHA512Salt = ConvToHexString(Bytes)
    End If
    Set asc = Nothing
    Set enc = Nothing
End Function

Public Function FileToSHA256(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-256 hash
    'Set a reference to mscorlib 4.0 64-bit

    Dim enc, Bytes, outstr As String, pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFullPath) 'returned as a byte array
    Bytes = enc.ComputeHash_2((Bytes))

    If bB64 = True Then
       FileToSHA256 = ConvToBase64String(Bytes)
    Else
       FileToSHA256 = ConvToHexString(Bytes)
    End If

    Set enc = Nothing

End Function

Public Function FileToSHA384(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-384 hash
    'Set a reference to mscorlib 4.0 64-bit

    Dim enc, Bytes, outstr As String, pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.SHA384Managed")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFullPath) 'returned as a byte array
    Bytes = enc.ComputeHash_2((Bytes))

    If bB64 = True Then
       FileToSHA384 = ConvToBase64String(Bytes)
    Else
       FileToSHA384 = ConvToHexString(Bytes)
    End If

    Set enc = Nothing

End Function

Public Function FileToSHA512(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-512 hash
    'Set a reference to mscorlib 4.0 64-bit

    Dim enc, Bytes, outstr As String, pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Convert the string to a byte array and hash it
    Bytes = GetFileBytes(sFullPath) 'returned as a byte array
    Bytes = enc.ComputeHash_2((Bytes))

    If bB64 = True Then
       FileToSHA512 = ConvToBase64String(Bytes)
    Else
       FileToSHA512 = ConvToHexString(Bytes)
    End If

    Set enc = Nothing

End Function

'Private Function GetFileBytes(ByVal sPath As String) As Byte()
'    'makes byte array from file
'    'Set a reference to mscorlib 4.0 64-bit
'
'    Dim lngFileNum As Long, bytRtnVal() As Byte, bTest
'
'    lngFileNum = FreeFile
'
'    If LenB(Dir(sPath)) Then ''// Does file exist?
'
'        Open sPath For Binary Access Read As lngFileNum
'
'        'a zero length file content will give error 9 here
'
'        ReDim bytRtnVal(0 To LOF(lngFileNum) - 1&) As Byte
'        Get lngFileNum, , bytRtnVal
'        Close lngFileNum
'    Else
'        Err.Raise 53 'File not found
'    End If
'
'    GetFileBytes = bytRtnVal
'
'    Erase bytRtnVal
'
'End Function

Function ConvToBase64String(vIn As Variant) As Variant
    'used to produce a base-64 output
    'Set a reference to mscorlib 4.0 64-bit

    Dim oD As Object

    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    Set oD = Nothing
End Function

Function ConvToHexString(vIn As Variant) As Variant
    Dim oD As Object
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    Set oD = Nothing
End Function

Function GetFileSize(sFilePath As String, nSize As Long) As Boolean
    'use this to test for a zero file size
    'takes full path as string in sFilePath
    'returns file size in bytes in nSize
    'Make a reference to Scripting Runtime

    Dim fs As FileSystemObject, F As File

    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.fileexists(sFilePath) Then
        Set F = fs.GetFile(sFilePath)
        nSize = F.Size
        GetFileSize = True
        Exit Function
    End If

End Function

Function SelectFile2(Optional sTitle As String = "") As String
    'opens a file-select dialog and on selection
    'returns its full path string in the function name
    'If Cancel or OK without selection, returns empty string

    Dim fd As FileDialog, sPathOnOpen As String, sOut As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'do not include backslash here
    sPathOnOpen = Application.DefaultFilePath

    'set the file-types list on the dialog and other properties
    With fd
        .Filters.Clear
        'the first filter line below sets the default on open (here all files are listed)
        .Filters.Add "All Files", "*.*"
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls;*.xltx;*.xltm;*.xlt;*.xml;*.ods"
        .Filters.Add "Word documents", "*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.odt"

        .AllowMultiSelect = False
        .InitialFileName = sPathOnOpen
        .Title = sTitle
        .InitialView = msoFileDialogViewList 'msoFileDialogViewSmallIcons
        .Show

        If .SelectedItems.Count = 0 Then
            'MsgBox "Canceled without selection"
            Exit Function
        Else
            sOut = .SelectedItems(1)
            'MsgBox sOut
        End If
    End With

    SelectFile2 = sOut

End Function





