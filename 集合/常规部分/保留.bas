Attribute VB_Name = "保留"
'Option Explicit
'
''Function Filein(ByVal fl As File)
''Dim filemd5 As String
''Dim rngmd As Range
''Dim FilePath As String
''Dim t As Integer       '用于标记hash是用那种方式生成的
''Dim J As Integer        '用于标记非电子书文件的路径是否存在特殊字符
''Dim p As Integer           '是否存在特殊字符
''Dim flr As Variant        '用于记录数组查找返回的值
''Dim faddress As String '用于记录同名文件检索到的第一个文件的位置
''Dim ckn As Integer '用于标记文件已经经过检查 'ckn=0表示文件未经过文件名检查,=1表示经过检查
''Dim exet As Integer
'
'If Len(fso.GetExtensionName(fl.Path)) = 0 Or fl.Size = 0 Or fl.Attributes = 34 Then GoTo 100 '文件扩展名为空/隐藏文件的跳过 ,34表示hidden属性,注意不要直接使用hidden来表示属性,无法识别
'exet = fso.GetExtensionName(fl.Path) '限制添加进目录的文件类型 ' ucase为转换大小函数
'If Not exet Like "EPUB" And Not exet Like "PDF" And Not exet Like "MOBI" And Not exet Like "DO*" And Not exet Like "XL*" And Not exet Like "PP*" And Not exet Like "AC*" And Not exet Like "TX*" Then GoTo 100 '文件类型筛选
'FilePath = fl.Path
'J = 0
't = 0
'p = Errcode(fl.ShortName, fl.ParentFolder, 0)
'
'If p > 0 Then
'    If exet Like "EPUB" Or exet Like "PDF" Or exet Like "MOBI" Then '限定于pdf和mobi,epub使用md5 ,考虑到其他的文件都是易编辑的文件,文件的hash动态变化
'        If f = 1 Then
'        t = 1
'        GoTo 1010
'        End If
'1011
'        If ShellxExist = 1 Then '判断powershell是否存在, Environ("SystemRoot")系统环境变量(C:\windows,假设系统安装在C盘)
'            DoEvents
'            t = 1
'            filemd5 = UCase(Hashpowershell(FilePath)) 'ucase,转换成大写
'            If Len(filemd5) = 2 Then GoTo 1002 '没有获取到文件hash
'        Else
'            If fl.Size < 20971520 Then '由于采用adodb.stream方式计算hash的速度太慢,所以需要限制文件的大小, 不超过20M
'                DoEvents
'                t = 1
'                filemd5 = UCase(GetMD5Hash_File(FilePath, fl.Size))
'            Else
'1002
'                flc = flc + 1
'                t = 0
'                arrcode(flc) = "ERC"       '异常字符标记 '修改-无法返回有效值,就进行文件名比较
'                arrmd5(flc) = "UC"
'                If Tagfnansi = True And Tagfpansi = True Then 'd
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EDC"
'                ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = ""
'                    arrcm(flc) = "ENC" '出现的位置
'                ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EPC" '出现的位置
'                End If
'            End If
'        End If
'        Else
'        J = 1
'        If f = 1 Then GoTo 1010
'    End If
'
'ElseIf p = 0 Then                                                             '不存在乱码
'    If exet Like "EPUB" Or exet Like "PDF" Or exet Like "MOBI" Then
'        t = 2
'        If f = 1 Then GoTo 1010
'1012
'        DoEvents                                                    '由于计算hash的速度会很慢,所以需要用doevents来减缓代码执行的假死的状况
'        filemd5 = UCase(FileHashes(FilePath)) '此模块计算hash的效果最好,可以突破2G大小的限制
'    Else
'        J = 2
'    End If
'ElseIf p = -1 Then GoTo 100 '无法读取文件路径
'End If
'
''--------------------------------------------------------------------------------------判断文件路径是否存在非ansi编码字符/是否计算md5/计算md5的方式
'
'If elow > 5 Then '当已经存在有目录的时候
'    If J > 0 Then
'1010
'    With ThisWorkbook.Sheets("书库")
'        Set rngmd = .Range("c6:c" & .[c65536].End(xlUp).Row).Find(fl.Name) '检查是否同名文件已存在
'    End With
'    If rngmd Is Nothing Then '检查文件名是否相同
'        If f = 1 And ckn = 0 And J = 0 Then     '进行文件更新的时候电子书格式的文件先进行文件名判断,如果不存在同名文件就进行md5计算
'            ckn = 1
'            If t = 1 Then
'                GoTo 1011 '重新获取文件的详细信息
'            ElseIf t = 2 Then
'                GoTo 1012
'            End If
'        End If
'1003
'        If w > 0 Then
'            flr = Filter(arrbase, fl.Name) '先进行文件名判断(减少不必要的循环)
'            If UBound(flr) >= 0 Then
'                For h = 1 To flc        '先进行文件添加前的比较
'                    If fl.Name = arrbase(h) And fl.Size = arrsizeb(h) And fl.DateLastModified = arredit(h) Then GoTo 100
'                Next
'            End If
'        End If
'
'        w = w + 1
'        If J = 2 Then
'            flc = flc + 1
'            arrmd5(flc) = ""
'            arrcode(flc) = ""
'            arrcm(flc) = ""
'            arrfnansi(flc) = ""
'            arrfpansi(flc) = ""
'        ElseIf J = 1 Then
'            flc = flc + 1
'            arrmd5(flc) = ""
'            arrcode(flc) = "ERC"
'            If Tagfnansi = True And Tagfpansi = True Then 'd
'                arrfnansi(flc) = errcodenx
'                arrfpansi(flc) = errcodepx
'                arrcm(flc) = "EDC"
'            ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                arrfnansi(flc) = errcodenx
'                arrfpansi(flc) = ""
'                arrcm(flc) = "ENC" '出现的位置
'            ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                arrfnansi(flc) = ""
'                arrfpansi(flc) = errcodepx
'                arrcm(flc) = "EPC" '出现的位置
'            End If
'        End If
'    Else
'
'        If f = 1 Then GoTo 100 '文件已存在目录(更新的方式,不在继续检查比较)
'
'        If rngmd.Offset(0, 4).Value <> fl.Size Then '文件大小
'            faddress = rngmd.address '检索到的文件名相同的第一个位置
'            Do
'                With ThisWorkbook.Sheets("书库") '如果文件名相同,文件大小不一样,则执行下一个同名文件名的大小比较
'                    Set rngmd = .Range("c6:c" & .[c65536].End(xlUp).Row).FindNext(rngmd) '检查下一个同名文件
'                End With
'                If rngmd.Offset(0, 4) = fl.Size Then
'                    Exit Do
'                    GoTo 1006 '下一个文件的大小和这个文件一致
'                End If
'            Loop Until faddress = rngmd.address '循环完所有的目录,回到第一个位置
'            w = w + 1
'            If J = 2 Then
'                flc = flc + 1
'                arrmd5(flc) = ""
'                arrcode(flc) = ""
'                arrcm(flc) = ""
'                arrfnansi(flc) = ""
'                arrfpansi(flc) = ""
'            ElseIf J = 1 Then
'                flc = flc + 1
'                arrmd5(flc) = ""
'                arrcode(flc) = "ERC"
'                If Tagfnansi = True And Tagfpansi = True Then 'd
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EDC"
'                ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = ""
'                    arrcm(flc) = "ENC" '出现的位置
'                ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EPC" '出现的位置
'                End If
'            End If
'        Else
'1006
'            If rngmd.Offset(0, 5) <> fl.DateLastModified Then '文件修改的时间
'                w = w + 1
'                If J = 2 Then
'                    flc = flc + 1
'                    arrmd5(flc) = "DP"
'                    arrcode(flc) = ""
'                    arrcm(flc) = ""
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = ""
'                ElseIf J = 1 Then
'                    flc = flc + 1
'                    arrmd5(flc) = "DP"
'                    arrcode(flc) = "ERC"
'                    If Tagfnansi = True And Tagfpansi = True Then 'd
'                        arrfnansi(flc) = errcodenx
'                        arrfpansi(flc) = errcodepx
'                        arrcm(flc) = "EDC"
'                    ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                        arrfnansi(flc) = errcodenx
'                        arrfpansi(flc) = ""
'                        arrcm(flc) = "ENC" '出现的位置
'                    ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                        arrfnansi(flc) = ""
'                        arrfpansi(flc) = errcodepx
'                        arrcm(flc) = "EPC" '出现的位置
'                    End If
'                End If
'            Else
'                ls = ls + 1
'                GoTo 100
'            End If
'        End If
'        Set rngmd = Nothing
'        End If
'    End If
''---------------------------------------------------------不计算md5的方式来判断文件是否重叠
'    If t > 0 Then        '计算出md5后进行比较
'        With ThisWorkbook.Sheets("书库")
'            Set rngmd = .Range("k6:k" & .[k65536].End(xlUp).Row).Find(filemd5) '检查是否文件已存在(由于hash具有唯一性,不需要进行二次比较)
'        End With
'        If rngmd Is Nothing Then '文件不存在重叠/且路径不存在特殊字符
'1004
'            If n > 0 Then '有值的时候进行比较
'                flr = Filter(arrmd5, filemd5) 'filter函数,用于筛选数组
'                If UBound(flr) >= 0 Then GoTo 1005 '出现重叠的文件(没有重叠文件,值为-1)
'            End If
'            n = n + 1
'            If t = 1 Then
'                flc = flc + 1                             '记录有多少行的数据
'                arrmd5(flc) = filemd5
'                arrcode(flc) = "ERC"
'                If Tagfnansi = True And Tagfpansi = True Then 'd
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EDC"
'                ElseIf Tagfnansi = True And Tagfpansi = False Then 's
'                    arrfnansi(flc) = errcodenx
'                    arrfpansi(flc) = ""
'                    arrcm(flc) = "ENC" '出现的位置
'                ElseIf Tagfnansi = False And Tagfpansi = True Then 's
'                    arrfnansi(flc) = ""
'                    arrfpansi(flc) = errcodepx
'                    arrcm(flc) = "EPC" '出现的位置
'                End If
'            ElseIf t = 2 Then
'                flc = flc + 1                             '记录有多少行的数据
'                arrmd5(flc) = filemd5
'                arrcode(flc) = ""
'                arrcm(flc) = ""
'                arrfnansi(flc) = ""
'                arrfpansi(flc) = ""
'            End If
'            Set rngmd = Nothing
'        Else
'1005
'            If f = 1 Then GoTo 100 '文件已存在目录
'            Set rngmd = Nothing
'            dl = dl + 1
'            If t = 1 Then
'                fso.DeleteFile (FilePath) '和kill命令一样会直接删除掉文件
'            Else
'                DeleteFiles (FilePath) '如果文件相同则删除文件,移除到回收站(不支持异常字符的删除,kill命令同样存在一样的问题)
'                GoTo 100 '文件相同,执行下一个文件
'            End If
'            With ThisWorkbook.Sheets("目录")
'                .Cells(a, b + c + 2) = fd.DateCreated '删除操作后文件夹的时间发生变化
'            End With
'        End If
'    End If
''---------------------------------------------------------------------------------------------------------------计算md进行文件判断
'Else         '当表格还未写入数据的时候
'    If J > 0 Then
'        GoTo 1003
'    ElseIf t > 0 Then
'        GoTo 1004
'    End If
'End If
'
''-----------------------------------------------------------------------------------------------------------------------------------------判断文件是否存在在目录/添加的文件中是否存在重叠的文件/执行删除还是跳过
'
'With fl
'    arrfiles(flc) = .Path                  '数组存储文件的路径
'    arrbase(flc) = .Name                     '包含扩展名的文件名
'    arrextension(flc) = fso.GetExtensionName(.Path)                    '文件扩展名
'    arrparent(flc) = .ParentFolder '文件上一级目录
'    arrdate(flc) = .DateCreated '文件创建
'    arrsizeb(flc) = .Size         '文件大小
'    arredit(flc) = .DateLastModified '文件修改时间
'    If .Size < 1048576 Then
'        arrsize(flc) = Format(.Size / 1024, "0.00") & "KB"    '文件字节大于1048576显示"MB",否则显示"KB"
'    Else
'        arrsize(flc) = Format(.Size / 1048576, "0.00") & "MB"
'    End If
'End With
''---------------------------------------------------------------------------------------------------------------------写入数据
'100
''ckn = 0 '重置数据
''t = 0
''J = 0
''End Function
