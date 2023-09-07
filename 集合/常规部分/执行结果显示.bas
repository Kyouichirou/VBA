Attribute VB_Name = "执行结果显示"
Option Explicit

'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messagebox
'wLanguageId
'Specifies the language in which to display the text contained in the predefined push buttons. This value must be in the form returned by theMAKELANGID macro.
'For a list of the language identifiers supported by Win32, seeLanguage Identifiers. Note that each localized release of Windows and
'Windows NT typically contains resources only for a limited set of languages. Thus, for example, the U.S. version offers LANG_ENGLISH,
'the French version offers LANG_FRENCH, the German version offers LANG_GERMAN, and the Japanese version offers LANG_JAPANESE.
'Each version offers LANG_NEUTRAL. This limits the set of values that can be used with the wLanguageId parameter.
'Before specifying a language identifier, you should enumerate the locales that are installed on a system. '这里是关键, 安装在系统当中的语言包
Private Const LANG_NEUTRAL = &H0
Private Const LANG_AFRIKAANS = &H36
Private Const LANG_ALBANIAN = &H1C
Private Const LANG_ARABIC = &H1
Private Const LANG_BASQUE = &H2D
Private Const LANG_BELARUSIAN = &H23
Private Const LANG_BULGARIAN = &H2
Private Const LANG_CATALAN = &H3
Private Const LANG_CHINESE = &H4
Private Const LANG_CROATIAN = &H1A
Private Const LANG_CZECH = &H5
Private Const LANG_DANISH = &H6
Private Const LANG_DUTCH = &H13
Private Const LANG_ENGLISH = &H9
Private Const LANG_ESTONIAN = &H25
Private Const LANG_FAEROESE = &H38
Private Const LANG_FARSI = &H29
Private Const LANG_FINNISH = &HB
Private Const LANG_FRENCH = &HC
Private Const LANG_GERMAN = &H7
Private Const LANG_GREEK = &H8
Private Const LANG_HEBREW = &HD
Private Const LANG_HINDI = &H39
Private Const LANG_HUNGARIAN = &HE
Private Const LANG_ICELANDIC = &HF
Private Const LANG_INDONESIAN = &H21
Private Const LANG_ITALIAN = &H10
Private Const LANG_JAPANESE = &H11
Private Const LANG_KOREAN = &H12
Private Const LANG_LATVIAN = &H26
Private Const LANG_LITHUANIAN = &H27
Private Const LANG_MACEDONIAN = &H2F
Private Const LANG_MALAY = &H3E
Private Const LANG_NORWEGIAN = &H14
Private Const LANG_POLISH = &H15
Private Const LANG_PORTUGUESE = &H16
Private Const LANG_ROMANIAN = &H18
Private Const LANG_RUSSIAN = &H19
Private Const LANG_SERBIAN = &H1A
Private Const LANG_SLOVAK = &H1B
Private Const LANG_SLOVENIAN = &H24
Private Const LANG_SPANISH = &HA
Private Const LANG_SWAHILI = &H41
Private Const LANG_SWEDISH = &H1D
Private Const LANG_THAI = &H1E
Private Const LANG_TURKISH = &H1F
Private Const LANG_UKRANIAN = &H22
Private Const LANG_VIETNAMESE = &H2A
Private Const SUBLANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Const SUBLANG_SYS_DEFAULT = &H2
Private Const SUBLANG_ARABIC = &H1
Private Const SUBLANG_ARABIC_IRAQ = &H2
Private Const SUBLANG_ARABIC_EGYPT = &H3
Private Const SUBLANG_ARABIC_LIBYA = &H4
Private Const SUBLANG_ARABIC_ALGERIA = &H5
Private Const SUBLANG_ARABIC_MOROCCO = &H6
Private Const SUBLANG_ARABIC_TUNISIA = &H7
Private Const SUBLANG_ARABIC_OMAN = &H8
Private Const SUBLANG_ARABIC_YEMEN = &H9
Private Const SUBLANG_ARABIC_SYRIA = &HA
Private Const SUBLANG_ARABIC_JORDAN = &HB
Private Const SUBLANG_ARABIC_LEBANON = &HC
Private Const SUBLANG_ARABIC_KUWAIT = &HD
Private Const SUBLANG_ARABIC_UAE = &HE
Private Const SUBLANG_ARABIC_BAHRAIN = &HF
Private Const SUBLANG_ARABIC_QATAR = &H10
Private Const SUBLANG_CHINESE_TRADITIONAL = &H1
Private Const SUBLANG_CHINESE_SIMPLIFIED = &H2
Private Const SUBLANG_CHINESE_HONGKONG = &H3
Private Const SUBLANG_CHINESE_SINGAPORE = &H4
Private Const SUBLANG_DUTCH = &H1
Private Const SUBLANG_DUTCH_BELGIAN = &H2
Private Const SUBLANG_ENGLISH_US = &H1
Private Const SUBLANG_ENGLISH_UK = &H2
Private Const SUBLANG_ENGLISH_AUS = &H3
Private Const SUBLANG_ENGLISH_CAN = &H4
Private Const SUBLANG_ENGLISH_NZ = &H5
Private Const SUBLANG_ENGLISH_EIRE = &H6
Private Const SUBLANG_ENGLISH_SAFRICA = &H7
Private Const SUBLANG_ENGLISH_JAMAICA = &H8
Private Const SUBLANG_ENGLISH_CARRIBEAN = &H9
Private Const SUBLANG_FRENCH = &H1
Private Const SUBLANG_FRENCH_BELGIAN = &H2
Private Const SUBLANG_FRENCH_CANADIAN = &H3
Private Const SUBLANG_FRENCH_SWISS = &H4
Private Const SUBLANG_FRENCH_LUXEMBOURG = &H5
Private Const SUBLANG_GERMAN = &H1
Private Const SUBLANG_GERMAN_SWISS = &H2
Private Const SUBLANG_GERMAN_AUSTRIAN = &H3
Private Const SUBLANG_GERMAN_LUXEMBOURG = &H4
Private Const SUBLANG_GERMAN_LIECHTENSTEIN = &H5
Private Const SUBLANG_ITALIAN = &H1
Private Const SUBLANG_ITALIAN_SWISS = &H2
Private Const SUBLANG_KOREAN = &H1
Private Const SUBLANG_KOREAN_JOHAB = &H2
Private Const SUBLANG_NORWEGIAN_BOKMAL = &H1
Private Const SUBLANG_NORWEGIAN_NYNORSK = &H2
Private Const SUBLANG_PORTUGUESE = &H2
Private Const SUBLANG_PORTUGUESE_BRAZILIAN = &H1
Private Const SUBLANG_SPANISH = &H1
Private Const SUBLANG_SPANISH_MEXICAN = &H2
Private Const SUBLANG_SPANISH_MODERN = &H3
Private Const SUBLANG_SPANISH_GUATEMALA = &H4
Private Const SUBLANG_SPANISH_COSTARICA = &H5
Private Const SUBLANG_SPANISH_PANAMA = &H6
Private Const SUBLANG_SPANISH_DOMINICAN = &H7
Private Const SUBLANG_SPANISH_VENEZUELA = &H8
Private Const SUBLANG_SPANISH_COLOMBIA = &H9
Private Const SUBLANG_SPANISH_PERU = &HA
Private Const SUBLANG_SPANISH_ARGENTINA = &HB
Private Const SUBLANG_SPANISH_ECUADOR = &HC
Private Const SUBLANG_SPANISH_CHILE = &HD
Private Const SUBLANG_SPANISH_URUGUAY = &HE
Private Const SUBLANG_SPANISH_PARAGUAY = &HF
Private Const SUBLANGSPANISHBOLIVIA = &H10
Private Declare Function aMsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, _
ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long '注意区分Unicode字符和ANSI字符的区别
Private Declare Function uMsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutW" (ByVal hwnd As Long, _
ByVal lpText As Long, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long

Private Declare Function aMessageBoxEX Lib "user32.dll" Alias "MessageBoxExA" (ByVal hwnd&, ByVal cText As String, ByVal sTtile As String, ByVal sPattern As Long, ByRef iLang As Long) As Long
'iLang, 这里可以设置显示的语言, 但是需要注意的是系统必须安装对应的语言包, 英文例外

Declare Function aMessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hwnd&, ByVal cText As String, ByVal sTtile As String, ByVal sPattern As Long) As Long '这个为系统自动调用的api
Declare Function uMessageBox Lib "user32.dll" Alias "MessageBoxW" (ByVal hwnd&, ByVal cText As Long, ByVal sTtile As String, ByVal sPattern As Long) As Long
'--------------------------------------------------------------------------------------------------------------------------------'定时可自动关闭弹窗/ 弹窗
Public Timeset As Byte '用于控制时间sub的执行

Sub MsgShow(ByVal strText As String, ByVal signalx As String, ByVal timex As Integer, ByVal isUnicode As Boolean) '定时弹窗
    '过程,"弹出对话","对话框标题",图标类型,默认参数,N秒后自动关闭
    If isUnicode = True Then
        uMsgBoxTimeOut 0, StrPtr(strText), signalx, 64, 0, timex 'unicode和ansi分离开来
    Else
        aMsgBoxTimeOut 0, strText, signalx, 64, 0, timex
    End If
End Sub

Function Warning(ByVal wncode As Integer, ByVal cmfrom As Byte) '操作显示执行结果或过程
        Dim strx As String
        
        If wncode = 1 Then                '整理编码后换成select case '增减显示颜色,红色表示警告,绿色表示执行成功
            strx = "操作成功!"            '执行结果代码-对应的提示-选择结果输出-显示-清除结果
        ElseIf wncode = 2 Then
            strx = "!操作失败"
        ElseIf wncode = 3 Then
            strx = "!文件已被删除或移除出书库"
        ElseIf wncode = 4 Then
            strx = "!文件已添加"
        ElseIf wncode = 5 Then
            strx = "!设置没有修改"
        ElseIf wncode = 6 Then
            strx = "!信息为空"
        ElseIf wncode = 7 Then
            strx = "!程序尚未设置"
        ElseIf wncode = 8 Then
            strx = "!请稍后,处理中"
        End If
    End With
    ShowResult strx, cmCode
End Function

Function TimeClock() '自动清除提示信息
    If Timeset = 0 Then
       Timeset = 1
    ElseIf Timeset = 1 Then '控制执行
        LabelClear
        Exit Function
    ElseIf Timeset = 2 Then
        Exit Sub
    End If
    Application.OnTime Now + TimeValue("00:00:03"), "TimeClock" '定时3s
End Function

Sub LabelClear()
    UserForm3.Label57.Caption = ""
    ThisWorkbook.Sheets("书库").Label1.Caption = ""
End Sub

Function ShowResult(ByVal Result As String, ByVal cmfrom As Byte) '显示表/控制板操作的结果
    If cmfrom = 1 Then
        ThisWorkbook.Sheets("书库").Label1.Caption = Result
    ElseIf cmfrom = 2 Then
        UserForm3.Label57.Caption = Result
    End If
    Call TimeClock
End Function

'#define    MAKELANGID(p, s)       ((((WORD  )(s)) << 10) | (WORD  )(p))
Function MAKELANGIDs(ByVal usPrimaryLanguage As Integer, ByVal usSubLanguage As Integer) As Long '生成对应的语言id
    MAKELANGIDs = (usSubLanguage * 1024) Or usPrimaryLanguage '如英语, 9, 1, = 1033
End Function
