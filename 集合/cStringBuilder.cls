VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'如CopyMemory的声明，它是定义在Kernel32.dll中的RtlMoveMemory这个API，32位C函数库中的memcpy就是这个API的包装，
'如MSDN文档中所言，它的功能是将从Source指针所指处开始的长度为Length的内存拷贝到Destination所指的内存处。
'它不会管我们的程序有没有读写该内存所应有的权限，一但它想读写被系统所保护的内存时，我们就会得到著名的Access Violation Fault(内存越权访问错误)，
'甚至会引起更著名的general protection (GP) fault（通用保护错误） 。
'所以，在进行本系列文章里的实验时，请注意随时保存你的程序文件，在VB集成环境中将"工具"->"选项"中的"环境"选项卡里的"启动程序时"设为"保存改变"，
'并记住在"立即"窗口中执行危险代码之前一定要保存我们的工作成果。
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
#End If
'大型字符串的构建, 比较麻烦的是每次拼接需要复制数据造成的时间浪费, 如 a=a & b, a=a & c, 每次产生新的数据需要在新的内存为数据腾出空间, 再复制数据到新的地址
'You repeatedly copy stuff to the end of the string with the concatenation operator &. The bad news is, VB is not optimized for this. As you grow the string,
'VB repeatedly copies Big$ over and over again. This really degrades performance when repeated. VB constantly allocates new space and performs a copy.
'The problem gets worse as the string exceeds 64K in size. Small strings are stored in a 64K string cache. As the string becomes larger than the cache, performance drops considerably.
Private m_cString As String
Private m_iChunkSize As Long
Private m_iPosition As Long
Private m_iLen As Long

Friend Property Get sLength() As Long
   Length = m_iPosition \ 2
End Property

Friend Property Get sCapacity() As Long
   Capacity = m_iLen \ 2
End Property

Friend Property Get sChunkSize() As Long
   ChunkSize = m_iChunkSize \ 2   ' Return the unicode character chunk size:
End Property

Friend Property Let sChunkSize(ByVal iChunkSize As Long)
   m_iChunkSize = iChunkSize * 2  ' we are considering bytes: Set the chunksize.  We multiply by 2 because internally
End Property

Friend Property Get ToString() As String
   If m_iPosition > 0 Then ToString = Left$(m_cString, m_iPosition \ 2)
End Property

Friend Property Let TheString(ByRef sText As String)
   Dim lLen As Long
   
   lLen = LenB(sText)  ' Setting the string:
   If lLen = 0 Then
      m_cString = ""  'Clear
      m_iPosition = 0
      m_iLen = 0
   Else
      If m_iLen < lLen Then  ' Need to expand string to accommodate:
         Do
            m_cString = m_cString & Space$(m_iChunkSize \ 2)
            m_iLen = m_iLen + m_iChunkSize
         Loop While m_iLen < lLen
      End If
      CopyMemory ByVal StrPtr(m_cString), ByVal StrPtr(sText), lLen
      m_iPosition = lLen
   End If
End Property

Friend Sub Clear()
   m_cString = vbNullString
   m_iPosition = 0
   m_iLen = 0
End Sub

Friend Sub Append_NewLine(ByRef sText As String)
   Append sText
   Append vbCrLf
End Sub

Friend Sub Append(ByRef sText As String) ' 注意byref 和 byval 的区别, 优先选取byref
   Dim lLen As Long
   Dim lLenPlusPos As Long
 
   '-----------------------添加字符串到字符串
   lLen = LenB(sText)
   lLenPlusPos = lLen + m_iPosition
   If lLenPlusPos > m_iLen Then
      Dim lTemp As Long
      lTemp = m_iLen
      Do While lTemp < lLenPlusPos
         lTemp = lTemp + m_iChunkSize
      Loop
      m_cString = m_cString & Space$((lTemp - m_iLen) \ 2)
      m_iLen = lTemp
   End If
   CopyMemory ByVal UnsignedAdd(StrPtr(m_cString), m_iPosition), ByVal StrPtr(sText), lLen
   m_iPosition = m_iPosition + lLen
End Sub

Friend Sub Append_ByVal(ByVal sText As String)
   Append sText
End Sub

Friend Sub Insert(ByVal iIndex As Long, ByRef sText As String)
   Dim lLen As Long
   Dim lPos As Long
   Dim lSize As Long
   
   If (iIndex * 2 > m_iPosition) Then '越界错误
      Err.Raise 9
   Else
      lLen = LenB(sText)
      If (m_iPosition + lLen) > m_iLen Then
         m_cString = m_cString & Space$(m_iChunkSize \ 2)
         m_iLen = m_iLen + m_iChunkSize
      End If
      lPos = UnsignedAdd(StrPtr(m_cString), iIndex * 2)           ' Move existing characters from current position
      lSize = m_iPosition - iIndex * 2
      CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize ' moving from iIndex to iIndex + lLen
      CopyMemory ByVal lPos, ByVal StrPtr(sText), lLen            ' Insert new characters:
      m_iPosition = m_iPosition + lLen
   End If
End Sub

Friend Sub Insert_ByVal(ByVal iIndex As Long, ByVal sText As String)
   Insert iIndex, sText
End Sub

Friend Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
   Dim lSrc As Long
   Dim lDst As Long
   Dim lSize As Long

   If (iIndex * 2 > m_iPosition) Then             ' is iIndex within bounds?
      Err.Raise 9
   Else
      If ((iIndex + lLen) * 2 > m_iPosition) Then ' is there sufficient length?
         Err.Raise 9
      Else
         ' Need to copy characters from iIndex*2 to m_iPosition back by lLen chars
         lSrc = UnsignedAdd(StrPtr(m_cString), (iIndex + lLen) * 2)
         lDst = UnsignedAdd(StrPtr(m_cString), iIndex * 2)
         lSize = (m_iPosition - (iIndex + lLen) * 2)
         CopyMemory ByVal lDst, ByVal lSrc, lSize
         m_iPosition = m_iPosition - lLen * 2
      End If
   End If
End Sub

Friend Function Find(ByVal sToFind As String, Optional ByVal lStartIndex As Long = 1, Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
   
   '这里需要注意尽可能减少使用vbTextCompare作为对比参数, vbTextCompare存在一系列潜在的问题
   Dim lInstr As Long
   If (lStartIndex > 0) Then
      lInstr = InStr(lStartIndex, m_cString, sToFind, compare)
   Else
      lInstr = InStr(m_cString, sToFind, compare)
   End If
   If (lInstr < m_iPosition \ 2) Then
      Find = lInstr
   End If
End Function

Friend Sub HeapMinimize()
   Dim iLen As Long
   ' Reduce the string size so only the minimal chunks
   ' are allocated:
   If (m_iLen - m_iPosition) > m_iChunkSize Then
      iLen = m_iLen
      Do While (iLen - m_iPosition) > m_iChunkSize
         iLen = iLen - m_iChunkSize
      Loop
      m_cString = Left$(m_cString, iLen \ 2)
      m_iLen = iLen
   End If
End Sub

'------------------------------------LongPtr过度数据类型, 32/64, 非具体数据类型
#If Win64 Then
Private Function UnsignedAdd(Start As LongPtr, Incr As Long) As LongPtr
#Else
Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
#End If
   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
End Function

Private Sub Class_Initialize()
   m_iChunkSize = 16384        'The default allocation: 8192 characters.
End Sub

Private Sub Class_Terminate()
    m_cString = vbNullString
    m_iChunkSize = 0
    m_iPosition = 0
    m_iLen = 0
End Sub
