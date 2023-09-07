VERSION 1.0 CLASS
BEGIN
MultiUse = -1        'True
End
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = FALSE
Attribute VB_Creatable = FALSE
Attribute VB_PredeclaredId = FALSE
Attribute VB_Exposed = FALSE
Option Explicit
' @description: 计算文件的crc32值
Private CRCTable(0 To 255) As Long

Public Function CalcCRC32(ByVal FilePath As String, ByVal cmfrom As Byte) As Long
    Dim ByteArray() As Byte
    Dim Limit       As Long
    Dim CRC         As Long
    Dim Temp1       As Long        'long类型最大支持2G
    Dim Temp2       As Long
    Dim i           As Long
    Dim intFF       As Integer
    
    '原生的vba的操作的方法是不支持Unicode的
    '注意不要使用filelen,不仅大小限制,而且也不支持非ansi
    Limit = fso.GetFile(FilePath).Size
    CRC = -1
    Limit = Limit - 1
    '优先使用这种方式计算,open不支持非ansi字符,总体计算还是非常慢,计算300M+的文件需要10+s,而install计算md5,只需要0.0几秒即可
    If cmfrom = 0 Then
        intFF = FreeFile
        Open FilePath For Binary Access Read As #intFF
        ReDim ByteArray(Limit)
        Get #intFF, , ByteArray
        Close #intFF
        
        For i = 0 To Limit
            If CRC < 0 Then
                Temp1 = CRC And &H7FFFFFFF
                Temp1 = Temp1 \ 256
                Temp1 = (Temp1 Or &H800000) And &HFFFFFF
            Else
                Temp1 = (CRC \ 256) And &HFFFFFF
            End If
            Temp2 = ByteArray(i)        ' get the byte
            Temp2 = CRCTable((CRC Xor Temp2) And &HFF)
            CRC = Temp1 Xor Temp2
        Next i
    ElseIf cmfrom > 0 Then
        '采用这种方法计算速度太慢
        With CreateObject("ADODB.Stream")
            .Mode = 3
            .type = 1
            .Open
            .LoadFromFile (FilePath)
            For i = 0 To Limit
                If CRC < 0 Then
                    Temp1 = CRC And &H7FFFFFFF
                    Temp1 = Temp1 \ 256
                    Temp1 = (Temp1 Or &H800000) And &HFFFFFF
                Else
                    Temp1 = (CRC \ 256) And &HFFFFFF
                End If
                ' get the byte
                Temp2 = AscB(.Read(1))
                '如果a、b两个值不相同，则异或结果为1。如果a、b两个值相同，异或结果为0
                Temp2 = CRCTable((CRC Xor Temp2) And &HFF)
                'https://baike.baidu.com/item/%E5%BC%82%E6%88%96/10993677?fr=aladdin
                CRC = Temp1 Xor Temp2
            Next
            .Close
        End With
    End If
    CRC = CRC Xor &HFFFFFFFF
    CalcCRC32 = CRC
End Function

Private Sub Class_Initialize()
    Dim i           As Integer
    Dim j           As Integer
    Dim Limit       As Long
    Dim CRC         As Long
    Dim Temp1       As Long
    
    Limit = &HEDB88320
    For i = 0 To 255
        CRC = i
        For j = 8 To 1 Step -1
            If CRC < 0 Then
                Temp1 = CRC And &H7FFFFFFF
                Temp1 = Temp1 \ 2
                Temp1 = Temp1 Or &H40000000
            Else
                Temp1 = CRC \ 2
            End If
            If CRC And 1 Then
                CRC = Temp1 Xor Limit
            Else
                CRC = Temp1
            End If
        Next j
        CRCTable(i) = CRC
    Next i
End Sub


'调用方法-计算文件的crc32-计算的速度太慢
'Private CRC32      As New cCRC32
'debug.ptint Hex$(CRC32.CalcCRC32(strx, 0))