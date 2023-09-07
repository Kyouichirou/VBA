Option Explicit

Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte
    
    '把BYTE類型變量左移1位的函數，參數Byt是待移位的字節，函數返回移位結果
    '
    '(Byt And &H7F)的作用是屏蔽最高位。 *2：左移一位
    
    ShLB_By1Bit = (Byt And &H7F) * 2
    
End Function

Function UEFCheckUTF8NoBom(ByRef bufAll() As Byte) as string
    
    ' 检查文本文件的编码类型, 注意Windows下的utf-8, bom的存在可能

    Dim i           As Long
    Dim cOctets     As Long        '可以容纳UTF-8编码字符的字节大小 4bytes
    Dim bAllAscii   As Boolean        '如果全部为ASCII，说明不是UTF-8
    Dim fmt         as string
    bAllAscii = TRUE
    cOctets = 0
    
    For i = 0 To UBound(bufAll)
        If (bufAll(i) And &H80) <> 0 Then
            'ASCII用7位储存，最高位为0，如果这里相与非0，就不是ASCII
            '对于单字节的符号，字节的第一位设为0，后面7位为这个符号的unicode码。
            '因此对于英语字母，UTF-8编码和ASCII码是相同的
            bAllAscii = FALSE
        End If
        
        '对于n字节的符号（n>1），第一个字节的前n位都设为1，第n+1位设为0，后面字节的前两位一律设为10
        'cOctets = 0 表示本字节是leading byte
        If cOctets = 0 Then
            If bufAll(i) >= &H80 Then
                '计数：是cOctets字节的符号
                Do While (bufAll(i) And &H80) <> 0
                    'bufAll(i)左移一位
                    bufAll(i) = ShLB_By1Bit(bufAll(i))
                    cOctets = cOctets + 1
                Loop
                
                'leading byte至少应为110x xxxx
                cOctets = cOctets - 1
                If cOctets = 0 Then
                    '返回默认编码
                    fmt = "UEF_ANSI"
                    Exit Function
                End If
            End If
        Else
            '非leading byte形式必须是 10xxxxxx
            If (bufAll(i) And &HC0) <> &H80 Then
                '返回默认编码
                fmt = "UEF_ANSI"
                Exit Function
            End If
            '准备下一个byte
            cOctets = cOctets - 1
        End If
        
    Next i
    
    '文本结束.  不应有任何多余的byte 有即为错误 返回默认编码
    If cOctets > 0 Then
        fmt = "UEF_ANSI"
        Exit Function
    End If
    
    '如果全是ascii.  需要注意的是使用相应的code pages做转换
    If bAllAscii Then
        fmt = "UEF_ANSI"
    Else
         '修成正果 终于格式全部正确 返回UTF8 No BOM编码格式
        fmt = "UEF_UTF8NB"
    End If
    UEFCheckUTF8NoBom = fmt
End Function