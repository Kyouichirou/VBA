Attribute VB_Name = "���ļ�ɾ��"
Option Explicit

Sub MFilesDele()   '����ļ�ɾ��
    Dim yesno As Variant
    Dim Arrow() As Integer, arr() As Integer
    Dim slc As Variant, slr As Variant
    Dim i As Integer, k As Integer, p As Integer, tfile As String, excodex As Integer, j As Integer
    
    On Error GoTo 100
    With ThisWorkbook.Sheets("���")    'ֻ�е�ѡ�����������1, ѡ��������c��ʱ,�Ž��в���
        k = Selection.Cells.Count
        If k < 2 Then
            .Label1.Caption = "ѡ����������2"
            Exit Sub
        ElseIf k > 10 Then
            .Label1.Caption = "ѡ������������Χ"
            Exit Sub
        End If
        For Each slc In Selection.Columns              '�ж���ѡ��������Ƿ�����Ҫ��,����Ҫ�󳬹�����,����ѡ��c��
            If slc.Column <> 2 Then
                .Label1.Caption = "ѡ�����򳬳���Χ,��ѡ��c��"
                Exit Sub
            End If
        Next
        i = 1
        ReDim Arrow(1 To k)     '��ȡ�������е��к�
        For Each slr In Selection.rows
            j = slr.Row
            If j < 6 Then
                .Label1.Caption = "ѡ���ļ�����,�����²���"                '��ֹ�����
                Exit Sub
            End If
            Arrow(i) = j ''��ѡ��������кŷŽ���������ʱ����
            i = i + 1
        Next
        '---------------------------------------------------------------------------ѡ�������ж�
        ReDim arr(1 To k)
        arr = Down(Arrow)                                         'ɾ����Ҫ���õ���ɾ��,�����������������,�������ֵ�����������������
        yesno = MsgBox("�Ƿ�ɾ�������ļ�?_", vbYesNo) '�Ƿ�ɾ���ļ�
        If yesno = vbYes Then '�ļ�������ִ��ɾ������
            excodex = 1
        Else
            excodex = 0
        End If
        For p = 1 To k
            tfile = .Range("e" & arr(p))
            If Len(.Cells(arr(p), "ab")) > 0 Then p = 1
            If excodex = 1 Then
                FileDeleExc tfile, arr(p), p, 0, .Cells(arr(p), "d") 'ִ��ɾ�������ļ�
            Else
                DeleMenu arr(p) '�Ƴ����
            End If
        Next
    End With
100
End Sub

Function Down(xi() As Integer) As Integer()  '������
    Dim i As Integer, j As Integer, a As Integer, d() As Integer, m As Byte, n As Byte
    
    m = LBound(xi)
    n = UBound(xi)
    ReDim d(m To n)
    ReDim Down(m To n)
    d = xi
    If m = n Then
        Down = d
        Exit Function 'ֻ��һ������
    End If

    For i = m To n - 1
        For j = i + 1 To n
            If d(j) > d(i) Then
                a = d(j): d(j) = d(i): d(i) = a
            End If
        Next
    Next
    Down = d
End Function

Private Function Downx(xi() As Integer) As Integer() '����-ʹ��sortlist��ʵ������
    Dim sortlist As Object
    Dim i As Integer, k As Byte, arrTemp() As Integer
    'https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8
    i = UBound(xi())
    ReDim arrTemp(1 To i)
    ReDim Downx(1 To i)
    Set sortlist = CreateObject("System.Collections.ArrayList") 'ע������createobject("System.Collections.SortedList")'ArrayList
    If sortlist Is Nothing Then MsgBox "�޷���������": Exit Function
    With sortlist
        For k = 1 To i
            .Add xi(k)
        Next
        .sort
        i = i - 1
        For k = 0 To i
            arrTemp(k + 1) = sortlist(k)
        Next
    End With
    Downx = arrTemp
    Set sortlist = Nothing
End Function
