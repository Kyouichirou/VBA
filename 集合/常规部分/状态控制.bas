Attribute VB_Name = "״̬����"
Option Explicit
'��Ҫ����״̬��sub,���ڶ��������,�����ڱ���������������ʱ����ɽ�����ֹ
Private Appspy As Eventspy

Sub EveSpy() '�����������excel���¼�����-�����жϹ������Ƿ������С��
    Set Appspy = New Eventspy
End Sub

Function RecData() As Boolean 'ִ�д洢�ļ����Ӽ��
    Dim FilePath As String
                                        'ÿ�μ���ʱ,���conn�����Ƿ�����
    If Conn.State = adStateClosed Then
       On Error GoTo 100
       Set Conn = Nothing
       FilePath = ThisWorkbook.Sheets("temp").Range("ab3").Value
       If fso.fileexists(FilePath) = False Then                '����ļ��Ƿ����
          RecData = False
          MsgBox "!���������쳣,�����޷���������", vbCritical, "Warning"
          Exit Function
        End If
       Conn.Open "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";extended properties=""excel 12.0;HDR=YES""" '�����ݴ洢�ļ�
       RecData = True
    End If
    RecData = True
    Exit Function
100
    Debug.Print Err.Number
    RecData = False '��������ӹ����г��ִ���
End Function

Sub LockWorkSheet() '��������ֹ�ֶ��޸�,ֻ�������ȥ�޸�
    Dim i As Byte, k As Byte
    
    With ThisWorkbook
        k = .Worksheets.Count
        For i = 1 To k
            .Worksheets(i).Protect "123", UserInterfaceOnly:=True
        Next
    End With
End Sub

Sub UnLockWorkSheet() '�������
    Dim i As Byte, k As Byte
    
    With ThisWorkbook
        k = .Worksheets.Count
        For i = 1 To k
            .Worksheets(i).Unprotect ("123")
        Next
    End With
End Sub
