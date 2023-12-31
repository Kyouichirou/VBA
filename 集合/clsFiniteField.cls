VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFiniteField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const UseInverseLUT = False

Private Const m_nMaxOrderShift As Long = 14
Private Const m_nMaxOrder As Long = 2 ^ m_nMaxOrderShift

Private m_nOrder As Long
Private m_nPolynomial As Long
Private m_nPrimitiveRoot As Long

Private m_nDiscreteLog(m_nMaxOrder - 1) As Integer
Private m_nDiscreteExp(m_nMaxOrder - 1) As Integer
#If UseInverseLUT Then
Private m_nInverse(m_nMaxOrder - 1) As Integer
#End If

Friend Property Get Order() As Long
    Order = m_nOrder
End Property

Friend Property Get MinimalPolynomial() As Long
    MinimalPolynomial = m_nPolynomial
End Property

Friend Property Get PrimitiveRoot() As Long
    PrimitiveRoot = m_nPrimitiveRoot
End Property

Friend Function pFindPrimitiveRoot() As Long
    Dim nPrimeFactor(7) As Long
    Dim t(m_nMaxOrderShift - 1) As Long
    Dim nPrimeFactorCount As Long
    Dim nPrimitiveRoot As Long
    Dim i As Long, j As Long, k As Long
    Dim m As Long
    i = m_nOrder - 1
    If (i And 1&) = 0 Then
        nPrimeFactor(0) = i \ 2&
        nPrimeFactorCount = 1
        Do Until i And 1&
            i = i \ 2&
        Loop
    End If
    m = Sqr(i) + 2
    j = 3
    Do Until j > m
        k = i \ j
        If k * j = i Then
            Do
                i = k
                k = i \ j
            Loop While k * j = i
            nPrimeFactor(nPrimeFactorCount) = (m_nOrder - 1) \ j
            nPrimeFactorCount = nPrimeFactorCount + 1
            m = Sqr(i) + 2
        End If
        j = j + 2
    Loop
    If i > 1 Then
        nPrimeFactor(nPrimeFactorCount) = (m_nOrder - 1) \ i
        nPrimeFactorCount = nPrimeFactorCount + 1
    End If
    For nPrimitiveRoot = 2 To m_nOrder - 1
        t(0) = nPrimitiveRoot
        j = nPrimitiveRoot
        For i = 1 To m_nMaxOrderShift - 1
            j = Multiply(j, j)
            t(i) = j
        Next i
        For i = 0 To nPrimeFactorCount - 1
            m = nPrimeFactor(i)
            j = 1
            For k = 0 To m_nMaxOrderShift - 1
                If m And 1& Then j = Multiply(j, t(k))
                m = m \ 2&
                If m = 0 Then Exit For
            Next k
            If j = 1 Then Exit For
        Next i
        If i >= nPrimeFactorCount Then
            pFindPrimitiveRoot = nPrimitiveRoot
            Exit Function
        End If
    Next nPrimitiveRoot
End Function

Friend Function Init(ByVal nOrder As Long, Optional ByVal nPolynomial As Long, Optional ByVal nPrimitiveRoot As Long) As Boolean
    Dim i As Long, j As Long, k As Long
    If nOrder <= 1 Or nOrder > m_nMaxOrder Then Exit Function
    If nOrder = 2 Then
        m_nOrder = 2
        m_nPolynomial = 2
        m_nPrimitiveRoot = 1
        m_nDiscreteLog(0) = 0
        m_nDiscreteLog(1) = 0
        m_nDiscreteExp(0) = 1
        m_nDiscreteExp(1) = 1
#If UseInverseLUT Then
        m_nInverse(0) = 0
        m_nInverse(1) = 1
#End If
    ElseIf (nOrder And (nOrder - 1)) = 0 Then
        If nPolynomial = 0 Then Exit Function
        m_nOrder = nOrder
        m_nPolynomial = nPolynomial
        If nPrimitiveRoot = 0 Then nPrimitiveRoot = pFindPrimitiveRoot
        m_nPrimitiveRoot = nPrimitiveRoot
        m_nDiscreteLog(0) = 0
        m_nDiscreteLog(1) = 0
        m_nDiscreteLog(nPrimitiveRoot) = 1
        m_nDiscreteExp(0) = 1
        m_nDiscreteExp(nOrder - 1) = 1
        m_nDiscreteExp(1) = nPrimitiveRoot
        j = nPrimitiveRoot
        For i = 2 To nOrder - 2
            j = Multiply(j, nPrimitiveRoot)
            m_nDiscreteLog(j) = i
            m_nDiscreteExp(i) = j
        Next i
#If UseInverseLUT Then
        m_nInverse(0) = 0
        m_nInverse(1) = 1
        For i = 1 To nOrder - 2
            m_nInverse(m_nDiscreteExp(i)) = m_nDiscreteExp(nOrder - 1 - i)
        Next i
#End If
    Else
        m_nOrder = nOrder
        m_nPolynomial = nOrder
        If nPrimitiveRoot = 0 Then nPrimitiveRoot = pFindPrimitiveRoot
        m_nPrimitiveRoot = nPrimitiveRoot
        m_nDiscreteLog(0) = 0
        m_nDiscreteLog(1) = 0
        m_nDiscreteLog(nPrimitiveRoot) = 1
        m_nDiscreteExp(0) = 1
        m_nDiscreteExp(nOrder - 1) = 1
        m_nDiscreteExp(1) = nPrimitiveRoot
        j = nPrimitiveRoot
        For i = 2 To nOrder - 2
            j = (j * nPrimitiveRoot) Mod nOrder
            m_nDiscreteLog(j) = i
            m_nDiscreteExp(i) = j
        Next i
#If UseInverseLUT Then
        m_nInverse(0) = 0
        m_nInverse(1) = 1
        For i = 1 To nOrder - 2
            m_nInverse(m_nDiscreteExp(i)) = m_nDiscreteExp(nOrder - 1 - i)
        Next i
#End If
    End If
    Init = True
End Function

Friend Function Add(ByVal n1 As Long, ByVal n2 As Long) As Long
    If (m_nOrder And (m_nOrder - 1)) = 0 Then
        Add = n1 Xor n2
    Else
        n1 = n1 + n2
        If n1 >= m_nOrder Then n1 = n1 - m_nOrder
        Add = n1
    End If
End Function

Friend Function Negative(ByVal n1 As Long) As Long
    If (m_nOrder And (m_nOrder - 1)) = 0 Or n1 = 0 Then
        Negative = n1
    Else
        Negative = m_nOrder - n1
    End If
End Function

Friend Function Subtract(ByVal n1 As Long, ByVal n2 As Long) As Long
    If (m_nOrder And (m_nOrder - 1)) = 0 Then
        Subtract = n1 Xor n2
    Else
        n1 = n1 - n2
        If n1 < 0 Then n1 = n1 + m_nOrder
        Subtract = n1
    End If
End Function

Friend Function MultiplyLUT(ByVal n1 As Long, ByVal n2 As Long) As Long
    If n1 = 0 Or n2 = 0 Then Exit Function
    n1 = m_nDiscreteLog(n1) + m_nDiscreteLog(n2)
    If n1 >= m_nOrder - 1 Then n1 = n1 - m_nOrder + 1
    MultiplyLUT = m_nDiscreteExp(n1)
End Function

Friend Function Multiply(ByVal n1 As Long, ByVal n2 As Long) As Long
    Dim i As Long, j As Long
    If n1 = 0 Or n2 = 0 Then Exit Function
    If m_nOrder = 2 Then
        Multiply = 1
    ElseIf (m_nOrder And (m_nOrder - 1)) = 0 Then
        i = 1
        Do
            If n2 And i Then
                n2 = n2 Xor i
                j = j Xor n1
            End If
            If n2 = 0 Then Exit Do
            i = i + i
            n1 = n1 + n1
            If n1 And m_nOrder Then n1 = (n1 Xor m_nPolynomial) And (m_nOrder - 1)
        Loop
        Multiply = j
    Else
        Multiply = (n1 * n2) Mod m_nOrder
    End If
End Function

Friend Function InverseLUT(ByVal n As Long) As Long
    If n = 0 Then Exit Function
#If UseInverseLUT Then
    InverseLUT = m_nInverse(n)
#Else
    InverseLUT = m_nDiscreteExp(m_nOrder - 1 - m_nDiscreteLog(n))
#End If
End Function

Friend Function Inverse(ByVal b As Long) As Long
    Dim a As Long, q As Long, t As Long
    Dim y As Long, Y1 As Long
    If b = 0 Then Exit Function
    If m_nOrder = 2 Then
        Inverse = 1
    ElseIf (m_nOrder And (m_nOrder - 1)) = 0 Then
        Debug.Assert False
    Else
        a = m_nOrder
        y = 1
        Do While b <> 0
            q = a \ b
            t = a - q * b
            a = b
            b = t
            t = Y1 - q * y
            Y1 = y
            y = t
        Loop
        Y1 = Y1 Mod m_nOrder
        If Y1 < 0 Then Y1 = Y1 + m_nOrder
        Inverse = Y1
    End If
End Function

Friend Function DivideLUT(ByVal n1 As Long, ByVal n2 As Long) As Long
    If n1 = 0 Or n2 = 0 Then Exit Function
    n1 = m_nDiscreteLog(n1) - m_nDiscreteLog(n2)
    If n1 < 0 Then n1 = n1 + m_nOrder - 1
    DivideLUT = m_nDiscreteExp(n1)
End Function

Friend Function Divide(ByVal n1 As Long, ByVal n2 As Long) As Long
    Dim a As Long, q As Long, t As Long
    Dim y As Long, Y1 As Long
    If n1 = 0 Or n2 = 0 Then Exit Function
    If m_nOrder = 2 Then
        Divide = 1
    ElseIf (m_nOrder And (m_nOrder - 1)) = 0 Then
        Debug.Assert False
    Else
        a = m_nOrder
        y = 1
        Do While n2 <> 0
            q = a \ n2
            t = a - q * n2
            a = n2
            n2 = t
            t = Y1 - q * y
            Y1 = y
            y = t
        Loop
        Y1 = (n1 * Y1) Mod m_nOrder
        If Y1 < 0 Then Y1 = Y1 + m_nOrder
        Divide = Y1
    End If
End Function

Friend Function DiscreteLog(ByVal n As Long) As Long
    DiscreteLog = m_nDiscreteLog(n)
End Function

Friend Function DiscreteExp(ByVal n As Long) As Long
    n = n Mod (m_nOrder - 1)
    If n < 0 Then n = n + m_nOrder - 1
    DiscreteExp = m_nDiscreteExp(n)
End Function
