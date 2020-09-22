Attribute VB_Name = "modConv"
Option Explicit

' Converting DWords, Words, and Bytes

' Original code/concept Bruce McKinney

' You can write functions in Visual Basic® that split and combine
' DWords, Words, and Bytes. These functions can be used to do such
' things as taking two integer values and transferring them to a
' long integer value, where the first integer is the high word and
' the second integer is the low word.

Private Type TwoInts
    Lo As Integer
    Hi As Integer
End Type

Private Type OneLong
    dw As Long
End Type

Private Type TwoBytes
    Lo As Byte
    Hi As Byte
End Type

Private Type OneInt
    w As Integer
End Type

'Declare Function lobyte Lib "Tlbinf32" (ByVal Word As Integer) As Byte
'Declare Function hibyte Lib "Tlbinf32" (ByVal Word As Integer) As Byte
'Declare Function loword Lib "Tlbinf32" (ByVal DWord As Long) As Integer
'Declare Function hiword Lib "Tlbinf32" (ByVal DWord As Long) As Integer
'Declare Function makelong Lib "Tlbinf32" (ByVal WordLo As Integer, _
'                                          ByVal WordHi As Integer) As Long
'Declare Function makeword Lib "Tlbinf32" (ByVal ByteLo As Byte, _
'                                          ByVal ByteHi As Byte) As Integer

Function LoByte(ByVal Word As Integer) As Byte
    Dim LoHi As TwoBytes
    Dim Both As OneInt
    Both.w = Word
    LSet LoHi = Both
    LoByte = LoHi.Lo
End Function

Function HiByte(ByVal Word As Integer) As Byte
    Dim LoHi As TwoBytes
    Dim Both As OneInt
    Both.w = Word
    LSet LoHi = Both
    HiByte = LoHi.Hi
End Function

Function LoWord(ByVal DWord As Long) As Integer
    Dim LoHi As TwoInts
    Dim Both As OneLong
    Both.dw = DWord
    LSet LoHi = Both
    LoWord = LoHi.Lo
End Function

Function HiWord(ByVal DWord As Long) As Integer
    Dim LoHi As TwoInts
    Dim Both As OneLong
    Both.dw = DWord
    LSet LoHi = Both
    HiWord = LoHi.Hi
End Function

Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
    Dim LoHi As TwoBytes
    Dim Both As OneInt
    LoHi.Lo = LoByte
    LoHi.Hi = HiByte
    LSet Both = LoHi
    MakeWord = Both.w
End Function

Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    Dim LoHi As TwoInts
    Dim Both As OneLong
    LoHi.Lo = LoWord
    LoHi.Hi = HiWord
    LSet Both = LoHi
    MakeDWord = Both.dw
End Function

Function LShiftWord(ByVal Word As Integer, ByVal NumBits As Integer) As Integer
    If (NumBits >= 1) And (NumBits <= 15) Then
        Dim dw As Long
        dw = Word * (2 ^ NumBits)
        If dw And &H8000& Then
            LShiftWord = CInt(dw And &H7FFF&) Or &H8000
        Else
            LShiftWord = dw And &HFFFF&
        End If
    End If
End Function

Function RShiftWord(ByVal Word As Integer, ByVal NumBits As Integer) As Integer
    If (NumBits >= 1) And (NumBits <= 15) Then
        Dim dw As Long
        dw = Word \ (2 ^ NumBits)
        If dw And &H8000& Then
            RShiftWord = CInt(dw And &H7FFF&) Or &H8000
        Else
            RShiftWord = dw And &HFFFF&
        End If
    End If
End Function

Function LShiftDWord(ByVal DWord As Long, ByVal NumBits As Integer) As Long
    If (NumBits >= 1) And (NumBits <= 31) Then
        Dim ti As TwoInts, ol As OneLong, i As Integer
        ol.dw = DWord
        LSet ti = ol
        For i = 1 To NumBits
            ti.Hi = LShiftWord(ti.Hi, 1)
            If ti.Lo And &H80 Then ti.Hi = ti.Hi + 1
            ti.Lo = LShiftWord(ti.Lo, 1)
        Next i
        LSet ol = ti
        LShiftDWord = ol.dw
    End If
End Function

Function RShiftDWord(ByVal DWord As Long, ByVal NumBits As Integer) As Long
    If (NumBits >= 1) And (NumBits <= 31) Then
        Dim ti As TwoInts, ol As OneLong, i As Integer
        ol.dw = DWord
        LSet ti = ol
        For i = 1 To NumBits
            ti.Lo = RShiftWord(ti.Lo, 1)
            If ti.Hi And 1 Then ti.Lo = ti.Lo + &H80
            ti.Hi = RShiftWord(ti.Hi, 1)
        Next i
        LSet ol = ti
        RShiftDWord = ol.dw
    End If
End Function

' Set or clear iBitPos bit in iValue according to fTest expression.
Sub SetBit(ByVal fTest As Boolean, iValue As Integer, ByVal iBitPos As Integer)
    If fTest Then
        iValue = LoWord(iValue Or (2 ^ iBitPos))
    Else
        iValue = LoWord(iValue And Not (2 ^ iBitPos))
    End If
End Sub

Sub SetBitLong(ByVal fTest As Boolean, iValue As Long, ByVal iBitPos As Integer)
    If fTest Then
        iValue = iValue Or (2 ^ iBitPos)
    Else
        iValue = iValue And Not (2 ^ iBitPos)
    End If
End Sub

' Get state of iBitPos bit in iValue
Function GetBit(ByVal iValue As Long, ByVal iBitPos As Integer) As Boolean
    GetBit = iValue And (2 ^ iBitPos)
End Function

