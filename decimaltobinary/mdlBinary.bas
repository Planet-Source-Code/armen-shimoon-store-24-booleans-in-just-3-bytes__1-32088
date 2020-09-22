Attribute VB_Name = "mdlBinary"
'2 Functions to convert from dec to bin or from bin to dec
'Parts taken from DecimalToBinary from PSC. I improved upon the
'the part that checks if the number is odd.
'The binary to decimal written totally by me.



Public Function DecimalToBinary(nByte As Integer) As String
Dim tBin As String
Dim OddNum As Boolean
Dim OddDouble As Double
Dim FoundFirstNumber As Boolean
Dim CurrentVal As Integer
Dim CurrentPos As Double

CurrentPos = 1
CurrentVal = CInt(nByte)



OddDouble = CDbl(Right(CurrentVal, 1))
OddDouble = OddDouble / 2
If Len(CStr(OddDouble)) > 1 Then
    CurrentVal = CurrentVal - 1
    OddNum = True
End If


Do While CurrentPos <> 0
    Do
        If FoundFirstNumber = False Then
            If CurrentVal <= 2 ^ CurrentPos Then
                If CurrentPos <> 0 Then
                    CurrentVal = CurrentVal - 2 ^ CurrentPos
                        If CurrentVal >= 0 Then
                            tBin = tBin & "1"
                        Else
                            CurrentVal = CurrentVal + 2 ^ CurrentPos
                        End If
                End If
                FoundFirstNumber = True
                Exit Do
            End If
            CurrentPos = CurrentPos + 1
        Else
            CurrentPos = CurrentPos - 1
            If CurrentPos = 0 Then Exit Do
            CurrentVal = CurrentVal - 2 ^ CurrentPos
            If CurrentVal >= 0 Then
                tBin = tBin & "1"
            Else
                CurrentVal = CurrentVal + 2 ^ CurrentPos
                tBin = tBin & "0"
            End If
        End If
    Loop
Loop

If OddNum = True Then tBin = tBin & "1" Else: tBin = tBin & "0"

For i = 1 To 8 - Len(tBin)
    tBin = "0" & tBin
Next i

DecimalToBinary = tBin
End Function

Public Function BinaryToDecimal(BinaryString As String) As Integer
Dim tChar As String
Dim fNum As Integer

For i = 1 To 8
    tChar = Mid(BinaryString, i, 1)
    If CInt(tChar) = 1 Then
        Select Case i
            Case 1
                fNum = fNum + 128
            Case 2
                fNum = fNum + 64
            Case 3
                fNum = fNum + 32
            Case 4
                fNum = fNum + 16
            Case 5
                fNum = fNum + 8
            Case 6
                fNum = fNum + 4
            Case 7
                fNum = fNum + 2
            Case 8
                fNum = fNum + 1
        End Select
    End If
Next i

BinaryToDecimal = fNum
End Function
