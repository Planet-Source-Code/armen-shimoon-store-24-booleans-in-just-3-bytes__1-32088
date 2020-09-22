VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bitwise Boolean Compression"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   3360
      TabIndex        =   25
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   3360
      TabIndex        =   24
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   23
      Left            =   1680
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   22
      Left            =   1680
      TabIndex        =   22
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   21
      Left            =   1680
      TabIndex        =   21
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   20
      Left            =   1680
      TabIndex        =   20
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   19
      Left            =   1680
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   18
      Left            =   1680
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   17
      Left            =   1680
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   16
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   15
      Left            =   1680
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   13
      Left            =   1680
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   12
      Left            =   1680
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit %"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "STORE 24 BOOLEAN STATEMENTS IN JUST 3 BYTES OF MEMORY!!!!!"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   3135
      Left            =   3240
      TabIndex        =   26
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call LoadBits
End Sub



Private Sub Command2_Click()
Dim bString(2) As String
Dim bDec(2) As Byte

For i = 0 To 7
    bString(0) = bString(0) & chkBit(i).Value
Next i

For i = 8 To 15
    bString(1) = bString(1) & chkBit(i).Value
Next i

For i = 16 To 23
    bString(2) = bString(2) & chkBit(i).Value
Next i

For i = 0 To 2
    bDec(i) = BinaryToDecimal(bString(i))
Next i

Open App.Path & "\options.txt" For Binary As #1

For i = 0 To 2
    Put #1, , bDec(i)
Next i

Close #1


End Sub

Private Sub Form_Load()

For i = 0 To 23
    chkBit(i).Caption = "Bit " & i
Next i


Call LoadBits
End Sub

Sub LoadBits()
Dim tChar As String
Dim tAsc As Integer
Dim binstring As String
Dim lCounter As Integer
Dim tString As String

Open App.Path & "\options.txt" For Input As #1
    Input #1, tString
Close #1

lCounter = 0
For i = 1 To Len(tString)
    lCounter = lCounter + 1
    tChar = Mid(tString, i, 1)
    tAsc = Asc(tChar)
    binstring = DecimalToBinary(tAsc)
    
    If lCounter = 1 Then
        For b = 1 To Len(binstring)
            chkBit(b - 1).Value = Mid(binstring, b, 1)
        Next b
    End If
    
    If lCounter = 2 Then
                For b = 1 To Len(binstring)
                    chkBit(b + 7).Value = Mid(binstring, b, 1)
                Next b
    End If
    
    If lCounter = 3 Then
                For b = 1 To Len(binstring)
                    chkBit(b + 15).Value = Mid(binstring, b, 1)
                Next b
    End If
Next i
End Sub
