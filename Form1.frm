VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mythware-kill"
   ClientHeight    =   4040
   ClientLeft      =   110
   ClientTop       =   440
   ClientWidth     =   7990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4040
   ScaleWidth      =   7990
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   850
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   6130
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   22
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   5770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "进程名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag As Boolean, Pp2, Pp3, PZ1, PZ2
Private Sub Command1_Click()
    File
    If Flag = True Then
        PZ1 = Environ("UserProfile") & "\Desktop\"
        Rem-----------------------------------------------------------------
        'taskkill
        PZ2 = "taskkill /f /t /im" & Space(1) & Chr(34) & Pp3 & Chr(34)
        Open PZ1 & "Kill.bat" For Output As #1
            Print #1, PZ2
        Close
        Call Shell(PZ1 & "Kill.bat")
        'delet
        PZ2 = "rd/s/q" & Space(1) & Chr(34) & Pp2 & Chr(34)
        Open PZ1 & "Delet.bat" For Output As #1
            Print #1, PZ2
        Close
        Call Shell(PZ1 & "Delet.bat")
    End If
End Sub
Sub File()
    Dim Pp$, i2%
    Pp = Text1.Text
    i2 = 1
        For i2 = Len(Pp) To 1 Step -1
            If Mid(Pp, i2, 1) = "\" Then
                Flag = True
                Exit For
            End If
        Next
        Pp3 = Right(Pp, Len(Pp) - i2)
        Text1.Text = Pp3
        Pp2 = Left(Pp, i2 - 1)
End Sub


