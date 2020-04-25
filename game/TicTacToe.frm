VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   Icon            =   "TicTacToe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7110
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Choose 
      Appearance      =   0  'Flat
      Caption         =   "Let me choose who goes first."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Points2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Points1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   5400
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton AIShape 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Microsoft JhengHei UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AI"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton PlayerShape 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Microsoft JhengHei UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton grid 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label LblOutPut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Click me START!"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NN = 0 'None
Const PL = 1 'Player
Const AI = 2 'AI
Const DGPL = 3 'Danger to player
Const DGAI = 4 'Danger to AI

Const WaitStart = 0
Const PlayersTurn = 1
Const AIsTurn = 2

Const DangerColor = &HD9D9EF
Const SafeColor = &H80000016
Const WinColor = &HE1C2AE
Const O = "O"
Const X = "X"

Dim First As Integer
Dim Status(1 To 9) As Integer
Dim GameStatus As Integer

Private Sub DANGER()
    If Status(1) = Status(2) And Status(1) = AI And Status(3) = NN Then Status(3) = DGPL
    If Status(1) = Status(3) And Status(1) = AI And Status(2) = NN Then Status(2) = DGPL
    If Status(2) = Status(3) And Status(2) = AI And Status(1) = NN Then Status(1) = DGPL
    
    If Status(4) = Status(6) And Status(4) = AI And Status(5) = NN Then Status(5) = DGPL
    If Status(4) = Status(5) And Status(4) = AI And Status(6) = NN Then Status(6) = DGPL
    If Status(6) = Status(5) And Status(5) = AI And Status(4) = NN Then Status(4) = DGPL
    
    If Status(9) = Status(7) And Status(9) = AI And Status(8) = NN Then Status(8) = DGPL
    If Status(7) = Status(8) And Status(7) = AI And Status(9) = NN Then Status(9) = DGPL
    If Status(9) = Status(8) And Status(8) = AI And Status(7) = NN Then Status(7) = DGPL
    
    If Status(1) = Status(4) And Status(1) = AI And Status(7) = NN Then Status(7) = DGPL
    If Status(4) = Status(7) And Status(4) = AI And Status(1) = NN Then Status(1) = DGPL
    If Status(7) = Status(1) And Status(7) = AI And Status(4) = NN Then Status(4) = DGPL
    
    If Status(2) = Status(5) And Status(2) = AI And Status(8) = NN Then Status(8) = DGPL
    If Status(5) = Status(8) And Status(5) = AI And Status(2) = NN Then Status(2) = DGPL
    If Status(8) = Status(2) And Status(8) = AI And Status(5) = NN Then Status(5) = DGPL
    
    If Status(3) = Status(6) And Status(3) = AI And Status(9) = NN Then Status(9) = DGPL
    If Status(6) = Status(9) And Status(6) = AI And Status(3) = NN Then Status(3) = DGPL
    If Status(9) = Status(3) And Status(9) = AI And Status(6) = NN Then Status(6) = DGPL
    
    If Status(7) = Status(5) And Status(7) = AI And Status(3) = NN Then Status(3) = DGPL
    If Status(5) = Status(3) And Status(5) = AI And Status(7) = NN Then Status(7) = DGPL
    If Status(3) = Status(7) And Status(3) = AI And Status(5) = NN Then Status(5) = DGPL
    
    If Status(1) = Status(5) And Status(1) = AI And Status(9) = NN Then Status(9) = DGPL
    If Status(5) = Status(9) And Status(5) = AI And Status(1) = NN Then Status(1) = DGPL
    If Status(1) = Status(9) And Status(1) = AI And Status(5) = NN Then Status(5) = DGPL
    
    
    If Status(1) = Status(2) And Status(1) = PL And Status(3) = NN Then Status(3) = DGAI
    If Status(1) = Status(3) And Status(1) = PL And Status(2) = NN Then Status(2) = DGAI
    If Status(2) = Status(3) And Status(2) = PL And Status(1) = NN Then Status(1) = DGAI
    
    If Status(4) = Status(6) And Status(4) = PL And Status(5) = NN Then Status(5) = DGAI
    If Status(4) = Status(5) And Status(4) = PL And Status(6) = NN Then Status(6) = DGAI
    If Status(6) = Status(5) And Status(5) = PL And Status(4) = NN Then Status(4) = DGAI
    
    If Status(9) = Status(7) And Status(9) = PL And Status(8) = NN Then Status(8) = DGAI
    If Status(7) = Status(8) And Status(7) = PL And Status(9) = NN Then Status(9) = DGAI
    If Status(9) = Status(8) And Status(8) = PL And Status(7) = NN Then Status(7) = DGAI
    
    If Status(1) = Status(4) And Status(1) = PL And Status(7) = NN Then Status(7) = DGAI
    If Status(4) = Status(7) And Status(4) = PL And Status(1) = NN Then Status(1) = DGAI
    If Status(7) = Status(1) And Status(7) = PL And Status(4) = NN Then Status(4) = DGAI
    
    If Status(2) = Status(5) And Status(2) = PL And Status(8) = NN Then Status(8) = DGAI
    If Status(5) = Status(8) And Status(5) = PL And Status(2) = NN Then Status(2) = DGAI
    If Status(8) = Status(2) And Status(8) = PL And Status(5) = NN Then Status(5) = DGAI
    
    If Status(3) = Status(6) And Status(3) = PL And Status(9) = NN Then Status(9) = DGAI
    If Status(6) = Status(9) And Status(6) = PL And Status(3) = NN Then Status(3) = DGAI
    If Status(9) = Status(3) And Status(9) = PL And Status(6) = NN Then Status(6) = DGAI
    
    If Status(7) = Status(5) And Status(7) = PL And Status(3) = NN Then Status(3) = DGAI
    If Status(5) = Status(3) And Status(5) = PL And Status(7) = NN Then Status(7) = DGAI
    If Status(3) = Status(7) And Status(3) = PL And Status(5) = NN Then Status(5) = DGAI
    
    If Status(1) = Status(5) And Status(1) = PL And Status(9) = NN Then Status(9) = DGAI
    If Status(5) = Status(9) And Status(5) = PL And Status(1) = NN Then Status(1) = DGAI
    If Status(1) = Status(9) And Status(1) = PL And Status(5) = NN Then Status(5) = DGAI
End Sub

Private Function JUDGE()
    If Status(1) = PL And Status(2) = PL And Status(3) = PL Then JUDGE = 1
    If Status(4) = PL And Status(5) = PL And Status(6) = PL Then JUDGE = 1
    If Status(7) = PL And Status(8) = PL And Status(9) = PL Then JUDGE = 1
    If Status(1) = PL And Status(4) = PL And Status(7) = PL Then JUDGE = 1
    If Status(2) = PL And Status(5) = PL And Status(8) = PL Then JUDGE = 1
    If Status(3) = PL And Status(6) = PL And Status(9) = PL Then JUDGE = 1
    If Status(1) = PL And Status(5) = PL And Status(9) = PL Then JUDGE = 1
    If Status(7) = PL And Status(5) = PL And Status(3) = PL Then JUDGE = 1
    
    If Status(1) = AI And Status(2) = AI And Status(3) = AI Then JUDGE = 2
    If Status(4) = AI And Status(5) = AI And Status(6) = AI Then JUDGE = 2
    If Status(7) = AI And Status(8) = AI And Status(9) = AI Then JUDGE = 2
    If Status(1) = AI And Status(4) = AI And Status(7) = AI Then JUDGE = 2
    If Status(2) = AI And Status(5) = AI And Status(8) = AI Then JUDGE = 2
    If Status(3) = AI And Status(6) = AI And Status(9) = AI Then JUDGE = 2
    If Status(1) = AI And Status(5) = AI And Status(9) = AI Then JUDGE = 2
    If Status(7) = AI And Status(5) = AI And Status(3) = AI Then JUDGE = 2
    If JUDGE = 1 Or JUDGE = 2 Then Exit Function
    If (Status(1) = 1 Or Status(1) = 2) And (Status(2) = 1 Or Status(2) = 2) And (Status(3) = 1 Or Status(3) = 2) And (Status(4) = 1 Or Status(4) = 2) And (Status(5) = 1 Or Status(5) = 2) And (Status(6) = 1 Or Status(6) = 2) And (Status(7) = 1 Or Status(7) = 2) And (Status(8) = 1 Or Status(8) = 2) And (Status(9) = 1 Or Status(9) = 2) Then JUDGE = 3
End Function

Private Sub CALC()
    For i = 1 To 9
        Select Case Status(i)
            Case NN
                grid(i).Caption = ""
                grid(i).BackColor = SafeColor
            Case PL
                grid(i).Caption = PlayerShape.Caption
                grid(i).BackColor = SafeColor
            Case AI
                grid(i).Caption = AIShape.Caption
                grid(i).BackColor = SafeColor
            Case DGPL
                grid(i).Caption = ""
                grid(i).BackColor = DangerColor
            Case DGAI
                grid(i).Caption = ""
                grid(i).BackColor = WinColor
        End Select
    Next i
End Sub

Private Sub grid_Click(Index As Integer)
    If (Status(Index) = 0 Or Status(Index) > 2) And GameStatus = PlayersTurn Then
        Status(Index) = PL
        DANGER
        CALC
        If JUDGE = 1 Then
            MsgBox "Player WiNS!"
            GameStatus = WaitStart
            LblOutPut.Caption = "Click me START!"
            Points1.Text = Val(Points1.Text) + 1
        ElseIf JUDGE = 3 Then
            MsgBox "DRAW!"
            GameStatus = WaitStart
            LblOutPut.Caption = "Click me START!"
        Else
            GameStatus = AIsTurn
            AImove
        End If
    End If
End Sub

Private Sub LblOutPut_Click()
    If GameStatus = WaitStart Then
        Start_Game
        If Choose.Value = 0 Then First = Int(Rnd * 2) + 1
        If Choose.Value = 1 Then First = MsgBox("Do you want to be FIRST?", vbOKCancel)
        GameStatus = First
        If First = 1 Then
            LblOutPut.Caption = "You first."
            PlayerShape.Caption = O
            AIShape.Caption = X
        ElseIf First = 2 Then
            LblOutPut.Caption = "AI first."
            PlayerShape.Caption = X
            AIShape.Caption = O
            AImove
        End If
    End If
End Sub
Private Sub Start_Game()
    For i = 1 To 9
        Status(i) = NN
    Next i
    CALC
End Sub
Private Sub AImove()
    Status(getAInum) = AI
    DANGER
    CALC
    If JUDGE = 2 Then
        MsgBox "AI wins!"
        GameStatus = WaitStart
        LblOutPut.Caption = "Click me START!"
        Points2.Text = Val(Points2.Text) + 1
    ElseIf JUDGE = 3 Then
        MsgBox "DRAW!"
        GameStatus = WaitStart
        LblOutPut.Caption = "Click me START!"
    Else
        GameStatus = PlayersTurn
    End If
End Sub
Private Function getAInum()
    If Status(1) = PL And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = PL And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = PL And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = PL Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = PL And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = AI And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = PL Then
        getAInum = 2
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = PL And Status(4) = NN And Status(5) = AI And Status(6) = NN And Status(7) = PL And Status(8) = NN And Status(9) = NN Then
        getAInum = 2
        Exit Function
    End If
    If Status(1) = AI And Status(2) = PL And Status(3) = AI And Status(4) = PL And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 9
        Exit Function
    End If
    If Status(1) = AI And Status(2) = PL And Status(3) = NN And Status(4) = PL And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = AI And Status(2) = PL And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 9
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = PL And Status(9) = NN Then
        getAInum = 9
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = PL And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = NN And Status(8) = NN And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = PL And Status(8) = PL And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    '========================================================================================
    If Status(1) = AI And Status(2) = PL And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = PL And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = AI And Status(2) = PL And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = PL Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = AI And Status(2) = NN And Status(3) = PL And Status(4) = PL And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = NN Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = AI And Status(2) = NN And Status(3) = NN And Status(4) = PL And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = NN And Status(9) = PL Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = PL And Status(8) = NN And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = PL And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = NN And Status(8) = NN And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = PL And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = PL And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = PL And Status(4) = NN And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = PL And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = PL And Status(5) = NN And Status(6) = NN And Status(7) = NN And Status(8) = PL And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = NN And Status(8) = PL And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = NN And Status(2) = PL And Status(3) = NN And Status(4) = NN And Status(5) = NN And Status(6) = PL And Status(7) = NN And Status(8) = NN And Status(9) = AI Then
        getAInum = 5
        Exit Function
    End If
    If Status(1) = PL And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = AI And Status(6) = NN And Status(7) = NN And Status(8) = PL And Status(9) = NN Then
        getAInum = 7
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = PL And Status(4) = NN And Status(5) = AI And Status(6) = NN And Status(7) = NN And Status(8) = PL And Status(9) = NN Then
        getAInum = 9
        Exit Function
    End If
    If Status(1) = NN And Status(2) = NN And Status(3) = NN And Status(4) = NN And Status(5) = AI And Status(6) = PL And Status(7) = PL And Status(8) = NN And Status(9) = NN Then
        getAInum = 9
        Exit Function
    End If
    '======================================================================================
    For i = 1 To 9
        If Status(i) = DGPL Then
            getAInum = i
            Exit Function
        End If
    Next i
    For i = 1 To 9
        If Status(i) = DGAI Then
            getAInum = i
            Exit Function
        End If
    Next i
    For i = 1 To 9
        If (i = 1 Or i = 3 Or i = 7 Or i = 9) And Status(i) = NN Then
            getAInum = i
            Exit Function
        End If
    Next i
    For i = 1 To 9
        If (i = 2 Or i = 4 Or i = 6 Or i = 8) And Status(i) = NN Then
            getAInum = i
            Exit Function
        End If
    Next i
    getAInum = i
End Function
