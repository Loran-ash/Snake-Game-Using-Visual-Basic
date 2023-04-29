VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10215
   ClientLeft      =   3345
   ClientTop       =   1275
   ClientWidth     =   14145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   14145
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   5760
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   12960
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   4440
      Width           =   374
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '平面
      Height          =   492
      Index           =   0
      Left            =   484
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   480
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   10200
      TabIndex        =   4
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   492
      Left            =   12240
      Picture         =   "20+++.frx":0000
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   492
   End
   Begin VB.Image Image3 
      Height          =   372
      Left            =   12840
      Top             =   8640
      Width           =   372
   End
   Begin VB.Label Label3 
      Caption         =   "這裡你要自己拉一個label one當作開始"
      Height          =   852
      Left            =   10200
      TabIndex        =   3
      Top             =   3720
      Width           =   2532
   End
   Begin VB.Image Image1 
      Height          =   2292
      Left            =   10440
      Picture         =   "20+++.frx":0E84
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   612
      Left            =   11040
      TabIndex        =   2
      Top             =   4920
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim snk As String '記錄蛇身體所有的位置
Dim sdir As Integer  '行進方向
Dim shead As Integer  '頭的座標
Dim apple As Integer   '蘋果座標

Private Sub Form_Load()
    For i = 0 To 360
        If i > 0 Then
            Load Command1(i)  'load可以動態產生物件
            If i Mod 19 = 0 Then
                Command1(i).Left = Command1(0).Left
                Command1(i).Top = Command1(i - 9).Top + Command1(i).Height
            Else
                Command1(i).Left = Command1(i - 1).Left + Command1(i - 1).Width
                Command1(i).Top = Command1(i - 1).Top

            End If
        End If
        If i Mod 2 = 0 Then
            Command1(i).BackColor = RGB(115, 167, 12)
        Else
            Command1(i).BackColor = RGB(162, 209, 73)
        End If
       ' Command1(i).Caption = i
        Command1(i).Visible = True
    Next
    For i = 0 To 18
        Command1(i).BackColor = vbBlack
    Next
    For i = 342 To 360
        Command1(i).BackColor = vbBlack
    Next
    For i = 0 To 342 Step 19
        Command1(i).BackColor = vbBlack
    Next
    For i = 18 To 360 Step 19
        Command1(i).BackColor = vbBlack
    Next
    
End Sub

Private Sub NewApple()
    apple = -1
    Do
        x = Int(Rnd * 361)
        If Command1(x).BackColor <> vbBlack And Command1(x).BackColor <> vbBlue Then
            apple = x
        End If
    Loop Until apple <> -1
    Command1(apple).Picture = Image2

End Sub

Private Sub Label1_Click()
    Command1(175).BackColor = vbBlue
    Command1(176).BackColor = vbBlue
    Command1(177).BackColor = vbBlue
    snk = ChrW(175) & ChrW(176) & ChrW(177)
    sdir = 1            '一開始預設向右
    shead = 177         '一開始預設蛇首117
    NewApple
    
    Timer1.Interval = 400
End Sub

'第一題，完成方向鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        sdir = -19
    ElseIf KeyCode = vbKeyDown Then
        sdir = 19
    ElseIf KeyCode = vbKeyLeft Then
        sdir = -1
    Else
        sdir = 1
    End If
    
End Sub

Private Sub Timer1_Timer()
    shead = shead + sdir
    
    '第二題，撞牆
    If Command1(shead).BackColor = vbBlack Then
        MsgBox "hit wall, game over"
        Timer1.Interval = 0
        Exit Sub
    End If
    '第三題，咬到自己
    If Command1(shead).BackColor = vbBlue Then
        MsgBox "bite self"
        Timer1.Interval = 0
        Exit Sub
    End If

    '第四題，吃蘋果len+1
    '更新snk座標，
    '如果吃到蘋果，長度+1，留住最後一截尾巴，
    '若沒吃到，消掉最後一截尾巴
    '繪圖，頭前進一格
    Command1(shead).BackColor = vbBlue
    If shead = apple Then
        NewApple
        snk = snk & ChrW(shead)
        Command1(shead).Picture = Image3
    Else
        '先記錄尾巴座標
        stail = AscW(Mid(snk, 1, 1))
        snk = Mid(snk, 2, Len(snk)) & ChrW(shead)  'snk削掉尾巴最後一格
        '繪圖，snk削掉尾巴最後一格，還是要符合馬賽克磚
        i = stail
        If i Mod 2 = 0 Then
            
            Command1(i).BackColor = RGB(115, 167, 12)
        Else
            
            Command1(i).BackColor = RGB(162, 209, 73)
        End If
    End If

    
    Label2 = Len(snk) '目前蛇的長度
End Sub

