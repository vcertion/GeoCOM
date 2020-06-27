VERSION 5.00
Begin VB.Form woautomesure 
   Caption         =   "自动测量参数设置"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   6075
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "点位遮挡"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   5895
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "放弃观测"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "等待"
         Height          =   495
         Left            =   1080
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "秒后重测"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "测距模式"
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   5895
      Begin VB.OptionButton Option4 
         Caption         =   "标准测量"
         Height          =   615
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "精密测量"
         Height          =   615
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "观测次数"
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "基准点观测数"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "感测点观测数"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame 度盘设置 
      Caption         =   "度盘设置"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.OptionButton Option2 
         Caption         =   "盘左感测"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "盘左盘右观测"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
End
Attribute VB_Name = "woautomesure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
         If Option1.Value = True Then
            ChangeF = True
        ElseIf Option2.Value = True Then
            ChangeF = False

        End If

        If Option3.Value = True Then
            MeasureMode = True
        ElseIf Option4.Value = True Then
            MeasureMode = False
        End If

        If Option5.Value = True Then
            wait = True
            wait_1 = Text3.Text

        ElseIf Option6.Value = True Then
            wait = False
        End If
        MsgBox "设置成功"
    End Sub


Private Sub Command2_Click()
Unload Me
End Sub

