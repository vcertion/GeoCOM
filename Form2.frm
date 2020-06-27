VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "目标设置"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8790
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "测点信息"
      Height          =   7095
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "刷新"
         Height          =   495
         Left            =   1800
         TabIndex        =   22
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "采集"
         Height          =   495
         Left            =   480
         TabIndex        =   21
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "删除"
         Height          =   495
         Left            =   1800
         TabIndex        =   20
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "学习"
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "(0:基准点 1：目标点 2：后视点）"
         Height          =   615
         Left            =   2760
         TabIndex        =   29
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "(m)"
         Height          =   375
         Left            =   2760
         TabIndex        =   28
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "(m)"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "(m)"
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "(m)"
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "(rad)"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "(rad)"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "类型"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "斜距"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "H"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "E"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "N"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "竖直角"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "水平角"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "点名"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据浏览"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6495
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   11456
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
 '采集数据进入数据库，在开始自动测量时调用这些数据
  Set conn = New ADODB.Connection
          conn.open ConStr
        conn.Execute "insert into 点位学习表 (点名,类型,N坐标,E坐标,H坐标,距离,水平角,竖直角)  values ('" & Text1.Text & "','" & TextBox8.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')"
        bind
        conn.Close
End Sub

Private Sub Command2_Click()
 Set conn = New ADODB.Connection
conn.open ConStr
conn.Execute "delete from 点位学习表 where 点名='" & Text1.Text & "'"
conn.Close
End Sub

Private Sub Command3_Click()
  Dim Hz As Double
        Dim v As Double
        Dim dSearchHz As Double
        dSearchHz = 0.08
        Dim dSearchV As Double
        dsearcjhv = 0.08
        Dim result As GRC_TYPE
        Dim ATRState As ON_OFF_TYPE
        Dim POSMode As Long
        POSMode = AUT_PRECISE
        Dim ATRMode As Long
        Dim OnOff As Long
        Dim rc As Integer
        Dim OnlyAngel As TMC_HZ_V_ANG
        Dim SlopeDistance As Double
        Dim Timeoutpar As AUT_TIMEOUT
        Dim TryAgain As Boolean
        TryAgain = True
        Dim nComeTimeOut, nOldComTimeOut As Short
        hrc = VB_AUS_SetUserAtrState(0)
        result = VB_AUS_GetUserAtrState(ATRState) 'ATR模式必须设置为微调模式
        If ATRState = 1 Then
            result = VB_AUT_FineAdjust3(dSearchHz, dSearchV, False)
            Select Case result
                Case 0
                    result = VB_AUS_SetUserLockState(0)
                    result = VB_AUT_LockIn()
                    result = VB_TMC_SetEdmMode(EDM_CONT_EXACT)
                    result = VB_TMC_DoMeasure(TMC_DEF_DIST, 1)
                    result = VB_TMC_GetSimpleMea(300, OnlyAngle, SlopeDistance, 0)
                    Select Case result
                        Case 0
                            Text2.Text = OnlyAngel.dHz
                            Text3.Text = OnlyAngel.dV
                            Text7.Text = SlopeDistance
                        Case Else
                            MsgBox result
                    End Select
                    result = VB_TMC_GetCoordinate(100, COORDINATE, 1)

                    Select Case result
                        Case 0
                            Text4.Text = COORDINATE.dN
                            Text5.Text = COORDINATE.dE
                            Text6.Text = COORDINATE.dH
                        Case Else
                            MsgBox result
                    End Select
                    VB_AUS_SetUserLockState (1)

                Case 8710
                    MsgBox "没有搜索到目标", "提示"
                Case 8711
                     MsgBox "搜索到多目标", "提示"
                Case 8712
                    MsgBox "环境条件不足", "提示"
            End Select
        End If


        VB_TMC_DoMeasure (TMC_CLEAR)
        VB_TMC_SetEdmMode (EMD_SINGLE_STANDARD)
        rc = VB_AUS_SetUserAtrState(OFF) '锁定模式将自动重置
End Sub

Private Sub Command4_Click()
bind
End Sub

Private Sub DataGrid1_Click()
  PN = DataGrid1.Columns(0).Text

        Text1.Text = PN

        If DataGrid1.Columns(6).Text > "" Then
           ha = DataGrid1.Columns(6).Text
            Text2.Text = ha
        Else
           Text2.Text = ""
       End If
       If DataGrid1.Columns(7).Text > "" Then
         Va = DataGrid1.Columns(7).Text
            Text3.Text = Va
        Else
            Text3.Text = ""
        End If
          If DataGrid1.Columns(2).Text > "" Then
            nc = DataGrid1.Columns(2).Text
            Text4.Text = nc
        Else
            Text4.Text = ""
        End If
        If DataGrid1.Columns(3).Text > "" Then
            Ec = DataGrid1.Columns(3).Text
            Text5.Text = Ec
        Else
            Text5.Text = ""
        End If
        If DataGrid1.Columns(4).Text > "" Then

            Hc = DataGrid1.Columns(4).Text
            Text6.Text = Hc
        Else
            Text6.Text = ""
        End If
        If DataGrid1.Columns(5).Text > "" Then

            ds = DataGrid1.Columns(5).Text
            Text7.Text = ds
        Else
            Text7.Text = ""
        End If
        If DataGrid1.Columns(1).Text > "" Then

            t = DataGrid1.Columns(1).Text
            Text8.Text = t
        Else
            Text8.Text = ""
        End If
        
End Sub

Private Sub Form_Load()
bind
End Sub

Public Sub bind()
Dim str As String
Set Rs = New ADODB.Recordset
str = ConStr
Set Cnn = New ADODB.Connection
Cnn.ConnectionString = str
Cnn.open
Cnn.CursorLocation = adUseClient
Rs.open "select * from 点位学习表", Cnn
Set DataGrid1.DataSource = Rs
End Sub

