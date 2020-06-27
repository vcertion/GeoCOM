VERSION 5.00
Begin VB.Form WOstart 
   Caption         =   "start"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6855
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   6855
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "12:00:00"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "开始时间"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "WOstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time_1 As String
Private Sub Command1_Click()

time_1 = Text1.Text
If Not IsDate(time_1) Then
MsgBox "时间格式错误"
Else
time_1 = CDate(time_1)
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub init(i As Integer)
        Dim Hz As Double
        Dim ty As Integer
        Dim v As Double
        Dim n As Double
        Dim e As Double
        Dim h As Double
        Dim name As String
        Set myconn = New ADODB.Connection
        Set Rs = New ADODB.Recordset
        myconn.open ConStr
       Rs.open "select * from 点位学习表", myconn
        If Not (Rs Is Nothing) Then
        If Not (Rs.RecordCount = 0) Then
         
              Do While Not (Rs.EOF)
                name = Rs.Fields("点名")
                ty = Rs.Fields("类型")
                n = Rs.Fields("N坐标")
                e = Rs.Fields("E坐标")
                h = Rs.Fields("H坐标")
                Hz = Rs.Fields("水平角")
                v = Rs.Fields("竖直角")
               start name, n, e, h, Hz, v, i, ty
               Rs.MoveNext
       Loop
       Else
       MsgBox "无数据"
       End If
       End If
        Rs.Close
        myconn.Close
    End Sub
Public Sub start(name As String, n As Double, e As Double, h As Double, Hz As Double, v As Double, i As Integer, ty As Integer)
        Dim ds As Double
        ds = 0
        Dim hv As Double
        hv = 0
        Dim ha As Double
        ha = 0
        Dim N_2 As Double
        N_2 = 0
        Dim E_2 As Double
        E_2 = 0
        Dim H_2 As Double
        H_2 = 0
        Dim p As Integer
        p = i
        Dim OnOff As Long
        Dim dSearchHz As Double
        dSearchHz = 0.08
        Dim dSearchV As Double
        dsearchhv = 0.08
        Dim POSMode As Long
        POSMode = AUT_PRECISE
        Dim ATRMode As Long
        Dim result As Integer
        Dim rc As Integer
        Dim hrc As Integer
        Dim TryAgain As Boolean
        Dim Timeoutpar As AUT_TIMEOUT
        Dim OnlyAngle As TMC_HZ_V_ANG
        Dim COORDINATE As TMC_COORDINATE
        trtagain = True

        hrc = VB_COM_SetTimeOut(wait_1)


        hrc = VB_AUT_SetATRStatus(0)

        While TryAgain Or i <> 0
            rc = VB_AUT_MakePositioning4(Hz, v, POSMode, AUT_TARGET, False)
            Select Case (rc)
                Case 8704
                    '定位一个或两个轴时超时。
                    'hrc = VB_AUT_SetTimeout(Timeoutpar)
                Case 0
                    result = VB_AUT_GetATRStatus(ATRState) 'ATR模式必须设置为微调模式
                 If ATRState = 1 Then
                        result = VB_AUT_FineAdjust3(dSearchHz, dSearchV, False)
                        Select Case result
                            Case 0
                               
                                result = VB_AUT_SetLockStatus(1)
                                result = VB_AUT_LockIn()
                                result = VB_TMC_SetEdmMode(EDM_CONT_EXACT)
                                result = VB_TMC_DoMeasure(TMC_DEF_DIST, 0)
                                result = VB_TMC_GetSimpleMea(300, OnlyAngle, SlopeDistance, 0)
                                result = VB_TMC_GetCoordinate1(100, COORDINATE, 1)
                                result = VB_AUT_SetLockStatus(0)
                                Select Case result
                                    Case 0
                                        hv = OnlyAngle.dV + hv
                                        ha = OnlyAngle.dHz + ha
                                        ds = SlopeDistance + ds
                                        N_2 = COORDINATE.dN + N_2
                                        E_2 = COORDINATE.dE + E_2
                                        H_2 = COORDINATE.dH + H_2
                                        i = i - 1
                                        If ChangeF Then
                                            result = VB_AUT_ChangeFace4(1, 0, False)
                                            Select Case result
                                                Case 0
                                                Case Else
                                                    MsgBox result
                                            End Select
                                        End If
                                    Case Else
                                        MsgBox result
                                End Select


                            Case 8710
                                MsgBox "没有搜索到目标"
                            Case 8711
                                MsgBox "搜索到多目标"
                            Case 8712
                                MsgBox "环境条件不足"
                        End Select
                    End If
                Case 3077
                    '通讯连接超时
                    TryAgain = False
                    MsgBox rc
                Case GRC_AUTANGLE_ERROR
                    '角度测量误差:
                    '倾斜校正关
                    hrc = VB_TMC_SetInclineSwitch(OFF)
                Case Else
                    '无法精确定位
                    TryAgain = False
            End Select
       Wend
        hrc = VB_AUT_SetATRStatus(1)
        Insert n, e, h, i, N_2, E_2, H_2, p, name, ty, hv, ha, ds
    End Sub

    Private Sub Insert(n As Double, e As Double, h As Double, i As Integer, N_2 As Double, E_2 As Double, H_2 As Double, i_2 As Integer, name As String, ty As Integer, hv As Double, ha As Double, ds As Double)
        Dim str As String
        N_2 = N_2 \ i
        E_2 = E_2 \ i
        H_2 = H_2 \ i
        hv = hv \ i
        ha = ha \ i
        ds = ds \ i
        dN = n - N_2
        dE = e - E_2
        dH = h - H_2
         myconn.Execute "insert into 第" & QS & "期原始观测数据表 (点名,类型,时间,距离,水平角,竖直角,N坐标,E坐标,H坐标,N坐标差,E坐标差,H坐标差)  values ('" & name & "','" & ty & "','" & Now & "','" & ds & "','" & ha & "','" & hv & "','" & N_2 & "','" & E_2 & "','" & H_2 & "','" & dN & "','" & dE & "','" & dH & "')"
    End Sub



Private Sub Timer1_Timer()
If time_1 = Time Then
    Set mycon = New ADODB.Connection
     mycon.open ConStr
    mycon.Execute "CREATE TABLE 第" & QS & "期原始观测数据表(点名 VARCHAR(10) PRIMARY KEY,类型 INT,N坐标 FLOAT,E坐标 FLOAT,H坐标 FLOAT,N坐标差 FLOAT,E坐标差 FLOAT,H坐标差 FLOAT,时间 VARCHAR(20),距离 FLOAT,水平角 VARCHAR(MAX),竖直角 VARCHAR(MAX))"
    mycon.Execute "CREATE TABLE 第" & QS & "期差分数据表(点名 VARCHAR(10) PRIMARY KEY,类型 INT,N坐标 FLOAT,E坐标 FLOAT,H坐标 FLOAT,N坐标差 FLOAT,E坐标差 FLOAT,H坐标差 FLOAT,距离 FLOAT,水平角 VARCHAR(MAX),竖直角 VARCHAR(MAX))"
    init (measure_time)
    QS = QS + 1
     MsgBox "已开始"
     Exit Sub
  End If
End Sub
