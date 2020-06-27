VERSION 5.00
Begin VB.Form WOstationsetting 
   Caption         =   "测站信息设置"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   5715
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton c 
      Caption         =   "设置"
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "棱镜"
      Height          =   1095
      Left            =   240
      TabIndex        =   21
      Top             =   4800
      Width           =   5175
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   1320
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "棱镜参数"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "大气改正"
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   5175
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   3480
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   1440
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   3480
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1440
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "湿度"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "大气压"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "温度"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "EDM发射器波长"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "仪器参数"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "后方交会"
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1320
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "仪器高"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "H"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "E"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "N"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "测站名称"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "WOstationsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public str As String
Private Sub c_Click()
Set conn = New ADODB.Connection
str = ConStr
        conn.ConnectionString = str
        Dim Insert As String
        Dim NullStation As TMC_STATION
        Dim rc1 As Integer
        Dim rc2 As Integer
        Dim rc3 As Integer
        Dim AtmCoor As TMC_ATMOS_TEMPERATURE
        Dim PrismCorr As Double
        
        station = Text1.Text
        n = Text2.Text
        Es = Text3.Text
        Ho = Text4.Text
        Hi = Text5.Text
        Time = Now
conn.open
On Error GoTo err1:
conn.Execute "insert into 测站信息表 (测站,N坐标,E坐标,H坐标,测站高,开始时间) values ('" & station & "','" & n & "','" & Es & "','" & Ho & "','" & Hi & "','" & Time & "')"
MsgBox "测站建立成功"
Exit Sub
err1:
MsgBox "测站建立失败，请检查数据库"
Resume Next
conn.Close
        '建站
        NullStation.dN0 = Text2.Text
        NullStation.dH0 = Text3.Text
        NullStation.dE0 = Text4.Text
        rc1 = VB_TMC_SetStation(NullStation)
        
     '大气改正
        AtmCoor.dLambda = Text6.Text
        AtmCoor.dPressure = Text7.Text
        AtmCoor.dDryTemperature = Text9.Text
        AtmCoor.dWetTemperature = Text8.Text
        rc2 = VB_TMC_SetAtmCorr(AtmCoor)
        '棱镜参数
         PrismCorr = Text9.Text
        rc3 = VB_TMC_SetPrismCorr(PrismCorr)

        If rc1 <> 0 Then
           MsgBox "测站建立失败"
        ElseIf rc2 <> 0 Then
           MsgBox "大气改正失败"
        ElseIf rc3 <> 0 Then
            MsgBox "棱镜参数失败"
        End If
        If rc1 = 0 And rc2 = 0 And rc3 = 0 Then
            MsgBox "测站建立成功"
        End If


conn.Close

End Sub

Private Sub Command1_Click()
'后方交会
Dim x(3) As Double
Dim y(3) As Double
Dim ha(3) As Double
Dim hv(3) As Double
Dim H(3) As Double
Dim d(3) As Double
Dim i As Integer
Dim v As Double
Dim j As Double


i = 0

Set sqlcon = New ADODB.Connection
Set sqlre = New ADODB.Recordset
sqlcon.open ConStr
sqlre.open "select * from 点位学习表 where 类型=3", sqlcon
If Not (sqlre Is Nothing) Then
If Not (sqlre.RecordCount = 0) Then
Do While Not (sqlre.EOF)
   x(i) = sqlre.Fields("N坐标")
   y(i) = sqlre.Fields("E坐标")
   ha(i) = sqlre.Fields("水平角")
   hv(i) = sqlre.Fields("竖直角")
   d(i) = sqlre.Fields("距离")
   H(i) = sqlre.Fields("H坐标")
   sqlre.MoveNext
   i = i + 1
   
   Loop
   HJ x(), y(), ha()
   sh = GC(d(), hv(), j, v, H())
   Text4.Text = sh
   Else
   MsgBox "无数据"
   End If
   End If
   sqlre.Close
   sqlcon.Close
   detaX = n - xp_2
   detaY = Es - yp_2
   detah = Ho - sh
   MsgBox "N值差:" & detaX & "E值差:" & detaY & "H值差:" & detah

End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub HJ(x() As Double, y() As Double, ha() As Double)
        Dim i As Integer
        Dim na As Integer
        na = 0
        Dim nb As Integer
        nb = 1
        Dim nc As Integer
        nc = 2
        Dim ctg As Double
        ctg = 0
        Dim pi_2 As Double
        pi_2 = 1.5707963267949
        Dim p As Double
        p = 0
        Dim xp As Double
        xp = 0
        Dim yp As Double
        yp = 0
        Dim sp As Double
        sp = 0

        For i = 0 To 2
            ctg = ((x(nb) - x(na)) * (x(nc) - x(na)) + (y(nb) - y(na)) * (y(nc) - y(na))) / ((x(nb) - x(na)) * (y(nc) - y(na)) - (x(nc) - x(na)) * (y(nb) - y(na)))
            p = 1# / (ctg + Tan(pi_2 + ha(i)))
            xp = x(i) * p + xp
            yp = y(i) * p + yp
            sp = sp + p
            na = na + 1
            If na > 2 Then
                na = 0
            End If
            nb = nb + 1
            If nb > 2 Then
                nb = 0
            End If
         
            nc = nc + 1
            If nc > 2 Then
                nc = 0
            End If
        Next

        xp_2 = xp / sp


        yp_2 = yp / sp
        Text2.Text = xp_2
        Text3.Text = yp_2


    End Sub
   '高程
Private Function GC(d() As Double, hv() As Double, i As Double, v As Double, H() As Double)
Dim sh As Double
Dim sh_1 As Double
sh_1 = 0
Dim sh_2 As Double
sh_2 = 0
Dim c As Double
sh = d(0) * Sin(hv(0)) + i - v
Dim j As Integer
For j = 0 To 2
sh_1 = H(j) - (d(j) * Sin(hv(j)) + i - v) + sh_1
'c = (sh - sh_1) / (d(j) * d(j) * Cos(hv(j)) * Cos(hv(j)))
'sh_2 = sh_1 * (1 + c) + sh_2
j = j + 1
Next j
GC = sh_1 / 3
End Function


