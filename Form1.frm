VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   Caption         =   "自动变形监测系统"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   10395
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   840
      Top             =   4920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu manige 
      Caption         =   "工程管理"
      Begin VB.Menu new 
         Caption         =   "新建工程..."
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "打开工程..."
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "保存..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu c 
      Caption         =   "参数设置"
      Begin VB.Menu Communication 
         Caption         =   "通讯参数..."
      End
      Begin VB.Menu d 
         Caption         =   "定向..."
      End
      Begin VB.Menu s 
         Caption         =   "测站信息..."
      End
      Begin VB.Menu x 
         Caption         =   "限差设置..."
      End
   End
   Begin VB.Menu au 
      Caption         =   "自动测量"
      Begin VB.Menu MB 
         Caption         =   "目标设置..."
      End
      Begin VB.Menu a 
         Caption         =   "自动测量参数设置..."
      End
      Begin VB.Menu v 
         Caption         =   "自动测量"
      End
   End
   Begin VB.Menu date 
      Caption         =   "数据"
      Begin VB.Menu CF 
         Caption         =   "数据差分"
      End
      Begin VB.Menu data 
         Caption         =   "数据预览..."
      End
      Begin VB.Menu deta 
         Caption         =   "误差曲线"
      End
      Begin VB.Menu output 
         Caption         =   "数据导出..."
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub a_Click()
woautomesure.Show
End Sub

Private Sub CF_Click()
Set mycon = New ADODB.Connection
Set Rs = New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim hv(10) As Double
Dim ds(10) As Double
Dim ha(10) As Double
Dim c1, c2, c3 As Double
mycon.open ConStr
Rs.open "select*from 点位学习表 where type=0", mycon
i = Rs.RecordCount
j = 0
If Not (Rs Is Nothing) Then
 If (Rs.BOF And Rs.EOF) Then
        MsgBox "无数据"
        Else
         
              Do While Not (Rs.EOF)
              hv(j) = Rs.Fields("竖直角")
              ha(j) = Rs.Fields("水平角")
              ds(j) = Rs.Fields("距离")
              j = j + 1
               Rs.MoveNext
               Loop
               End If
               End If
Rs.Close

init hv, ha, ds

 
 mycon.Close
End Sub

Private Sub init(hv() As Double, ha() As Double, ds() As Double)
Set Rs = New ADODB.Recordset
Dim hv_2(10) As Double
Dim ha_2(10) As Double
Dim ds_2(10) As Double
Dim c(2) As Double
Dim j As Integer
j = 0
Rs.open "select*from 第" & QS & "期原始观测数据表 where type=0", mycon
If Not (Rs Is Nothing) Then
 If (Rs.BOF And Rs.EOF) Then
        MsgBox "无数据"
        Else
         
              Do While Not (Rs.EOF)
              hv_2(j) = Rs.Fields("竖直角")
              ha_2(j) = Rs.Fields("水平角")
              ds_2(j) = Rs.Fields("距离")
              j = j + 1
               Rs.MoveNext
               Loop
               End If
               End If
Rs.Close
For i = 0 To 2
c(i) = 0
Next i
For i = 0 To j
c(0) = (hv_2(i) - hv(i)) \ hv_2(i) + c(0)
Next i
For i = 0 To j
c(1) = (ha_2(i) - ha(i)) \ ha_2(i) + c(1)
Next i
For i = 0 To j
c(2) = (ds_2(i) - ds(i)) \ ds_2(i) + c(2)
Next i
c(0) = c(0) \ (j + 1)
c(1) = c(1) \ (j + 1)
c(2) = c(2) \ (j + 1)
start c
End Sub

Private Sub start(c() As Double)
Dim name As String
Dim ha(100) As Double
Dim hv(100) As Double
Dim ds(100) As Double
Dim n  As Double
Dim e As Double
Dim h As Double
Dim i As Integer
Dim j As Integer
j = 0
Set Rs = New ADODB.Recordset
Rs.open "select*from 第" & QS & "期原始观测数据表 where type=1", mycon
If Not (Rs Is Nothing) Then
 If (Rs.BOF And Rs.EOF) Then
        MsgBox "无数据"
        Else
         
              Do While Not (Rs.EOF)
              name = Rs.Fields("点名")
              hv(j) = Rs.Fields("竖直角") * (1 + c(0))
              ha(j) = Rs.Fields("水平角") * (1 + c(1))
              ds(j) = Rs.Fields("距离") * (1 + c(2))
              n = ds(i) * Cos(hv(i)) * Cos(ha(i))
              e = ds(i) * Cos(hv(i)) * Sin(ha(i))
              h = ds(i) * Sin(hv(i))
              mycon.Execute "insert into 第" & QS & "期差分数据表(点名,竖直角,水平角，距离,N坐标,E坐标,H坐标) values('" & name & "','" & hv(j) & "','" & ha(j) & "','" & ds(j) & "','" & n & "','" & e & "','" & h & "')"
              j = j + 1
               Rs.MoveNext
               Loop
               End If
               End If
Rs.Close
End Sub
Private Sub Communication_Click()
WOcommunication.Show
End Sub

Private Sub d_Click()
 WOdirection.Show
End Sub

Private Sub data_Click()
dataview.Show
End Sub

Private Sub deta_Click()
'activeX
 Dim i As Integer
 Dim j As Integer
 Dim t(1000) As Integer
 Dim y(1000) As Double
 Dim x(1000) As Double
 Dim h(1000) As Double
 
        Set myconn = New ADODB.Connection
        Set Rs = New ADODB.Recordset
        myconn.open ConStr
       Rs.open "select * from 第" & QS & "期原始观测数据表", myconn
       i = Rs.RecordCount
       For j = 0 To i
       t(j) = j + 1
       Next
       j = 0
        If Not (Rs Is Nothing) Then
        If (Rs.BOF And Rs.EOF) Then
        MsgBox "无数据"
        Else
         
              Do While Not (Rs.EOF)
              x(j) = Rs.Fields("N坐标差") * 1000
              y(j) = Rs.Fields("E坐标差") * 1000
              h(j) = Rs.Fields("H坐标差") * 1000
              j = j + 1
               Rs.MoveNext
               Loop
    chart x, t, 0
    chart y, t, 1
    chart h, t, 2
        End If
        End If
        Rs.Close
        myconn.Close
  
    

End Sub

Private Sub m_Click()
Form1.Show
End Sub

Private Sub MB_Click()
Form1.Show
End Sub

Private Sub new_Click()
Set Cnn = New ADODB.Connection
Dim str As String
str = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=;Data Source="
Cnn.ConnectionString = str
If Cnn.State <> 1 Then
Cnn.open
End If

CommonDialog1.DialogTitle = "请选择要打开的文件名"
CommonDialog1.InitDir = "c:\" '缺省打开路径
CommonDialog1.Filter = "数据库文件(*.mdf)|*.mdf|所有文件(*.*)|*.*"
CommonDialog1.ShowSave '保存文件
filename_select = CommonDialog1.FileName  '文件名


filename_select = ExtractionFileName2(filename_select)
QS = 1
Cnn.Execute "create database " + filename_select
Cnn.Close
FileName = filename_select
ConStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & FileName & ";Data Source="
str = ConStr
Cnn.ConnectionString = str
Cnn.open
Cnn.Execute "CREATE TABLE 点位学习表(点名 VARCHAR(10) PRIMARY KEY,类型 INT,N坐标 FLOAT,E坐标 FLOAT,H坐标 FLOAT,距离 FLOAT,水平角 VARCHAR(MAX),竖直角 VARCHAR(MAX))"
Cnn.Execute "CREATE TABLE 测站信息表(测站 VARCHAR(10) PRIMARY KEY,类型 INT,N坐标 FLOAT,E坐标 FLOAT,H坐标 FLOAT,测站高 FLOAT,开始时间 VARCHAR(20))"
MsgBox "数据库创建成功"
Cnn.Close
Exit Sub


End Sub

Private Sub open_Click()
' 设置“CancelError”为 True
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' 设置标志
CommonDialog1.Flags = cdlOFNHideReadOnly
' 设置过滤器
CommonDialog1.Filter = "数据库文件(*.mdf)|*.mdf|所有文件(*.*)|*.*"
' 指定缺省的过滤器
CommonDialog1.FilterIndex = 2
' 显示“打开”对话框
CommonDialog1.ShowOpen
' 显示选定文件的名字
MsgBox CommonDialog1.FileName
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub

Public Function ExtractionFileFormat(ByVal CompletePath As String) As String '全路径提取扩展名
Dim t As Variant
If InStr(1, CompletePath, ".") = 0 Or InStr(1, CompletePath, ":\") = 0 Then Exit Function
t = Split(CompletePath, ".")
If InStr(1, t(UBound(t)), "\") = 0 Then ExtractionFileFormat = t(UBound(t))
End Function



Public Function ExtractionFileName(ByVal CompletePath As String) As String   '全路径提取文件名(带扩展名)
Dim t As Variant
If InStr(1, CompletePath, ":\") = 0 Or Right$(CompletePath, 1) = "\" Then Exit Function
t = Split(CompletePath, "\")
ExtractionFileName = t(UBound(t))
End Function

Public Function ExtractionFileName2(ByVal CompletePath As String) As String '全路径提取文件名(不带扩展名)
ExtractionFileName2 = ExtractionFileName(CompletePath)
ExtractionFileName2 = Mid(ExtractionFileName2, 1, Len(ExtractionFileName2) - Len(ExtractionFileFormat(CompletePath)) - 1)
End Function




Private Sub output_Click()
WOoutput.Show
End Sub

Private Sub s_Click()
WOstationsetting.Show
End Sub

Private Sub save_Click()
CommonDialog1.DialogTitle = "请选择要打开的文件名"
CommonDialog1.InitDir = "c:\" '缺省打开路径
CommonDialog1.Filter = "数据库文件(*.mdf)|*.mdf|所有文件(*.*)|*.*"
CommonDialog1.ShowSave '保存文件
End Sub

Private Sub v_Click()
WOstart.Show
End Sub

Private Sub chart(a() As Double, b() As Integer, i As Integer)
Dim ma As Object
Dim c() As Double
 Set ma = CreateObject("Matlab.Application")
 Call ma.putfullmatrix("x", "base", a, c)
 Call ma.putfullmatrix("y", "base", b, c)
 If i = 0 Then
 ma.Execute "p=polyfit(x,y,1);yfit=polyval(p,x);plot(x,y,'r*',x,yfit,'b-');title('误差曲线');xlabel('▲N/毫米');ylabel('t/周期');grid on"
 ElseIf i = 1 Then
  ma.Execute "p=polyfit(x,y,1);yfit=polyval(p,x);plot(x,y,'r*',x,yfit,'b-');title('误差曲线');xlabel('▲E/毫米');ylabel('t/周期');grid on"
  ElseIf i = 2 Then
   ma.Execute "p=polyfit(x,y,1);yfit=polyval(p,x);plot(x,y,'r*',x,yfit,'b-');title('误差曲线');xlabel('▲H/毫米');ylabel('t/周期');grid on"
   End If
End Sub



