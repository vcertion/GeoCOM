VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   Caption         =   "�Զ����μ��ϵͳ"
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
         Name            =   "����"
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
      Caption         =   "���̹���"
      Begin VB.Menu new 
         Caption         =   "�½�����..."
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "�򿪹���..."
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "����..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu c 
      Caption         =   "��������"
      Begin VB.Menu Communication 
         Caption         =   "ͨѶ����..."
      End
      Begin VB.Menu d 
         Caption         =   "����..."
      End
      Begin VB.Menu s 
         Caption         =   "��վ��Ϣ..."
      End
      Begin VB.Menu x 
         Caption         =   "�޲�����..."
      End
   End
   Begin VB.Menu au 
      Caption         =   "�Զ�����"
      Begin VB.Menu MB 
         Caption         =   "Ŀ������..."
      End
      Begin VB.Menu a 
         Caption         =   "�Զ�������������..."
      End
      Begin VB.Menu v 
         Caption         =   "�Զ�����"
      End
   End
   Begin VB.Menu date 
      Caption         =   "����"
      Begin VB.Menu CF 
         Caption         =   "���ݲ��"
      End
      Begin VB.Menu data 
         Caption         =   "����Ԥ��..."
      End
      Begin VB.Menu deta 
         Caption         =   "�������"
      End
      Begin VB.Menu output 
         Caption         =   "���ݵ���..."
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
Rs.open "select*from ��λѧϰ�� where type=0", mycon
i = Rs.RecordCount
j = 0
If Not (Rs Is Nothing) Then
 If (Rs.BOF And Rs.EOF) Then
        MsgBox "������"
        Else
         
              Do While Not (Rs.EOF)
              hv(j) = Rs.Fields("��ֱ��")
              ha(j) = Rs.Fields("ˮƽ��")
              ds(j) = Rs.Fields("����")
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
Rs.open "select*from ��" & QS & "��ԭʼ�۲����ݱ� where type=0", mycon
If Not (Rs Is Nothing) Then
 If (Rs.BOF And Rs.EOF) Then
        MsgBox "������"
        Else
         
              Do While Not (Rs.EOF)
              hv_2(j) = Rs.Fields("��ֱ��")
              ha_2(j) = Rs.Fields("ˮƽ��")
              ds_2(j) = Rs.Fields("����")
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
Rs.open "select*from ��" & QS & "��ԭʼ�۲����ݱ� where type=1", mycon
If Not (Rs Is Nothing) Then
 If (Rs.BOF And Rs.EOF) Then
        MsgBox "������"
        Else
         
              Do While Not (Rs.EOF)
              name = Rs.Fields("����")
              hv(j) = Rs.Fields("��ֱ��") * (1 + c(0))
              ha(j) = Rs.Fields("ˮƽ��") * (1 + c(1))
              ds(j) = Rs.Fields("����") * (1 + c(2))
              n = ds(i) * Cos(hv(i)) * Cos(ha(i))
              e = ds(i) * Cos(hv(i)) * Sin(ha(i))
              h = ds(i) * Sin(hv(i))
              mycon.Execute "insert into ��" & QS & "�ڲ�����ݱ�(����,��ֱ��,ˮƽ�ǣ�����,N����,E����,H����) values('" & name & "','" & hv(j) & "','" & ha(j) & "','" & ds(j) & "','" & n & "','" & e & "','" & h & "')"
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
       Rs.open "select * from ��" & QS & "��ԭʼ�۲����ݱ�", myconn
       i = Rs.RecordCount
       For j = 0 To i
       t(j) = j + 1
       Next
       j = 0
        If Not (Rs Is Nothing) Then
        If (Rs.BOF And Rs.EOF) Then
        MsgBox "������"
        Else
         
              Do While Not (Rs.EOF)
              x(j) = Rs.Fields("N�����") * 1000
              y(j) = Rs.Fields("E�����") * 1000
              h(j) = Rs.Fields("H�����") * 1000
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

CommonDialog1.DialogTitle = "��ѡ��Ҫ�򿪵��ļ���"
CommonDialog1.InitDir = "c:\" 'ȱʡ��·��
CommonDialog1.Filter = "���ݿ��ļ�(*.mdf)|*.mdf|�����ļ�(*.*)|*.*"
CommonDialog1.ShowSave '�����ļ�
filename_select = CommonDialog1.FileName  '�ļ���


filename_select = ExtractionFileName2(filename_select)
QS = 1
Cnn.Execute "create database " + filename_select
Cnn.Close
FileName = filename_select
ConStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & FileName & ";Data Source="
str = ConStr
Cnn.ConnectionString = str
Cnn.open
Cnn.Execute "CREATE TABLE ��λѧϰ��(���� VARCHAR(10) PRIMARY KEY,���� INT,N���� FLOAT,E���� FLOAT,H���� FLOAT,���� FLOAT,ˮƽ�� VARCHAR(MAX),��ֱ�� VARCHAR(MAX))"
Cnn.Execute "CREATE TABLE ��վ��Ϣ��(��վ VARCHAR(10) PRIMARY KEY,���� INT,N���� FLOAT,E���� FLOAT,H���� FLOAT,��վ�� FLOAT,��ʼʱ�� VARCHAR(20))"
MsgBox "���ݿⴴ���ɹ�"
Cnn.Close
Exit Sub


End Sub

Private Sub open_Click()
' ���á�CancelError��Ϊ True
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' ���ñ�־
CommonDialog1.Flags = cdlOFNHideReadOnly
' ���ù�����
CommonDialog1.Filter = "���ݿ��ļ�(*.mdf)|*.mdf|�����ļ�(*.*)|*.*"
' ָ��ȱʡ�Ĺ�����
CommonDialog1.FilterIndex = 2
' ��ʾ���򿪡��Ի���
CommonDialog1.ShowOpen
' ��ʾѡ���ļ�������
MsgBox CommonDialog1.FileName
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
End Sub

Public Function ExtractionFileFormat(ByVal CompletePath As String) As String 'ȫ·����ȡ��չ��
Dim t As Variant
If InStr(1, CompletePath, ".") = 0 Or InStr(1, CompletePath, ":\") = 0 Then Exit Function
t = Split(CompletePath, ".")
If InStr(1, t(UBound(t)), "\") = 0 Then ExtractionFileFormat = t(UBound(t))
End Function



Public Function ExtractionFileName(ByVal CompletePath As String) As String   'ȫ·����ȡ�ļ���(����չ��)
Dim t As Variant
If InStr(1, CompletePath, ":\") = 0 Or Right$(CompletePath, 1) = "\" Then Exit Function
t = Split(CompletePath, "\")
ExtractionFileName = t(UBound(t))
End Function

Public Function ExtractionFileName2(ByVal CompletePath As String) As String 'ȫ·����ȡ�ļ���(������չ��)
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
CommonDialog1.DialogTitle = "��ѡ��Ҫ�򿪵��ļ���"
CommonDialog1.InitDir = "c:\" 'ȱʡ��·��
CommonDialog1.Filter = "���ݿ��ļ�(*.mdf)|*.mdf|�����ļ�(*.*)|*.*"
CommonDialog1.ShowSave '�����ļ�
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
 ma.Execute "p=polyfit(x,y,1);yfit=polyval(p,x);plot(x,y,'r*',x,yfit,'b-');title('�������');xlabel('��N/����');ylabel('t/����');grid on"
 ElseIf i = 1 Then
  ma.Execute "p=polyfit(x,y,1);yfit=polyval(p,x);plot(x,y,'r*',x,yfit,'b-');title('�������');xlabel('��E/����');ylabel('t/����');grid on"
  ElseIf i = 2 Then
   ma.Execute "p=polyfit(x,y,1);yfit=polyval(p,x);plot(x,y,'r*',x,yfit,'b-');title('�������');xlabel('��H/����');ylabel('t/����');grid on"
   End If
End Sub



