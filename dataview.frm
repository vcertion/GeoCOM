VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form dataview 
   Caption         =   "������ͼ"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6390
   ScaleWidth      =   10680
   Begin VB.Frame frame2 
      Caption         =   "����Ԥ��"
      Height          =   6015
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   7815
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5655
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9975
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
   Begin VB.Frame frame1 
      Caption         =   "�����б�"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   10186
         _Version        =   393217
         HideSelection   =   0   'False
         Style           =   4
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "dataview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim str As String
Set Rs = New ADODB.Recordset
str = ConStr
Set Cnn = New ADODB.Connection
Cnn.ConnectionString = str
Cnn.open
Rs.open "select * from ��λѧϰ��", Cnn, adOpenStatic, adLockReadOnly
Set DataGrid1.DataSource = Rs


Dim nodeinex As Node

Set nodeinex = TreeView1.Nodes.Add(, , "��1", "����")
nodeinex.Expanded = True
Set nodeinex = TreeView1.Nodes.Add("��1", tvwChild, "��1", "��վ����")

nodeinex.Sorted = True
Set nodeinex = TreeView1.Nodes.Add("��1", tvwChild, "��2", "��λ����")
nodeinex.Sorted = True
Set nodeinex = TreeView1.Nodes.Add(, , "��2", "ʵʱ���")
nodeinex.Expanded = True
For i = 1 To QS - 1
Set nodeinex = TreeView1.Nodes.Add("��2", tvwChild, "��" & i + 2, "��" & i & "�ڲ���ԭʼ����")
nodeinex.Sorted = True
Next i
For i = 1 To QS - 1
Set nodeinex = TreeView1.Nodes.Add("��2", tvwChild, "��" & i + 4, "��" & i & "�ڲ������")
nodeinex.Sorted = True
Next i
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim str As String

If TreeView1.SelectedItem.Text = "��վ����" Then
  str = "��վ��Ϣ��"
  bind str
End If
If TreeView1.SelectedItem.Text = "��λ����" Then
  str = "��λѧϰ��"
  bind str
End If
If TreeView1.SelectedItem.Text = "��1�ڲ���ԭʼ����" Then
  str = "��1��ԭʼ�۲����ݱ�"
  bind str
End If
If TreeView1.SelectedItem.Text = "��2�ڲ���ԭʼ����" Then
  str = "��2��ԭʼ�۲����ݱ�"
  bind str
End If

If TreeView1.SelectedItem.Text = "��3�ڲ���ԭʼ����" Then
  str = "��3��ԭʼ�۲����ݱ�"
  bind str
End If
If TreeView1.SelectedItem.Text = "��1�ڲ������" Then
  str = "��1�ڲ�����ݱ�"
  bind str
End If
If TreeView1.SelectedItem.Text = "��2�ڲ������" Then
  str = "��2�ڲ�����ݱ�"
  bind str
End If
If TreeView1.SelectedItem.Text = "��3�ڲ������" Then
  str = "��3�ڲ�����ݱ�"
  bind str
End If
End Sub

Private Sub bind(s As String)
Dim str As String
Set Rs = New ADODB.Recordset
str = ConStr
Set Cnn = New ADODB.Connection
Cnn.ConnectionString = str
Cnn.open
Cnn.CursorLocation = adUseClient
Rs.open "select * from " & s, Cnn
Set DataGrid1.DataSource = Rs
End Sub

