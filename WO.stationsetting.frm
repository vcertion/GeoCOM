VERSION 5.00
Begin VB.Form WOstationsetting 
   Caption         =   "��վ��Ϣ����"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   5715
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton c 
      Caption         =   "����"
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "�⾵"
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
         Caption         =   "�⾵����"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������"
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
         Caption         =   "ʪ��"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "����ѹ"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "�¶�"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "EDM����������"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "�󷽽���"
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
         Caption         =   "������"
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
         Caption         =   "��վ����"
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
conn.Execute "insert into ��վ��Ϣ�� (��վ,N����,E����,H����,��վ��,��ʼʱ��) values ('" & station & "','" & n & "','" & Es & "','" & Ho & "','" & Hi & "','" & Time & "')"
MsgBox "��վ�����ɹ�"
Exit Sub
err1:
MsgBox "��վ����ʧ�ܣ��������ݿ�"
Resume Next
conn.Close
        '��վ
        NullStation.dN0 = Text2.Text
        NullStation.dH0 = Text3.Text
        NullStation.dE0 = Text4.Text
        rc1 = VB_TMC_SetStation(NullStation)
        
     '��������
        AtmCoor.dLambda = Text6.Text
        AtmCoor.dPressure = Text7.Text
        AtmCoor.dDryTemperature = Text9.Text
        AtmCoor.dWetTemperature = Text8.Text
        rc2 = VB_TMC_SetAtmCorr(AtmCoor)
        '�⾵����
         PrismCorr = Text9.Text
        rc3 = VB_TMC_SetPrismCorr(PrismCorr)

        If rc1 <> 0 Then
           MsgBox "��վ����ʧ��"
        ElseIf rc2 <> 0 Then
           MsgBox "��������ʧ��"
        ElseIf rc3 <> 0 Then
            MsgBox "�⾵����ʧ��"
        End If
        If rc1 = 0 And rc2 = 0 And rc3 = 0 Then
            MsgBox "��վ�����ɹ�"
        End If


conn.Close

End Sub

Private Sub Command1_Click()
'�󷽽���
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
sqlre.open "select * from ��λѧϰ�� where ����=3", sqlcon
If Not (sqlre Is Nothing) Then
If Not (sqlre.RecordCount = 0) Then
Do While Not (sqlre.EOF)
   x(i) = sqlre.Fields("N����")
   y(i) = sqlre.Fields("E����")
   ha(i) = sqlre.Fields("ˮƽ��")
   hv(i) = sqlre.Fields("��ֱ��")
   d(i) = sqlre.Fields("����")
   H(i) = sqlre.Fields("H����")
   sqlre.MoveNext
   i = i + 1
   
   Loop
   HJ x(), y(), ha()
   sh = GC(d(), hv(), j, v, H())
   Text4.Text = sh
   Else
   MsgBox "������"
   End If
   End If
   sqlre.Close
   sqlcon.Close
   detaX = n - xp_2
   detaY = Es - yp_2
   detah = Ho - sh
   MsgBox "Nֵ��:" & detaX & "Eֵ��:" & detaY & "Hֵ��:" & detah

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
   '�߳�
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


