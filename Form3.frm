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
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "ȡ��"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "��ʼʱ��"
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
MsgBox "ʱ���ʽ����"
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
       Rs.open "select * from ��λѧϰ��", myconn
        If Not (Rs Is Nothing) Then
        If Not (Rs.RecordCount = 0) Then
         
              Do While Not (Rs.EOF)
                name = Rs.Fields("����")
                ty = Rs.Fields("����")
                n = Rs.Fields("N����")
                e = Rs.Fields("E����")
                h = Rs.Fields("H����")
                Hz = Rs.Fields("ˮƽ��")
                v = Rs.Fields("��ֱ��")
               start name, n, e, h, Hz, v, i, ty
               Rs.MoveNext
       Loop
       Else
       MsgBox "������"
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
                    '��λһ����������ʱ��ʱ��
                    'hrc = VB_AUT_SetTimeout(Timeoutpar)
                Case 0
                    result = VB_AUT_GetATRStatus(ATRState) 'ATRģʽ��������Ϊ΢��ģʽ
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
                                MsgBox "û��������Ŀ��"
                            Case 8711
                                MsgBox "��������Ŀ��"
                            Case 8712
                                MsgBox "������������"
                        End Select
                    End If
                Case 3077
                    'ͨѶ���ӳ�ʱ
                    TryAgain = False
                    MsgBox rc
                Case GRC_AUTANGLE_ERROR
                    '�ǶȲ������:
                    '��бУ����
                    hrc = VB_TMC_SetInclineSwitch(OFF)
                Case Else
                    '�޷���ȷ��λ
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
         myconn.Execute "insert into ��" & QS & "��ԭʼ�۲����ݱ� (����,����,ʱ��,����,ˮƽ��,��ֱ��,N����,E����,H����,N�����,E�����,H�����)  values ('" & name & "','" & ty & "','" & Now & "','" & ds & "','" & ha & "','" & hv & "','" & N_2 & "','" & E_2 & "','" & H_2 & "','" & dN & "','" & dE & "','" & dH & "')"
    End Sub



Private Sub Timer1_Timer()
If time_1 = Time Then
    Set mycon = New ADODB.Connection
     mycon.open ConStr
    mycon.Execute "CREATE TABLE ��" & QS & "��ԭʼ�۲����ݱ�(���� VARCHAR(10) PRIMARY KEY,���� INT,N���� FLOAT,E���� FLOAT,H���� FLOAT,N����� FLOAT,E����� FLOAT,H����� FLOAT,ʱ�� VARCHAR(20),���� FLOAT,ˮƽ�� VARCHAR(MAX),��ֱ�� VARCHAR(MAX))"
    mycon.Execute "CREATE TABLE ��" & QS & "�ڲ�����ݱ�(���� VARCHAR(10) PRIMARY KEY,���� INT,N���� FLOAT,E���� FLOAT,H���� FLOAT,N����� FLOAT,E����� FLOAT,H����� FLOAT,���� FLOAT,ˮƽ�� VARCHAR(MAX),��ֱ�� VARCHAR(MAX))"
    init (measure_time)
    QS = QS + 1
     MsgBox "�ѿ�ʼ"
     Exit Sub
  End If
End Sub
