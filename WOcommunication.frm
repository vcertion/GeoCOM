VERSION 5.00
Begin VB.Form WOcommunication 
   Caption         =   "ͨѶ����"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   ScaleHeight     =   4350
   ScaleWidth      =   5790
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "˫�Ჹ��"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "�ر�"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "��"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ͨѶ����"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "�������Ӵ���"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "������"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "���ں�"
         Height          =   280
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "WOcommunication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
       Dim ePort As Integer
        Dim eRate As Integer
        Dim nRetries As Integer
        Dim rc As Integer
        
        Select Case Text1.Text
        Case "COM_1"""
         ePort = 0
        Case "COM_2"
        ePort = 1
        Case "COM_3"
        ePort = 2
        Case "COM_4"
        ePort = 3
        End Select
        Select Case Text2.Text
        Case "COM_BAUD_38400"
        eRate = 0
        Case "COM_BAUD_19200"
        eRate = 1
        Case "COM_BAUD_9600"
        eRate = 2
         Case "COM_BAUD_4800"
        eRate = 3
         Case "COM_BAUD_2400"
        eRate = 4
        End Select
        
        nRetries = Text3.Text

        VB_COM_SwitchOnTPS (0)
        
        If RC_OK = VB_COM_Init Then
           MsgBox "��ʼ���ɹ�"
            rc = VB_COM_OpenConnection(ePort, eRate, nRetries)
            Select Case rc
                Case "0"
                   MsgBox "���ӳɹ�"
                Case "3103"
                    MsgBox "�˿�����ʹ�û򲻴���"
                Case "3105"
                   MsgBox "GeoCOMδ�ܼ�⵽TPS"
                Case "2"
                   MsgBox "�Ƿ��Ĳ���"
            End Select
            If Option1.Value = True Then
                VB_TMC_SetInclineSwitch (1)
            ElseIf Option2.Value = True Then
                VB_TMC_SetInclineSwitch (0)
            End If

        End If


End Sub


Private Sub Command2_Click()
Unload Me

End Sub
