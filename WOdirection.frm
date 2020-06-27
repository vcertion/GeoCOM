VERSION 5.00
Begin VB.Form WOdirection 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "棱镜高"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "bE"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "bN"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "水平角"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "WOdirection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        Dim Hz As Double
        Hz = 0
        Dim h As TMC_HEIGHT
        'Dim bN As Double
        'Dim bE As Double
        Dim result As Double
        h.dHr = Text4.Text

        If Not Text1.Text = "" Then
            Hz = Text1.Text
            result = VB_TMC_SetOrientation(Hz)
        ElseIf Not Text2.Text = "" And Not Text3.Text = "" Then
            Hz = GetDirection(Text2.Text, Text3.Text)
            result = VB_TMC_SetOrientation(Hz)
        End If

        'VB_TMC_DoMeasure(TMC_CLEAR)

        Select Case result
            Case 0
            Case Else
                MsgBox result
        End Select
        result = VB_TMC_SetHeight(h)
        Select Case result
            Case 0
            Case Else
                MsgBox result
        End Select
        MsgBox "设置成功"
End Sub

Private Sub Command2_Click()
Unload Me

End Sub
 Private Function GetDirection(bE As Double, bN As Double)
        Dim Dx As Double
        Dim Dy As Double
        Dim angle As Double
        Dim angle1 As Double
        Dim stationE As Double
        Dim stationN As Double
        Dim direction1 As Double
        Dim direction2 As Double
        direction1 = bE
        direction2 = bN
        stationE = GetE()
        stationN = GetN()
        Dy = direction1 - stationE
        Dx = direction2 - stationN

        If Dx <> 0 Then
            angle1 = Abs(Dy / Dx)
            angle = Atan(angle1)
        ElseIf Dx = 0 And Dy > 0 Then
            angle = PI / 2
        ElseIf Dx = 0 And Dy < 0 Then
            angle = PI / 2 + PI

        End If

        If Dy = 0 And Dx > 0 Then
            angle = 0
        End If

        If Dy = 0 And Dx < 0 Then
            angle = PI
        End If

        If Dx > 0 And Dy > 0 Then
            angle = angle
        End If

        If Dx < 0 And Dy > 0 Then
            angle = PI - angle
        End If

        If Dx < 0 And Dy < 0 Then
            angle = angle + PI
        End If

        If Dx > 0 And Dy < 0 Then
            angle = PI + PI - angle
        End If
        GetDirection = angle
    End Function



