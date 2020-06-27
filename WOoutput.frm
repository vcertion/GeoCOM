VERSION 5.00
Begin VB.Form WOoutput 
   Caption         =   "数据导出"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6225
   LinkTopic       =   "Form2"
   ScaleHeight     =   2790
   ScaleWidth      =   6225
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "C:\Users\Administrator\Desktop"
      Top             =   960
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "WOoutput.frx":0000
      Left            =   1320
      List            =   "WOoutput.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "输出位置"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "输出表"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "WOoutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim xlsApp As Object
    Dim Cnn As New ADODB.Connection
    Dim Rs As ADODB.Recordset
     
    Cnn.ConnectionString = ConStr
    If Cnn.State <> ADODB.ObjectStateEnum.adStateClosed Then Cnn.Close
    Cnn.open
     
    Set Rs = New ADODB.Recordset
    With Rs
        Set .ActiveConnection = Cnn
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .open "SELECT * FROM " & Combo1.Text
    End With
    If Rs.EOF Then Exit Sub
    Set xlsApp = CreateObject("Excel.Application")
'    xlsApp.Visible = True
    xlsApp.Workbooks.Add
    xlsApp.Sheets("sheet1").Select
    xlsApp.ActiveSheet.Range("A1").CopyFromRecordset Rs
 
 
    If xlsApp.ActiveWorkbook.Saved = False Then
        xlsApp.ActiveWorkbook.SaveAs Text1.Text & "\" & Combo1.Text & ".xls"
    End If
    xlsApp.Quit
     
    Rs.Close
    Set Rs = Nothing
    Set xlsApp = Nothing
MsgBox "导出成功"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


