VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'大气折射气象改正 d 实测距离 p实测大气压 e实测水汽压 t实测大气温度
Public Function JLQXGZ(d As Double, p As Double, e As Double, t As Double)
 Dim n As Double
 Dim d_2 As Double
 n = 1 + (2.97 / (1 + 0.003661 * t) * p - 0.414 / (1 + 0.003661 * t) * e) * 0.0000000001
 d_2 = d * 1.0029 / n
 JLQXGZ = d_2 - d
End Function
'球气改正 d为实测斜距 k为实测竖直折光系数 R地球半径
Public Function QQCGZ(d As Double)
Dim k As Double
k = 0.14
Dim R As Integer
R = 6371
QQCGZ = (1 - k) / (2 * R) * d * d
End Function
'垂直角的气象改正 d水平距离 R地球半径 k垂直折光系数
Public Function CZQXGZ(d As Double)
Dim R As Double
R = 6371
Dim k As Double
k = 0.14
CZQXGZ = (k * d) / (2 * R)
End Function
'水平角气象改正
Public Function SPJQXGZ()

End Function
