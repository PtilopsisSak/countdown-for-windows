VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   6585
   ClientTop       =   2205
   ClientWidth     =   11145
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   990
      Left            =   360
      Top             =   6000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   7
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   1455
      Index           =   2
      Left            =   3600
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   1695
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "天"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'窗口透明常数
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2 ''使用此参数，透明度有效，透明颜色无效
Const LWA_COLORKEY = &H1 '指定色变透明'使用此参数，透明度无效，透明颜色有效

Dim lx, bm, ly, targetday, targettime, high, f1, TITLE1, f2, x1
Dim mode As Boolean
Dim c(3), n
Dim co
Dim ori(1 To 8)




Private Sub Form_Load()
Dim strFileName As String
Dim lngHandle As Long
If Dir(App.Path + "\setting.txt") = "" Then
Dim strWrite As String

strFileName = App.Path + "\setting.txt"
lngHandle = FreeFile()
strWrite = "20201217      '格式：如20200106表示2020年1月6日"
strWrite = strWrite & vbCrLf & "255255255     '各个颜色通道如000000000代表黑色"
strWrite = strWrite & vbCrLf & "171500        '格式：如170000表示17时15分0秒"
strWrite = strWrite & vbCrLf & "高考"
strWrite = strWrite & vbCrLf & "'上一行全部读取作为标题"
strWrite = strWrite & vbCrLf & "120           '数字字号大小"
strWrite = strWrite & vbCrLf & "6585          '窗体默认加载x坐标"
strWrite = strWrite & vbCrLf & "2205          '窗体默认加载y坐标"
    Open strFileName For Output As lngHandle
    Print #lngHandle, strWrite
    Close lngHandle




End If
f2 = 30
x1 = 1000

f1 = 120
targetday = "20200106"
targettime = "000000"
co = RGB(255, 255, 255)
    Dim strAll As String '所读取的文本文件的所有内容
    Dim strLine As String '在循环中存放每行的内容
    strFileName = App.Path + "\setting.txt" '获得文件的句柄
    
    lngHandle = FreeFile()
    
    Open strFileName For Input As lngHandle

    Do While Not EOF(lngHandle) '循环直到文件尾
            k = k + 1

        Line Input #lngHandle, strLine '每次读取一行存放在strLine变量中
        ori(k) = strLine
    Select Case k
    Case 1
        targetday = strLine
    Case 2
        co = RGB(Val(Mid(strLine, 1, 3)), Val(Mid(strLine, 4, 3)), Val(Mid(strLine, 7, 3)))
        If Val(Mid(strLine, 4, 3)) > 20 Then
        co1 = RGB(Val(Mid(strLine, 1, 3)), Val(Mid(strLine, 4, 3)) - 5, Val(Mid(strLine, 7, 3)))
        Else
        co1 = RGB(Val(Mid(strLine, 1, 3)), Val(Mid(strLine, 4, 3)) + 5, Val(Mid(strLine, 7, 3)))
        End If
    Case 3
        targettime = strLine
    Case 4
        TITLE1 = strLine
    Case 6
        f1 = Val(Mid(strLine, 1, 5))
        f2 = f1 \ 4
        x1 = f1 * 10
    Case 7
        Form1.Left = Val(Mid(strLine, 1, 5))
    Case 8
        Form1.Top = Val(Mid(strLine, 1, 5))
    End Select
    Loop
    
    Close lngHandle

Dim rtn As Long
Me.BackColor = co1
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, co1, 150, LWA_COLORKEY
'SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 150, LWA_ALPHA
mode = True
Call countdown
Call labelload
Form1.Width = 4 * (f1 * 20 + x1) + 1600
high = f1 * 20 + 400
Form1.Height = high
End Sub


Private Sub labelload()
If c(0) >= 0 Then z = Len(CStr(c(0))) Else z = 1
For i = 0 To 3
Label3(i).ForeColor = co
Label3(i).FontSize = f2
Label3(i).Left = i * (f1 * 20 + x1) + 12 * f1 * z
Label3(i).Top = f1 * 20 - f2 * 20 + (120 - f1) * 4
Label3(i).Width = f2 * 24
Label3(i).Height = f2 * 24
Label3(i).Caption = Mid("天时分秒", i + 1, 1)
Label3(i).FontBold = 1
Label3(i).Font = "微软雅黑"


Next
Label2.ForeColor = co

i = 0
Label4(i).ForeColor = co
Label4(i).FontSize = f1
Label4(i).Left = 0
Label4(i).Top = (120 - f1) * 4
Label4(i).Width = f1 * 20 + x1 + 11 * f1
Label4(i).Height = f1 * 20 + x1
Label4(i).Font = "微软雅黑"

k = 0

For i = 1 To 3
Label4(i).ForeColor = co
Label4(i).FontSize = f1
Label4(i).Left = (i - 1) * (f1 * 20 + x1) + 11 * f1 * z + x1 - Int(300 * f1 / 100)
Label4(i).Top = (120 - f1) * 4
Label4(i).Width = f1 * 20 + x1
Label4(i).Height = f1 * 20 + x1
Label4(i).Font = "微软雅黑"

k = 0


Next

End Sub

Private Sub Label4_DblClick(Index As Integer)
End

End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
lx = X
ly = Y
bm = True
End If
End Sub
Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
MX = X
MY = Y
If bm = True Then
    Form1.Left = Form1.Left + (MX - lx)
    Form1.Top = Form1.Top + (MY - ly)
End If
End If
End Sub
Private Sub Label4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Index = 0 Then
 bm = False
 Dim strFileName As String
Dim lngHandle As Long
Dim strWrite As String

strFileName = App.Path + "\setting.txt"
lngHandle = FreeFile()
strWrite = ori(1)
strWrite = strWrite & vbCrLf & ori(2)
strWrite = strWrite & vbCrLf & ori(3)
strWrite = strWrite & vbCrLf & ori(4)
strWrite = strWrite & vbCrLf & ori(5)
strWrite = strWrite & vbCrLf & ori(6)
strWrite = strWrite & vbCrLf & CStr(Form1.Left) & "          '窗体默认加载x坐标"
strWrite = strWrite & vbCrLf & CStr(Form1.Top) & "          '窗体默认加载y坐标"
    Open strFileName For Output As lngHandle
    Print #lngHandle, strWrite
    Close lngHandle





 
 End If
End Sub



Public Sub countdown()

d = Date
a = Year(d)
b = Month(d)
c1 = Day(d)
tempday = -1
If (a = Val(Mid(targetday, 1, 4)) And b = Val(Mid(targetday, 5, 2)) And c1 = Val(Mid(targetday, 7, 2))) Then
tempday = tempday + 1
Else

Do Until (a = Val(Mid(targetday, 1, 4)) And b = Val(Mid(targetday, 5, 2)) And c1 = Val(Mid(targetday, 7, 2))) Or tempday > 299
a = Year(d)
b = Month(d)
c1 = Day(d)
d = d + 1
tempday = tempday + 1
Loop

End If

sec = Val(Mid(targettime, 5, 2)) - Second(Time)
If sec < 0 Then sec = sec + 60: Mit = -1
Mit = Val(Mid(targettime, 3, 2)) - Minute(Time) + Mit
If Mit < 0 Then Mit = Mit + 60: hor = -1
hor = Val(Mid(targettime, 1, 2)) - Hour(Time) + hor
If hor < 0 Then hor = hor + 24: tempday = tempday - 1
Label2.Caption = "距离" + TITLE1 + "仅有："


c(0) = tempday
c(1) = hor
c(2) = Mit
c(3) = sec

For i = 0 To 3
If c(i) < 10 And i > 0 Then a = "0" + CStr(c(i)) Else a = CStr(c(i))
Label4(i).Caption = a
Next
If tempday > 298 Or tempday < 0 Then Call addup
End Sub

Public Sub addup()

d = Date
a = Year(d)
b = Month(d)
c1 = Day(d)
tempday = -1
If (a = Val(Mid(targetday, 1, 4)) And b = Val(Mid(targetday, 5, 2)) And c1 = Val(Mid(targetday, 7, 2))) Then
tempday = tempday + 1
Else

Do Until (a = Val(Mid(targetday, 1, 4)) And b = Val(Mid(targetday, 5, 2)) And c1 = Val(Mid(targetday, 7, 2))) Or tempday > 99
a = Year(d)
b = Month(d)
c1 = Day(d)
d = d - 1
tempday = tempday + 1
Loop

End If
If tempday > 99 Then Exit Sub

sec = Second(Time) - Val(Mid(targettime, 5, 2))
If sec < 0 Then sec = sec + 60: Mit = -1
Mit = Minute(Time) - Val(Mid(targettime, 3, 2)) + Mit
If Mit < 0 Then Mit = Mit + 60: hor = -1
hor = Hour(Time) - Val(Mid(targettime, 1, 2)) + hor
If hor < 0 Then hor = hor + 24: tempday = tempday - 1
Label2.Caption = TITLE1 + "之后："

c(0) = tempday
c(1) = hor
c(2) = Mit
c(3) = sec

For i = 0 To 3
If c(i) < 10 And i > 0 Then a = "0" + CStr(c(i)) Else a = CStr(c(i))
Label4(i).Caption = a
Next
mode = False
End Sub

Public Sub diffdraw(i As Integer)
If i <> 0 Then
c(i) = c(i) - 1

If c(i) < 0 Then
If i = 1 Then c(i) = c(i) + 24 Else c(i) = c(i) + 60
flag = True
End If
If c(i) < 10 Then a = "0" + CStr(c(i)) Else a = CStr(c(i))
Label4(i).Caption = a
If flag Then Call diffdraw(i - 1)

Else
c(i) = c(i) - 1
a = CStr(c(i))
Label4(i).Caption = a
Call labelload
If c(i) < 0 Then
Call addup
End If
End If

End Sub
Public Sub adddraw(i As Integer)
If i <> 0 Then
c(i) = c(i) + 1

If c(i) > 59 Then
If i = 1 Then c(i) = c(i) - 24 Else c(i) = c(i) - 60
flag = True
End If
If c(i) < 10 Then a = "0" + CStr(c(i)) Else a = CStr(c(i))
Label4(i).Caption = a
If flag Then Call adddraw(i - 1)

Else
c(i) = c(i) + 1
Call labelload
a = CStr(c(i))
Label4(i).Caption = a
End If

End Sub


Private Sub Timer1_Timer()
If mode Then Call diffdraw(3) Else Call adddraw(3)
n = n + 1
If n > 300 Then n = 0: Call countdown
End Sub



