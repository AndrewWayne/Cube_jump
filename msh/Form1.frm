VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11340
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   3000
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7440
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   2265
      Left            =   -1080
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   1200
      Top             =   1800
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   0
      Top             =   3600
      Width           =   15135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vy(1 To 10) As Double
Dim Accy(1 To 10) As Double
Dim Accx(1 To 10) As Double
Dim object(1 To 10) As Shape
Dim Tops(1 To 2) As Double
Dim ground As Shape
Dim num As Integer
Dim k As Boolean
Dim v0 As Double
Dim g0 As Double
Dim pts, score As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 97 And Shape2.Top + Shape2.Height >= Shape1.Top Then
    Accy(1) = g0
    Vy(1) = v0
    Shape2.Top = Shape2.Top - 1
    k = True
Else
    If KeyAscii = 97 And k Then
        Accy(1) = g0
        Vy(1) = v0
        Shape2.Top = Shape2.Top - 1
        k = False
    End If
End If

End Sub

Private Sub Form_Load()
    MsgBox ("按A跳，连按2段跳，躲避黄色方块")
    pts = 0
    k = False
    Randomize
    num = 1
    g0 = 8000000
    Accy(1) = g0
    v0 = -150000
    Tops(1) = 2640
    Tops(2) = 1640
    'Label1.Caption = Str(Tops(1))
End Sub

Private Sub Timer1_Timer()
        If Shape2.Top + Shape2.Height < Shape1.Top Then
            Shape2.Top = Shape2.Top + Vy(1) * Timer1.Interval * 0.001
            Vy(1) = Vy(1) + Accy(1) * Timer1.Interval * 0.001
        Else
            Shape2.Top = Shape1.Top - Shape2.Height
            Vy(1) = 0
            Accy(1) = 0
            k = False
        End If
End Sub

Private Sub Timer2_Timer()
    If Shape3.Left + Shape3.Width > 0 Then
        Shape3.Left = Shape3.Left - 200
    Else
        Shape3.Left = 10000 + Int(Rnd * 5000)
        decision = 1 + Int(Rnd * 2)
        'Print (Int(decision))
        Shape3.Top = Tops(decision)
        score = 1
    End If
End Sub
Function inRect(X As Integer, Y As Integer) As Boolean
    XX1 = Shape3.Left
    XX2 = Shape3.Left + Shape3.Width
    YY1 = Shape3.Top + Shape3.Height
    YY2 = Shape3.Top
    inRect = X > XX1 And X < XX2 And Y > YY2 And Y < YY1
End Function
Private Sub Timer3_Timer()
    Dim X1 As Integer
    Dim Y1 As Integer
    Dim X2 As Integer
    Dim Y2 As Integer
    X1 = Shape2.Left
    Y1 = Shape2.Top + Shape2.Height
    X2 = Shape3.Left + Shape3.Width
    Y2 = Shape3.Top
    If X2 < X1 And Y2 > Y1 Then
        pts = pts + score
        Label1.Caption = Str(pts)
        score = 0
    End If
    X1 = Shape2.Left
    X2 = Shape2.Left + Shape2.Width
    Y1 = Shape2.Top + Shape2.Height
    Y2 = Shape2.Top
    If inRect(X1, Y1) Or inRect(X1, Y2) Or inRect(X2, Y1) Or inRect(X2, Y2) Then
        ans = MsgBox("你已经死了，爬爬爬,是否重来", vbYesNo)
        Print "本次成绩" & pts
        If ans = 6 Then
            Shape3.Left = -10000
            pts = 0
            Label1.Caption = Str(pts)
        Else
            End
        End If
    End If
End Sub
