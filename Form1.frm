VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H007F5F09&
   BorderStyle     =   0  'None
   Caption         =   "Borderless..."
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press escape to end."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Borderless skinnable form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skinned borderless rezizable form example."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Image rightup 
      Height          =   225
      Left            =   1680
      MousePointer    =   6  'Size NE SW
      Picture         =   "Form1.frx":0000
      Top             =   4680
      Width           =   225
   End
   Begin VB.Image rightdown 
      Height          =   225
      Left            =   1920
      MousePointer    =   8  'Size NW SE
      Picture         =   "Form1.frx":0314
      Top             =   4680
      Width           =   225
   End
   Begin VB.Image leftdown 
      Height          =   225
      Left            =   2160
      MousePointer    =   6  'Size NE SW
      Picture         =   "Form1.frx":0628
      Top             =   4680
      Width           =   225
   End
   Begin VB.Image leftup 
      Height          =   225
      Left            =   1440
      MousePointer    =   8  'Size NW SE
      Picture         =   "Form1.frx":093C
      Top             =   4680
      Width           =   225
   End
   Begin VB.Image right 
      Height          =   60
      Left            =   1080
      MousePointer    =   9  'Size W E
      Picture         =   "Form1.frx":0C50
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   225
   End
   Begin VB.Image leftpic 
      Height          =   60
      Left            =   600
      MousePointer    =   9  'Size W E
      Picture         =   "Form1.frx":0D54
      Stretch         =   -1  'True
      Top             =   4635
      Width           =   225
   End
   Begin VB.Image down 
      Height          =   225
      Left            =   960
      MousePointer    =   7  'Size N S
      Picture         =   "Form1.frx":0E58
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   75
   End
   Begin VB.Image uppe 
      Height          =   225
      Left            =   840
      MousePointer    =   7  'Size N S
      Picture         =   "Form1.frx":0F8C
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub down_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF006, 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then End
End Sub

Private Sub Form_Load()
DoEvents
Me.Show


End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012, 0 'Move the form
End Sub

Private Sub Form_Resize()
On Error Resume Next
Label2.Top = 5
Label2.Left = 15
leftup.Top = 0
leftup.Left = 0
uppe.Top = 0
uppe.Left = 15
uppe.Width = Me.ScaleWidth - 30
rightup.Left = Me.ScaleWidth - 15
rightup.Top = 0
right.Left = Me.ScaleWidth - 15
right.Top = 15
right.Height = Me.ScaleHeight - 30
rightdown.Left = Me.ScaleWidth - 15
rightdown.Top = Me.ScaleHeight - 15
down.Top = Me.ScaleHeight - 15
down.Left = 15
down.Width = Me.ScaleWidth - 30
leftdown.Left = 0
leftdown.Top = Me.ScaleHeight - 15
leftpic.Left = 0
leftpic.Top = 15
leftpic.Height = Me.ScaleHeight - 30
ShapeTheForm Me

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012, 0
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012, 0
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012, 0
End Sub

Private Sub leftdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF007, 0
End Sub

Private Sub leftpic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF001, 0
End Sub

Private Sub leftup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF004, 0
End Sub

Private Sub right_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF002, 0
End Sub

Private Sub rightdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF008, 0 'THIS is importent SIZE_SE is nowhere in winapi
End Sub

Private Sub rightdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.ScaleWidth < 300 Then Me.ScaleWidth = 300
If Me.ScaleHeight < 300 Then Me.ScaleHeight = 300
ShapeTheForm Me
End Sub

Private Sub rightup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF005, 0
End Sub

Private Sub uppe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF003, 0 'Move the form
End Sub

