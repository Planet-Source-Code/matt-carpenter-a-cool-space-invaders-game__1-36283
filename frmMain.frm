VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6720
      Top             =   5280
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   960
      TabIndex        =   2
      Top             =   5445
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   360
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   1
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press ESC to end game"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kills: 0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim RF_DELAY As Integer

Dim EBE(100) As Boolean

Dim HEALTH As Integer


Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type Ship
  x As Integer
  y As Integer
  speed As Integer
  Frame As Integer
  turnspeed As Integer
  bulletspeed As Integer
End Type

Dim Kills As Long


Dim Enemies(5) As Ship
Dim You As Ship

Dim ES As Long  'Enemy Ship Face
Dim ES_DC As Long 'Enemy Ship Face DC
Dim ES_M As Long  'Enemy ship Mask
Dim ES_M_DC As Long  'Enemy Ship Mask DC

Dim YS As Long  'Your ships face
Dim YS_DC As Long 'Your Ships Face's DC
Dim YS_M As Long 'Your ship's Mask
Dim YS_M_DC As Long 'Your ship's Mask's DC

Dim YBULLET As Long
Dim YBULLETDC  As Long
Dim EBULLET As Long
Dim EBULLETDC As Long

Dim Black As Long
Dim BlackDC As Long

Dim BackBuffer As Long
Dim BackBufferDC As Long

Dim CD As Long
Dim CDDC As Long

Private Type tStars
  x As Integer
  y As Integer
  color As Long
  speed As Integer
End Type

Dim Stars(100) As tStars

Private Type tBullet
  ENABLED As Boolean
  x As Integer
  y As Integer
End Type
Dim Bullets(100) As tBullet
Dim eBullets(100) As tBullet

Dim Exp(9) As Long
Dim ExpDC(9) As Long

Private Type Explosion
  ENABLED As Boolean
  Frame As Integer
  x As Integer
  y As Integer
  speed As Integer
End Type

Dim Explo(100) As Explosion



Private Sub Form_Load()
'///////////////////// FORM_LOAD //////////////////////////////

Dim sleepval As Integer
sleepval = InputBox("Type a number between 0 and 20 to slow the game down some for faster computers (0 for slower cpu's, 20 for faster cpu's) 0 works fine for my 550 mhz cpu.", "Space Invaders", "0")
MsgBox "Instructions: " & vbCrLf & "Arrow keys to move" & vbCrLf & "Space bar to shoot."



Form1.Show
DoEvents

Dim RFTimeout As Long
RFTimeout = 3


ProgressBar1.Value = Val(100)
HEALTH = 100


LoadScreen 'Load the 'load screen'
LoadImages 'Load Images

'Set some variables
Randomize Timer
For i = 0 To 5
  Enemies(i).x = Rnd * Picture1.ScaleWidth - 15
  Enemies(i).y = 1
  Enemies(i).speed = Rnd * 5
  Enemies(i).turnspeed = 1
  If Enemies(i).speed = 0 Then Enemies(i).speed = 3
Next i

For i = 0 To 100
  Stars(i).x = Rnd * Picture1.ScaleWidth
  Stars(i).y = Rnd * Picture1.ScaleHeight
  c = Int(Rnd * 2)
  If c = 1 Or c = 2 Then
    Stars(i).color = RGB(150, 150, 170)
    Stars(i).speed = 2
  Else
    Stars(i).color = RGB(244, 244, 255)
    Stars(i).speed = 4
  End If
Next i


You.y = Picture1.ScaleHeight - 35
You.x = 200
You.turnspeed = 4
You.bulletspeed = 7




Dim esc As Boolean

Dim x As Integer

'Game Loop''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Do Until esc = True
Sleep sleepval

If GetAsyncKeyState(vbKeyEscape) < 0 Then esc = True
If GetAsyncKeyState(vbKeyLeft) < 0 Then You.x = You.x - You.turnspeed
If GetAsyncKeyState(vbKeyRight) < 0 Then You.x = You.x + You.turnspeed
If GetAsyncKeyState(vbKeyUp) < 0 Then You.y = You.y - You.turnspeed
If GetAsyncKeyState(vbKeyDown) < 0 Then You.y = You.y + You.turnspeed

If You.y < 0 Then You.y = Picture1.ScaleHeight
If You.y > Picture1.ScaleHeight Then You.y = 0
If You.x < 0 Then You.x = Picture1.ScaleWidth
If You.x > Picture1.ScaleWidth Then You.x = 0

If RF_DELAY = RFTimeout Then
RF_DELAY = 0
If GetAsyncKeyState(vbKeySpace) < 0 Then
  '///// FIRE GUN //////
  'Find and initiate a bullet
  For i = 0 To 100
  If Bullets(i).ENABLED = False Then
    'Initate the bullet
    Bullets(i).ENABLED = True
    Bullets(i).x = You.x + 15
    Bullets(i).y = You.y
    Exit For
  End If
  Next i
End If
End If
RF_DELAY = RF_DELAY + 1

  
'Blt black to backbuffer
BitBlt BackBufferDC, 0, 0, 500, 500, BlackDC, 0, 0, vbSrcCopy

'Draw Stars
For i = 0 To 100
  SetPixel BackBufferDC, Stars(i).x, Stars(i).y, Stars(i).color
  Stars(i).y = Stars(i).y + Stars(i).speed
  If Stars(i).y > Picture1.ScaleHeight Then
    Randomize Timer
    Stars(i).x = Rnd * Picture1.ScaleWidth
    Stars(i).y = 0
  End If
Next i


'bitblt your ship to backbuffer
BitBlt BackBufferDC, You.x, You.y, 30, 30, YS_M_DC, 0, 0, vbSrcAnd
BitBlt BackBufferDC, You.x, You.y, 30, 30, YS_DC, 0, 0, vbSrcPaint
BitBlt BackBufferDC, You.x, You.y, 30, 30, YS_DC, 0, 0, vbSrcCopy

'Bitblt Enemies
For i = 0 To 5
  BitBlt BackBufferDC, Enemies(i).x, Enemies(i).y, 30, 30, ES_M_DC, 0, 0, vbSrcAnd
  BitBlt BackBufferDC, Enemies(i).x, Enemies(i).y, 30, 30, ES_DC, 0, 0, vbSrcPaint
  Enemies(i).y = Enemies(i).y + Enemies(i).speed
  If Enemies(i).x > You.x Then
    Enemies(i).x = Enemies(i).x - Enemies(i).turnspeed
  Else
    Enemies(i).x = Enemies(i).x + Enemies(i).turnspeed
  End If
  If Enemies(i).y > Picture1.ScaleHeight Then
    Randomize Timer
    Enemies(i).x = Rnd * Picture1.ScaleWidth - 15
    Enemies(i).y = -30
    Enemies(i).speed = Rnd * 5
    If Enemies(i).speed = 0 Then Enemies(i).speed = 4
  End If
Next i

'Draw & Update your bullets
For i = 0 To 100
If Bullets(i).ENABLED = True Then
  BitBlt BackBufferDC, Bullets(i).x, Bullets(i).y, 2, 2, YBULLETDC, 0, 0, vbSrcCopy
  Bullets(i).y = Bullets(i).y - You.bulletspeed
  If Bullets(i).y < 0 Then
    'Disable bullet
    Bullets(i).ENABLED = False
  End If
  
  'Check for collsion
  For x = 0 To 5
    If Bullets(i).x < Enemies(x).x + 30 And Bullets(i).x > Enemies(x).x Then
      If Bullets(i).y < Enemies(x).y + 30 And Bullets(i).y > Enemies(x).y Then
        'Destroy Enemy Ship & Disable Bullet
        Bullets(i).ENABLED = False
        Kills = Kills + 1
        Label3.Caption = "Kills: " & Kills
        Label3.Refresh
        
        'Add to health
        On Error Resume Next
        HEALTH = HEALTH + 1
        ProgressBar1.Value = Val(HEALTH)
        'Replace with an explotion animation
        For ii = 0 To 100
          If Explo(ii).ENABLED = False Then
            Explo(ii).ENABLED = True
            Explo(ii).Frame = 0
            Explo(ii).speed = Enemies(x).speed
            Explo(ii).x = Enemies(x).x
            Explo(ii).y = Enemies(x).y - 27
          Exit For
          End If
        Next ii
        
        Randomize Timer
        xx = Int(Rnd * Picture1.ScaleWidth)
        Enemies(x).x = xx
        Enemies(x).y = -30
        
      End If
    End If
  Next x
  
End If
Next i

'Initiate Enemy Bullets
Randomize Timer

x = Int(Rnd * 5)
If x = 4 Then 'Initiate Bullet


y = Int(Rnd * 6)
For i = 0 To 100
  If eBullets(i).ENABLED = False Then
    eBullets(i).x = Enemies(y).x + 15
    eBullets(i).y = Enemies(y).y + 30
    eBullets(i).ENABLED = True
    Exit For
  End If
Next i
End If




'Draw and update enemy bullets
For i = 0 To 100
If eBullets(i).ENABLED = True Then
  
  BitBlt BackBufferDC, eBullets(i).x, eBullets(i).y, 2, 2, EBULLETDC, 0, 0, vbSrcCopy
  eBullets(i).y = eBullets(i).y + You.bulletspeed
  If eBullets(i).y > Picture1.ScaleHeight Then
    'Disable enemies bullet
    eBullets(i).ENABLED = False
  End If
  
  'Check for collison with your ship
  If eBullets(i).x < You.x + 30 And eBullets(i).x > You.x Then
    If eBullets(i).y < You.y + 30 And eBullets(i).y > You.y Then
      'Theres a collison with you. Lower your health & disable EB
      eBullets(i).ENABLED = False
      HEALTH = HEALTH - 6
      If HEALTH < 0 Then
      '=========================== GAME OVER SCRIPT HERE ========
      esc = True
      Else
      ProgressBar1.Value = Val(HEALTH)
      End If
    End If
  End If
End If
Next i

'Draw and update explosions
For i = 0 To 100
If Explo(i).ENABLED = True Then

  BitBlt BackBufferDC, Explo(i).x, Explo(i).y, 60, 75, ExpDC(Explo(i).Frame), 0, 0, vbSrcCopy
  Explo(i).Frame = Explo(i).Frame + 1
  Explo(i).y = Explo(i).y + Explo(i).speed
  If Explo(i).Frame = 10 Then
    'Disable explotion
    Explo(i).ENABLED = False
  End If
End If
Next i

  
'Do some stuff to make the game better (upgrates, etc...)
If Kills = 40 Then
  Kills = 41
  You.bulletspeed = 12
End If

If Kills = 80 Then
  Kills = 81
  RFTimeout = 2
  RF_DELAY = 0
End If

If Kills = 100 Then
  Kills = 101
  RFTimeout = 1
  RF_DELAY = 0
End If

If Kills = 200 Then
  Kills = 201
  You.turnspeed = You.turnspeed + 5
End If





'Bitblt backbuffer to surface
BitBlt Picture1.hdc, 0, 0, 500, 500, BackBufferDC, 0, 0, vbSrcCopy


Picture1.Refresh





Loop

Dim GO As Long
Dim GODC As Long
GO = LoadImage(0, App.Path & "\gameover.bmp", 0, 0, 0, &H10)
GODC = CreateCompatibleDC(Me.hdc)
SelectObject GODC, GO
BitBlt Picture1.hdc, 50, 50, 500, 500, GODC, 0, 0, vbSrcCopy
Picture1.Refresh






'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
Dim p As POINTAPI
GetCursorPos p
t = Me.Top
l = Me.Left
Me.Move (p.x * 15), (p.y * 15)

End If




  
End Sub

Private Sub Label1_Click()
End

End Sub



Private Sub LoadImages()
'/////////////// Load Images /////////////
'Load images here
ES = LoadImage(0, App.Path & "\enemyship.bmp", 0, 0, 0, &H10)
ES_DC = CreateCompatibleDC(Me.hdc)
ES_M = LoadImage(0, App.Path & "\enemyshipmask.bmp", 0, 0, 0, &H10)
ES_M_DC = CreateCompatibleDC(Me.hdc)
SelectObject ES_DC, ES
SelectObject ES_M_DC, ES_M

YS = LoadImage(0, App.Path & "\youship.bmp", 0, 0, 0, &H10)
YS_DC = CreateCompatibleDC(Me.hdc)
YS_M = LoadImage(0, App.Path & "\youshipmask.bmp", 0, 0, 0, &H10)
YS_M_DC = CreateCompatibleDC(Me.hdc)
SelectObject YS_DC, YS
SelectObject YS_M_DC, YS_M

BackBuffer = LoadImage(0, App.Path & "\backbuffer.bmp", 0, 0, 0, &H10)
BackBufferDC = CreateCompatibleDC(Me.hdc)
SelectObject BackBufferDC, BackBuffer

CD = LoadImage(0, App.Path & "\backbuffer.bmp", 0, 0, 0, &H10)
CDDC = CreateCompatibleDC(Me.hdc)
SelectObject CDDC, CD

Black = LoadImage(0, App.Path & "\backbuffer.bmp", 0, 0, 0, &H10)
BlackDC = CreateCompatibleDC(Me.hdc)
SelectObject BlackDC, Black

YBULLET = LoadImage(0, App.Path & "\ybullet.bmp", 0, 0, 0, &H10)
YBULLETDC = CreateCompatibleDC(Me.hdc)
SelectObject YBULLETDC, YBULLET

EBULLET = LoadImage(0, App.Path & "\ebullet.bmp", 0, 0, 0, &H10)
EBULLETDC = CreateCompatibleDC(Me.hdc)
SelectObject EBULLETDC, EBULLET

For i = 0 To 9
Exp(i) = LoadImage(0, App.Path & "\" & i & ".bmp", 0, 0, 0, &H10)
ExpDC(i) = CreateCompatibleDC(Me.hdc)
SelectObject ExpDC(i), Exp(i)
Next i


End Sub





Private Sub LoadScreen()
Dim ls As Long, lsdc As Long

ls = LoadImage(0, App.Path & "\loading.bmp", 0, 0, 0, &H10)
lsdc = CreateCompatibleDC(Me.hdc)
SelectObject lsdc, ls
BitBlt Picture1.hdc, 50, 50, 500, 500, lsdc, 0, 0, vbSrcCopy



End Sub


Private Sub Timer1_Timer()
msg.Caption = ""
Timer1.ENABLED = False

End Sub
