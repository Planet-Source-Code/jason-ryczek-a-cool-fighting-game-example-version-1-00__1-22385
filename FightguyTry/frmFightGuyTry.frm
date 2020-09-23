VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Fighting Game"
   ClientHeight    =   3750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFightGuyTry.frx":0000
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4320
      Top             =   1980
   End
   Begin VB.Label lblDied 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Died"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   675
      Left            =   1500
      TabIndex        =   4
      Top             =   2460
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4620
      TabIndex        =   3
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Jason Ryczek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label lblPause 
      BackStyle       =   0  'Transparent
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   675
      Left            =   2040
      TabIndex        =   0
      Top             =   1980
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GameInProgress As Boolean
Dim GuyGames As Integer, CompGames As Integer
Dim guyX As Long, guyY As Long
Dim enemyX As Long, enemyY As Long
Dim EnemyPos As Long
Dim EnemyLife As Long
Dim DCenemyLifeBar As Long
Dim DCenemyMask As Long
Dim DCenemySprite As Long
Dim FightPos As Long
Dim guyLife As Long
Dim Movement As Boolean
Dim DCguyMask As Long
Dim DCguySprite As Long
Dim DCguyLifeBar As Long
Const Tile_Horiz = 30
Const Tile_Vert = 40
Const BackHeight = 250
Const BackWidth = 400

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        FightPos = 4
    Case vbKeyDown
        FightPos = 3
    Case vbKeyShift
        If FightPos = 3 Then
            FightPos = 5
        Else
            FightPos = 2
        End If
    Case vbKeySpace
        If FightPos = 3 Then
            FightPos = 6
        Else
            FightPos = 1
        End If
    Case vbKeyLeft
        guyX = guyX - 5
    Case vbKeyRight
        guyX = guyX + 5
    Case vbKeyUp And vbKeyLeft
        FightPos = 4
        guyX = guyX - 5
    Case vbKeyUp And vbKeyRight
        FightPos = 4
        guyX = guyX + 5
    Case vbKeyDown And vbKeyLeft
        FightPos = 3
        guyX = guyX - 5
    Case vbKeyDown And vbKeyRight
        FightPos = 3
        guyX = guyX + 5
    Case vbKeyP
        If Timer1.Enabled = True Then
            lblPause.Visible = True
            Timer1.Enabled = False
        Else
            lblPause.Visible = False
            Timer1.Enabled = True
        End If
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp
        FightPos = 0
    Case vbKeyDown
        FightPos = 0
    Case vbKeyShift
        FightPos = 0
    Case vbKeySpace
        FightPos = 0
End Select
End Sub

Private Sub Form_Load()
GuyGames = 0: CompGames = 0
GameInProgress = False
DCenemyLifeBar = GenerateDC(App.Path & "\images\enLifeBar.bmp")
DCenemyMask = GenerateDC(App.Path & "\images\enmask.bmp")
DCenemySprite = GenerateDC(App.Path & "\images\ensprite.bmp")
DCguyLifeBar = GenerateDC(App.Path & "\images\LifeBar.bmp")
DCguyMask = GenerateDC(App.Path & "\images\mask.bmp")
DCguySprite = GenerateDC(App.Path & "\images\sprite.bmp")
FightPos = 0
EnemyPos = 0
enemyX = Me.ScaleWidth - 90
enemyY = Me.ScaleHeight - Tile_Vert
guyX = 50
guyY = Me.ScaleHeight - Tile_Vert
guyLife = 100
EnemyLife = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteGeneratedDC DCenemyLifeBar
DeleteGeneratedDC DCenemyMask
DeleteGeneratedDC DCenemySprite
DeleteGeneratedDC DCguyLifeBar
DeleteGeneratedDC DCguyMask
DeleteGeneratedDC DCguySprite
End Sub

Private Sub lblDied_Click()
New_Game
End Sub

Private Sub mnuAbout_Click()
Timer1.Enabled = False
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHelp_Click()
Timer1.Enabled = False
frmHelp.Show
End Sub

Private Sub mnuNew_Click()
New_Game
End Sub

Private Sub mnuPause_Click()
If GameInProgress = True Then
    If Timer1.Enabled = True Then
        lblPause.Visible = True
        Timer1.Enabled = False
    Else
        lblPause.Visible = False
        Timer1.Enabled = True
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim move_time As Integer
'-Drawing the Fighters
Me.Cls
FightingEnemy EnemyPos, Me, DCenemyMask, vbSrcAnd
FightingEnemy EnemyPos, Me, DCenemySprite, vbSrcPaint
EnemyAI_Fighting guyX, FightPos
FightingGuy FightPos, Me, DCguyMask, vbSrcAnd
FightingGuy FightPos, Me, DCguySprite, vbSrcPaint
HitDetection guyX, enemyX, FightPos, EnemyPos
If FightPos = 5 Then FightPos = 3
If EnemyPos = 5 Then EnemyPos = 3
If Movement = True Then
    Movement = False
    FightPos = 0
    EnemyPos = 0
End If
'-Life Bars
BitBlt Me.hdc, Me.ScaleWidth - 110, 10, EnemyLife, 10, DCenemyLifeBar, 0, 0, vbSrcCopy
BitBlt Me.hdc, 10, 10, guyLife, 10, DCguyLifeBar, 0, 0, vbSrcCopy
' show who's ahead with wins
Label2.Caption = "Player - " & GuyGames
Label3.Caption = "Computer - " & CompGames
End Sub

Sub FightingGuy(ByVal Fight_Position_Number As Long, ByVal DestinationObj As Object, ByVal SourceDC As Long, ByVal CopyAndPaint As Long)
Select Case Fight_Position_Number
    Case 0 '-standing still
        Movement = False
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 1 '-Punch
        Movement = True
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 2 '-Kick
        Movement = True
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 3 '-Crouch
        Movement = False
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 4 '-Jump
        Movement = True
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 5 '-low kick
        Movement = False
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 6 '-upper cut
        Movement = True
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case Else '-standing still
        Fight_Position_Number = 0
        Movement = False
        BitBlt DestinationObj.hdc, guyX, guyY, Tile_Horiz, Tile_Vert, SourceDC, Fight_Position_Number * Tile_Horiz, 0, CopyAndPaint
End Select
End Sub

Sub FightingEnemy(ByVal Enemy_Position_Number As Long, ByVal DestinationObj As Object, ByVal SourceDC As Long, ByVal CopyAndPaint As Long)
Select Case Fight_Position_Number
    Case 0 '-standing still
        Movement = False
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 1 '-Punch
        Movement = True
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 2 '-Kick
        Movement = True
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 3 '-Crouch
        Movement = False
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 4 '-Jump
        Movement = True
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 5 '-low kick
        Movement = False
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case 6 '-upper cut
        Movement = True
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
    Case Else '-standing still
        Fight_Position_Number = 0
        Movement = False
        BitBlt DestinationObj.hdc, enemyX, enemyY, Tile_Horiz, Tile_Vert, SourceDC, Enemy_Position_Number * Tile_Horiz, 0, CopyAndPaint
End Select
End Sub

Sub EnemyAI_Fighting(ByVal guy_X As Long, ByVal Fight_Position As Long)
If enemyX > guy_X Then
    enemyX = enemyX - 5
    EnemyPos = 0
ElseIf enemyX < guy_X + 50 Then
    EnemyPos = Round((Rnd * 5) + 1, 1)
    If guy_X + 30 > enemyX Then enemyX = enemyX + 5
    If guy_X + 30 < enemyX Then enemyX = enemyX - 5
End If

'If guy_X > enemyX Then
'    enemyX = enemyX + 5
'ElseIf guy_X + 30 < enemyX Then
'    enemyX = enemyX - 5
'    Select Case Fight_Position
'        Case 0
'            If (guy_X - 20 >= enemyX) And (guy_X <= enemyX) Then
'                EnemyPos = Round(((4 * Rnd) + 2), 1)
'            Else
'                'EnemyPos = Round(((4 * Rnd) + 2), 1)
'            End If
'        Case 1 To 6
'            If (guy_X >= enemyX - 40) And (guy_X <= enemyX + 10) Then
'                enemyX = enemyX + Rnd * 5
'            Else
'                EnemyPos = Round(((5 * Rnd) + 1), 1)
'            End If
'        Case Else
'            EnemyPos = Round(((5 * Rnd) + 1), 1)
'    End Select
'End If
End Sub

Sub HitDetection(ByVal guy_X As Long, ByVal enemy_X As Long, ByVal guy_FP As Long, ByVal enemy_FP As Long)
If (guy_X >= enemy_X - 20) And (guy_X <= enemy_X) Then
    If (guy_FP > 0) And (guy_FP <> 3) Then
        EnemyLife = EnemyLife - Round(((2 * (guy_FP + 2) * Rnd) + 5), 1)
        enemyX = enemyX + 7
    End If
    If (enemy_FP > 0) And (enemy_FP <> 3) Then
        guyX = guyX - 7
        guyLife = guyLife - Round(((2 * (enemy_FP + 2) * Rnd) + 5), 1)
    End If
End If

If EnemyLife <= 0 Then
    lblDied.Visible = True
    lblDied.Caption = "You Won!"
    Timer1.Enabled = False
    GameInProgress = False
    GuyGames = GuyGames + 1
ElseIf guyLife <= 0 Then
    lblDied.Visible = True
    lblDied.Caption = "You Died!"
    Timer1.Enabled = False
    GameInProgress = False
    CompGames = CompGames + 1
End If
End Sub

Sub New_Game()
FightPos = 0
EnemyPos = 0
enemyX = Me.ScaleWidth - 90
enemyY = Me.ScaleHeight - Tile_Vert
guyX = 50
guyY = Me.ScaleHeight - Tile_Vert
guyLife = 100
EnemyLife = 100
Timer1.Enabled = True
lblDied.Visible = False
GameInProgress = True
End Sub
