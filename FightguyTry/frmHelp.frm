VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help..."
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3780
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   780
      Picture         =   "frmHelp.frx":0000
      ScaleHeight     =   600
      ScaleWidth      =   3150
      TabIndex        =   0
      Top             =   0
      Width           =   3150
   End
   Begin VB.Label Label3 
      Caption         =   "By Jason Ryczek"
      Height          =   255
      Left            =   3180
      TabIndex        =   4
      Top             =   4020
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   420
      TabIndex        =   3
      Top             =   660
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   $"frmHelp.frx":6302
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
frmHelp.Hide
End Sub
