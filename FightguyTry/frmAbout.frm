VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Fighting Game"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2475
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "About this silly little program - :)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4035
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
frmAbout.Hide
End Sub
