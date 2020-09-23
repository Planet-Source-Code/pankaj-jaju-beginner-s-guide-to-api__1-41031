VERSION 5.00
Begin VB.Form frmAPI 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Idiot's Guide To Windows API"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAPI 
      Appearance      =   0  'Flat
      Caption         =   "Miscll. API Examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1785
      Width           =   3990
   End
   Begin VB.CommandButton CmdAPI 
      Appearance      =   0  'Flat
      Caption         =   "System API Examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   1245
      Width           =   3990
   End
   Begin VB.CommandButton CmdAPI 
      Appearance      =   0  'Flat
      Caption         =   "String API Examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   705
      Width           =   3990
   End
   Begin VB.CommandButton CmdAPI 
      Appearance      =   0  'Flat
      Caption         =   "Drawing API Examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   3990
   End
End
Attribute VB_Name = "frmAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer:- Pankaj Jaju
'E-mail:- pankaj_jaju@rediffmail.com

Private Sub CmdAPI_Click(Index As Integer)
    Select Case Index
        Case 0
            frmDrawing.Show
        Case 1
            frmString.Show
        Case 2
            frmSystem.Show
        Case 3
            frmMiscll.Show
    End Select
End Sub
