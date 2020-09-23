VERSION 5.00
Begin VB.Form frmDrawing 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drawing APIs"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "Pie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   225
      TabIndex        =   5
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "Chord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   225
      TabIndex        =   2
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "RoundRect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   225
      TabIndex        =   7
      Top             =   3090
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "Ellipse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   225
      TabIndex        =   3
      Top             =   1410
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "Rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   225
      TabIndex        =   6
      Top             =   2670
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "LineTo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   225
      TabIndex        =   4
      Top             =   1830
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "Arc To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   570
      Width           =   1215
   End
   Begin VB.CommandButton CmdDraw 
      Appearance      =   0  'Flat
      Caption         =   "Arc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer:- Pankaj Jaju
'E-mail:- pankaj_jaju@rediffmail.com

'********************************************************************
'Some Important Points To Know
'.hDC = Handle Provided By WinOS To The Device Context Of An Object
'Handle = Unique Identifier(Integer) Given By The WinOS And Used By
'         Various Programs To Identify That Object(For eg:- Form)
'Decice Context = Interface(Link) Between Applications & Devices
'                 (Devices Such as Display Moniter, Printers etc)
'********************************************************************

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'USED IN GetCursorPos
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Sub CmdDraw_Click(Index As Integer)
    'Note:At Form Load
    '   The DrawWidth Of The Form = 5
    '   The ForeColor Of The Form = vbred
    Me.Cls 'Clear The Form Before Next Drawing
    
    Select Case Index
        Case 0
            Arc Me.hdc, 110, 40, 250, 15, 350, 55, 110, 40
        Case 1
            ArcTo Me.hdc, 110, 40, 250, 15, 350, 55, 110, 40
        Case 2
            Chord Me.hdc, 110, 35, 205, 110, 155, 135, 85, 5
        Case 3
            Ellipse Me.hdc, 110, 15, 200, 230 'Ellipse
            Ellipse Me.hdc, 210, 15, 310, 115 'Circle
            Ellipse Me.hdc, 210, 150, 330, 200 'Ellipse
        Case 4
            Dim P As POINTAPI
            '1) Move Current Position To Specified Position.
            '2) Also Returns Previous Position (in type POINTAPI)
            '3) Since The Starting Point Is (0,0),We Have To Move
            '   The Point To Our Desired Position
            MoveToEx Me.hdc, 210, 10, P
            
            'Draws A Line From Current Position To Specified Position
            LineTo Me.hdc, 110, 230
            LineTo Me.hdc, 320, 230
            LineTo Me.hdc, 210, 10
        Case 5
            Pie Me.hdc, 150, 50, 300, 220, 270, 100, 150, 50
        Case 6
            Rectangle Me.hdc, 110, 15, 320, 230
        Case 7
            '80,80 Is Width & Height Of Ellipse To Draw Rounded Corners
            RoundRect Me.hdc, 110, 15, 320, 230, 80, 80
    End Select
End Sub

Private Sub Form_Load()
    Me.AutoRedraw = True 'Not Necessary, But Good Idea While Working With Drawings
    Me.DrawWidth = 5 'So That The Drawing Can Be Thick
    Me.ForeColor = vbRed 'The Color Of Drawing
    frmAPI.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAPI.Visible = True
End Sub
