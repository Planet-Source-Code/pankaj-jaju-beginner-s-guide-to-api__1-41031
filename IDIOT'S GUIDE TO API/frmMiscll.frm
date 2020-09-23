VERSION 5.00
Begin VB.Form frmMiscll 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Miscll. APIs"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCursor 
      Appearance      =   0  'Flat
      Caption         =   "Hide Cursor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   675
      Width           =   1365
   End
   Begin VB.CommandButton CmdCursor 
      Appearance      =   0  'Flat
      Caption         =   "Swap MouseButtons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   1695
      Width           =   1365
   End
   Begin VB.CommandButton CmdCursor 
      Appearance      =   0  'Flat
      Caption         =   "Set Cursor Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   1185
      Width           =   1365
   End
   Begin VB.CommandButton CmdEnable 
      Appearance      =   0  'Flat
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   4
      Top             =   2295
      Width           =   1365
   End
   Begin VB.CommandButton CmdFlash 
      Appearance      =   0  'Flat
      Caption         =   "Flash Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   5
      Top             =   2805
      Width           =   1365
   End
   Begin VB.CommandButton CmdCaption 
      Appearance      =   0  'Flat
      Caption         =   "Set Caption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   7
      Top             =   3825
      Width           =   1365
   End
   Begin VB.CommandButton CmdMoveWin 
      Appearance      =   0  'Flat
      Caption         =   "Move Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   6
      Top             =   3315
      Width           =   1365
   End
   Begin VB.CommandButton CmdCursor 
      Appearance      =   0  'Flat
      Caption         =   "CursorPos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   1365
   End
   Begin VB.TextBox TxtCaret 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   135
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2850
   End
   Begin VB.CommandButton CmdCaret 
      Appearance      =   0  'Flat
      Caption         =   "Hide Caret"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   16
      Top             =   3885
      Width           =   1365
   End
   Begin VB.CommandButton CmdShell 
      Appearance      =   0  'Flat
      Caption         =   "ShellAbout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   11
      Top             =   1560
      Width           =   1365
   End
   Begin VB.CommandButton CmdMsgBox 
      Appearance      =   0  'Flat
      Caption         =   "MessageBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   9
      Top             =   630
      Width           =   1365
   End
   Begin VB.CommandButton CmdTickCnt 
      Appearance      =   0  'Flat
      Caption         =   "TickCount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   13
      Top             =   2490
      Width           =   1365
   End
   Begin VB.CommandButton CmdSleep 
      Appearance      =   0  'Flat
      Caption         =   "Sleep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   12
      Top             =   2025
      Width           =   1365
   End
   Begin VB.CommandButton CmdMulDiv 
      Appearance      =   0  'Flat
      Caption         =   "Multiply Divide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   10
      Top             =   1095
      Width           =   1365
   End
   Begin VB.CommandButton CmdDrawIcon 
      Appearance      =   0  'Flat
      Caption         =   "Draw Icon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   15
      Top             =   3420
      Width           =   1365
   End
   Begin VB.CommandButton CmdTxtOut 
      Appearance      =   0  'Flat
      Caption         =   "Text Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   14
      Top             =   2955
      Width           =   1365
   End
   Begin VB.CommandButton CmdBeep 
      Appearance      =   0  'Flat
      Caption         =   "Beep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1620
      TabIndex        =   8
      Top             =   165
      Width           =   1365
   End
End
Attribute VB_Name = "frmMiscll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer:- Pankaj Jaju
'E-mail:- pankaj_jaju@rediffmail.com

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'USED IN GetCursorPos
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Sub CmdBeep_Click()
    Beep 500, 250
End Sub

Private Sub CmdCaption_Click()
    Dim StrCaption As String
    StrCaption = InputBox$("Enter A Caption", "Check Out The Window Caption", "Pankaj Jaju")
    Call SetWindowText(Me.hwnd, StrCaption)
    'Other Example:-
    'Call SetWindowText(CmdCaption.hwnd, StrCaption)
End Sub

Private Sub CmdCaret_Click()
'Caret Is The Flashing Line Thing In TextBox
    If CmdCaret.Caption = "Hide Caret" Then
        MsgBox "The Focus Is Now Set In The TextBox" & vbCrLf & _
               "Now Type Anything In The TextBox," & vbCrLf & _
               "Check That Its Caret(Flashing Line Thing) Is Gone", , "Hide Caret"
        TxtCaret.SetFocus
        HideCaret TxtCaret.hwnd
        CmdCaret.Caption = "Show Caret"
    Else
        MsgBox "Check That Its Caret(Flashing Line Thing) Is Now Showing", , "Show Caret"
        TxtCaret.SetFocus
        ShowCaret TxtCaret.hwnd
        CmdCaret.Caption = "Hide Caret"
    End If
End Sub

Private Sub CmdCursor_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim P As POINTAPI
            MsgBox "Place The Cursor At Any Place On The Screen", , "GetCursorPos"
            GetCursorPos P
            MsgBox "The Cursor Is At " & P.x & "," & P.y, , "GetCursorPos"
        Case 1
            If CmdCursor(1).Caption = "Show Cursor" Then
                ShowCursor True
                CmdCursor(1).Caption = "Hide Cursor"
            Else
                ShowCursor False
                CmdCursor(1).Caption = "Show Cursor"
            End If
        Case 2
            MsgBox "The Cursor Postion Will Be Set To (0,0)", , "Set CursorPos"
            Call SetCursorPos(0, 0)
        Case 3
            If CmdCursor(3).Caption = "Swap MouseButtons" Then
                Call SwapMouseButton(True)
                CmdCursor(3).Caption = "Normalize MouseButtons"
            Else
                Call SwapMouseButton(False)
                CmdCursor(3).Caption = "Swap MouseButtons"
            End If
    End Select
End Sub

Private Sub CmdDrawIcon_Click()
    Me.Cls
    DrawIcon Me.hdc, 85, 290, CLng(Me.Icon)
End Sub

Private Sub CmdEnable_Click()
    Dim Ctr As Integer
    If CmdEnable.Caption = "Enable" Then
        For Ctr = 0 To 3
            Call EnableWindow(CmdCursor(Ctr).hwnd, True)
        Next
        CmdEnable.Caption = "Disable"
    Else
        For Ctr = 0 To 3
            Call EnableWindow(CmdCursor(Ctr).hwnd, False)
        Next
        CmdEnable.Caption = "Enable"
    End If
End Sub

Private Sub CmdFlash_Click()
    Call FlashWindow(Me.hwnd, True)
End Sub

Private Sub CmdMoveWin_Click()
    If Me.Left = 0 Then
        MoveWindow Me.hwnd, 294, 105, 212, 392, True 'True Is For Repainting The Form
    Else
        MoveWindow Me.hwnd, 0, 0, 212, 392, True
    End If
End Sub

Private Sub CmdMsgBox_Click()
    Dim ButtonClick As Long
    ButtonClick = MessageBox(Me.hwnd, "Don't U Think API Is Cooooooool !!!", "API Created MsgBox", 0)
    'In Place Of 0 U Can Also Use The Following
    '1 = OK CANCEL
    '2 = ABORT RETRY IGNORE
    '3 = YES NO CANCEL
    '4 = YES NO
    '5 = RETRY CANCEL
    '6 = CANCEL TRYAGAIN CONTINUE
End Sub

Private Sub CmdMulDiv_Click()
    Dim MD As Long
    'If Using Only VB-Power The Following Will Creates An Overflow
    'MD = 123456789 * 987 / 654321
    
    'Using The Same With MulDiv API It Can Overcome That Overflow Problem
    MD = MulDiv(123456789, 987, 654321)
    MsgBox "The Result Is " & MD, , "Math Using API"
End Sub

Private Sub CmdShell_Click()
    ShellAbout Me.hwnd, "Idiot's Guide To Windows API", "Created By Pankaj Jaju" & vbCrLf & "This Article Is Meant For Beginners", CLng(Me.Icon)
End Sub

Private Sub CmdSleep_Click()
    Sleep 1000
    MsgBox "I Slept For Just A Second", , "Sleep API"
End Sub

Private Sub CmdTickCnt_Click()
    Dim Ret As Long, Hr As Integer, Mn As Integer
    Ret = GetTickCount() 'Returns No. Of Milliseconds Passed Since The System Was Started
    Hr = CInt(Ret / 3600000)
    Mn = CInt(Ret / 60000) - (Hr * 60)
    MsgBox "Windows Has Been Running For " & Hr & " Hour And " & Mn & " Minutes", , "Window Run Time"
    
    'Other Example:-
    'Dim StartTime As Long, FinishTime As Long, Ctr As Integer
    'StartTime = GetTickCount()
    'For Ctr = 0 To Screen.FontCount - 1
    '    Combo1.AddItem Screen.Fonts(Ctr)
    'Next
    'FinishTime = GetTickCount()
    'MsgBox "It Took " & (FinishTime - StartTime) & " MilliSeconds To Load All The Fonts", , "Font Load"
End Sub

Private Sub CmdTxtOut_Click()
    Me.Cls 'Clear The Form
    'U Can Use Me.ForeColor=vbYellow Instead Of SetTextColor API
    SetTextColor Me.hdc, CLng(vbYellow)
    Me.FontBold = True
    TextOut Me.hdc, 17, 300, "Code Provided By Pankaj Jaju", Len("Code Provided By Pankaj Jaju")
End Sub

Private Sub Form_Load()
    frmAPI.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAPI.Visible = True
End Sub
