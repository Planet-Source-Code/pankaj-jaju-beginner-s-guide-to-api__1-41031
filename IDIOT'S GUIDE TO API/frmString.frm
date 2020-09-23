VERSION 5.00
Begin VB.Form frmString 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String APIs"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   1740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdStrCase 
      Appearance      =   0  'Flat
      Caption         =   "Lower Case"
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
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrCpy 
      Appearance      =   0  'Flat
      Caption         =   "StrCpy"
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
      Left            =   195
      TabIndex        =   8
      Top             =   3795
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrCmp 
      Appearance      =   0  'Flat
      Caption         =   "StrCmpi"
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
      Left            =   195
      TabIndex        =   7
      Top             =   3300
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrCmp 
      Appearance      =   0  'Flat
      Caption         =   "StrCmp"
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
      Left            =   195
      TabIndex        =   6
      Top             =   2880
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrCase 
      Appearance      =   0  'Flat
      Caption         =   "Upper Case"
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
      Left            =   195
      TabIndex        =   1
      Top             =   630
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrLen 
      Appearance      =   0  'Flat
      Caption         =   "StrLen"
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
      Left            =   195
      TabIndex        =   10
      Top             =   4725
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrCpy 
      Appearance      =   0  'Flat
      Caption         =   "StrCpyn"
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
      Left            =   195
      TabIndex        =   9
      Top             =   4215
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrIs 
      Appearance      =   0  'Flat
      Caption         =   "IsLower"
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
      Left            =   195
      TabIndex        =   4
      Top             =   1965
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrIs 
      Appearance      =   0  'Flat
      Caption         =   "IsUpper"
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
      Left            =   195
      TabIndex        =   5
      Top             =   2385
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrIs 
      Appearance      =   0  'Flat
      Caption         =   "IsAplha"
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
      Left            =   195
      TabIndex        =   2
      Top             =   1125
      Width           =   1350
   End
   Begin VB.CommandButton CmdStrIs 
      Appearance      =   0  'Flat
      Caption         =   "IsAplhaNum"
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
      Left            =   195
      TabIndex        =   3
      Top             =   1545
      Width           =   1350
   End
End
Attribute VB_Name = "frmString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer:- Pankaj Jaju
'E-mail:- pankaj_jaju@rediffmail.com

Private Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Private Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Private Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Sub CmdStrCase_Click(Index As Integer)
    Dim Str1 As String
    Select Case Index
        Case 0
            Str1 = InputBox$(vbCrLf & vbCrLf & "Enter A String" & vbCrLf & vbCrLf & "To See Real Work, Enter In Upper Case Only", "CharLower", "PANKAJ JAJU", 3300, 3500)
            If Str1 <> "" Then 'If OK Is Pressed
                MsgBox "Lower Case Of " & Str1 & " = " & CharLower(CStr(Str1)), , "Lower Case"
            End If
        Case 1
            Str1 = InputBox$(vbCrLf & vbCrLf & "Enter A String" & vbCrLf & vbCrLf & "To See Real Work, Enter In Lower Case Only", "CharUpper", "pankaj jaju", 3300, 3500)
            If Str1 <> "" Then 'If OK Is Pressed
                MsgBox "Upper Case Of " & Str1 & " = " & CharUpper(CStr(Str1)), , "Upper Case"
            End If
    End Select
End Sub

Private Sub CmdStrCmp_Click(Index As Integer)
    Dim Str1 As String, Str2 As String
    Dim Ret As Long
    Str1 = InputBox$("Enter First String", "String Compare", "Pankaj Jaju", 3300, 3500)
    If Str1 = "" Then Exit Sub 'If Cancel Is Pressed
    Str2 = InputBox$("Enter Second String", "String Compare", "PANKAJ JAJU", 3300, 3500)
    If Str2 = "" Then Exit Sub 'If Cancel Is Pressed
    Select Case Index
        Case 0
            Ret = lstrcmp(Str1, Str2) 'Case Sensitive
        Case 1
            Ret = lstrcmpi(Str1, Str2) 'Case InSensitive
    End Select
    If Ret = 0 Then
        MsgBox "Both The Strings Are Same", vbInformation, "String Compare"
    Else
        MsgBox "Strings Are Different", vbCritical, "String Compare"
    End If
End Sub

Private Sub CmdStrCpy_Click(Index As Integer)
    Dim Str1 As String, Str2 As String
    Dim Ret As Long
    Str2 = InputBox$("Enter A String", "String Copy", "Pankaj Jaju", 3300, 3500)
    If Str2 = "" Then Exit Sub 'If Cancel Is Pressed
    
    'Must Do One Of The Following
    '1)- Make Room To Receive Copied String --OR--
    '2)- The Lenght Of Destination String = Length Of Source String
    '    i.e Dim Str1 As String * 25, Str2 As String * 25
    'Or Else The Function(s) Will Fail
    Str1 = Space(Len(Str2))
    
    Select Case Index
        Case 0
            lstrcpy Str1, Str2 'Copy Entire String
            MsgBox "The Copied String Is " & Str1, , "String Copy"
        Case 1
            Ret = InputBox(vbCrLf & "How Many Chars To Copy" & vbCrLf & vbCrLf & "Keep Numbers Less Than The Len Of String", "lstrcpyn", 7, 3300, 3500)
            If Ret < 0 Or Ret > (Len(Str2) + 1) Then Exit Sub
            lstrcpyn Str1, Str2, Ret 'Copy Number Of Chars(1 Less For '\0') As Specified By Ret
            MsgBox "The NCopied String Is " & Str1, , "NString Copy"
    End Select
End Sub

Private Sub CmdStrIs_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            MsgBox "PRESS ANY KEY ON YOUR KEYBOARD TO CHECK IF IT IS 'IsAlpha'" & vbCrLf & "For Eg:- 'A' or 1 or Shift+1", vbInformation, "IsAlpha"
        Case 1
            MsgBox "PRESS ANY KEY ON YOUR KEYBOARD TO CHECK IF IT IS 'IsAlphaNumeric'" & vbCrLf & "For Eg:- 'A' or 1 or Shift+1", vbInformation, "IsAlphaNumeric"
        Case 2
            MsgBox "PRESS ANY CHARACTER KEY ON YOUR KEYBOARD TO CHECK IF IT IS OF 'LOWER CASE'" & vbCrLf & "For Eg:- 'A' or Shift+'A'", vbInformation, "IsLower"
        Case 3
            MsgBox "PRESS ANY CHARACTER KEY ON YOUR KEYBOARD TO CHECK IF IT IS OF 'UPPER CASE'" & vbCrLf & "For Eg:- 'A' or Shift+'A'", vbInformation, "IsUpper"
    End Select
End Sub

Private Sub CmdStrIs_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Ret As Long
    Select Case Index
        Case 0
            Ret = IsCharAlpha(CByte(KeyAscii)) 'A-Z KeyBoard
        Case 1
            Ret = IsCharAlphaNumeric(CByte(KeyAscii)) 'A-Z And 0-9 KeyBoard
        Case 2
            Ret = IsCharLower(CByte(KeyAscii)) 'A-Z KeyBoard
        Case 3
            Ret = IsCharUpper(CByte(KeyAscii)) 'A-Z KeyBoard(Using Shift)
    End Select
    If Ret = 0 Then
        MsgBox "False", vbCritical, "IsChar"
    Else
        MsgBox "True", vbInformation, "IsChar"
    End If
End Sub

Private Sub CmdStrLen_Click()
    Dim Str1 As String
    Str1 = InputBox$("Enter A String", "lstrlen", "Pankaj Jaju", 3300, 3500)
    MsgBox "The Length Of String Is " & lstrlen(Str1), , "String Length"
End Sub

Private Sub Form_Load()
    frmAPI.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAPI.Visible = True
End Sub
