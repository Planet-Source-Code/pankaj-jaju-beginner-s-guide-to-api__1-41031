VERSION 5.00
Begin VB.Form frmSystem 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System APIs"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   1830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Free Space"
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
      Left            =   229
      TabIndex        =   1
      Top             =   660
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Sys Time"
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
      Index           =   10
      Left            =   229
      TabIndex        =   10
      Top             =   4440
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Local Time"
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
      Index           =   9
      Left            =   229
      TabIndex        =   9
      Top             =   4020
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Win Dir"
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
      Index           =   8
      Left            =   229
      TabIndex        =   8
      Top             =   3600
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "VersionInfo"
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
      Left            =   229
      TabIndex        =   7
      Top             =   3180
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "User Name"
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
      Left            =   229
      TabIndex        =   6
      Top             =   2760
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Comp Name"
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
      Left            =   229
      TabIndex        =   0
      Top             =   240
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Drive Type"
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
      Left            =   229
      TabIndex        =   2
      Top             =   1080
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Temp Path"
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
      Left            =   229
      TabIndex        =   5
      Top             =   2340
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "System Dir"
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
      Left            =   229
      TabIndex        =   4
      Top             =   1920
      Width           =   1365
   End
   Begin VB.CommandButton CmdSystem 
      Appearance      =   0  'Flat
      Caption         =   "Short Path"
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
      Left            =   229
      TabIndex        =   3
      Top             =   1500
      Width           =   1365
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer:- Pankaj Jaju
'E-mail:- pankaj_jaju@rediffmail.com

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'USED IN GetVersionEx
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'USED IN GetLocalTime & GetSystemTime
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Sub CmdSystem_Click(Index As Integer)
    'Must Do One Of The Following
    '1)- Make Room To Receive Results i.e Str1=Space(255) --OR--
    '2)- Specify The Lenght Of String i.e Dim Str1 As String * 255
    'Or Else The Function(s) Will Fail And VB-IDE May Crash
    Dim LSTime As SYSTEMTIME
    
    Select Case Index
        Case 0
            Dim CompName As String * 255
            Call GetComputerName(CompName, 255)
            CompName = Left$(CompName, InStr(CompName, vbNullChar)) 'Remove Null Spaces
            MsgBox "The Computer's Name = " & CompName, , "Computer Name"
        Case 1
            Dim FreeSpace As Long, TotSpace As Long
            Dim SecPerClus As Long, BytPerSec As Long
            Dim FreeClus As Long, TotClus As Long
            
            Call GetDiskFreeSpace("C:\", SecPerClus, BytPerSec, FreeClus, TotClus)
            
            FreeSpace = SecPerClus * BytPerSec * (FreeClus / 1024) '/1024 To Convert Bytes To KB
            TotSpace = SecPerClus * BytPerSec * (TotClus / 1024)   'Also To Counter Overflow(if any)
            
            MsgBox "Total Free Space = " & FreeSpace & " kilobytes" & vbCrLf & vbCrLf _
            & "Total Space = " & TotSpace & " kilobytes", , "Free Space In  Drive C:\"
        Case 2
            Dim DriveType As String * 4
            '2=Floppy  3=HardDisk  5=CD-ROM
            DriveType = InputBox$("Enter Any Drive Type" & vbCrLf & "For Ex: C:\", "DriveType", "C:\")
            Select Case GetDriveType(DriveType)
                Case 2
                    MsgBox DriveType & " = Floppy", , "DriveType"
                Case 3
                    MsgBox DriveType & " = HardDisk", , "DriveType"
                Case 5
                    MsgBox DriveType & " = CD-ROM", , "DriveType"
                Case Else
                    'If You Had Some Other Stuff Then You Check It Out Yourself
                End Select
        Case 3
            Dim PathName As String * 255
            Call GetShortPathName(App.Path, PathName, 255)
            MsgBox "The ShortPath Of APP.Path = " & PathName, , "ShortPath Name(DOS Style)"
        Case 4
            Dim SysDir As String * 255
            Call GetSystemDirectory(SysDir, 255)
            MsgBox "The Path Of System Directory Is " & SysDir, , "System Directory Path"
        Case 5
            Dim TempDir As String * 255
            Call GetTempPath(255, TempDir)
            MsgBox "The Path Of Temporary Directory Is " & TempDir, , "Temporary Directory Path"
        Case 6
            Dim UserName As String * 255
            Call GetUserName(UserName, 255)
            MsgBox "The UserName Is " & UserName, , "UserName"
        Case 7
            Dim OS As OSVERSIONINFO
            OS.dwOSVersionInfoSize = Len(OS) 'Must Do This
            Call GetVersionEx(OS)
            MsgBox "The Version Of This OS Is " & OS.dwMajorVersion _
                    & "." & OS.dwMinorVersion & vbCrLf & "Platform ID Is " & OS.dwPlatformId, , "Version Info"
            'OS.dwPlatformId = 0 i.e Win32 on Windows 3.1
            'OS.dwPlatformId = 1 i.e Win32 on Windows 95(If OS.dwMinorVersion = 0)
            '                              OR Windows 98(If OS.dwMinorVersion > 0)
            'OS.dwPlatformId = 2 i.e Win32 on Windows NT/Windows 2000/XP
        Case 8
            Dim WinDir As String * 255
            Call GetWindowsDirectory(WinDir, 255)
            MsgBox "The Path Of Windows Directory Is " & WinDir, , "Windows Directory Path"
        Case 9
            Call GetLocalTime(LSTime)
            MsgBox "Year = " & LSTime.wYear & vbCrLf & _
            "Month = " & LSTime.wMonth & vbCrLf & _
            "Day = " & LSTime.wDay & vbCrLf & _
            "DayOfWeek = " & LSTime.wDayOfWeek & vbCrLf & _
            "Hour = " & LSTime.wHour & vbCrLf & _
            "Minute = " & LSTime.wMinute & vbCrLf & _
            "Seconds = " & LSTime.wSecond & vbCrLf & _
            "MilliSeconds = " & LSTime.wMilliseconds, , "LocalTime"
        Case 10
            Call GetSystemTime(LSTime)
            MsgBox "Year = " & LSTime.wYear & vbCrLf & _
            "Month = " & LSTime.wMonth & vbCrLf & _
            "Day = " & LSTime.wDay & vbCrLf & _
            "DayOfWeek = " & LSTime.wDayOfWeek & vbCrLf & _
            "Hour = " & LSTime.wHour & vbCrLf & _
            "Minute = " & LSTime.wMinute & vbCrLf & _
            "Seconds = " & LSTime.wSecond & vbCrLf & _
            "MilliSeconds = " & LSTime.wMilliseconds, , "SystemTime"
    End Select
End Sub

Private Sub Form_Load()
    frmAPI.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAPI.Visible = True
End Sub
