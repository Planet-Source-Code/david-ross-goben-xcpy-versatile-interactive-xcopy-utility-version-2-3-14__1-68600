VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmXcopy 
   Caption         =   "XCPY"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   Icon            =   "frmXcopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4500
      TabIndex        =   44
      ToolTipText     =   "Clear history of recent configurations"
      Top             =   7020
      Width           =   1155
   End
   Begin VB.ComboBox cboRecent 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   43
      ToolTipText     =   "Recent configurations. Select from dropdown list to activate selection"
      Top             =   7020
      Width           =   4155
   End
   Begin VB.CheckBox chkNewOnly 
      Caption         =   "Copy only &New Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   17
      ToolTipText     =   "/N Include only files that do not exist  in the destination path"
      Top             =   3760
      Width           =   2415
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   4800
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   180
      TabIndex        =   40
      Top             =   6060
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBatch 
      Caption         =   "Add Sequence to a BATCH File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   22
      ToolTipText     =   "Add the selected xcopy options to a batch file"
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdShortcut 
      Caption         =   "Create Shortcut on Desktop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   21
      ToolTipText     =   "Create a shortcut on the desktop for the selected operation"
      Top             =   5220
      Width           =   2415
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Command Line Format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   20
      ToolTipText     =   "View shortcut command line template for this setup"
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filepath Parameters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Index           =   2
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox Picture1 
         Height          =   1935
         Left            =   60
         ScaleHeight     =   1875
         ScaleWidth      =   5355
         TabIndex        =   1
         Top             =   240
         Width           =   5415
         Begin VB.CheckBox chkAppend 
            Caption         =   "Append new Log Entries to Log file"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   42
            ToolTipText     =   "/L=+ Append Log entries to Log File instead of over-writing it"
            Top             =   1680
            Width           =   3255
         End
         Begin VB.CommandButton cmdDownload 
            Caption         =   "Download vbZip10.dll from Internet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3120
            TabIndex        =   9
            ToolTipText     =   "Download vbZip10.dll from VB Accelerator Werb Site to enable backup Zipping"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtZIP 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   11
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkZip 
            Caption         =   "Save to ZIP Filename:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   10
            ToolTipText     =   "Checked will include the folder itself, otherwise just its contents are copied"
            Top             =   960
            Width           =   1875
         End
         Begin VB.CheckBox chkIncludeFolder 
            Caption         =   "Include selected path folder in copy"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   5
            ToolTipText     =   "Checked will include the folder itself, otherwise just its contents are copied"
            Top             =   300
            Width           =   3375
         End
         Begin VB.TextBox txtLogfile 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   780
            TabIndex        =   13
            Text            =   "(Path to write log entries)"
            ToolTipText     =   "/L= Location to store  log file. Blank or enbraced=ignore"
            Top             =   1320
            Width           =   4155
         End
         Begin VB.TextBox txtTo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   780
            TabIndex        =   7
            Text            =   "(Folder path to copy into)"
            ToolTipText     =   "/D= Folder path to store the copied data"
            Top             =   615
            Width           =   4155
         End
         Begin VB.TextBox txtFrom 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            TabIndex        =   3
            Text            =   "(Folder path to copy out of)"
            ToolTipText     =   "/S= Source path (with mask) for files/folders to copy"
            Top             =   0
            Width           =   4155
         End
         Begin VB.CommandButton cmdLogfile 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4980
            TabIndex        =   14
            ToolTipText     =   "Browse for Error Log file path"
            Top             =   1320
            Width           =   315
         End
         Begin VB.CommandButton cmdBrowseTo 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4980
            TabIndex        =   8
            ToolTipText     =   "Browse for DESTINATION path"
            Top             =   630
            Width           =   315
         End
         Begin VB.CommandButton cmdBrowseFrom 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4980
            TabIndex        =   4
            ToolTipText     =   "Browse for SOURCE Path"
            Top             =   0
            Width           =   315
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   660
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   60
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Log:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&To:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   6
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&From:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   2
            Top             =   60
            Width           =   420
         End
      End
   End
   Begin VB.CheckBox chkRstArchive 
      Caption         =   "&Reset Archive Flags"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   18
      ToolTipText     =   "/R Reset archive bits on source files"
      Top             =   4040
      Width           =   1995
   End
   Begin VB.CheckBox chkSubdirs 
      Caption         =   "&Include Subdirectories"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   19
      ToolTipText     =   "/I Include subdirectories of source path"
      Top             =   4320
      Width           =   1995
   End
   Begin VB.CheckBox chkModified 
      Caption         =   "Copy only &Modified Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   16
      ToolTipText     =   "/M Copy only modified files"
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Optional Parameters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   3240
      Width           =   2595
   End
   Begin VB.CommandButton cmdHelp 
      Height          =   555
      Left            =   5100
      MousePointer    =   1  'Arrow
      Picture         =   "frmXcopy.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Command Line Invokation Help"
      Top             =   240
      Width           =   555
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   555
      Left            =   180
      TabIndex        =   25
      Top             =   6420
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   979
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   325
      FullHeight      =   37
   End
   Begin VB.CommandButton cmdXcopy 
      Caption         =   "Start Copy"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   23
      ToolTipText     =   "Begin Xcopy process"
      Top             =   5220
      Width           =   1155
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4500
      MousePointer    =   1  'Arrow
      TabIndex        =   24
      ToolTipText     =   "Exit program"
      Top             =   5220
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   2820
      TabIndex        =   30
      Top             =   3240
      Width           =   2895
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StartTime:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elapsed TIme:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lblStartTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblEndTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label lblElapsed 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   960
         Width           =   1275
      End
   End
   Begin VB.Label lblPcent 
      Alignment       =   2  'Center
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5220
      TabIndex        =   41
      Top             =   6120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblFiles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   39
      Top             =   5700
      Width           =   465
   End
   Begin VB.Label lblFolders 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblFolders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   5280
      Width           =   675
   End
   Begin VB.Image imgCNC 
      Height          =   225
      Left            =   4920
      Picture         =   "frmXcopy.frx":2BC4
      Stretch         =   -1  'True
      Top             =   4860
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XCpy (Interactive XCopy Utility)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   37
      Top             =   300
      Width           =   4305
   End
   Begin VB.Label lblErrors 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   5760
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Errors encountered:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   27
      Top             =   5820
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblStatus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   4860
      Width           =   615
   End
End
Attribute VB_Name = "frmXcopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private AviFile As String       'path to AVI file extracted from Resources
Private CmdLineMode As Boolean  'True if running from command line
Private Batch As String         'Batch file data accumulator
Private BatchPath As String     'Batch file location
Private HaveZip As Boolean      'True if we have the vbZip10.DLL
Public m_Zip As cZip            'Zip class object
Private Checking As Boolean     'true if combobox being processed

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Set up XP buttons
'*******************************************************************************
Private Sub Form_Initialize()
  Call FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Start ball rolling
'*******************************************************************************
Private Sub Form_Load()
  Dim Cmd As String, S As String, Fld(2) As String
  Dim I As Integer, F As Integer
  Dim Bol As Boolean
  
  Set fso = New FileSystemObject                  'set up file I/O object
  Set m_Zip = New cZip                            'create Zip resource
  Set LogData = New Collection                    'init log collection
  Me.Picture1.BorderStyle = 0                     'flatten backgound
  VBRemoveMenuItem Me, rmClose                    'disable the X button
  Me.Caption = "XCPY " & GetAppVersion()          'set up form ribbon banner
  
  Me.lblStatus.Caption = vbNullString             'erase status for now
  Me.imgCheck(0).Picture = LoadPicture()          'blank out checkmarks
  Me.imgCheck(1).Picture = LoadPicture()
  Me.imgCheck(2).Picture = LoadPicture()
'
' set up size and control positions
'
  Me.Height = (Me.Height - Me.ScaleHeight) + Me.Animation1.Top + 60
  Me.Animation1.Top = Me.cmdHelp.Top
  Me.Animation1.Left = Me.Frame1(2).Left
  Me.cmdView.Left = Me.lblStatus.Left
  Me.cmdShortcut.Left = Me.lblStatus.Left
  Me.cmdBatch.Left = Me.lblStatus.Left
'
' Set recent history button
'
  Me.cboRecent.Top = Me.ProgressBar1.Top
  Me.cmdClear.Top = Me.cboRecent.Top
'
' download button for downloading the vbZip10.dll
'
  With Me.cmdDownload
    .Height = Me.txtTo.Height
    .Left = Me.txtTo.Left
    .Width = Me.txtTo.Width
    .Top = Me.chkZip.Top
  End With
'
' init the default Zip filename if ZIP processing is supported
'
  S = Format(Now, "Short Date")                   'get the locale date
  F = 0                                           'init format field index (0-2)
  For I = 1 To Len(S)
    If IsNumeric(Mid$(S, I, 1)) Then              'grab just digits
      Fld(F) = Fld(F) & Mid$(S, I, 1)             'grab to current field
    Else
      F = F + 1                                   'bump to next field if not digit (/ or - or whatever)
    End If
  Next I
  '
  'ensure 1-digit fields become 2 with pre-pended "0"
  '
  For F = 0 To 2
    If Len(Fld(F)) = 1 Then Fld(F) = "0" & Fld(F)
  Next F
  '
  'stuff default zip name
  '
  Me.txtZIP.Text = "Zip_" & Fld(0) & Fld(1) & Fld(2) & ".zip"
'
' now check for vbZip10.dll
'
  HaveZip = fso.FileExists(AddSlash(App.Path) & "vbzip10.dll")  'set true if we found it
  Me.cmdDownload.Visible = Not HaveZip            'make download button visible if not found
'
' see if command line parameters provided
'
  Cmd = Trim$(Command$)                           'get any command line parameters
  If CBool(Len(Cmd)) Then                         'anything?
    If Left$(Cmd, 1) <> "/" Then Cmd = "/H"       'default for help if bad info
    If InStr(1, Cmd, "/H", vbTextCompare) Then
      Call HelpInfo
      Unload Me
      Exit Sub
    End If
    
    Call ProcessCmd(Cmd)
    CmdLineMode = True                            'mark command line mode
  End If
'
' load AVI file (File Copy movie) to temp folder location
'
  AviFile = GetTempName("AVI")                    'get a temp file name, based upon "AVI"
  Call LoadResource(101, AviFile)                 'extract AVI from resources to that file
'
' set up recent history
'
  F = CInt(GetSetting(App.Title, "Settings", "HistCnt", "0"))
  For I = 0 To F - 1
    Me.cboRecent.AddItem GetSetting(App.Title, "Settings", "Hist" & CStr(I), vbNullString)
  Next I
  Checking = True
  If CBool(F) Then
    Me.cboRecent.ListIndex = 0
  Else
    Me.cboRecent.ListIndex = -1
  End If
  Checking = False
'
' auto-launch xcopy and close app if command line parameters present
'
  If Me.cmdXcopy.Enabled Then
    If Len(Me.txtFrom.Text) > 0 And Len(Me.txtTo.Text) > 0 Then
      If Left$(Me.txtFrom.Text, 1) <> "(" And Left$(Me.txtTo.Text, 1) <> "(" Then
        Me.Show
        Me.cmdXcopy.Value = True                  'automode launch
        Unload Me                                 'then unload self
        Exit Sub                                  'and close app
      End If
    End If
  ElseIf CmdLineMode Then
'
' check to see if the drives are available in case removable
'
    Bol = False
    S = Trim$(Me.txtFrom.Text)          'check source drivepath
    If CBool(Len(S)) Then
      If Left$(S, 1) <> "(" Then        'ignore if embraced
        If Not IsDriveReady(S) Then
          MsgBox "Source Drive '" & S & "' is not ready. Please check it.", vbOKOnly Or vbExclamation, "Drive Not Ready"
        Else
          Bol = True
        End If
      End If
    End If
    
    If Bol Then
      S = Trim$(Me.txtTo.Text)          'check destination drivepath
      If CBool(Len(S)) Then
        If Left$(S, 1) <> "(" Then      'ignore if embraced
          If Not IsDriveReady(S) Then
            MsgBox "Destination Drive '" & S & "' is not ready. Please check it.", vbOKOnly Or vbExclamation, "Drive Not Ready"
            Bol = False
          End If
        End If
      End If
    End If
    
    If Not Bol Then
      Unload Me                                   'exit if src or dst drives not available
      Exit Sub
    End If
  End If
'
' set up display if not auto-mode
'
  CmdLineMode = False                             'turn off automode if on
  Me.cmdXcopy.Enabled = False                     'else disable XCOPY button
  Me.cmdView.Enabled = False
  Me.cmdShortcut.Enabled = False
  Me.cmdBatch.Enabled = False
  Me.chkIncludeFolder.Enabled = False
  Batch = vbNullString                            'batch data accumulator if building batches
  BatchPath = vbNullString                        'path to batch file
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  Dim S As String

  S = "XCPY " & GetAppVersion()
  Select Case Me.WindowState
    Case vbMinimized
      If Me.lblPcent.Visible Then
        S = "Xcpy [" & Me.lblPcent.Caption & "] from " & Me.txtFrom.Text
      End If
  End Select
  If Me.Caption <> S Then
    Me.Caption = S
    Me.Refresh
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Remove avi file
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim Idx As Integer, Cnt As Integer
  
  On Error Resume Next
  Me.Animation1.Close                       'make sure animation closed
  If Len(AviFile) Then Kill AviFile         'remove animation file from temp folder
  
  With Me.cboRecent                         'save hostory list
    Cnt = .ListCount                        'get count
    If Cnt > 25 Then Cnt = 25               'save only up to 25 entries
    SaveSetting App.Title, "Settings", "HistCnt", CStr(Cnt)
    For Idx = 0 To Cnt - 1
      SaveSetting App.Title, "Settings", "Hist" & CStr(Idx), .List(Idx)
    Next Idx
  End With
  
  Set fso = Nothing                         'remove resource objects
  Set m_Zip = Nothing
  Set LogData = Nothing                     'and collection
End Sub

'*******************************************************************************
' Function Name     : StripQuotes
' Purpose           : Strip quotes from a string
'*******************************************************************************
Private Function StripQuotes(InString As String) As String
  Dim S As String
  Dim I As Integer
  
  S = Trim$(InString)                           'grab text
  I = InStr(1, S, Chr$(34))                     'find a quote
  Do While I
    If I = 1 Then
      S = Mid$(S, 2)
    ElseIf I = Len(S) Then
      S = Left$(S, I - 1)
    Else
      S = Left$(S, I - 1) & Mid$(S, I + 1)
    End If
    I = InStr(1, S, Chr$(34))                   'find all quoted strings
  Loop
  StripQuotes = S
End Function

'*******************************************************************************
' Subroutine Name   : cmdHelp_Click
' Purpose           : Display Help
'*******************************************************************************
Private Sub cmdHelp_Click()
  Call HelpInfo
End Sub

'*******************************************************************************
' Subroutine Name   : HelpInfo
' Purpose           : Provide Help for using the program
'*******************************************************************************
Private Sub HelpInfo()
  CenterMsgBoxOnForm Me, _
        App.Title & " " & GetAppVersion() & " " & App.LegalCopyright & vbCrLf & vbCrLf & _
         " Command Line Options:" & vbCrLf & vbCrLf & _
         "    /H[elp]                        This help message" & vbCrLf & _
         "    /M[odifiedOnly]           Copy only modified files" & vbCrLf & _
         "    /N[ewOnly]                 Copy only new or newer files" & vbCrLf & _
         "    /R[esetArchive]          Reset Archive bit on source files" & vbCrLf & _
         "    /I[ncludeSubdirs]        Include Subdirectories" & vbCrLf & _
         "    /S[rc]=srcpath           Source path" & vbCrLf & _
         "    /D[st]=dstpath           Destination path" & vbCrLf & _
         "    /Z[ip]=Zipfile               Name of ZIP file to send data to" & vbCrLf & _
         "    /L[ogFile]=LogData    Output log file path" & vbCrLf & _
         "    /L[ogFile]=+LogData Append new log entries to log file", _
         vbOKOnly Or vbInformation, "XCPY Help"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBrowseFrom_Click
' Purpose           : Browse for Source Path
'*******************************************************************************
Private Sub cmdBrowseFrom_Click()
  Dim S As String
  
  S = DirBrowser(Me.hwnd, ViewDirsOnly, "Select Source Path")
  If CBool(Len(S)) Then Me.txtFrom.Text = S
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBrowseTo_Click
' Purpose           : Browser for Destination Path
'*******************************************************************************
Private Sub cmdBrowseTo_Click()
  Dim S As String
  
  S = Trim$(Me.txtFrom.Text)
  If Left$(S, 1) = "(" Then S = vbNullString
  S = DirBrowser(Me.hwnd, ViewDirsOnly, "Select Desintation Path", S)
  If CBool(Len(S)) Then Me.txtTo.Text = S
End Sub

'*******************************************************************************
' Subroutine Name   : cmdLogFile_Click
' Purpose           : Browser for Log Path
'*******************************************************************************
Private Sub cmdLogFile_Click()
  LogFile = SelectFile()
  If CBool(Len(LogFile)) Then Me.txtLogfile.Text = LogFile
End Sub

'*******************************************************************************
' Subroutine Name   : cmdQuit_Click
' Purpose           : Check for Quit or Stop
'*******************************************************************************
Private Sub cmdQuit_Click()
  If Me.cmdQuit.Caption = "Close" Then
    Unload Me
  Else
    CopyCancel = 1  'set flag non-zero
    Cancel = True   'was STOP label
    Me.cmdXcopy.Enabled = False
  Screen.MousePointer = vbHourglass
  DoEvents
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdXcopy_Click
' Purpose           : Perform XCOPY
'*******************************************************************************
Private Sub cmdXcopy_Click()
  Dim Idx As Long
  Dim S As String, T As String
  Dim Bol As Boolean
  
  If Me.cmdXcopy.Caption = "Pause" Then
    Me.cmdXcopy.Caption = "Continue"
    Me.cmdXcopy.ToolTipText = "Continue processing"
    Pausing = True
    Exit Sub
  ElseIf Me.cmdXcopy.Caption = "Continue" Then
    Me.cmdXcopy.Caption = "Pause"
    Me.cmdXcopy.ToolTipText = "Pause Processing"
    Pausing = False
    Exit Sub
  End If
  T = GetCmdData()
  S = GetCmdLine()
  AddItemToHistory Mid$(S, Len(T) + 2)    'add to history list
  
'  Me.cmdXcopy.Enabled = False             'disable XCOPY button and all associated
  Me.cmdView.Visible = False
  Me.cmdShortcut.Visible = False
  Me.cmdBatch.Visible = False
  Me.chkRstArchive.Enabled = False
  Me.chkSubdirs.Enabled = False
  Me.chkModified.Enabled = False
  Me.chkNewOnly.Enabled = False
  Me.txtFrom.Enabled = False
  Me.chkIncludeFolder.Enabled = False
  Me.txtTo.Enabled = False
  Me.txtLogfile.Enabled = False
  Me.cmdBrowseFrom.Enabled = False
  Me.cmdBrowseTo.Enabled = False
  Me.cmdLogfile.Enabled = False
  Me.chkZip.Enabled = False
  Me.chkAppend.Enabled = False
  Me.txtZIP.Enabled = False
  Me.txtZIP.BackColor = &H8000000F
  Me.cboRecent.Visible = False
  Me.cmdClear.Visible = False
  
  Me.lblFolders.Caption = vbNullString
  Me.lblFiles.Caption = vbNullString
  Me.lblErrors.Caption = "0"              'init no errors
  Me.cmdQuit.Caption = "STOP"             'change Quit to STOP
  Me.cmdQuit.ToolTipText = "Stop current operation"
  Me.cmdXcopy.Caption = "Pause"
  Me.cmdXcopy.ToolTipText = "Pause Processing"
  Me.cmdXcopy.Enabled = True
  Cancel = False                          'init cancel to no
  
  With Me.Animation1                      'start animation
    .AutoPlay = True
    .Open AviFile
    .Visible = True
    DoEvents
  '
  ' copy data
  '
    S = Me.txtFrom.Text                                 'source path
    If Me.chkIncludeFolder.Value = vbChecked Then S = AddSlash(S)
    Idx = Xcopy(S, Me.txtTo.Text, Me.chkSubdirs.Value, Me.chkModified.Value, Me.chkNewOnly.Value, True)
    frmXcopy.lblStatus.Caption = vbNullString           'remove status text
    .Close                                              'close animation
    .AutoPlay = False
    .Visible = False
  End With
  
  If Not CmdLineMode Then                               'if not command line mode...
    If Idx > 0 And CmdLineMode = False Then CenterMsgBoxOnForm Me, CStr(Idx) & " Errors were found"
    Me.cmdXcopy.Enabled = True
    Me.lblStatus.Caption = vbNullString
    Me.cmdView.Enabled = True
    Me.cmdShortcut.Enabled = True
    Me.cmdBatch.Enabled = True
    Me.cmdView.Visible = True
    Me.cmdShortcut.Visible = True
    Me.cmdBatch.Visible = True
    Me.chkRstArchive.Enabled = True
    Me.chkSubdirs.Enabled = True
    Me.chkModified.Enabled = True
    Me.chkNewOnly.Enabled = True
    Me.txtFrom.Enabled = True
    Me.chkIncludeFolder.Enabled = True
    Me.txtTo.Enabled = True
    Me.txtLogfile.Enabled = True
    Me.cmdBrowseFrom.Enabled = True
    Me.cmdBrowseTo.Enabled = True
    Me.cmdLogfile.Enabled = True
    Me.cboRecent.Visible = True
    Me.cmdClear.Visible = True
    
    Me.chkZip.Enabled = HaveZip
    If Me.chkZip.Value = vbChecked And HaveZip Then
      Me.txtZIP.Enabled = True
      Me.txtZIP.BackColor = &H80000005
    End If
    
    Me.cmdQuit.Caption = "Close"                  'reset button label
    Me.cmdQuit.ToolTipText = "Exit program"
    Me.cmdXcopy.Caption = "Start Copy"
    Me.cmdXcopy.ToolTipText = "Begin Xcopy process"
    
    If CBool(Len(LogFile)) Then                   'log file defined?
      On Error Resume Next
      Bol = fso.FileExists(LogFile)               'does log file exist?
      If CBool(Err.Number) Then Bol = False       'error, so assume NO
      On Error GoTo 0
      Me.chkAppend.Enabled = Bol                  'enable checkbox in case was not enabled before
    End If
  End If
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Subroutine Name   : txtFrom_Change
' Purpose           : When txtfrom or txtto change, enable XCOPY button as needed
'*******************************************************************************
Private Sub txtFrom_Change()
  Dim Bol As Boolean
  Dim S As String
  
  S = Me.txtFrom.Text
  Me.cmdXcopy.Enabled = TestText()                'check enablement of execute button
  Bol = CBool(Len(S))                             'now check for enabling checkmark
  If Bol Then
    Bol = fso.FolderExists(S)
  End If
  Me.chkIncludeFolder.Enabled = Bol               'can enable Include Folder option
  If Bol Then
    Me.imgCheck(0).Picture = Me.imgCNC.Picture    'add checkmark
  Else
    Me.imgCheck(0).Picture = LoadPicture()        'else ensure it is clear
  End If
  If Len(S) < 4 Then                              'if path len <3, then do not allow path inclusion
    Me.chkIncludeFolder.Enabled = False
    Me.chkIncludeFolder.Value = vbUnchecked
  End If
  
  Me.cmdView.Enabled = Me.cmdXcopy.Enabled        'enable/disable buttons
  Me.cmdShortcut.Enabled = Me.cmdXcopy.Enabled
  Me.cmdBatch.Enabled = Me.cmdXcopy.Enabled
End Sub

'*******************************************************************************
' Function Name     : TestText
' Purpose           : Return True if both From and To data is valid
'*******************************************************************************
Private Function TestText() As Boolean
  TestText = CBool(Len(Me.txtFrom.Text)) And CBool(Len(Me.txtTo.Text))
  If TestText Then
    TestText = fso.FolderExists(Me.txtFrom.Text) And fso.FolderExists(Me.txtTo.Text)
  End If
End Function

'*******************************************************************************
' Subroutine Name   : txtFrom_GotFocus
' Purpose           : Select full text when we get focus
'*******************************************************************************
Private Sub txtFrom_GotFocus()
  With Me.txtFrom
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtTo_Change
' Purpose           : When txtfrom or txtto change, enable XCOPY button as needed
'*******************************************************************************
Private Sub txtTo_Change()
  Dim Bol As Boolean
  
  Me.cmdXcopy.Enabled = TestText()                'set execute button enablement
  Bol = CBool(Len(Me.txtTo.Text))                 'check for checkmark enablement
  If Bol Then
    Bol = fso.FolderExists(Me.txtTo.Text)
  End If
  If Bol Then
    Me.imgCheck(1).Picture = Me.imgCNC.Picture    'can show check
  Else
    Me.imgCheck(1).Picture = LoadPicture()        'else ensure it is gone
  End If
  Bol = Me.cmdXcopy.Enabled
  Me.cmdView.Enabled = Bol
  Me.cmdShortcut.Enabled = Bol
  Me.cmdBatch.Enabled = Bol
  Me.chkZip.Enabled = Bol And HaveZip
  If Me.chkZip.Enabled Then
    If Me.chkZip.Value = vbChecked Then
      Me.txtZIP.Enabled = True
      Me.txtZIP.BackColor = &H80000005
    Else
      Me.txtZIP.Enabled = False
      Me.txtZIP.BackColor = &H8000000F
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : txtTo_GotFocus
' Purpose           : Select full text when we get focus
'*******************************************************************************
Private Sub txtTo_GotFocus()
  With Me.txtTo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtLogFile_Change
' Purpose           : When txtfrom or txtto change, enable XCOPY button as needed
'*******************************************************************************
Private Sub txtLogFile_Change()
  Dim Bol As Boolean
  
  LogFile = Trim$(Me.txtLogfile.Text)
  If Left$(LogFile, 1) = "(" Then LogFile = vbNullString
  Bol = CBool(Len(LogFile))
  If Bol Then                                   'if acceptable data present
    Me.imgCheck(2).Picture = Me.imgCNC.Picture
    On Error Resume Next
    Bol = fso.FileExists(LogFile)               'does file already exist?
    If CBool(Err.Number) Then Bol = False       'error, so assume NO
    On Error GoTo 0
  Else
    Me.imgCheck(2).Picture = LoadPicture()
  End If
    
  Me.chkAppend.Enabled = Bol                    'enable if log file already exists
  
  Bol = TestText()
  Me.cmdXcopy.Enabled = Bol
  Me.cmdView.Enabled = Bol
  Me.cmdShortcut.Enabled = Bol
  Me.cmdBatch.Enabled = Bol
End Sub

'*******************************************************************************
' Subroutine Name   : txtLogFile_GotFocus
' Purpose           : Select full text when we get focus
'*******************************************************************************
Private Sub txtLogFile_GotFocus()
  With Me.txtLogfile
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Function Name     : SelectFile
' Purpose           : Select a file I/O folder
'*******************************************************************************
Private Function SelectFile() As String
  Dim sOpen As SelectedFile

  On Error Resume Next
  FileDialog.sFilter = "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
  ' See Standard CommonDialog Flags for all options
  FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
  FileDialog.sDlgTitle = "Select Log File"
  FileDialog.sInitDir = App.Path & "\"
  sOpen = ShowOpen(Me)
  If sOpen.bCanceled Then Exit Function
  SelectFile = AddSlash(sOpen.sLastDirectory) & sOpen.sFiles(1)
End Function

'*******************************************************************************
' Subroutine Name   : cmdView_Click
' Purpose           : Display command line format for this data
'*******************************************************************************
Private Sub cmdView_Click()
  Dim S As String
  
  S = GetCmdLine()
'
' copy data to clipboard
'
  With Clipboard
    .Clear
    .SetText S, vbCFText
  End With
'
' display data
'
  CenterMsgBoxOnForm Me, "Command line format (Copied also to clipboard):" & vbCrLf & _
                          S, _
                          vbOKOnly, _
                          "Command Line Format"
End Sub

'*******************************************************************************
' Function Name     : GetCmdLine
' Purpose           : Get DOS command Line / Shortcut rendition of options
'*******************************************************************************
Private Function GetCmdLine() As String
  Dim S As String, SF As String, SD As String, SL As String, SZ As String
  
  SF = Trim$(Me.txtFrom.Text)
  If Me.chkIncludeFolder.Value = vbChecked Then SF = AddSlash(SF)
  SD = Trim$(Me.txtTo.Text)

  SZ = Me.txtZIP.Text
  If Not HaveZip Or Me.chkZip.Value = vbUnchecked Then SZ = vbNullString
  
  SL = Trim$(Me.txtLogfile.Text)
  If Left$(SL, 1) = "(" Then
    SL = vbNullString
  ElseIf Me.chkAppend.Value = vbChecked Then
    SL = "+" & SL
  End If
  
  If CBool(InStr(1, SF, " ")) Then SF = """" & SF & """"
  If CBool(InStr(1, SD, " ")) Then SD = """" & SD & """"
  If CBool(Len(SZ)) Then
    If CBool(InStr(1, SZ, " ")) Then SZ = """" & SZ & """"
  End If
  If CBool(Len(SL)) Then
    If CBool(InStr(1, SL, " ")) Then SL = """" & SL & """"
  End If
  
  S = GetCmdData & " "
  If Me.chkModified.Value = vbChecked Then S = S & "/M "
  If Me.chkNewOnly.Value = vbChecked Then S = S & "/N "
  If Me.chkRstArchive.Value = vbChecked Then S = S & "/R "
  If Me.chkSubdirs.Value = vbChecked Then S = S & "/I "
  S = S & "/S=" & SF & " /D=" & SD
  If CBool(Len(SZ)) Then S = S & " /Z=" & SZ
  If CBool(Len(SL)) Then S = S & " /L=" & SL
  GetCmdLine = S
End Function

'*******************************************************************************
' Function Name     : GetCmdData
' Purpose           : Gather path to XCPY
'*******************************************************************************
Private Function GetCmdData() As String
  Dim S As String

  S = AddSlash(App.Path) & App.EXEName & ".exe"
  If InStr(1, S, " ") Then S = """" & S & """"
  GetCmdData = S
End Function

'*******************************************************************************
' Subroutine Name   : cmdShortcut_Click
' Purpose           : Create Desktop shortcut for current Item
'*******************************************************************************
Private Sub cmdShortcut_Click()
  Dim S As String, Cmd As String, Ttl As String
  Dim sc As cShortcut
    
  Ttl = vbNullString
  Ttl = Trim$(InputBox("Enter Title for This Desktop Shortcut:", "Enter Shortcut Title", "Xcopy Data"))
  If Not CBool(Len(Ttl)) Then Exit Sub
  
  Cmd = GetCmdData()
  S = GetCmdLine()
  S = Mid$(S, Len(Cmd) + 2)
  AddItemToHistory S                'add to history list
  If Left$(Cmd, 1) = """" Then Cmd = Mid$(Cmd, 2, Len(Cmd) - 2)
  
  Set sc = New cShortcut
  Call sc.CreateShortcut(Ttl, Cmd, S, "Backup '" & Me.txtFrom.Text & "' to '" & Me.txtTo.Text & "'")
  Set sc = Nothing
  CenterMsgBoxOnForm Me, "Desktop shortcut '" & Ttl & "' created.", vbOKOnly Or vbInformation, "Shortcut Created"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBatch_Click
' Purpose           : Add selection to a batch file
'*******************************************************************************
Private Sub cmdBatch_Click()
  Dim sOpen As SelectedFile
  Dim Ts As TextStream
  Dim S As String, T As String, Ary() As String
  Dim I As Long
  
  If Not CBool(Len(BatchPath)) Then     'batch file not yet set?
    On Error Resume Next
    With FileDialog
      .sFilter = "Batch files (*.bat)|*.bat"
      .flags = OFS_FILE_SAVE_FLAGS        'See Standard CommonDialog Flags for all options
      .sDlgTitle = "Create BATCH File"
      .sDefFileExt = "bat"
      .sInitDir = App.Path & "\"
      .sFIle = "XcpyBatch.bat"
    End With
    sOpen = ShowOpen(Me)
    If sOpen.bCanceled Then Exit Sub
    BatchPath = sOpen.sLastDirectory & "\" & sOpen.sFiles(1)
    '
    ' because we are new, supply a title to the batch
    '
    Batch = "@Echo XCPY Backup Batch file" & vbCrLf & _
            "@Echo ." & vbCrLf
    '
    ' see if batch file already exists
    '
    If fso.FileExists(BatchPath) Then
      I = InStrRev(BatchPath, "\")
      S = Mid$(BatchPath, I + 1)
      Select Case CenterMsgBoxOnForm(Me, "Batch File '" & S & "' Already Exists. Overwrite?" & vbCrLf & _
                                        "Yes=Overwrite, No=Append, Cancel=Stop", vbYesNoCancel Or vbQuestion, S & " Already Exists")
        Case vbCancel     'user want to cancel
          Exit Sub
        Case vbYes        'user want to over-write
        Case vbNo         'user wants to append
          Set Ts = fso.OpenTextFile(BatchPath, ForReading, True)  'read batch file
          Ary = Split(Ts.ReadAll, vbCrLf)                         'slipt to array
          Ts.Close
          For I = UBound(Ary) To 0 Step -1                        'find end of data
            If StrComp("@exit", Ary(I), vbTextCompare) = 0 Then
              If CBool(I) Then
                ReDim Preserve Ary(I - 1)                         'resize it
                Batch = Join(Ary, vbCrLf) & vbCrLf                'set initial data
              End If
              Exit For
            End If
          Next I
      End Select
    End If
    
    T = GetCmdData()
    S = GetCmdLine()
    AddItemToHistory Mid$(S, Len(T) + 2)                          'add to history list
    Batch = Batch & "Start " & S & vbCrLf                         'add command line
    
    On Error Resume Next
    '
    ' write the new data to the batch
    '
    Set Ts = fso.OpenTextFile(BatchPath, ForWriting, True)
    If CBool(Err.Number) Then
      CenterMsgBoxOnForm Me, "Cannot Create selected batch file:" & vbCrLf & _
                              BatchPath & vbCrLf & vbCrLf & _
                              "Error Description:" & vbCrLf & _
                              "(" & CStr(Err.Number) & ") " & Err.Description, _
                              vbOKOnly Or vbExclamation, _
                              "Cannot Create Batch"
      BatchPath = vbNullString
      Exit Sub
    End If
    On Error GoTo 0
    Ts.Write Batch & "@Exit" & vbCrLf
    Ts.Close

    I = InStrRev(BatchPath, "\")
    S = Mid$(BatchPath, I + 1)  'grab just filename
    CenterMsgBoxOnForm Me, "Once you are finished, you can execute your batch file by launching" & vbCrLf & _
                            S & " from the " & Left$(BatchPath, I) & " location.", _
                            vbOKOnly Or vbInformation, _
                            "Reminder"
  Else  'batch file already opened from previous use in this session
    Batch = Batch & "Start " & GetCmdLine() & vbCrLf
    Set Ts = fso.OpenTextFile(BatchPath, ForWriting, True)
    Ts.Write Batch & "@Exit" & vbCrLf
    Ts.Close
    
    CenterMsgBoxOnForm Me, "Added sequence to " & BatchPath & vbCrLf & vbCrLf & _
                           "File So Far:" & vbCrLf & Batch, _
                           vbOKOnly Or vbInformation, "Operation Successful"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : chkZip_Click
' Purpose           : Toggle ZIP option
'*******************************************************************************
Private Sub chkZip_Click()
  On Error Resume Next
  If Me.chkZip.Value = vbChecked Then
    Me.txtZIP.Enabled = True
    Me.txtZIP.BackColor = &H80000005
    Me.txtZIP.SetFocus
  Else
    Me.txtZIP.Enabled = False
    Me.txtZIP.BackColor = &H8000000F
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : txtZIP_GotFocus
' Purpose           : Select all text when we get focus
'*******************************************************************************
Private Sub txtZIP_GotFocus()
  With Me.txtZIP
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDownload_Click
' Purpose           : Dowload VBZIP10.DLL
'*******************************************************************************
Private Sub cmdDownload_Click()
  If CheckInternetConnect() Then
    If CenterMsgBoxOnForm(Me, "You are about to download ""Info-ZIP_Zip_DLL_(Renamed_vbzip10_dll).zip""." & vbCrLf & _
                          "Select Save Or Open (to run WinZip immediately), but save the vbZip10.dll" & vbCrLf & _
                          "to your XCPY program folder. Once unzipped to this folder, the next time" & vbCrLf & _
                          "you run XCPY, you will be able to ZIP selected data to a" & vbCrLf & _
                          "<SelectedFileName>.zip file, ready for permanent archiving." & vbCrLf & vbCrLf & _
                          "NOTE: If the File Download dialog pops up and then disappears, try again." & vbCrLf & _
                          "This can take up to three tries sometimes, for whatever reason." & vbCrLf & vbCrLf & _
                          "Do you want to continue?", vbYesNo Or vbQuestion, "Ready to Download") = vbNo Then Exit Sub
    On Error Resume Next
    OpenFilePath Me.hwnd, "http://www.vbaccelerator.com/home/VB/Code/Libraries/Compression/Introduction_to_the_Info-ZIP_Libraries/Info-ZIP_Zip_DLL_(Renamed_vbzip10_dll).zip"
    If CBool(Err.Number) Then
      CenterMsgBoxOnForm Me, "Error. Are you connected to the internet?", vbOKOnly Or vbExclamation, "DownLoad Error"
    End If
  Else
    CenterMsgBoxOnForm Me, "Internet Connection not detected", vbOKOnly Or vbExclamation, "No Internet"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : tmrDelay_Timer
' Purpose           : 1-second delay
'*******************************************************************************
Private Sub tmrDelay_Timer()
  Me.tmrDelay.Enabled = False
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClear_Click
' Purpose           : Erase Recent List
'*******************************************************************************
Private Sub cmdClear_Click()
  If CenterMsgBoxOnForm(Me, "Verify clearing history of previous configurations.", _
     vbYesNo Or vbQuestion, _
     "Verify History Delete") = vbNo Then
    Exit Sub
  End If
  Me.cboRecent.Clear
  SaveSetting App.Title, "Settings", "HistCnt", "0"
  Me.cboRecent.ToolTipText = "Recent configurations. Select from dropdown list to activate selection"
End Sub

'*******************************************************************************
' Subroutine Name   : cboRecent_Click
' Purpose           : Selection made for item
'*******************************************************************************
Private Sub cboRecent_Click()
  Dim S As String
  
  If Checking Then Exit Sub
  Checking = True
  With Me.cboRecent
    S = .List(.ListIndex)                               'get selected item
    .ToolTipText = S
    .RemoveItem .ListIndex                              'remove current selection
    .AddItem S, 0                                       'add item
    .ListIndex = 0
    Call ProcessCmd(S)
  End With
  Checking = False
End Sub

'*******************************************************************************
' Subroutine Name   : AddItemToHistory
' Purpose           : Add entry to History
'*******************************************************************************
Private Sub AddItemToHistory(Itm As String)
  Dim Idx As Integer
  
  With Me.cboRecent
    For Idx = 0 To .ListCount - 1
      If StrComp(Itm, .List(Idx), vbTextCompare) = 0 Then 'found a match?
        .RemoveItem Idx                                   'yes, so remove it
        Exit For
      End If
    Next Idx
    Checking = True
    .AddItem Itm, 0                                       'add to start of list
    .ToolTipText = Itm
    .ListIndex = 0
    Checking = False
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : ProcessCmd
' Purpose           : Process command data and extract options
'*******************************************************************************
Private Sub ProcessCmd(Txt As String)
  Dim Cmd As String, S As String, T As String
  Dim I As Integer
  
  Me.chkModified.Value = vbUnchecked
  Me.chkNewOnly.Value = vbUnchecked
  Me.chkRstArchive.Value = vbUnchecked
  Me.chkSubdirs.Value = vbUnchecked
  Me.chkIncludeFolder.Value = vbUnchecked
  Me.chkAppend.Value = vbUnchecked
  Me.txtTo.Text = "(Folder path to copy out of)"
  Me.txtFrom.Text = "(Folder path to copy into)"
  Me.txtLogfile.Text = "(Path to write log entries)"
  Me.imgCheck(0).Picture = LoadPicture()
  Me.imgCheck(1).Picture = LoadPicture()
  Me.imgCheck(2).Picture = LoadPicture()
  
  Cmd = Txt
  If CBool(Len(Cmd)) Then                         'anything?
    Cmd = Cmd & "/"                               'append command delimiter
    I = InStr(2, Cmd, "/")                        'find next slash
    Do While I                                    'do wile a second slash exists
      S = Trim$(Left$(Cmd, I - 1))                'get just data to work on during this cycle
      Cmd = Mid$(Cmd, I)                          'strip left stuff off
      Select Case UCase$(Mid$(S, 2, 1))           'parse parameter
        Case "M"                                  'Copy Modified files only
          Me.chkModified.Value = vbChecked
        Case "N"                                  'Copy new files only
          Me.chkNewOnly.Value = vbChecked
        Case "R"                                  'Reset Archive bits
          Me.chkRstArchive.Value = vbChecked
        Case "I"                                  'inclide subdirectories
          Me.chkSubdirs.Value = vbChecked
        Case "S"                                  'Source path
          I = InStr(1, S, "=")
          If CBool(I) Then
            T = Trim$(StripQuotes(Mid$(S, I + 1)))
            If Right$(T, 1) = "\" Then
              Select Case Mid$(T, 2, 1)
                Case "\"                          'if network path...
                  I = InStr(3, T, "\")            'find end of root
                  I = InStr(I + 1, T, "\")        'find end of drive off root
                  If CBool(I) Then                'if drive followed by backslash
                    If I <> Len(T) Then           'if not root...
                      Me.chkIncludeFolder.Value = vbChecked 'then OK to include sourcepath
                    End If
                  End If
                Case ":"                          'drivepath
                  If Len(T) <> 3 Then             'if not specifying something like "C:\"...
                    Me.chkIncludeFolder.Value = vbChecked 'allow including sourcepath
                  End If
              End Select
              T = Left$(T, Len(T) - 1)            'strip trailing slash, regardless
            End If
            Me.txtFrom.Text = T
            If CBool(Len(Me.txtFrom.Text)) Then Me.imgCheck(0).Picture = Me.imgCNC.Picture
          End If
        Case "D"                                  'destination path
          I = InStr(1, S, "=")
          If CBool(I) Then
            Me.txtTo.Text = Trim$(StripQuotes(Mid$(S, I + 1)))
            If CBool(Len(Me.txtTo.Text)) Then Me.imgCheck(1).Picture = Me.imgCNC.Picture
          End If
        Case "Z"
          I = InStr(1, S, "=")
          If CBool(I) Then
            Me.txtZIP.Text = Trim$(StripQuotes(Mid$(S, I + 1)))
            If HaveZip And CBool(Len(Me.txtZIP.Text)) Then Me.chkZip.Value = vbChecked
          End If
        Case "L"                                  'log file
          I = InStr(1, S, "=")
          If CBool(I) Then
            LogFile = Trim$(StripQuotes(Mid$(S, I + 1)))
            If CBool(Len(LogFile)) Then
              If Left$(LogFile, 1) = "+" Then
                LogFile = LTrim$(Mid$(LogFile, 2))
                Me.chkAppend.Value = vbChecked
              End If
              Me.txtLogfile.Text = LogFile
              Me.imgCheck(2).Picture = Me.imgCNC.Picture
            End If
          End If
      End Select
      I = InStr(2, Cmd, "/")                      'now find next block
    Loop
  End If
End Sub
'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************
