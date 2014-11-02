VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Index 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Index"
   ClientHeight    =   9390
   ClientLeft      =   555
   ClientTop       =   1530
   ClientWidth     =   14490
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   16563
      AllowCustomize  =   0   'False
      _Version        =   327682
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   12135
         Left            =   -1200
         ScaleHeight     =   12105
         ScaleWidth      =   4545
         TabIndex        =   1
         Top             =   -480
         Width           =   4575
         Begin VB.CommandButton cmdPosition 
            Appearance      =   0  'Flat
            BackColor       =   &H00AA5A00&
            Caption         =   "&Position"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3120
            UseMaskColor    =   -1  'True
            Width           =   2055
         End
         Begin VB.PictureBox Home 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   1080
            Picture         =   "index.frx":0000
            ScaleHeight     =   81
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   185
            TabIndex        =   5
            Top             =   960
            Width           =   2775
         End
         Begin VB.CommandButton cmdVoters 
            Appearance      =   0  'Flat
            BackColor       =   &H00AA5A00&
            Caption         =   "&Voters"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5160
            UseMaskColor    =   -1  'True
            Width           =   2055
         End
         Begin VB.CommandButton cmdCandidates 
            Appearance      =   0  'Flat
            BackColor       =   &H00AA5A00&
            Caption         =   "&Candidates"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1560
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Add Candidates"
            Top             =   4200
            Width           =   2055
         End
         Begin VB.CommandButton cmdLogs 
            Appearance      =   0  'Flat
            BackColor       =   &H00AA5A00&
            Caption         =   "&Logs"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   6240
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCandidates_Click()
    Load Candidate
    
    Unload Log
    Unload Voter
    Unload Position
End Sub

Private Sub cmdLogs_Click()
    Load Log
    
    Unload Candidate
    Unload Voter
    Unload Position
End Sub

Private Sub cmdPosition_Click()
    Load Position
    
    Unload Log
    Unload Voter
    Unload Candidate
End Sub

Private Sub cmdVoters_Click()
    Load Voter
    
    Unload Log
    Unload Candidate
    Unload Position
End Sub

Private Sub Home_Click()
    Unload Candidate
    Unload Log
    Unload Voter
    Unload Position
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Security.Show
End Sub
