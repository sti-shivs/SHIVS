VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Candidate 
   BackColor       =   &H80000005&
   Caption         =   "Candidates"
   ClientHeight    =   9375
   ClientLeft      =   1845
   ClientTop       =   1620
   ClientWidth     =   15780
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCandSearchBox 
      Height          =   375
      Left            =   12840
      TabIndex        =   22
      Top             =   2160
      Width           =   2655
   End
   Begin VB.ComboBox cboCandSearchKey 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10320
      TabIndex        =   21
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "College"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4200
      TabIndex        =   19
      Top             =   7680
      Width           =   3615
      Begin VB.ComboBox cboCandCollege 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Position"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4200
      TabIndex        =   17
      Top             =   6600
      Width           =   3615
      Begin VB.ComboBox cboCandPosition 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdCandAdd 
      Caption         =   "Add Candidate"
      Height          =   855
      Left            =   11880
      TabIndex        =   16
      Top             =   8880
      Width           =   3615
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Birthday"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   11880
      TabIndex        =   14
      Top             =   7680
      Width           =   3615
      Begin VB.Label lblCandBirthday 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Year Level"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8040
      TabIndex        =   12
      Top             =   8760
      Width           =   3615
      Begin VB.Label lblCandYearLvl 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Course"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8040
      TabIndex        =   10
      Top             =   7680
      Width           =   3615
      Begin VB.Label lblCandCourse 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gender"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   11880
      TabIndex        =   8
      Top             =   6600
      Width           =   3615
      Begin VB.Label lblCandGender 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ID Number"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8040
      TabIndex        =   6
      Top             =   6600
      Width           =   3615
      Begin VB.Label lblCandID 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   7680
      Width           =   3615
      Begin VB.ComboBox cboCandName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Electoral Party"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   6600
      Width           =   3615
      Begin VB.ComboBox cboCandElectoralParty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCandidates 
      Height          =   3735
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   6588
      _Version        =   393216
      BackColorSel    =   11164160
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Candidates"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Candidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCandElectoralParty_Click()
    If Record.State = 1 Then Record.Close
    
    SQL = "SELECT id FROM wp_bp_groups WHERE name = '" & cboCandElectoralParty.Text & "'"
    
    Record.Open SQL, Connect
    
    cboCandName.Enabled = True
    
    Call PollCandidate.DisplayElectoralPartyMembers(Record!ID, cboCandName)
End Sub

Private Sub cboCandName_Click()
    Dim tmpID As Integer    'this will hold the ID of the candidate
    
    If Record.State = 1 Then Record.Close
    
    SQL = "SELECT ID FROM wp_users WHERE display_name = '" & cboCandName.Text & "'"
    
    Record.Open SQL, Connect
    
    tmpID = Record!ID
    
    'ID Number
    Call PollCandidate.DisplayElectoralPartyMembersIDNumber(tmpID, lblCandID)
    
    'Course
    Call PollCandidate.DisplayElectoralPartyMembersCourse(tmpID, lblCandCourse)
    
    'Year Level
    Call PollCandidate.DisplayElectoralPartyMembersYearLvl(tmpID, lblCandYearLvl)
    
    'Gender
    Call PollCandidate.DisplayElectoralPartyMembersGender(tmpID, lblCandGender)
    
    'Birthdate
    Call PollCandidate.DisplayElectoralPartyMembersBirthday(tmpID, lblCandBirthday)
End Sub

Private Sub Form_Load()
    Call HeaderFor.CandidatesGrid
    
    Call PollCandidate.DisplayCandidates(grdCandidates)
    
    Call PollCandidate.DisplayElectoralParty(cboCandElectoralParty)
End Sub
