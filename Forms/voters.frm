VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Voter 
   BackColor       =   &H80000005&
   Caption         =   "Voters"
   ClientHeight    =   9360
   ClientLeft      =   2175
   ClientTop       =   1485
   ClientWidth     =   15780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grdVoters 
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
      Caption         =   "Voters"
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
      Width           =   720
   End
End
Attribute VB_Name = "Voter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call HeaderFor.VoterGrid
    Call PollVoter.DisplayToGrid(grdVoters)
End Sub
