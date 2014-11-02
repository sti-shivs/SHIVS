VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Position 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Candidates"
   ClientHeight    =   9375
   ClientLeft      =   135
   ClientTop       =   1770
   ClientWidth     =   15795
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
   ScaleWidth      =   15795
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPosAdd 
      Caption         =   "Add Position"
      Height          =   855
      Left            =   11880
      TabIndex        =   10
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8040
      TabIndex        =   9
      Top             =   7680
      Width           =   3615
      Begin MSComCtl2.DTPicker dtePosDateEnd 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109641729
         CurrentDate     =   41923
      End
      Begin MSComCtl2.DTPicker dtePosTimeEnd 
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109641730
         CurrentDate     =   41923
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8040
      TabIndex        =   8
      Top             =   6600
      Width           =   3615
      Begin MSComCtl2.DTPicker dtePosDateStart 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109641729
         CurrentDate     =   41923
      End
      Begin MSComCtl2.DTPicker dtePosTimeStart 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109641730
         CurrentDate     =   41923
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   6600
      Width           =   3615
      Begin VB.TextBox txtPosName 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Intended for"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   7680
      Width           =   3615
      Begin VB.ComboBox cboPosType 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Description"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   4200
      TabIndex        =   4
      Top             =   6600
      Width           =   3615
      Begin VB.TextBox txtPosDesc 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.TextBox txtPosSearchBox 
      Height          =   375
      Left            =   12840
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.ComboBox cboPosSearchKey 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid grdPositions 
      Height          =   3735
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   6588
      _Version        =   393216
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
      Caption         =   "Positions"
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
      Width           =   990
   End
End
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call HeaderFor.PositionsGrid
    Call PollPosition.DisplayToGrid(grdPositions)
End Sub
