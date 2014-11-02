VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Log 
   BackColor       =   &H80000005&
   Caption         =   "Logs"
   ClientHeight    =   9375
   ClientLeft      =   2475
   ClientTop       =   1725
   ClientWidth     =   15825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   15825
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboLogsSearchKey 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox cboLogsSearchBox 
      Height          =   375
      Left            =   12840
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid grdLogs 
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
      Caption         =   "Logs"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   525
   End
End
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call HeaderFor.LogGrid
    Call PollResult.DisplayToGrid(grdLogs)
End Sub
