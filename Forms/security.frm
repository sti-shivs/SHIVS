VERSION 5.00
Begin VB.Form Security 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security"
   ClientHeight    =   5655
   ClientLeft      =   5850
   ClientTop       =   3705
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6090
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1560
      Picture         =   "security.frx":0000
      ScaleHeight     =   1215
      ScaleWidth      =   3015
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00AA5A00&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Username"
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
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Password"
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
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   1035
   End
End
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mUsername As String
Private mPassword As String
Private mCancel As Boolean

Private Sub cmdCancel_Click()
    mCancel = True
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mCancel = False
    mUsername = Me.txtUsername.Text
    mPassword = Me.txtPassword.Text
    
    Unload Me
End Sub

Public Function GetUserInfo(ByRef Username As String, ByRef Password As String, Owner As Object) As Boolean
    Me.txtUsername.Text = Username
    
    Me.Show vbModal, Owner
    
    Username = mUsername
    Password = mPassword
    
    GetUserInfo = Not mCancel
End Function

Private Sub Form_Activate()
    If Len(Me.txtUsername.Text) > 0 Then Me.txtPassword.SetFocus
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub txtUsername_GotFocus()
    txtUsername.SelStart = 0
    txtUsername.SelLength = Len(txtUsername.Text)
End Sub

