VERSION 5.00
Begin VB.Form LogIn 
   Caption         =   "LogIn"
   ClientHeight    =   1965
   ClientLeft      =   5670
   ClientTop       =   3705
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4650
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00000000&
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   120
         Picture         =   "LogIn.frx":0000
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "LogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdLogin_Click()
  Call CommonHelper.toDefaultSkin(txtUser)
  Call CommonHelper.toDefaultSkin(txtPassword)
  If (hasValidForm) Then
    If UserSession.hasValidCredential(txtUser, txtPassword) Then
      Unload Me
      frmMain.Show
      If UserSession.Role = "User" Then
        frmMain.mnSupplier.Visible = False
        frmMain.mnRegisterItem.Visible = False
        frmMain.mnReports.Visible = False
        frmMain.mnRegUsers.Visible = False
      Else
        frmMain.mnSupplier.Visible = True
        frmMain.mnRegisterItem.Visible = True
        frmMain.mnReports.Visible = True
        frmMain.mnRegUsers.Visible = True
      End If
    Else
      MsgBox "Username and Password do not match!", vbCritical
    End If
  End If
End Sub
Private Function hasValidForm() As Boolean
   If (Not CommonHelper.hasValidValue(txtUser)) Then
     Call CommonHelper.sendWarning(txtUser, "Please enter a username")
     hasValidForm = False
     
   ElseIf (Not CommonHelper.hasValidValue(txtPassword)) Then
     Call CommonHelper.sendWarning(txtPassword, "Please enter a password")
     hasValidForm = False
   Else
     hasValidForm = True
     
   End If
End Function

Private Sub Command2_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo) = vbYes Then
End

End If





End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CmdLogin_Click
End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CmdLogin_Click
End If
End Sub
