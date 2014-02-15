VERSION 5.00
Begin VB.Form LogIn 
   Caption         =   "LogIn"
   ClientHeight    =   3015
   ClientLeft      =   5670
   ClientTop       =   3705
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4140
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00000000&
         Caption         =   "&Login"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username:"
         Height          =   255
         Left            =   240
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
  Call CommonHelper.toDefaultSkin(TxtUser)
  Call CommonHelper.toDefaultSkin(TxtPassword)
  If (hasValidForm) Then
    If UserSession.hasValidCredential(TxtUser, TxtPassword) Then
    Unload Me
    frmMain.Show
    Else
    MsgBox "Username and Password do not match!", vbCritical
    End If
  End If
End Sub
Private Function hasValidForm() As Boolean
   If (Not CommonHelper.hasValidValue(TxtUser)) Then
     Call CommonHelper.sendWarning(TxtUser, "Please enter a username")
     hasValidForm = False
     
   ElseIf (Not CommonHelper.hasValidValue(TxtPassword)) Then
     Call CommonHelper.sendWarning(TxtPassword, "Please enter a password")
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
