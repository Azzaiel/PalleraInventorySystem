VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   8685
   ClientTop       =   1965
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   495
         Left            =   1080
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtOldPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm Password"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "New Password"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Old Password"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private rs As ADODB.Recordset
Private Sub cmdSubmit_Click()
Call CommonHelper.toDefaultSkin(txtOldPassword)
Call CommonHelper.toDefaultSkin(txtNewPassword)
Call CommonHelper.toDefaultSkin(txtConfirmPassword)
If hasValidForm Then
    Set rs = UserSession.getUserRS(txtOldPassword)
    If rs.RecordCount = 1 Then
       rs!Password = txtNewPassword
       rs.Update
       MsgBox "Password Updated", vbInformation
       Unload Me
    Else
      MsgBox "Old Password is incorrect", vbCritical
    End If
End If


End Sub

Private Function hasValidForm() As Boolean
   If (Not CommonHelper.hasValidValue(txtOldPassword)) Then
     Call CommonHelper.sendWarning(txtOldPassword, "Please enter Current Password ")
     hasValidForm = False
     
   ElseIf (Not CommonHelper.hasValidValue(txtNewPassword)) Then
     Call CommonHelper.sendWarning(txtNewPassword, "Please enter a New Password")
     hasValidForm = False
     
   ElseIf (Not CommonHelper.hasValidValue(txtConfirmPassword)) Then
     Call CommonHelper.sendWarning(txtConfirmPassword, "Password Confirmation Required")
     hasValidForm = False
   ElseIf txtConfirmPassword <> txtNewPassword Then
     MsgBox "NewPassword and Confirm Password must match"
     hasValidForm = False
   Else
     hasValidForm = True
     
   End If
End Function


