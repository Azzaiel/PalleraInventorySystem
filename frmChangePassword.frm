VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "Change Password Form"
   ClientHeight    =   3555
   ClientLeft      =   8685
   ClientTop       =   1965
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   Picture         =   "frmChangePassword.frx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   2760
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
         Left            =   720
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Passowrd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Image Image5 
         Height          =   4215
         Left            =   0
         Picture         =   "frmChangePassword.frx":59ED
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15975
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

Private Sub cmdCancel_Click()
Unload Me
End Sub

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


