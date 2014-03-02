VERSION 5.00
Begin VB.Form frmEntePayment 
   Caption         =   "Payment Form"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3540
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmbCancel 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmbSubmit 
      Cancel          =   -1  'True
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
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPayment 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter Payment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Total Cost:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblTotalCost 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmEntePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCancel_Click()
  Unload Me
End Sub

Private Sub cmbSubmit_Click()
 If Val(txtPayment) >= frmItemSell.totalCost Then
   frmItemSell.payment = Val(txtPayment)
   Unload Me
 Else
   MsgBox "Please enter an ammount equal or greater dan the total cost"
 End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtPayment) > 11)) Then
    KeyAscii = 0
    Beep
  End If
  If KeyAscii = 13 Then
    Call cmbSubmit_Click
  End If
End Sub
