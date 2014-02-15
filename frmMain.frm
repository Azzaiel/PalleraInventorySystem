VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Pallera's  Inventory System"
   ClientHeight    =   7575
   ClientLeft      =   420
   ClientTop       =   1635
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Pending Orders"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu mnSupplier 
      Caption         =   "Supplier"
      Begin VB.Menu mnReSupplier 
         Caption         =   "Regster Supplier"
      End
      Begin VB.Menu mnRegItemType 
         Caption         =   "Register Item Type"
      End
   End
   Begin VB.Menu mnInventory 
      Caption         =   "Inventory"
      Begin VB.Menu mnRegisterItem 
         Caption         =   "Register  Item"
      End
      Begin VB.Menu mnOder 
         Caption         =   "Order Itmes"
      End
   End
   Begin VB.Menu mnUsers 
      Caption         =   "Account"
      Begin VB.Menu mnChangePass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnRegUsers 
         Caption         =   "Regiter Users"
      End
      Begin VB.Menu mnLogout 
         Caption         =   "Logout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

lblWelcome.Caption = "Welcome " & UserSession.Name & " you are logged in as " & UserSession.Role
End Sub

Private Sub mnChangePass_Click()
  frmChangePassword.Show vbModal
End Sub

Private Sub mnLogout_Click()
   Unload Me
   LogIn.Show vbModal
End Sub

Private Sub mnOder_Click()
  frmOrder.Show vbModal
End Sub

Private Sub mnRegisterItem_Click()
  frmItemReg.Show vbModal
End Sub

Private Sub mnRegItemType_Click()
  frmItemReg.Show vbModal
End Sub

Private Sub mnRegUsers_Click()
  frmAccnt.Show vbModal
End Sub

Private Sub mnReSupplier_Click()
  frmSupplier.Show vbModal
End Sub
