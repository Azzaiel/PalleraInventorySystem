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
      Begin VB.Menu mnRegUsers 
         Caption         =   "Regiter Users"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
