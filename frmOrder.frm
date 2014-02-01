VERSION 5.00
Begin VB.Form frmOrder 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtRetailPrice 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Status"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Order ID"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Received By"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Received Date"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  By"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Suppliers:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Retail Price"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         Caption         =   "Order  Date"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblCreatedDate_Click()

End Sub
