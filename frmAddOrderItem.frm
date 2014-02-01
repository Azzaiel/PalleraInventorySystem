VERSION 5.00
Begin VB.Form frmAddOrderItem 
   Appearance      =   0  'Flat
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4815
      Begin VB.Label lblSuplier 
         Caption         =   "XXXXXXXXXXXXXXXXXXX"
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
         Left            =   1320
         TabIndex        =   15
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblOrderID 
         Caption         =   "XXXXXXXXXXXXXXXXXXX"
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
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Suplier:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Order ID:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbItems 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Retail Price"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Item Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FF00&
         Caption         =   "Total Price"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblUnitPrice 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblTotalPrice 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAddOrderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
