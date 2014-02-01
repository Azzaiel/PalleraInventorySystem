VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin MSComCtl2.DTPicker dtOrderDate 
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Format          =   106627075
         CurrentDate     =   41671
      End
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
         Width           =   2655
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2655
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
Option Explicit
Private rs As ADODB.Recordset
Private suplierIdList As Variant
Private itemTypeIdList As Variant
Private tempRs As ADODB.Recordset
Private Sub lblCreatedDate_Click()

End Sub

Private Sub cmbSupplier_Click()
   cmbItemType.Clear
  Set tempRs = DataCrudDao.getItemTypeRSBySupplierID(Val(suplierIdList(cmbSupplier.ListIndex)))
  ReDim itemTypeIdList(0 To tempRs.RecordCount) As Long
  Dim index As Integer
  index = 0
   While Not tempRs.EOF
    cmbItemType.AddItem tempRs!ITEM_TYPE_NAME
    itemTypeIdList(index) = tempRs!id
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub

Private Sub Form_Load()
  dtOrderDate.CustomFormat = Constants.DEFAULT_FORMAT
  dtOrderDate = Now
  Call populateLov
End Sub
Private Sub populateLov()
  Set tempRs = DataCrudDao.getSupplierRS("", "", "")
  cmbSupplier.Clear
  ReDim suplierIdList(0 To tempRs.RecordCount) As Long
  Dim index As Integer
  index = 0
  While Not tempRs.EOF
    cmbSupplier.AddItem tempRs!Name
    suplierIdList(index) = tempRs!id
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub

