VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItemReg 
   Caption         =   "Item Registration Form"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   17085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   17085
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Item Type Form"
      Height          =   5295
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtCriticalLevel 
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtRetailPrice 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox cmbSupplier 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Type Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Critical Level"
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
         Left            =   240
         TabIndex        =   42
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblQuantity 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   240
         TabIndex        =   36
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:"
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
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Active:"
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
         Left            =   240
         TabIndex        =   32
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label txtActive 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         Left            =   240
         TabIndex        =   30
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Retail Price"
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
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Suppliers:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Type:"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Created by:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Created date:"
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
         Left            =   240
         TabIndex        =   25
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Last mod by:"
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
         Left            =   240
         TabIndex        =   24
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Last mod date:"
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
         Left            =   240
         TabIndex        =   23
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   5415
         Left            =   0
         Picture         =   "frmItemReg.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.ComboBox cmbSupplierName 
      Height          =   315
      Left            =   12960
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtItemCodeSearch 
      Height          =   285
      Left            =   8040
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdActivation 
      Caption         =   "De-Activate"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmbClear 
      Caption         =   "Clear"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmbClose 
      Caption         =   "Close"
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
      Left            =   5040
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmbEdit 
      Caption         =   "Edit"
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
      Left            =   2640
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmbNewRec 
      Caption         =   "New"
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
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgItems 
      Height          =   4215
      Left            =   6240
      TabIndex        =   12
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7435
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Form"
      Height          =   1455
      Left            =   6240
      TabIndex        =   13
      Top             =   120
      Width           =   10695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6720
         TabIndex        =   40
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:"
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
         Left            =   5400
         TabIndex        =   41
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
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
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label aaa 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code Search"
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
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
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
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   1575
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   5415
         Left            =   0
         Picture         =   "frmItemReg.frx":750D4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11295
      End
   End
   Begin VB.Image Image5 
      Height          =   6015
      Left            =   -120
      Picture         =   "frmItemReg.frx":EA1A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17175
   End
End
Attribute VB_Name = "frmItemReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private suplierIdList As Variant
Private itemTypeIdList As Variant
Private tempRs As ADODB.Recordset

Private Sub cmbClear_Click()
  Call clearForm
  Call toogelInsertMode(False)
End Sub

Private Sub cmbClose_Click()
    Unload Me
End Sub

Private Sub cmbEdit_Click()
  If (hasValidFormValue(Val(rs!id))) Then
    Set tempRs = DataCrudDao.getItemRSByID(rs!id)
    tempRs!ITEM_CODE = txtItemCode
    tempRs!supplier_id = suplierIdList(cmbSupplier.ListIndex)
    tempRs!ITEM_TYPE_ID = itemTypeIdList(cmbItemType.ListIndex)
    tempRs!Name = txtItemName
    tempRs!CRITICAL_LEVEL = Val(txtCriticalLevel)
    tempRs!RETAIL_PRICE = txtRetailPrice
    tempRs!UNIT_PRICE = txtUnitPrice
    tempRs!CREATED_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Updated!! ", vbInformation
    Call populateDataGrid
  End If

End Sub

Private Sub cmbNewRec_Click()
  If (cmbNewRec.Caption = "New") Then
     toogelInsertMode (True)
  Else
    If (hasValidFormValue) Then
      Call toogelInsertMode(False)
      Set tempRs = DataCrudDao.getFakeItemsRS
      tempRs.AddNew
      tempRs!ITEM_CODE = txtItemCode
      tempRs!supplier_id = suplierIdList(cmbSupplier.ListIndex)
      tempRs!ITEM_TYPE_ID = itemTypeIdList(cmbItemType.ListIndex)
      tempRs!Name = txtItemName
      tempRs!RETAIL_PRICE = txtRetailPrice
      tempRs!CRITICAL_LEVEL = Val(txtCriticalLevel)
      tempRs!QUANTITY = 0
      tempRs!UNIT_PRICE = txtUnitPrice
      tempRs!CREATED_BY = UserSession.getLoginUser
      tempRs!CREATED_DATE = Now
      tempRs!LAST_MOD_DATE = Now
      tempRs!active = txtActive
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      MsgBox "Record Created", vbInformation
      Call populateDataGrid
      Call toogelInsertMode(False)
      Call clearForm
    End If
  End If
End Sub
Private Function hasValidFormValue(Optional itemID As Integer = -1) As Boolean
  Dim isValid As Boolean
  isValid = True
  If CommonHelper.hasValidValue(txtItemCode) = False Then
    isValid = False
    MsgBox "Please enter an Item Code", vbCritical
  ElseIf DataCrudDao.isItemCodeAlreadyUsed(txtItemCode, itemID) Then
    isValid = False
    MsgBox "ItemCodeAlready In use", vbCritical
  ElseIf (CommonHelper.hasValidValue(txtItemName) = False) Then
    isValid = False
    MsgBox "Please enter an Item Name", vbCritical
  ElseIf (CommonHelper.hasValidValue(txtCriticalLevel) = False) Then
    isValid = False
    MsgBox "Please enter an Critical level", vbCritical
  ElseIf (CommonHelper.hasValidValue(txtRetailPrice) = False) Then
    isValid = False
    MsgBox "Please enter an Retail Price", vbCritical
  ElseIf (CommonHelper.hasValidValue(txtUnitPrice) = False) Then
    isValid = False
    MsgBox "Please enter an Retail Price", vbCritical
  End If
  hasValidFormValue = isValid
End Function

Private Sub cmbSupplier_Click()
If cmbSupplier.ListIndex >= 0 Then
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
End If
End Sub

Private Sub cmdActivation_Click()

  Set tempRs = DataCrudDao.getItemRSByID(rs!id)
    
  If rs!active = "N" Then
    cmdActivation.Caption = "De-Activate"
    tempRs!active = "Y"
  Else
    tempRs!active = "N"
    cmdActivation.Caption = "Activate"
  End If
  tempRs.Update
  Call DbInstance.closeRecordSet(tempRs)
  MsgBox "Status Update "
    
  Call clearForm
  Call populateDataGrid


End Sub

Private Sub cmdSearch_Click()
Call populateDataGrid
End Sub

Private Sub dgItems_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub

Private Sub Form_Load()
  Call populateLov
  Call populateDataGrid
End Sub
Private Sub populateLov()
  Set tempRs = DataCrudDao.getSupplierRS("", "", "")
  cmbSupplier.Clear
  cmbSupplierName.Clear
  ReDim suplierIdList(0 To tempRs.RecordCount) As Long
  Dim index As Integer
  index = 0
  While Not tempRs.EOF
    cmbSupplier.AddItem tempRs!Name
    cmbSupplierName.AddItem tempRs!Name
    suplierIdList(index) = tempRs!id
    index = index + 1
    tempRs.MoveNext
  Wend
  Call DbInstance.closeRecordSet(tempRs)
End Sub
Private Sub populateDataGrid()
  Set rs = DataCrudDao.getItemReg(txtItemCodeSearch)
  Set dgItems.DataSource = rs
  dgItems.Refresh
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
 End If
Call formatDataGrid
End Sub

Private Sub Image4_Click()

End Sub

Private Sub txtCriticalLevel_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtCriticalLevel) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txtItemCodeSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdSearch_Click
End If

End Sub

Private Sub txtRetailPrice_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtRetailPrice) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txtSearchItemType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdSearch
End If
End Sub

Private Sub txtSearchSuppliers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdSearch
End If
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
  If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtUnitPrice) > 11)) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub showSelectedData()
 txtItemCode = CommonHelper.extractStringValue(rs!ITEM_CODE)
 cmbSupplier.Text = CommonHelper.extractStringValue(rs!SUPPLIER)
 cmbItemType.Text = CommonHelper.extractStringValue(rs!ITEM_TYPE)
 txtItemName = CommonHelper.extractStringValue(rs!ITEM_NAME)
 lblQuantity = Val(CommonHelper.extractStringValue(rs!QUANTITY))
 txtRetailPrice = CommonHelper.extractStringValue(rs!RETAIL_PRICE)
 txtUnitPrice = CommonHelper.extractStringValue(rs!UNIT_PRICE)
 txtActive = CommonHelper.extractStringValue(rs!active)
 lblCreatedBy = CommonHelper.extractStringValue(rs!CREATED_BY)
 lblCreatedDate = CommonHelper.extractDateValue(rs!CREATED_DATE)
 lblLatModBy = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
 lblLastModDate = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
 txtCriticalLevel = CommonHelper.extractStringValue(rs!CRITICAL_LEVEL)

End Sub

Private Sub formatDataGrid()
  With dgItems
    .Columns(0).Visible = False
  End With
End Sub

Private Sub clearForm()

txtItemCode = ""
cmbSupplier.ListIndex = -1
cmbItemType.ListIndex = -1
txtItemName = ""
txtRetailPrice = ""
txtUnitPrice = ""
lblCreatedBy = ""
lblCreatedDate = ""
lblLastModDate = ""
lblLatModBy = ""
txtCriticalLevel = ""

End Sub

Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    lblQuantity = 0
    cmbNewRec.Caption = "ADD"
    cmdActivation.Enabled = False
    cmbEdit.Enabled = False
  Else
    cmbNewRec.Caption = "New"
    cmdActivation.Enabled = True
    cmbEdit.Enabled = True
  End If
End Sub

