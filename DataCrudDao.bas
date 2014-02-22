Attribute VB_Name = "DataCrudDao"
Option Explicit
Public Function getSupplierRS(active As String, suplierName As String, salesContact As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "SELECT ID, Name, ACTIVE, COMPANY_PHONE_NUMBER, COMPANY_ADDRESS " & _
              "       , SALES_CONTACT, SALES_EMAIL, SALES_PHONE_NUMBER, CREATED_BY " & _
              "       , CREATED_DATE, LAST_MOD_BY, LAST_MOD_DATE " & _
              "FROM suppliers " & _
              "Where 1 = 1"
              
              
    If (CommonHelper.hasValidValue(active)) Then
       sqlQuery = sqlQuery & " And ACTIVE = '" & active & "' "
    End If
    
    If (CommonHelper.hasValidValue(suplierName)) Then
       sqlQuery = sqlQuery & " And Name like '" & suplierName & "%' "
    End If
    
    If (CommonHelper.hasValidValue(salesContact)) Then
       sqlQuery = sqlQuery & " And SALES_CONTACT like '" & salesContact & "%' "
    End If
              
    sqlQuery = sqlQuery & " Order By LAST_MOD_DATE desc"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getSupplierRS = rs

End Function

Public Function getItemTypeRS(itemType As String, Supplier_name As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, b.name as SUPPLIER_NAME, a.name as ITEM_TYPE_NAME, a.CREATED_BY " & _
              "       , a.CREATED_DATE, a.LAST_MOD_BY, a.LAST_MOD_DATE " & _
              "From supplier_item_types a, suppliers b " & _
              "Where a.SUPPLIER_ID = b.ID "
              
              
    
    If (CommonHelper.hasValidValue(Supplier_name)) Then
       sqlQuery = sqlQuery & " And b.name like '" & Supplier_name & "%' "
    End If

    
    If (CommonHelper.hasValidValue(itemType)) Then
       sqlQuery = sqlQuery & " And a.name like '" & itemType & "%' "
    End If
              
    sqlQuery = sqlQuery & " Order By LAST_MOD_DATE desc"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getItemTypeRS = rs

End Function
Public Function getFakeItemTypeRS() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT * " & _
             "FROM supplier_item_types " & _
             "Where 1 = 2 "
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
  Set getFakeItemTypeRS = rs
  
End Function

Public Function getItemTypeRSByID(id As Long) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT * " & _
             "FROM supplier_item_types " & _
             "Where ID = " & id
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
  Set getItemTypeRSByID = rs
  
End Function

Public Function getItemTypeRSBySupplierID(supplierID As Long) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT ID, name as ITEM_TYPE_NAME " & _
             "FROM supplier_item_types " & _
             "Where supplier_id = " & supplierID
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
  Set getItemTypeRSBySupplierID = rs
  
End Function
Public Function getAccount() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "SELECT ID, USERNAME, ROLE, FIRST_NAME, MIDDLE_NAME, LAST_NAME ,PASSWORD " & _
              "FROM users"
              
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getAccount = rs

End Function


Public Function getItemReg(Optional itemCode As String) As ADODB.Recordset

Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.ACTIVE, a.ITEM_CODE, b.Name as SUPPLIER , c.Name as ITEM_TYPE, a.Name as ITEM_NAME " & _
              "       , a.Quantity, a.RETAIL_PRICE, a.UNIT_PRICE, a.CREATED_BY , a.CREATED_DATE, a.LAST_MOD_BY " & _
              "       , a.LAST_MOD_DATE " & _
              "From items a, SUPPLIERS b, supplier_item_types c " & _
              "Where a.SUPPLIER_ID = b.ID " & _
              "      and a.ITEm_TYPE_ID = c.ID"
    If CommonHelper.hasValidValue(itemCode) Then
        sqlQuery = sqlQuery & " And a.item_code = '" & itemCode & "'"
    End If
    
            
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getItemReg = rs


End Function

Public Function getItemByItemsRS(itemTypeID As Long) As ADODB.Recordset

Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.ACTIVE, a.ITeM_CODE, b.Name as SUPPLIER , c.Name as ITEM_TYPE, a.Name as ITEM_NAME " & _
              ", a.RETAIL_PRICE, a.UNIT_PRICE, a.CREATED_BY , a.CREATED_DATE, a.LAST_MOD_DATE " & _
              ", a.LAST_MOD_DATE " & _
              "From items a, SUPPLIERS b, supplier_item_types c " & _
              "Where a.SUPPLIER_ID = b.ID " & _
              "      and a.ITEm_TYPE_ID = c.ID" & _
              "      and a.ITEm_TYPE_ID = " & itemTypeID

              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getItemByItemsRS = rs

End Function



   
Public Function getFakeItemsRS() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT * " & _
             "FROM items " & _
             "Where 1 = 2 "
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
  Set getFakeItemsRS = rs
  
End Function
Public Function getItemRSByID(itemID As Long) As ADODB.Recordset

Dim con As ADODB.Connection
Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT * " & _
             "FROM items " & _
             "Where ID =  " & itemID
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
  Set getItemRSByID = rs
  
End Function
Public Function getFakeOrdersRs() As ADODB.Recordset
 
 Dim con As ADODB.Connection
 Set con = DbInstance.getDBConnetion
 
 Dim sqlQuery As String
  
 sqlQuery = "Select * " & _
            "From orders " & _
            "Where 1 = 2"
              
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
  
 rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
 Set getFakeOrdersRs = rs

End Function
Public Function getOrderByIDRs(orderID As Integer) As ADODB.Recordset
 
 Dim con As ADODB.Connection
 Set con = DbInstance.getDBConnetion
 
 Dim sqlQuery As String
  
 sqlQuery = "Select * " & _
            "From orders " & _
            "Where id = " & orderID
              
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
  
 rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
 Set getOrderByIDRs = rs

End Function

Public Function getPendingOrdersRs() As ADODB.Recordset
 
 Dim con As ADODB.Connection
 Set con = DbInstance.getDBConnetion
 
 Dim sqlQuery As String
  
 sqlQuery = "Select a.ID as Order_Id,  b.name as Suplier_Name " & _
            "       , a.Status, a.Ordered_by, a.Order_Date " & _
            "       , a.RECIVED_BY, a.Recived_Date " & _
            "From orders a, suppliers b " & _
            "Where a.SUPLIER_ID = b.id "
              
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
  
 rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
 Set getPendingOrdersRs = rs

End Function
Public Function getFakeOrderItems() As ADODB.Recordset
 
 Dim con As ADODB.Connection
 Set con = DbInstance.getDBConnetion
 
 Dim sqlQuery As String
  
 sqlQuery = "Select * " & _
            "From order_items " & _
            "Where 1 = 2 "
              
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
  
 rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
 Set getFakeOrderItems = rs
 
End Function
Public Function getOrderItemsByID(orderItemID As Integer) As ADODB.Recordset
 
 Dim con As ADODB.Connection
 Set con = DbInstance.getDBConnetion
 
 Dim sqlQuery As String
  
 sqlQuery = "Select * " & _
            "From order_items " & _
            "Where ID = " & orderItemID
              
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
  
 rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
 Set getOrderItemsByID = rs
 
End Function
Public Function getOrderItemsByOrderID(orderID As Integer) As ADODB.Recordset
 
 Dim con As ADODB.Connection
 Set con = DbInstance.getDBConnetion
 
 Dim sqlQuery As String
  
 sqlQuery = "Select oi.id, i.id as ITEM_ID, sit.name as Item_Type, i.name, oi.retil_price " & _
            "       , oi.quantity, oi.retil_price *  oi.quantity as TOTAL_PRICE " & _
            "       , oi.CREATED_BY, oi.CREATED_DATE " & _
            "       , oi.LAST_MOD_BY, oi.LAST_MOD_DATE " & _
            "From order_items oi, items i, supplier_item_types sit " & _
            "Where oi.ITEM_ID = i.ID  " & _
            "      and sit.id = oi.ITEM_TYPE_ID " & _
            "      and oi.order_id = " & orderID
            
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
  
 rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
 Set getOrderItemsByOrderID = rs
 
End Function







