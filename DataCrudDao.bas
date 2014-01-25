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

Public Function getItemTypeRS(itemType As String, Supplier As String, Category As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "SELECT SUPPLIERS,CATEGORY ,ITEM_TYPE " & _
              "       , CREATED_DATE, LAST_MOD_BY, LAST_MOD_DATE " & _
              "FROM item_type "
              
              
    
    If (CommonHelper.hasValidValue(Supplier)) Then
       sqlQuery = sqlQuery & " And SUPPLIER like '" & Supplier & "%' "
    End If
    
    If (CommonHelper.hasValidValue(Category)) Then
       sqlQuery = sqlQuery & " And CATEGORY like '" & Category & "%' "
    End If
    
    If (CommonHelper.hasValidValue(itemType)) Then
       sqlQuery = sqlQuery & " And ITEM_TYPE like '" & itemType & "%' "
    End If
              
    sqlQuery = sqlQuery & " Order By LAST_MOD_DATE desc"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getItemTypeRS = rs

End Function
Public Function getFakeUserRS() As ADODB.Recordset


   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT * " & _
             "FROM users " & _
             "Where 1 = 2 "
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
   
  Set getFakeUserRS = rs
  
End Function
Public Function getAccount(ID As String, USERNAME As String, PASSWORD As String, ROLE As String, FIRSTNAME As String, LASTNAME As String, MIDDLENAME As String) As ADODB.Recordset

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



