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

