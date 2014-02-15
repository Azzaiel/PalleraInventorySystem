Attribute VB_Name = "UserSession"
Option Explicit
Public Name As String
Public Role As String
Public UserID As Integer
Private Username As String
Public Function getLoginUser() As String

  If (Username <> vbNullString) Then
   getLoginUser = Username
  Else
    getLoginUser = "System"
  End If
  
End Function

Public Function hasValidCredential(Uname As String, Password As String)
Dim con As ADODB.Connection
Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT id,username, role, first_name, last_name, middle_name " & _
             "FROM users " & _
             "WHERE username = '" & Uname & "' and Password = '" & Password & "'"
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
  If rs.RecordCount = 1 Then
    UserID = rs!ID
    Role = rs!Role
    Name = CommonHelper.extractStringValue(rs!First_name) & " " & CommonHelper.extractStringValue(rs!Middle_name) & " " & CommonHelper.extractStringValue(rs!Last_name)
    Username = CommonHelper.extractStringValue(rs!Username)
    hasValidCredential = True
  
  Else
    hasValidCredential = False
  End If
End Function

Public Function getUserRS(Password As String) As ADODB.Recordset
Dim con As ADODB.Connection
Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT * " & _
             "FROM users " & _
             "WHERE ID = " & UserID & " and Password = '" & Password & "'"
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
  Set getUserRS = rs
End Function


