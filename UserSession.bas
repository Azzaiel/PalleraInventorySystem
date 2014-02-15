Attribute VB_Name = "UserSession"
Option Explicit
Public Name As String
Public Role As String
Public userID As Integer
Private Username As String
Public Function getLoginUser() As String

  If (Username <> vbNullString) Then
   getLoginUser = Username
  Else
    getLoginUser = "System"
  End If
  
End Function

Public Function hasValidCredential(Username As String, Password As String)
Dim con As ADODB.Connection
Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  
  sqlQuery = "SELECT id,username, role, first_name, last_name, middle_name " & _
             "FROM users " & _
             "WHERE username = '" & Username & "' and Password = '" & Password & "'"
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
  If rs.RecordCount = 1 Then
    userID = rs!ID
    Role = rs!Role
    Name = rs!First_name & " " & rs!Middle_name & " " & rs!Last_name
    hasValidCredential = True
  
  Else
    hasValidCredential = False
  End If
End Function


