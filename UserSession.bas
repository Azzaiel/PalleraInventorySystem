Attribute VB_Name = "UserSession"
Option Explicit
Public Name As String
Public Role As String
Public UserID As Integer
Private username As String
Public Function getLoginUser() As String

  If (username <> vbNullString) Then
   getLoginUser = username
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
             "WHERE BINARY username = '" & Uname & "' and BINARY Password = '" & Password & "'"
              
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
  If rs.RecordCount = 1 Then
    UserID = rs!id
    Role = rs!Role
    Name = CommonHelper.extractStringValue(rs!FIRST_NAME) & " " & CommonHelper.extractStringValue(rs!MIDDLE_NAME) & " " & CommonHelper.extractStringValue(rs!LAST_NAME)
    username = CommonHelper.extractStringValue(rs!username)
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


