<div align="center">

## Dynamically generate an ODBC  DSN


</div>

### Description

Class object that can be compiled or copied and pasted into your application that will dynamically create ODBC DSN's for you.
 
### More Info
 
DSN Name, Server, Database, DSN Type, UserID and Password

0 if success, OSBC DSN is created


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Royce Powers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/royce-powers.md)
**Level**          |Advanced
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 6\.0, ASP \(Active Server Pages\) , VBA MS Access
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/royce-powers-dynamically-generate-an-odbc-dsn__1-24849/archive/master.zip)

### API Declarations

```
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv _
 As Long, phdbc As Long) As Integer
Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As _
 Long) As Integer
Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As _
 Long, ByVal szDSN As String, ByVal cbDSN As Integer, ByVal szUID As _
 String, ByVal cbUID As Integer, ByVal szAuthStr As String, ByVal _
 cbAuthStr As Integer) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As _
 Long) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc _
 As Long) As Integer
Declare Function SQLError Lib "odbc32.dll" (ByVal henv As _
 Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As _
 String, pfNativeError As Long, ByVal szErrorMsg As String, ByVal _
 cbErrorMsgMax As Integer, pcbErrorMsg As Integer) As Integer
Declare Function SQLConfigDataSource Lib "ODBCCP32" _
 (ByVal hwndParent As Long, ByVal fRequest As Long, _
 ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
```


### Source Code

```
IN MODULE (.BAS)
Option Explicit
Public Const vbAPINull As Long = 0&
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_SUCCESS_WITH_INFO As Long = 1
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv _
 As Long, phdbc As Long) As Integer
Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As _
 Long) As Integer
Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As _
 Long, ByVal szDSN As String, ByVal cbDSN As Integer, ByVal szUID As _
 String, ByVal cbUID As Integer, ByVal szAuthStr As String, ByVal _
 cbAuthStr As Integer) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As _
 Long) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc _
 As Long) As Integer
Declare Function SQLError Lib "odbc32.dll" (ByVal henv As _
 Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As _
 String, pfNativeError As Long, ByVal szErrorMsg As String, ByVal _
 cbErrorMsgMax As Integer, pcbErrorMsg As Integer) As Integer
Declare Function SQLConfigDataSource Lib "ODBCCP32" _
 (ByVal hwndParent As Long, ByVal fRequest As Long, _
 ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
In Class (.CLS)
Option Explicit
Public Enum peDSN_OPTIONS
 ODBC_ADD_DSN = 1
 ODBC_CONFIG_DSN = 2
 ODBC_ADD_SYS_DSN = 4
 ODBC_CONFIG_SYS_DSN = 5
End Enum
Public Function RegisterDataSource(iFunction As peDSN_OPTIONS, sDSNName As String, sServerName As String, sDatabasename As String, sUserID As String, sPassword As String) As Integer
 Dim sAttributes As String
 Dim iRetVal As Integer
 sAttributes = "DSN=" & sDSNName _
  & Chr$(0) & "Description=SQL Server on server " & sServerName _
  & Chr$(0) & "SERVER=" & sServerName _
  & Chr$(0) & "Database=" & sDatabasename _
  & Chr$(0) & Chr$(0)
 iRetVal = SQLConfigDataSource(vbAPINull, iFunction, "SQL Server", sAttributes)
End Function
```

