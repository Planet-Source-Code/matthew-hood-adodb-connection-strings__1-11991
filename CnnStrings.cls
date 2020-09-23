VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CnnStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'Wrapper Function to Generate OLEDB Connection Strings
'Dependencies: Microsoft DAO 3.51+ Object Library (For RefreshLinks Only)
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'*****************************************************************************************
'*****************************************************************************************

'*****************************************************************************************
'Microsoft Jet 4.0 Database.
'*****************************************************************************************
Public Function Jet4(ByVal Source As String, Optional ByVal Password As String) As String
Attribute Jet4.VB_Description = "Microsoft Jet 4.0 Database."
  Const cnnstr1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
  Const cnnstr2 As String = ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
  Const cnnstr3 As String = ";Jet OLEDB:Database Password="
  
  If Password = vbNullString Then Jet4 = cnnstr1 & Source & cnnstr2 Else Jet4 = cnnstr1 & Source & cnnstr2 & cnnstr3 & Password
End Function
'*****************************************************************************************
'Microsoft Jet 4.0 Database. (Shaped Heirachy)
'*****************************************************************************************
Public Function Jet4_Shaped(ByVal Source As String, Optional ByVal Password As String) As String
  Const cnnstr1 As String = "Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
  Const cnnstr2 As String = ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
  Const cnnstr3 As String = ";Jet OLEDB:Database Password="
  
  If Password = vbNullString Then Jet4_Shaped = cnnstr1 & Source & cnnstr2 Else Jet4_Shaped = cnnstr1 & Source & cnnstr2 & cnnstr3 & Password
End Function
'*****************************************************************************************
'Microsoft Visual FoxPro - Free Tables.
'*****************************************************************************************
Public Function VFP_FT(ByVal Source As String) As String
Attribute VFP_FT.VB_Description = "Microsoft Visual FoxPro - Free Tables."
  Const cnnstr1 As String = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=Visual FoxPro Tables;UID=;SourceDB="
  Const cnnstr2 As String = ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=No;Deleted=No;"""

  VFP_FT = cnnstr1 & Source & cnnstr2
End Function
Public Function VFP_FT_ODBC(ByVal Source As String) As String
  Const cnnstr1 As String = "ODBC;DSN=Visual FoxPro Tables;UID=;SourceDB="
  Const cnnstr2 As String = ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=No;Deleted=No;DATABASE="
  
  VFP_FT_ODBC = cnnstr1 & Source & cnnstr2
End Function
'*****************************************************************************************
'DBase IV Database Files.
'*****************************************************************************************
Public Function DBaseIV(ByVal Source As String) As String
  Const cnnstr1 As String = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=dBASE Files;DBQ="
  Const cnnstr2 As String = ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"""

  DBaseIV = cnnstr1 & Source & cnnstr2
End Function
'*****************************************************************************************
'Refreshes linked Jet Database Tables.
'*****************************************************************************************
Public Sub RefreshLinks(ByVal JetDB As String, ByVal Source As String)
On Error GoTo On_Error
  Dim db As Database
  Dim sCnn As String
  Dim tdf As TableDef
  
  Set db = DBEngine.OpenDatabase(JetDB)

  sCnn = CnnStrings.VFP_FT_ODBC(Source)
  If sCnn = vbNullString Then GoTo On_Exit
 
  For Each tdf In db.TableDefs
    If Left$(tdf.Name, 3) = "AC_" Then
      tdf.Connect = sCnn
      tdf.RefreshLink
    End If
  Next tdf
 
On_Exit:
On Error Resume Next
  db.Close
  Set db = Nothing
  Exit Sub
On_Error:
  Select Case Err.Number
    Case Else
      MsgBox Err.Number & ": " & Err.Description, vbExclamation, "RefreshLinks"
      Resume On_Exit
  End Select
End Sub
