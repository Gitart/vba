Attribute VB_Name = "Module9"
Option Compare Database
Option Explicit

'*********************** Code Start ***************************
'
Const JDS_DSN_name = "MDTS"
Const JDS_Server_name = "148.154.61.15"   ' Raw IP address is used to avoid NT _
                                                                           Domain name resolution probs.

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" _
Alias "RegEnumKeyExA" _
  (ByVal hKey As Long, _
   ByVal dwIndex As Long, _
   ByVal lpName As String, _
   lpcbName As Long, _
   ByVal lpReserved As Long, _
   ByVal lpClass As String, _
   lpcbClass As Long, _
   ByVal lpftLastWriteTime As String) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function SQLConfigDataSource Lib "odbccp32.dll" _
    (ByVal hwndParent As Long, _
    ByVal fRequest As Integer, _
    ByVal lpszDriver As String, _
    ByVal lpszAttributes As String) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

      Const HKEY_LOCAL_MACHINE = &H80000002

      Const ERROR_SUCCESS = 0&
      Const SYNCHRONIZE = &H100000
      Const STANDARD_RIGHTS_READ = &H20000
      Const STANDARD_RIGHTS_WRITE = &H20000
      Const STANDARD_RIGHTS_EXECUTE = &H20000
      Const STANDARD_RIGHTS_REQUIRED = &HF0000
      Const STANDARD_RIGHTS_ALL = &H1F0000
      Const KEY_QUERY_VALUE = &H1
      Const KEY_SET_VALUE = &H2
      Const KEY_CREATE_SUB_KEY = &H4
      Const KEY_ENUMERATE_SUB_KEYS = &H8
      Const KEY_NOTIFY = &H10
      Const KEY_CREATE_LINK = &H20
      Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                        KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or _
                        KEY_NOTIFY) And _
                        (Not SYNCHRONIZE))

      Const REG_DWORD = 4
      Const REG_BINARY = 3
      Const REG_SZ = 1
      
      Const ODBC_ADD_SYS_DSN = 4
      

Function Check_SDSN()

         '  Look for our System Data Source Name.  If we find it, then great!
         '  If not, then let's create one on the fly.
         
         Dim lngKeyHandle As Long
         Dim lngResult As Long
         Dim lngCurIdx As Long
         Dim strValue As String
         Dim classValue As String
         Dim timeValue As String
         Dim lngValueLen As Long
         Dim classlngValueLen As Long
         Dim lngData As Long
         Dim lngDataLen As Long
         Dim strResult As String
         Dim DSNfound As Long
         Dim syscmdresult As Long

         syscmdresult = SysCmd(acSysCmdSetStatus, "Looking for System DSN " & JDS_DSN_name & " ...")
         
         '  Let's open the registry key that contains all of the
         '  System Data Source Names.
         
         lngResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                 "SOFTWARE\ODBC\ODBC.INI", _
                  0&, _
                  KEY_READ, _
                  lngKeyHandle)

         If lngResult <> ERROR_SUCCESS Then
             MsgBox "ERROR:  Cannot open the registry key HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI." & vbCrLf & vbCrLf & _
                    "Please make sure that ODBC and the SQL Server ODBC drivers have been installed." & vbCrLf & _
                    "Contact call your MDTS System Administrator for more information."
             syscmdresult = SysCmd(acSysCmdClearStatus)
             Check_SDSN = -1
         End If
         
         ' Now that the key is open, Let's look among all of
         ' the possible system data source names for the one
         ' we want.

         lngCurIdx = 0
         DSNfound = False
         
         Do
            lngValueLen = 512
            classlngValueLen = 512
            strValue = String(lngValueLen, 0)
            classValue = String(classlngValueLen, 0)
            timeValue = String(lngValueLen, 0)
            lngDataLen = 512

            lngResult = RegEnumKeyEx(lngKeyHandle, _
                                     lngCurIdx, _
                                     strValue, _
                                     lngValueLen, _
                                     0&, _
                                     classValue, _
                                     classlngValueLen, _
                                     timeValue)
            lngCurIdx = lngCurIdx + 1

         If lngResult = ERROR_SUCCESS Then
         
           ' Is this our System Data Source Name?
         
           If strValue = JDS_DSN_name Then
           
             '  It is!  Let's assume everything is good and do nothing.
             
             DSNfound = True
             syscmdresult = SysCmd(acSysCmdClearStatus)
         
           End If
           
         End If

         Loop While lngResult = ERROR_SUCCESS And Not DSNfound
         
         Call RegCloseKey(lngKeyHandle)

         If Not DSNfound Then
         
           '  Our System Data Source Name doesn't exist, so let's
           '  try to create it on the fly.
         
           syscmdresult = SysCmd(acSysCmdSetStatus, "Creating System DSN " & JDS_DSN_name & "...")
         
           lngResult = SQLConfigDataSource(0, _
                                           ODBC_ADD_SYS_DSN, _
                                           "SQL Server", _
                                           "DSN=" & JDS_DSN_name & Chr(0) & _
                                           "Server=" & JDS_Server_name & Chr(0) & _
                                           "Database=SvCvMarketing" & Chr(0) & _
                                           "UseProcForPrepare=Yes" & Chr(0) & _
                                           "Description=MDTS Database" & Chr(0) & Chr(0))
                       
           If lngResult = False Then
           
             MsgBox "ERROR:  Could not create the System DSN " & JDS_DSN_name & "." & vbCrLf & vbCrLf & _
                    "Please make sure that the SQL Server ODBC drivers have been installed." & vbCrLf & _
                    "Contact your MDTS System Administrator for more information."
                    
             syscmdresult = SysCmd(acSysCmdClearStatus)
             Check_SDSN = -1

           End If
           
         End If
         
         syscmdresult = SysCmd(acSysCmdClearStatus)
         Check_SDSN = 0

End Function


