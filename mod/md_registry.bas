Attribute VB_Name = "md_registry"
Option Explicit

'# use to read/write registry #
Private Const REG_SZ As Long = 1 'REG_SZ represents a fixed-length text string.
Private Const REG_DWORD As Long = 4 'REG_DWORD represents data by a number that is 4 bytes long.

Public Const HKEY_CLASSES_ROOT = &H80000000 'The information stored here ensures that the correct program opens when you open a file by using Windows Explorer.
Public Const HKEY_CURRENT_USER = &H80000001 'Contains the root of the configuration information for the user who is currently logged on.
Public Const HKEY_LOCAL_MACHINE = &H80000002 'Contains configuration information particular to the computer (for any user).
Public Const HKEY_USERS = &H80000003 'Contains the root of all user profiles on the computer.

'Return values for all registry functions
Private Const ERROR_SUCCESS = 0
Private Const ERROR_NONE = 0

Private Const KEY_QUERY_VALUE = &H1 'Required to query the values of a registry key.
Private Const KEY_ALL_ACCESS = &H3F 'Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.

'API Calls for writing to Registry
'Close Registry Key
 Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
'Create Registry Key
 Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Open Registry Key
 Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Query a String Value
 Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'Query a Long Value
 Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
'Query a NULL Value
 Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
'Enumerate Sub Keys
 Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'Store a Value
 Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Delete Key
 Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
'# use to read/write registry #

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       
       Dim Data As Long
       Dim retval As Long 'Return value of RegQuery functions
       Dim lType As Long 'Determine data type of present data
       Dim lValue As Long 'Long value
       Dim sValue As String 'String value

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       retval = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, Data)
       
       If retval <> ERROR_NONE Then Error 5

       Select Case lType
           ' Determine strings
           Case REG_SZ:
               sValue = String(Data, 0)

               retval = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, Data)
               
               If retval = ERROR_NONE Then
                   vValue = Left$(sValue, Data - 1)
               Else
                   vValue = Empty
               End If
               
           ' Determine DWORDS
           Case REG_DWORD:
               retval = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, Data)
               
               If retval = ERROR_NONE Then vValue = lValue
           
           Case Else
               'all other data types not supported
               retval = -1
       End Select
    
QueryValueExError:
       QueryValueEx = retval
       Exit Function

   End Function

Public Function ReadKey(HKey, sKeyName As String, sValueName As String)
       
  Dim lRetVal As Long         'result of the API functions
  Dim HKeyR As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value

  lRetVal = RegOpenKeyEx(HKey, sKeyName, 0, KEY_QUERY_VALUE, HKeyR) 'Open Key to Query a value
  lRetVal = QueryValueEx(HKeyR, sValueName, vValue) 'Query (determine) the value stored
  
  RegCloseKey (HKeyR) 'Close the Key

  ReadKey = vValue 'whatever text was stored
End Function

Public Function ListKey(HKey, Key) As String()

    Dim strvalue As String 'Variable to hold current enumerated key
    Dim lDataLen As Long 'Length of data
    Dim lResult As Long 'Result of RegEnumKey
    Dim lValueLen As Long
    Dim lCurIdx As Long 'Current Index which gets incremented with each pass through the loop
    Dim lRetVal As Long 'Result of RegOpenKeyEx
    Dim hKeyResult As Long
    Dim k() As String
    lRetVal = RegOpenKeyEx(HKey, Key, 0, KEY_ALL_ACCESS, hKeyResult) 'Open key with Full Access Rights
    If lRetVal = ERROR_SUCCESS Then
      lCurIdx = 0 'Initialise loop counter
      lDataLen = 64 'data Length
      lValueLen = 64

    Do
      strvalue = String(lValueLen, 0) 'get current key's value
         lResult = RegEnumKey(hKeyResult, lCurIdx, strvalue, lDataLen) 'Enumerate keys
         If lResult = ERROR_SUCCESS Then 'if successful, add current enumerated key to the txtEnumKeys textbox
            ReDim Preserve k(lCurIdx) As String
            k(lCurIdx) = Replace(strvalue, Chr(0), "")
         End If

         lCurIdx = lCurIdx + 1 'Increment counter for next enumeration

    Loop While lResult = ERROR_SUCCESS 'continue while successful

         RegCloseKey hKeyResult 'Close key
    Else 'If lRetVal is unsuccessful
      MsgBox "Cannot Open Key"
    End If
    ListKey = k
    
End Function


