Attribute VB_Name = "md_folders"
Option Explicit

'# use to read folders #
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'# use to read folders #

Public Function Getpath_WINDOWS()
Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                    
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING
Dim Result                          'a variable for holding the the output of the function
TEMP = GetWindowsDirectory(WindirS, 255)     'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                 'holds final path
Getpath_WINDOWS = Result
End Function

Public Function Getpath_SYSTEM()
Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                        
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
Dim Result                          'a variable for holding the the output of the function
TEMP = GetSystemDirectory(WindirS, 255)      'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                 'holds final path
Getpath_SYSTEM = Result
End Function

Public Function Getpath_TEMP()
'this API(TEMP) is different from others(in placing arguments)

Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                        
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
Dim Result                          'a variable for holding the the output of the function
TEMP = GetTempPath(255, WindirS)            'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                'holds final path
Getpath_TEMP = Result
End Function

