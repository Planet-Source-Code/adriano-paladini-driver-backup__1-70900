Attribute VB_Name = "md_resource"

Public Function SaveRes( _
            ByVal iResourceNum As String, _
            ByVal sResourceType As String, _
            ByVal sDestFileName As String, _
            Optional ByVal dBytes As Double) As Long
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    
    'Example Call:
    ' iRetVal = SaveRes("101", "CUSTOM", "C:\myImage.gif")
    
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
        
    On Error GoTo SaveRes_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    
    'Get Free File Handle
    iFileNumOut = FreeFile
    
    
    '# BY adrianopaladini #
    '# This code is to correct an crazy error that occurs in the #
    '# program after compiled. It´s save 2 bytes to more in the files. #
    If dBytes > 0 Then
        If Hex(bytResourceData(UBound(bytResourceData))) = 0 Then
            ReDim Preserve bytResourceData(dBytes - 1)
        End If
    End If
    '# BY adrianopaladini #
    
    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut
        ''Write the resource to the file
        Put #iFileNumOut, , bytResourceData
    
    'Close the file
    Close #iFileNumOut
    
    'Return 0 for success
    SaveRes = 0
    
    Exit Function
SaveRes_err:
    'Return error number
    SaveRes = Err.Number
End Function

