Attribute VB_Name = "ms_StartUp"

'# Program start from here to not show error about xzip.dll #
Sub Main()
'# save to program folder, xzip.dll #
SaveRes "XZIP.DLL", "CUSTOM", App.Path & "\xzip.dll", 456536
DoEvents
'# register xzip.dll "
Register App.Path & "\xzip.dll"
DoEvents
'# finally show form #
frmMain.Show
End Sub

