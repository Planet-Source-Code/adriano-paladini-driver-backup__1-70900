'###############################
'#                             #
'#       Drivers BackUp        #
'#             By              #
'#  adrianopaladini@gmail.com  #
'#                             #
'#        version 1.0          #
'#                             #
'###############################



 To create ZIP and Self-installer i use
 xceedzip component. to use project rename 
 "xzip.ckk" in component folder to "xzip.dll"
 (this file is not required on compiled exe,
 because it has this file in resource.)


 The DriverBackup.exe contains in resource
 the icon, sfx, dlls, and setup.exe to make
 ZIP and EXE.

 
 To make changes in Self-Installer, use the
 setup.vbp in "self" folder, modify, compile
 and insert "setup.exe" in resource.res, open
 DriverBackUp.vbp and compile the main project.

 

 Files is resource.res:

 ICON.ICO      <- icon used in self-installer
 MSVBVM60.DLL  <- for run setup.exe in any windows
 SETUP.EXE     <- setup program. project in "self" folder
 SFX.BIN       <- binary for create self-installer
 XZIP.DLL      <- xceed dll to create zip and sfx file

