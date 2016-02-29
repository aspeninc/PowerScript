' ASPEN PowerScript Sample Program
'
' CHANGE.BAS
'
'
' Demonstrate how to read a change file from a PowerScript program
'
' PowerScript functions called:
'   ReadChangeFile( ChangeFileName$, SilentFlag& ):
'     Inputs:
'      ChangeFileName: full name of change file. If blank, the program
'                      will display File Open dialog
'      SilentFlag: 
'           1 - Apply change file w/o getting user confirmation
'           0 - Get user confirmation before applying changes
'      Output:
'           1 - Success
'           0 - Error
'

Sub Main

  nRetCode = ReadChangeFile( "", 0 )

Print nRetCode

End Sub