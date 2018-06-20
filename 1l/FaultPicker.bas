' ASPEN PowerScript Sample Program
'
' FAULTPICKER.BAS
'
' Demo the function FaultSelector()
'
' Version 1.0
' Category: OneLiner
'
Sub main

 dim nFltIdx(10) As long
 
 nCount = FaultSelector( nFltIdx, "My Fault Selector", "Please Select One Fault" )
 
 If nCount > 0 Then
  sTemp = ""
  For ii = 0 to nCount - 1
   sTemp = sTemp + Str(nFltIdx(ii)) + " "
  Next
  Print "Selected fault indices: " + sTemp
 End If

End Sub
