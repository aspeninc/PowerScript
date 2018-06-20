' ASPEN PowerScript Sample Program
'
' BrkDontDerate.BAS
'
' Search for all the breakers in the case and change all the breaker 
' setting "do not derate in reclosing operation flag" to true.
'
' Category: OneLiner
' Version 1.0
'
'========================== 
Const nFlagOld& = 1
Const nFlagNew& = 0
Sub main()
   nBrk = 0
   nError = 0
   BrkHnd& = 0
   nRet = GetEquipment( TC_BREAKER, BrkHnd& )
   Print "Error: ", ErrorString( )
   While GetEquipment( TC_BREAKER, BrkHnd& ) > 0
     Call GetData( BrkHnd&, BK_nDontDerate, nDontDerate& )
     If nDontDerate = nFlagOld Then
       If SetData( BrkHnd&, BK_nDontDerate, nFlagNew ) = 0 Then GoTo HasError
       If PostData( BrkHnd ) = 0 Then nError = nError + 1 Else nBrk = nBrk + 1
     End If   
   Wend
   Print nBrk & " breakers have been updated." & nError & " failed."
   Stop
   HasError:
   Print "Error: ", ErrorString( )
End Sub
