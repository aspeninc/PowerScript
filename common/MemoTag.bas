' ASPEN PowerScript sample program
'
' MEMOTAGS.BAS
'
' Demonstrate access memo and tag data from a PowerScript program
' 
'
Sub main()

   If GetEquipment( TC_PICKED, ObjHnd& ) = 0 Then
     Print "Please select an object"
     Exit Sub
   End If

   If EquipmentType( ObjHnd ) = TC_RLYGROUP Then
    RelayHnd& = 0
    While GetRelay( ObjHnd, RelayHnd ) > 0
	 Print "Memo: " + GetObjMemo( RelayHnd ) + _
	  chr(13) + chr(10) + "Tags: " + GetObjTags( RelayHnd )
    Wend
   Else
    Print "Memo: " + GetObjMemo( ObjHnd ) + _
	  Chr(13) + Chr(10) + "Tags: " + GetObjTags( ObjHnd )
   End If
   
   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
End Sub