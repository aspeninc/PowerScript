' ASPEN PowerScript Sample Program
'
' FindByTag.BAS
'
'
' Demonstrate the FindEquipmentByTag function
'
Sub main
  dim Options(10) As long
  
  ' OLR file
  OLRFile$ = "Sample30.olr"
  If 0 = LoadBinary( OLRFile$ ) Then 
    Print "Error opening OLR file"
    Stop
  End If

  ObjHnd& = 0
  
  While FindEquipmentByTag( "Line tag", TC_XFMR, ObjHnd ) > 0
  
   If EquipmentType( ObjHnd ) = TC_RLYGROUP Then
    RelayHnd& = 0
    While GetRelay( ObjHnd, RelayHnd ) > 0
	 Print "Memo: " + GetObjMemo( RelayHnd ) + _
	  Chr(13) + Chr(10) + "Tags: " + GetObjTags( RelayHnd )
    Wend
   Else
    Print "Memo: " + GetObjMemo( ObjHnd ) + _
	  Chr(13) + Chr(10) + "Tags: " + GetObjTags( ObjHnd )
   End If
  Wend
   Exit Sub
 HasError:
   Print "Error: ", ErrorString( )
   
End Sub
