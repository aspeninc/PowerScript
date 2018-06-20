' ASPEN Sample Script
' SETDATA.BAS
'
' Demo SetData and PostData script functions
'
' Version: 1.0
' Category: OneLiner

Sub main
 If GetEquipment( TC_PICKED, PickedHnd& ) = 0 Or _
    EquipmentType( PickedHnd& ) <> TC_XFMR Then
    Print "Please select a XFMR"
    Stop
  End If

  GoTo Skip
  sVal$ = "My"
  Call SetData( PickedHnd, XR_sName, sVal$ )
  Call PostData( PickedHnd )
  Call GetData( PickedHnd, XR_sName, sValNew$ )
  Print sValNew
  nID = XR_sCfgP
  sVal$ = "D"
  Call SetData( PickedHnd, nID, sVal$ )
  Call PostData( PickedHnd )
  Call GetData( PickedHnd, nID, sValNew$ )
  Print sValNew

  nID = XR_sCfgS
  sVal$ = "G"
  Call SetData( PickedHnd, nID, sVal$ )
  Call PostData( PickedHnd )
  Call GetData( PickedHnd, nID, sValNew$ )
  Print sValNew
sKip:
  nID = XR_nAuto
  nVal& = "1"
  Call SetData( PickedHnd, nID, nVal )
  Call PostData( PickedHnd )
  Call GetData( PickedHnd, nID, nValNew& )
  Print nValNew

End Sub
