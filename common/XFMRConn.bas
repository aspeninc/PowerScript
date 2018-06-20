' ASPEN PowerScrip sample program
'
' XFMRConn.BAS
'
' Check phase shift of all 2-winding transformers with wye-delta connection.
' When a transformer with high side lagging the low side is found, 
' this function converts it to make the high side lead.
'
' Version 1.0
'

Sub main()
  XFMRHnd& = 0
  count = 0
  fixed = 0
  PrintTTY( " " )
  PrintTTY("Fixed transformer with incorrect phase shift:")
  Do While GetEquipment( TC_XFMR, XFMRHnd ) = 1
    If GetData( XFMRHnd, XR_sCfg1, ConfigA$ ) = 0 Then GoTo HasError
    If GetData( XFMRHnd, XR_sCfg2, ConfigB$ ) = 0 Then GoTo HasError
    nPosA% = InStr(1, ConfigA$, "G")
    nPosB% = InStr(1, ConfigB$, "D")
    nPosC% = InStr(1, ConfigB$, "E")  
    If nPosA > 0 Then 
      count = count + 1
      If nPosB > 0 Then  ' Y_D_LAG
        If GetData( XFMRHnd, XR_dTap1, TapKVA ) = 0 Then GoTo HasError
        If GetData( XFMRHnd, XR_dTap2, TapKVB ) = 0 Then GoTo HasError
        If TapKVA < TapKVB Then
          If SetData( XFMRHnd, XR_sCfg2, "E" ) = 0 Then GoTo HasError
          If PostData( XFMRHnd ) = 0 Then GoTo HasError            
          PrintTTY(FullXFMRName( XFMRHnd ))
          fixed = fixed + 1
        End If
      End If
      
      If nPosC > 0 Then  ' Y_D_LEAD
        If GetData( XFMRHnd, XR_dTap1, TapKVA ) = 0 Then GoTo HasError
        If GetData( XFMRHnd, XR_dTap2, TapKVB ) = 0 Then GoTo HasError
        If TapKVA > TapKVB Then
          If SetData( XFMRHnd, XR_sCfg2, "D" ) = 0 Then GoTo HasError
          If PostData( XFMRHnd ) = 0 Then GoTo HasError
          PrintTTY(FullXFMRName( XFMRHnd ))
          fixed = fixed + 1
        End If 
      End If  
    End If 
  Loop
  If fixed = 0 Then 
    printTTY( "None" )
    Print "All wye-delta transformers in this OLR file have phase shift with high-side leading low-side."
  Else
    Print "Fixed " & Str(fixed) & " wye-delta transformers such that the high side leads the low side. Details are in TTY."
  End If
  exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Function FullXFMRName( ByVal XFMRHnd& ) As String
 Call GetData( XFMRHnd, XR_nBus1Hnd, Bus1Hnd& )
 Call GetData( XFMRHnd, XR_nBus2Hnd, Bus2Hnd& )
 Call GetData( XFMRHnd, XR_sID, XFMRID$ )
 sID = " T" 
 FullXFMRName$ = FullBusName( Bus1Hnd ) & "-" & FullBusName( Bus2Hnd ) & " " & XFMRID$ & sID
End Function