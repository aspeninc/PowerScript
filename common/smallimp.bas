' ASPEN PowerScript Sample Program
'
' SMALLIMP.BAS
'
' This program replaces all small line reactace X with 0.001 pu.
'
'
Sub main()

   ' Get line handle
   LineHnd& = 0
   LineCount = 0
   ChangeCount = 0
   While GetEquipment( TC_LINE, LineHnd& ) > 0 
      LineCount = LineCount + 1
      Call GetData( LineHnd, LN_dX, dX1# )
      Call GetData( LineHnd, LN_dR, dR1# )
      If (dX1 < 0.001) And (dR1 < 0.001) Then 
        ChangeCount = ChangeCount + 1
        Call GetData( LineHnd, LN_nBus1Hnd, Bus1Hnd& )
        Call GetData( LineHnd, LN_nBus2Hnd, Bus2Hnd& )
        Call GetData( LineHnd, LN_sID, sID$ )
        dX1 = 0.001
        Call SetData( LineHnd, LN_dX, dX1# )
        If PostData( LineHnd ) = 0 Then GoTo HasError
        sLine$ = FullBusName(Bus1Hnd&) & "-" & FullBusName(Bus2Hnd&) & sID
        PrintTTY( sLine$ )
      End If
   Wend
   Print "LineCount=", LineCount, "Changed=", ChangeCount, ". Details in TTY"
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
End Sub
