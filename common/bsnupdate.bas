' ASPEN PowerScript sample program
'
' BSNUPDATE.BAS
'
' Update bus number. Data in BSNFileName$
' must be in comma delimited format: 
'  OldNumber1,NewNumber1
'  OldNumber2,NewNumber2
' The list must be sorted by OldNumber in ascending order
'
'
Const BSNFileName$ = "c:\0tmp\bsn.txt"
  
Sub main()
  Open BSNFileName$ For Input As 1
  Count   = 0
  Done    = 0 
  BSNThis& = 0
  BusHnd& = 0	'This will make the NextBus cmd to seek the first bus
  While Not EOF(1)
   Line Input #1, ALine$
   Pos& = InStr( 1, ALine$, "," )
   If Pos& < 1 Then
     Done = 1
     exit Do
   End If
   sTmp$   = Left$(ALine$, Pos&-1)
   BSNOld& = Val(sTmp$)
   sTmp$   = Right$(Aline$, Len(ALine$)-Pos&)
   BSNNew& = Val(sTmp$)
   Do While (BSNThis& < BSNOld&)
    If NextBusByNumber( BusHnd& ) > 0 Then
     Call GetData( BusHnd&, BUS_nNumber, BSNThis& )
    Else
     exit Do
    End If
   Loop
   If Done = 1 Then exit Do
   If BSNThis& = BSNOld& Then
     Call SetData(BusHnd&, BUS_nNumber, BSNNew&)
     If PostData(BusHnd&) = 1 Then
      PrintTTY(ALine$ & ": OK")
      Count = Count + 1
     Else
      PrintTTY(ALine$ & ": FAILED. Duplicate bus number")
     End If
   Else
     PrintTTY(ALine$ & ": FAILED. Bus number not found")
   End If
  Wend
  If Count > 0 Then
    Print Count, " buses have been modified. Full list is in TTY"
  Else
    Print "No change made"
  End If
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub