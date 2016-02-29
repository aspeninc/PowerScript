' ASPEN Sample Script 
' IslandBus.BAS
'
' Make change to breaker data using Change file
'
' V 1.0
'
' Usage:
' 1- Read TTY Report that command Set Generator Reference Angle produces
' 2- Find all buses in the report that are listed as "Islanded from the reference bus"
' 3- Print a report with islanded bus no., bus name, kV, area number, zone number and bus location.
'    
'
Sub main()
  
 ' InputFile$  = FileOpenDialog( "", "Text File (*.txt)||", 0 )
 ' OutputFile$ = FileSaveDialog( "", "Text File (*.txt)||","txt", 16 )
 InputFile$  = InputBox("Enter input file full path name")
 OutputFile$ = InputBox("Enter output file full path name")
   
 If Len(InputFile) = 0 Then 
 Print "Bye"
  Stop
 End If
  
 Open InputFile For Input As 1
 Open OutputFile For Output As 2
 
 Print #2, "Island Bus Report"
 Print #2, "=================================================================="
 Print #2, "No.    Bus Name        kV Rating       Area    Zone       Location"

 nIslandBusCount = 0
 Do While Not EOF(1)
  Line Input #1, aLine$ ' Read a line of data.
  aRecord$ = aLine
   
  If InStr( aLine$, "Unable to determine the generator reference angle at " ) > 0 Then
   nIslandBusCount = nIslandBusCount + 1
   nStart = 53
   nEnd = InStr( aLine$, "kV." )
   BusInfo = Trim(Mid(aLine$,nStart,nEnd-nStart))   
   nEnd = Len(BusInfo)
   For ii& = 1 to nEnd
    temp = Trim(Mid(BusInfo$,nEnd-ii+1,1))
    If (StrComp(temp,"") = 0 ) Then exit For
   Next 
   nStart = nEnd-ii+2
   BusVolt# =  Val(Trim(Mid(BusInfo$,nStart,nEnd-nStart+1)))
   nEnd = nStart-1
   For ii& = 1 to nEnd
    temp = Trim(Mid(BusInfo$,ii,1))
    If (StrComp(temp,"") = 0 ) Then exit For
   Next     
   nStart = ii
   BusName$ = Trim(Mid(BusInfo$,nStart,nEnd-nStart+1))
   If findBusByName( BusName, BusVolt, nHnd& ) Then BusHnd = nHnd
   If GetData( BUsHnd&, BUS_nNumber, nNumber# ) = 0 Then GoTo HasError
   If GetData( BusHnd&, BUS_dKVnominal, dKVnominal# ) = 0 Then GoTo HasError
   If GetData( BusHnd&, BUS_nArea, nArea& ) = 0 Then GoTo HasError
   If GetData( BusHnd&, BUS_nZone, nZone& ) = 0 Then GoTo HasError
   If GetData( BusHnd&, BUS_sName, sName$ ) = 0 Then GoTo HasError
   If GetData( BUsHnd&, BUS_sLocation, sLocation$ ) = 0 Then GoTo HasError
   
   nLen1 = 7 - Len(Str(nNumber))
   nLen2 = 15 - Len(sName$)
   nLen3 = 16 - Len(Str(Format(dKVnominal,"#0.0")))
   nLen4 = 8 - Len(Str(nArea))
   nLen5 = 12 - Len(Str(nZone))
   LineOutput$ = Str(nNumber)
   For ii& = 1 to nLen1
    LineOutput$ = LineOutput$ + " "
   Next
   LineOutput$ = LineOutput$ + sName
   For ii& = 1 to nLen2
    LineOutput$ = LineOutput$ + " "
   Next
   LineOutput$ = LineOutput$ + Str(Format(dKVnominal,"#0.0"))
   For ii& = 1 to nLen3
    LineOutput$ = LineOutput$ + " "
   Next
   LineOutput$ = LineOutput$ + Str(nArea)
   For ii& = 1 to nLen4
    LineOutput$ = LineOutput$ + " "
   Next
   LineOutput$ = LineOutput$ + Str(nZone)
   For ii& = 1 to nLen5
    LineOutput$ = LineOutput$ + " "
   Next
   LineOutput$ = LineOutput$ + sLocation
   Print #2, LineOutput$
  End If
 Loop
 Print #2, "=================================================================="
 Print #2, "The total number of island buses is " + Str(nIslandBusCount)
 Print "The generated report has been saved to " + OutputFile$ + " with" + Str(nIslandBusCount) + " island buses included" 
 Close 1
 Close 2
Exit Sub
' Error handling
HasError:
Print ErrorString()
End Sub ' End of Sub Main()

