' ASPEN PowerScript Sample Program
'
' LISTBRANCH.BAS
'
' This program creates list of branches at the two ends of a line.
'
' User must select a relay group on the line before running this script
'
'
Sub main()
   ' Variable declaration
   Dim OutageList1(20) As Long
   Dim OutageList2(20) As Long
   
   
   ' Get picked relay group handle
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then 
      Print "Please select a relay group"
      Exit Sub
   End If
   ' Must be a relay group
   If EquipmentType( PickedHnd ) <> TC_RLYGROUP Then
      Print "Please select a relay group"
      Exit Sub
   End If


   ' Get the relay group branch handle
   If GetData( PickedHnd, RG_nBranchHnd, Branch1Hnd ) = 0 Then GoTo HasError
   
   ' Determine selected equipment type code and handle
   If GetData( Branch1Hnd, BR_nType,   TypeCode1 ) = 0 Then GoTo HasError
   If GetData( Branch1Hnd, BR_nHandle, FacilityHandle ) = 0 Then GoTo HasError
   
   ' Get the near bus handle
   If GetData( Branch1Hnd, BR_nBus1Hnd, Bus1Hnd ) = 0 Then GoTo HasError

   ' Get the far bus handle
   If GetData( Branch1Hnd, BR_nBus2Hnd, Bus2Hnd ) = 0 Then GoTo HasError

   If TypeCode1 = TC_LINE Then		' Line may have one or more tap
      ' Get line name
      If GetData( FacilityHandle, LN_sName, LineName ) = 0 Then GoTo HasError
      BusHnd = Bus1Hnd		' temporary near bus handle

      ' This do...loop will skip all taps on the lines
      Do 
         If GetData( Bus2Hnd, BUS_nTapBus, TapCode ) = 0 Then GoTo HasError
         If TapCode = 0 Then Exit Do			' reached a real line end
         ' Look for next line segment at the remote end
         BranchNext   = 0
         BranchHnd    = 0
         While GetBusEquipment( Bus2Hnd, TC_BRANCH, BranchHnd ) > 0
            If GetData( BranchHnd, BR_nBus2Hnd, RemoteBusHnd ) = 0 Then GoTo HasError
            If RemoteBusHnd = BusHnd Then ' Same line segment
              Branch2Hnd = BranchHnd
            Else				' Candidate for next line segment
              If GetData( BranchHnd, BR_nType, TypeCode2 ) = 0 Then GoTo HasError
              If TypeCode2 = TC_LINE Then 
                If BranchNext = 0 Then 
                  BranchNext = BranchHnd
                  BusHnd     = Bus2Hnd
                  Bus2Hnd    = RemoteBusHnd	
                Else
                  If GetData( BranchHnd, BR_nHandle, TmpHandle ) = 0 Then GoTo HasError
                  If GetData( TmpHandle, LN_sName, TmpLineName ) = 0 Then GoTo HasError
                  If TmpLineName = LineName Then 
                    BranchNext = BranchHnd
                    BusHnd     = Bus2Hnd
                    Bus2Hnd    = RemoteBusHnd	
                  End If
                End If
              End If
            End If 
         Wend
      Loop While BranchNext <> 0
   
   End If

   ' Near bus and branches 
   BusHnd     = Bus1Hnd
   BusName    = FullBusName( BusHnd )
   BranchThis = Branch1Hnd
   List1Len   = 0
   BranchHnd  = 0
   While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd ) > 0
     If BranchHnd <> BranchThis Then
       ' Store branch in outage list
       List1Len = List1Len + 1
       OutageList1(List1Len) = BranchHnd
       ' The code below prints branch info on the screen for visual check
       If GetData( BranchHnd, BR_nBus2Hnd, RemoteBusHnd ) = 0 Then GoTo HasError
       BusName2 = FullBusName( RemoteBusHnd )
       If GetData( BranchHnd, BR_nType, TypeCode ) = 0 Then GoTo HasError
       select case TypeCode
          case TC_LINE 
            BranchType = "LINE"
          case TC_XFMR
            BranchType = "XFMR"
          case TC_XFMR3
            BranchType = "XFMR3"
          case TC_PS
            BranchType = "SHIFTER"
          case TC_SWITCH
            BranchType = "SWITCH"
          case elst
            BranchType = "UNKNOWN"
       End Select
       aLine = "Sub " + BusName + ": " + BranchType + " to " + BusName2 + "; Handle=" + Str(BranchHnd)
       Print aLine
     End If
   Wend

   ' Far bus branches
   BusHnd     = Bus2Hnd
   BusName    = FullBusName( BusHnd )
   BranchThis = BranchHnd
   List2Len   = 0
   BranchHnd  = 0
   While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd ) > 0
     If BranchHnd <> BranchThis Then
       List2Len = List2Len + 1
       OutageList2(List2Len) = BranchHnd
       ' The code below prints branch info on the screen for visual check
       If GetData( BranchHnd, BR_nBus2Hnd, RemoteBusHnd ) = 0 Then GoTo HasError
       BusName2 = FullBusName( RemoteBusHnd )
       If GetData( BranchHnd, BR_nType, TypeCode ) = 0 Then GoTo HasError
       select case TypeCode
          case TC_LINE 
            BranchType = "LINE"
          case TC_XFMR
            BranchType = "XFMR"
          case TC_XFMR3
            BranchType = "XFMR3"
          case TC_PS
            BranchType = "SHIFTER"
          case TC_SWITCH
            BranchType = "SWITCH"
          case elst
            BranchType = "UNKNOWN"
       End Select
       aLine = "Sub " + BusName + ": " + BranchType + " to " + BusName2 + "; Handle=" + Str(BranchHnd)
       Print aLine
     End If
   Wend

   Print "Found " + Str(List1Len) + " branches at near end"
   Print "Found " + Str(List2Len) + " branches at far end"
   
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub
'===================================End of Main()====================================


