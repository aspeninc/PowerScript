' ASPEN PowerScript Sample Program
'
' REMOTELINEFLT.BAS
'
' Simulate fault on remote line(s).
' Lines with tap buses are handled correctly
'
'
' Global vars
dim FarBrHnd(50) As long
dim FarLnZ(50) As double
dim FarFarBsName(50) As String
dim RemoteBs(1) As String
dim CountFarBr As long
dim ThisLnZ As double
Begin Dialog REMOTELINEFLT 30,60,217,105, "Specify Remote Line Fault"
  Text 5,4,43,8,"Remote bus:"
  Text 163,56,10,8,"+j"
  Text 5,16,43,8,"Remote line:"
  Text 15,33,34,8,"Percent ="
  Text 66,32,146,8,"(Enter >100% to use Z base of the relay line)"
  GroupBox 121,45,87,28,"Fault Z (ohm)"
  GroupBox 6,45,110,28,"Phase connections"
  CheckBox 11,56,24,11,"3LG", .CheckBox1
  CheckBox 36,56,24,11,"2LG", .CheckBox2
  CheckBox 62,56,25,11,"1LG", .CheckBox3
  CheckBox 88,56,23,11,"L-L", .CheckBox4
  TextBox 128,55,27,11,.Edit1
  TextBox 172,55,27,11,.Edit2
  DropListBox 47,15,164,120,FarFarBsName(), .ComboBox1
  TextBox 47,31,17,11,.Edit3
  PushButton 49,84,71,13,"Simulate", .Button1
  CancelButton 127,84,38,13
  ListBox 47,3,164,11,RemoteBs(), .ListBox1
End Dialog

Sub main()
   If GetEquipment( TC_PICKED, PickedHnd& ) = 0 Or _
      EquipmentType( PickedHnd ) <> TC_RLYGROUP Then 
      Print "Please select a relay group on a line"
      Exit Sub
   End If

   ' Get the branch handle
   Call GetData( PickedHnd, RG_nBranchHnd, RlyBrHnd& )
   Call GetData( RlyBrHnd, BR_nType, BranchType& )
   If BranchType <> TC_LINE Then
      Print "Please select a relay group on a line"
      Exit Sub
   End If
   
   ThisLnZ = compuLineBrZ(RlyBrHnd, LastBrHnd&, FarBsHnd& )
   RemoteBs(0) = FullBusName(FarBsHnd)
   
   ' Find all lines at the far end
   BranchHnd& = 0
   CountFarBr = 0
   While GetBusEquipment( FarBsHnd, TC_BRANCH, BranchHnd& ) > 0
     If BranchHnd <> LastBrHnd Then
       Call GetData( BranchHnd, BR_nInservice, nFlag& )
       If nFlag = 1 Then
         Call GetData( BranchHnd, BR_nType, BrType& )
         If BrType = TC_LINE Then
           Call GetData( BranchHnd, BR_nHandle, LnHnd& )
           Call GetData( LnHnd, LN_sID, cktID$ )
           FarBrHnd(CountFarBr) = BranchHnd
           FarLnZ(CountFarBr)   = compuLineBrZ(BranchHnd, nDummy&, Bus2Hnd& )
           FarFarBsName(CountFarBr) = "To " + FullBusName(Bus2Hnd) + " L" + cktID$
           CountFarBr = CountFarBr + 1
         End If
       End If
     End If
   Wend
   
   If CountFarBr = 0 Then
     Print "Found no line at remote bus"
     Stop
   End If
   
   If CountFarBr > 1 Then
     FarFarBsName(CountFarBr+2) = "Shortest line at the remote bus"
     FarFarBsName(CountFarBr+3) = "Longest line at the remote bus"
   End If
   
   RunFault

End Sub


Sub runFault()
  dim dlg As REMOTELINEFLT
  Dim FltConn(5) As Long
  Dim FltOpt(15) As double
  Dim OutageOpt(5) As Long
  Dim OutageLst(30) As Long
  Dim vnShowRelay(4)
  
  dlg.Edit1 = 0
  dlg.Edit2 = 0
  dlg.Edit3 = 130
  dlg.CheckBox1 = true
  dlg.CheckBox3 = true
  ClearPrev     = 0 ' keep previous result
  For ii=0 to 14
    FltOpt(ii) = 0
  Next
  While 0 <> Dialog(dlg)
    RemoteIdx& = dlg.ComboBox1
    If RemoteIdx >= CountFarBr Then
      If RemoteIdx = CountFarBr Then   ' Shortes line
        zZ = 9999
        For ii=0 to CountFarBr-1
          If FarLnZ(ii) < zZ Then
            zZ = FarLnZ(ii) < zZ 
            nIdx = ii
          End If 
        Next
      Else                              ' Longest line
        zZ = 0
        For ii=0 to CountFarBr-1
          If FarLnZ(ii) > zZ Then
            zZ = FarLnZ(ii) < zZ 
            nIdx = ii
          End If 
        Next
      End If
      RemoteIdx = nIdx
    End If
    FltOpt(9)  = dlg.Edit3   ' Intermediate %
    If FltOpt(9) > 100 Then
      zZ# = (FltOpt(9)-100)*ThisLnZ/FarLnZ(RemoteIdx)
      FltOpt(9) = zZ#
      If FltOpt(9) > 99.5 Then FltOpt(9) = 99.5
    End If
  
    FltConn(1) = dlg.CheckBox1   ' Do 3PH
    FltConn(2) = dlg.CheckBox2
    FltConn(3) = dlg.CheckBox3   ' Do 1LG
    FltConn(4) = dlg.CheckBox4
    OutageOpt(1) = 0 ' With one outage at a time
    OutageOpt(2) = 0 ' With two outage at a time
    OutageOpt(3) = 0 ' With two outage at a time
    Rflt#        = dlg.Edit1 ' Fault R
    Xflt#        = dlg.Edit2 ' Fault X
  
    ' Simulate the fault
    BrHnd& = FarBrHnd(RemoteIdx)
    If DoFault( BrHnd, FltConn, FltOpt, OutageOpt, OutageLst, _
           Rflt, Xflt, ClearPrev ) = 0 Then
      Print "Error: ", ErrorString( )
    Else
      Call ShowFault( SF_LAST, 0, 4, 0, vnShowRelay )
    End If
  Wend
End Sub

Function compuLineBrZ( ByVal LineBrHnd&, ByRef RemoteBrHnd&, ByRef RemoteBsHnd& ) As double
   dim ProcessedHnd(100) As integer
  
   compuLineBrZ = 0
   If EquipmentType( LineBrHnd ) <> TC_BRANCH Then exit Sub
   Call GetData( LineBrHnd, BR_nHandle,  LineHnd& )
   If EquipmentType( LineHnd ) <> TC_LINE Then exit Sub
   
   Call GetData( LineBrHnd, BR_nBus1Hnd, Bus1Hnd& )
   Call GetData( LineBrHnd, BR_nBus2Hnd, Bus2Hnd& )
   
   RemoteBsHnd = Bus2Hnd
   BranchHnd& = 0
   While GetBusEquipment( RemoteBsHnd, TC_BRANCH, BranchHnd ) > 0
     Call GetData( BranchHnd, BR_nHandle,  TempHnd& )
     If TempHnd = LineHnd Then 
       RemoteBrHnd = BranchHnd
     End If
   Wend
   
   ' Get the branch bus handle
   Call GetData( LineHnd, LN_dR, dR# )
   Call GetData( LineHnd, LN_dX, dX# )
   Call GetData( LineHnd, LN_dR0, dR0# )
   Call GetData( LineHnd, LN_dX0, dX0# )
   Call GetData( LineHnd, LN_dLength, dLength# )
'   aLine$ = FullBusName(Bus1Hnd) + " - " + FullBusName(Bus2Hnd) + ": " + _
'                      "Z=" + Format(dR#,"0.00000") + "+j" + Format(dX#,"0.00000") + " " + _
'                      "Zo=" + Format(dR0#,"0.00000") + "+j" + Format(dX0#,"0.00000") + " " + _
'                      "L=" + Format(dLength#,"0.00000")
'   PrintTTY(" ")
'   PrintTTY(aLine$)
   
   ' Skip all taps on Bus2 side
   BusHnd&  = Bus2Hnd
   BusFHnd& = Bus1Hnd
   Do 
     Call GetData( BusHnd, BUS_nTapBus, TapCode& )
     If TapCode = 0 Then Exit Do ' Stop searching at the first Real bus
     BranchHnd& = 0
     While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd ) > 0
       Call GetData( BranchHnd, BR_nBus2Hnd, BusFarHnd )  ' Get the far end bus
       If BusFarHnd <> BusFHnd Then	' Not the same line
         Call GetData( BranchHnd, BR_nType, TypeCode )
         If TypeCode = TC_LINE Then 
           ' Found a continuation of the line. Calulate total impedance
           Call GetData( BranchHnd, BR_nHandle, LineHnd )
           Call GetData( LineHnd, LN_nInservice, nFlag& )
           nFound = 0
           For ii = 0 to nProcessed -1
             If LineHnd = ProcessedHnd(ii) Then nFound = 1
           Next
           If nFound = 0 Then
             If nProcessed >= 100 Then
               Print "Max number of segments reached. Abort"
               Stop
             End If
             ProcessedHnd(nProcessed) = LineHnd
             nProcessed = nProcessed + 1
           End If
           If nFound = 0 And nFlag = 1 Then
             Call GetData( LineHnd, LN_dR, dRn# )
             Call GetData( LineHnd, LN_dX, dXn# )
             Call GetData( LineHnd, LN_dR0, dR0n# )
             Call GetData( LineHnd, LN_dX0, dX0n# )
             Call GetData( LineHnd, LN_dLength, dL# )
             dLength = dLength + dL
             dR  = dR + dRn
             dX  = dX + dXn
             dR0 = dR0 + dR0n
             dX0 = dX0 + dX0n
             aLine$ = FullBusName(BusHnd) + " - " + FullBusName(BusFarHnd) + ": " + _
                      "Z=" + Format(dRn#,"0.00000") + "+j" + Format(dXn#,"0.00000") + " " + _
                      "Zo=" + Format(dR0n#,"0.00000") + "+j" + Format(dX0n#,"0.00000") + " " + _
                      "L=" + Format(dL#,"0.00000")
'             PrintTTY(aLine$)
             BusFHnd = BusHnd
             BusHnd  = BusFarHnd
             RemoteBsHnd = BusFarHnd
             RemoteBrHnd = BranchHnd
             GoTo ContinueDo1
           End If
         End If
       End If  
     Wend
     Exit Do    ' Stop searching when no more line is found
     ContinueDo1:
   Loop
'   aLine$ = "Z = " + Str(dR) + " + j" + Str(dX) + Chr(13) + " " + _
'            "Zo = " + Str(dR0) + " + j" + Str(dX0) + Chr(13) + " " + _
'            "Length = " + Str(dLength)
'   Print aLine$ + Chr(13) + "Result printed in TTY windows"
'   PrintTTY( aLine$ )
   compuLineBrZ = Sqr(dR*dR + dX*dX)
End Function

