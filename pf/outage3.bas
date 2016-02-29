' ASPEN PowerScript sample program
' OUTAGE.BAS
'
' Perform N-k contingency analysis
'
' PowerScript functions called:
'   FindBusByName()
'   FullBusName()
'   GetEquipment()
'   GetBusEquipment()
'   GetData()
'   GetFlow()
'   SetData()
'   PostData()
'

Type TreeCell
 nVal      As Long
 idxVHi As Long
 idxVLo As Long
 idxNHi As Long
 idxNLo As Long
 idxSibNext As Long
 idxSibPrev As Long
End Type

Sub main()
  Dim vnPFOption(10) As Long
  Dim vdPFCriteria(5) As Double
  Dim OutageList(50) As Long
  Dim OListSize As Long, OListSize2 As Long

  ' Make sure correct slack generator is selected
  If GetEquipment( TC_PICKED, SlackHnd& ) <= 0 Then
     Print "Must select Slack bus"
     Exit Sub
  End If
  nType& = EquipmentType( SlackHnd )
  If nType <> TC_BUS Then
     Print "Must select Slack bus"
     Exit Sub
  End If
  If GetBusEquipment( SlackHnd, TC_GEN, SlackGenHnd& ) <= 0 Then
     Print "Slack bus must have generator"
     Exit Sub
  End If
  If GetData( SlackGenHnd, GE_nFixedPQ, nPQFlag& ) = 0 Then GoTo HasError
  If nPQFlag = 1 Then
    Print "Bus with Fixed PQ generator cannot be slack"
  End If
  If GetData( SlackGenHnd, GE_nActive, nActiveFlag& ) = 0 Then GoTo HasError
  If nActiveFlag <> 1 Then
    Print "Slack generator not active"
    Exit Sub
  End If

Begin Dialog INPUTDLG 57,64, 183, 132, "Contigency Study"
  OptionGroup .GROUP_1
    OptionButton 108,72,20,8, "A"
    OptionButton 144,72,20,8, "B"
    OptionButton 108,80,20,8, "C"
    OptionButton 144,80,16,8, "D"
  GroupBox 8,8,88,72, "Scope"
  GroupBox 100,60,72,44, "Line current rating"
  GroupBox 100,12,72,44, "Voltage limit (PU)"
  OKButton 40,112,48,12
  CancelButton 100,112,44,12
  CheckBox 28,20,24,12, "N-1", .CheckBox_1
  CheckBox 60,20,28,12, "N-2", .CheckBox_2
  Text 8,80,60,12, "Report file name"
  TextBox 8,92,68,12, .EditBox_1
  TextBox 144,24,20,12, .EditBox_2
  TextBox 144,40,20,12, .EditBox_3
  Text 108,24,24,12, "High ="
  Text 112,44,24,8, "Low ="
  Text 104,92,48,8, "Threshold %="
  TextBox 152,88,16,12, .EditBox_4
  Text 12,36,28,12, "kV from"
  TextBox 44,36,16,12, .EditBox_5
  Text 64,36,8,12, "to"
  TextBox 76,36,16,12, .EditBox_6
  Text 12,56,80,8, "Skip names with prefix:"
  TextBox 20,64,68,12, .EditBox_7
End Dialog

  Dim dlg As InputDlg

  dlg.CheckBox_1 = 1
  dlg.Editbox_1  = "outage.rep"
  dlg.Editbox_2  = 0.95
  dlg.Editbox_3  = 1.05
  dlg.Editbox_4  = 85
  dlg.Editbox_5  = 0.0
  dlg.Editbox_6  = 999
  dlg.Editbox_7  = "ZIGZAG"
  dlg.GROUP_1    = 1
  button = Dialog( dlg )
  If button = 0 Then Exit Sub	' Canceled

  If dlg.CheckBox_1 = 0 And dlg.CheckBox_2 = 0 Then 
    Print "Must select scope"
    Exit Sub
  End If 

  OutFile$   = dlg.Editbox_1
  PUlow#     = dlg.Editbox_2
  PUhigh#    = dlg.Editbox_3
  DoN1       = dlg.CheckBox_1
  DoN2       = dlg.CheckBox_2
  LineRating = dlg.GROUP_1 + 1
  Threshold# = dlg.Editbox_4 / 100
  kVFrom     = dlg.Editbox_5
  kVTo       = dlg.Editbox_6
  SkipName$  = dlg.Editbox_7

  ' Build branch outage list
  Const LNListSIZE = 1000
  ' Make sure list size is adequate
  If GetData( HND_SYS, SY_nNOline, NOline& ) = 0 Then GoTo HasError
  If LNListSIZE - NOline < 50 Then 
    Print "Must increase LNListSIZE by ", 50 - LNListSIZE - NOline
    Exit Sub
  End If

  Dim LnList(LNListSIZE) As TreeCell

  ' Initialize tree
  For ii = 1 To LNListSIZE
    LnList(ii).idxVHi = -1
    LnList(ii).idxVLo = -1
    LnList(ii).idxNHi = -1
    LnList(ii).idxNLo = -1
    LnList(ii).idxSibNext = -1
    LnList(ii).idxSibPrev = -1
  Next

  ListSize& = 0
  BusHnd1&  = 0
  While NextBusByName( BusHnd1 ) > 0
    If GetData( BusHnd1, BUS_dKVnorminal, BuskV# ) = 0 Then GoTo HasError
    BranchHnd& = 0
    While BuskV >= kVFrom And BuskV <= kVTo And GetBusEquipment( BusHnd1, TC_BRANCH, BranchHnd ) > 0
      If GetData( BranchHnd, BR_nType, BrType& ) = 0 Then GoTo HasError
      If BrType <> TC_LINE Then GoTo Skip1 ' Look at lines only

      If GetData( BranchHnd, BR_nInService, BrFlag& ) = 0 Then GoTo HasError
      If BrFlag <> 1 Then GoTo Skip1 ' Line must be in-service
      If GetData( BranchHnd, BR_nHandle, LnHnd& ) = 0 Then GoTo HasError
      ' Check if this line has been recored before
      If  ListSize > 0 Then idxNext = 1 Else idxNext = -1
      Do While idxNext <> -1
        idxThis = idxNext
        If LnList(idxThis).nVal = LnHnd Then
          Exit Do
        ElseIf LnHnd > LnList(idxThis).nVal Then
          idxNext = LnList(idxThis).idxVHi
        Else
          idxNext = LnList(idxThis).idxVLo
        End If
      Loop 
      If idxNext <> -1 Then GoTo Skip1  ' This line has been recorded

      ' Store this line in the list
      ListSize = ListSize + 1
      LnList(ListSize).nVal = LnHnd

      If ListSize < 2 Then GoTo Skip1   ' Just store the first one

      ' Update tree pointers
      If LnHnd > LnList(idxThis).nVal Then
        LnList(idxThis).idxVHi = ListSize
      Else
        LnList(idxThis).idxVLo = ListSize
      End If

      ' Detect 3 terminal lines which have same name
      If GetData( LnHnd, LN_sName, LnName1$ ) = 0 Then GoTo HasError
      idxNext = 1
      Do While idxNext <> -1
        idxThis = idxNext
        LnHnd2& = LnList(idxThis).nVal
        If GetData( LnHnd2, LN_sName, LnName2$ ) = 0 Then GoTo HasError
        If LnName1$ > LnName2$ Then
          idxNext = LnList(idxThis).idxNHi
        ElseIf LnName1$ < LnName2$ Then
          idxNext = LnList(idxThis).idxNLo
        Else
          ' Update name links
          LnList(ListSize).idxSibNext = LnList(idxThis).idxSibNext
          LnList(idxThis).idxSibNext  = ListSize
          LnList(ListSize).idxSibPrev = idxThis
          Exit Do
        End If
      Loop
      If idxNext <> -1 Then GoTo Skip1
      ' Update tree pointers
      If LnName1$ > LnName2$ Then
        LnList(idxThis).idxNHi = ListSize
      Else
        LnList(idxThis).idxNLo = ListSize
      End If
    Skip1:
    Wend
  Wend

  Open OutFile$ For Output As 1

  ' Set PF criteria
  vdPFCriteria(1) = 20	' Max iterations
  vdPFCriteria(2) = 1	' MW tolerance
  vdPFCriteria(3) = 1	' MVAR tolerance
  vdPFCriteria(4) = 10	' MW adj. threshold
  vdPFCriteria(5) = 10	' MVAR adj. threshold
  ' Set PF options
  vnPFOption(1) = 1	' Use previous result
  vnPFOption(2) = 0	' Gen var limit
  vnPFOption(3) = 1	' Xfmr tap
  vnPFOption(4) = 0	' Area interchange
  vnPFOption(5) = 0	' Gen remove V control
  vnPFOption(6) = 0	' SVD
  vnPFOption(7) = 1	' Phase shifter
  vnPFOption(8) = 0	' Reset LTC
  vnPFOption(9) = 0	' Solution monitor
  ' PF method
  nMethod = 1		' Newton-Raphson


  Dim RatingLbl(4) As String
  RatingLbl(1) = "A"
  RatingLbl(2) = "B"
  RatingLbl(3) = "C"
  RatingLbl(4) = "D"

  ' Report Title page
  Print #1, "                            CONTINGENCY ANALYSIS REPORT"
  Print #1, "              GENERATED FROM ASPEN POWERFLOW BY A POWERSCRIPT PROGRAM"
  Print #1, ""
  Print #1, "                                 Date: ", Date()
  Print #1, ""
  Print #1, "Scope:"
  If DoN1 = 1 Then
    Print #1, "  N-1 contingency         [X]"
  Else
    Print #1, "  N-1 contingency         [ ]"
  End If
  If DoN2 = 1 Then
    Print #1, "  N-2 contingency         [X]"
  Else
    Print #1, "  N-2 contingency         [ ]"
  End If
  Print #1, "  From kV               = ", kVFrom
  Print #1, "  To kV                 = ", kVTo
  Print #1, "  Skip names            = ", SkipName
  Print #1, "Upper voltage limit(PU) = ", PUhigh
  Print #1, "Lower voltage limit(PU) = ", PUlow
  Print #1, "Use line rating         = ", RatingLbl(LineRating)
  Print #1, "Current threshold %     = ", Threshold
  Print #1, "Slack bus               = ", FullBusName( SlackHnd )
  Print #1, ""
  Print #1, ""


  CaseCount = 0
  FailCount = 0
  FailtStr  = ""
  For iiN1 = 1 To ListSize
    LineHnd1 = LnList(iiN1).nVal
    If GetData( LineHnd1, LN_sName, LnName$ ) = 0 Then GoTo HasError
    ' Skip all segments of multi terminal lines except first one
    If LnName <> "" And LnList(iiN1).idxSibPrev <> -1 Then GoTo continue1  
    ' Skip zizag xfmr
    If UCase(Left( LnName$, 6 )) = SkipName Then GoTo Continue1
    ' Take this line out of service
    If SetData( LineHnd1, LN_nInService, 0 ) = 0 Then GoTo HasError
    If PostData( LineHnd1 ) = 0 Then GoTo HasError
    OListSize = 1
    OutageList(OListSize) = LineHnd1
    ' Must outage all segment of a multiterminal line
    idxThis = iiN1
    Do While LnName$ <> "" And LnList(idxThis).idxSibNext > -1
      idxThis = LnList(idxThis).idxSibNext
      OListSize = OListSize + 1
      LnHndOut  = LnList(idxThis).nVal
      OutageList(OListSize) = LnHndOut
      ' Take this line out of service
      If SetData( LnHndOut, LN_nInService, 0 ) = 0 Then GoTo HasError
      If PostData( LnHndOut ) = 0 Then GoTo HasError
    Loop
    If DoN1 = 1 Then
      ' Print case title
      If GetData( LineHnd1, LN_sID,      LineID1$  ) = 0 Then GoTo HasError
      Print #1, ""
      CaseCount = CaseCount + 1
      Print #1, "====== Case #", CaseCount, _
                " (N-1) =============================================================================="
      Print #1, "Outages: "
      For ii = 1 To OListSize
        LineHnd& = OutageList(ii)
        If GetData( LineHnd, LN_nBus1Hnd, BusHnd1& ) = 0 Then GoTo HasError
        If GetData( LineHnd, LN_nBus2Hnd, BusHnd2& ) = 0 Then GoTo HasError
        If GetData( LineHnd, LN_sID,      LineID1$ ) = 0 Then GoTo HasError
        Print #1, FullBusName( BusHnd1 ), " - ", FullBusName( BusHnd2 ), " ", LineID1$
      Next
      Print #1, ""
      ' Do the power flow
      If DoPF( SlackHnd, vdPFCriteria, vnPFOption, nMethod ) = 0 Then 
        Print #1, "PowerFlow failed"
        FailCount = FailCount + 1
        FailStr   = FailtStr & ", #" & CaseCount
      Else
        Call PFReport( PUlow, PUhigh, LineRating, Threshold )
      End If
    End If
    If DoN2 = 1 Then
      For iiN2 = iiN1+1 To ListSize
        LineHnd2& = LNList(iiN2).nVal
        If GetData( LineHnd2, LN_sName, LnName$ ) = 0 Then GoTo HasError
        ' Skip all segments of multi terminal lines except first one
        If LnName <> "" And LnList(iiN2).idxSibPrev <> -1 Then GoTo Continue2  
        ' Skip zizag xfmr
        If UCase(Left( LnName$, 6 )) = SkipName Then GoTo Continue2
        If GetData( LineHnd2, LN_nInService, nActiveFlag& ) = 0 Then GoTo HasError
        If nActiveFlag <> 1 Then GoTo Continue2
        ' Take this line out of service
        If SetData( LineHnd2, LN_nInService, 0 ) = 0 Then GoTo HasError
        If PostData( LineHnd2 ) = 0 Then GoTo HasError
        OListSize2    = OListSize + 1
        OutageList(OListSize2) = LineHnd2
        ' Must outage all segment of a multiterminal line
        idxThis = iiN2
        Do While LnName$ <> "" And LnList(idxThis).idxSibNext > -1
          idxThis = LnList(idxThis).idxSibNext
          OListSize2 = OListSize2 + 1
          LnHndOut   = LnList(idxThis).nVal
          OutageList(OListSize2) = LnHndOut
          ' Take this line out of service
          If SetData( LnHndOut, LN_nInService, 0 ) = 0 Then GoTo HasError
          If PostData( LnHndOut ) = 0 Then GoTo HasError
        Loop
        ' Print case title
        Print #1, ""
        CaseCount = CaseCount + 1
        Print #1, "====== Case #", CaseCount, _
                  " (N-2) =============================================================================="
        Print #1, "Outages: "
        For ii = 1 To OListSize2
          LineHnd = OutageList(ii)
          If GetData( LineHnd, LN_nBus1Hnd, BusHnd1& ) = 0 Then GoTo HasError
          If GetData( LineHnd, LN_nBus2Hnd, BusHnd2& ) = 0 Then GoTo HasError
          If GetData( LineHnd, LN_sID,      LineID1$ ) = 0 Then GoTo HasError
        Print #1, FullBusName( BusHnd1 ), " - ", FullBusName( BusHnd2 ), " ", LineID1$
        Next
        Print #1, ""
        ' Do the power flow
        If DoPF( SlackHnd, vdPFCriteria, vnPFOption, nMethod ) = 0 Then 
          Print #1, "PowerFlow failed"
          FailCount = FailCount + 1
          FailStr   = FailtStr & ", #" & CaseCount
        Else
          Call PFReport( PUlow, PUhigh, LineRating, Threshold )
        End If
        ' Put lines back in service
        For ii=OListSize+1 To OListSize2
          LineHnd = OutageList(ii)
          If SetData( LineHnd, LN_nInService, 1 ) = 0 Then GoTo HasError
          If PostData( LineHnd ) = 0 Then GoTo HasError
        Next
      Continue2:
      Next
    End If
    ' Put this line back in service
    For ii=1 To OListSize
      LineHnd = OutageList(ii)
      If SetData( LineHnd, LN_nInService, 1 ) = 0 Then GoTo HasError
      If PostData( LineHnd ) = 0 Then GoTo HasError
    Next
  Continue1:
  Next

  ' Print summary
  Print #1, ""
  Print #1, "Summary:"
  Print #1, "  Analyzed ", CaseCount, " contigencies"
  If FailtCount > 0 Then
    Print #1, "  PowerFlow failed in following ", FailCount, " cases: ", FailStr
  End If

  Close
  Print "Analyzed ", CaseCount, " contigencies" & _
        Chr(10) & "Report is in ", OutFile$

  
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Sub PFReport( ByVal PUlow#, ByVal PUhigh#, ByVal nLineRating&, ByVal Threshold# )
  Dim Mag(16) As Double, Ang(16) As Double
  Dim Rating(4) As Double

Exit Sub


  ' Voltage report
  Print #1, "__Bus____________________Voltage(PU)_________Flag______"
  BusHnd& = 0
  While NextBusByName( BusHnd ) > 0
    If GetPFVoltage( BusHnd, Mag, Ang, ST_PU ) = 0 Then GoTo HasError
    FFlag$ = "  "
    If Mag(1) > PUhigh Then FFlag$ = "Over Voltage"
    If Mag(1) < PUlow Then  FFlag$ = "Under Voltage"
    Aline$ = FullBusName( BusHnd )
    Aline$ = Aline$ & Space( 30 - Len(Aline$) )
    Aline$ = Aline$ & Format( Mag(1), "#0.0##" )
    Aline$ = Aline$ & Space( 45 - Len(Aline$) )
    Aline$ = Aline$ & FFlag$
    Print #1, Aline$
  Wend

  ' Current report
  Print #1, ""
  Print #1, "__Line____________________________________________Current(A)_____Rating(A)______Flag______"
  LineHnd& = 0
  While GetEquipment( TC_LINE, LineHnd ) > 0
    If GetData( LineHnd, LN_sName,     LnName$ ) = 0 Then GoTo HasError
    If UCase(Left(LnName,6)) = "ZIGZAG" Then GoTo Continue3
    If GetData( LineHnd, LN_nInService, nFlag& ) = 0 Then GoTo HasError
    If nFlag <> 1 Then GoTo Continue3
    If GetPFCurrent( LineHnd, Mag, Ang, 0 ) = 0 Then GoTo HasError
    If GetData( LineHnd, LN_nBus1Hnd, BusHnd1& ) = 0 Then GoTo HasError
    If GetData( LineHnd, LN_nBus2Hnd, BusHnd2& ) = 0 Then GoTo HasError
    If GetData( LineHnd, LN_sID,      LineID$  ) = 0 Then GoTo HasError
    If GetData( LineHnd, LN_vdRating, Rating   ) = 0 Then GoTo HasError
    d1# = Mag(1)
    d2# = Rating(nLineRating)
    If Mag(1) < Rating(nLineRating)*Threshold Then FFlag$ = "" Else FFlag$ = "Overloaded"
    Aline$ = FullBusName( BusHnd1 ) + " - " + FullBusName( BusHnd2 ) + " " + LineID$
    Aline$ = Aline$ & Space( 50 - Len(Aline$) )
    Aline$ = Aline$ & Format( Mag(1), "#####0.0#" )
    Aline$ = Aline$ & Space( 65 - Len(Aline$) )
    Aline$ = Aline$ & Format( Rating(nLineRating), "#####0.0#" )
    Aline$ = Aline$ & Space( 80 - Len(Aline$) )
    Aline$ = Aline$ & FFlag$
    Print #1, Aline$
  Continue3:
  Wend
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Sub OutageTapBus( ByVal BusHnd1&, ByRef OutageList() As Long, ByRef OListSize& )
  BranchHnd& = 0
  While GetBusEquipment( BusHnd1, TC_BRANCH, BranchHnd ) > 0
    If GetData( BranchHnd, BR_nType, BrType& ) = 0 Then GoTo HasError
    If BrType <> TC_LINE Then GoTo Skip1 ' Look at lines only
    If GetData( BranchHnd, BR_nInService, BrFlag& ) = 0 Then GoTo HasError
    If BrFlag <> 1 Then GoTo Skip1 ' Must be in-service
    ' Outage this line
    If GetData( BranchHnd, BR_nHandle, LineHnd& ) = 0 Then GoTo HasError
    If SetData( LineHnd, LN_nInService, 0 ) = 0 Then GoTo HasError
    If PostData( LineHnd ) = 0 Then GoTo HasError
    OListSize = OListSize + 1
    OutageList( OListSize ) = LineHnd
  Skip1:
  Wend
  Exit Sub
HasError:
Print "Error: ", ErrorString()
Stop
End Sub