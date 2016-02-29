' ASPEN PowerScript sample program
' OUTAGE2.BAS
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
Sub main()
  Dim vnPFOption(10) As Long
  Dim vdPFCriteria(5) As Double
  Dim vBusFlag1(10000) As Integer, vBusFlag2(10000) As Integer

  ' Make sure correct slack generator is selected
  If GetEquipment( TC_PICKED, SlackHnd& ) = 0 Then
     Print "Must select Slack bus"
     Exit Sub
  End If
  nType& = EquipmentType( SlackHnd )
  If nType <> TC_BUS Then
     Print "Must select Slack bus"
     Exit Sub
  End If
  If GetBusEquipment( SlackHnd, TC_GEN, SlackGenHnd& ) = 0 Then GoTo HasError
  If GetData( SlackGenHnd, GE_nFixedPQ, nPQFlag& ) = 0 Then GoTo HasError
  If nPQFlag = 1 Then
    Print "Slack generator cannot be Fixed PQ"
  End If
  If GetData( SlackGenHnd, GE_nActive, nActiveFlag& ) = 0 Then GoTo HasError
  If nActiveFlag <> 1 Then
    Print "Slack generator not active"
    Exit Sub
  End If

Begin Dialog INPUTDLG 57,64, 185, 122, "Contigency Study"
  OptionGroup .GROUP_1
    OptionButton 16,52,20,8, "A"
    OptionButton 52,52,20,8, "B"
    OptionButton 16,60,20,8, "C"
    OptionButton 52,60,16,8, "D"
  GroupBox 8,8,72,28, "Scope"
  GroupBox 8,40,76,44, "Line current rating"
  GroupBox 100,40,68,44, "Voltage limit (PU)"
  OKButton 32,96,48,12
  CancelButton 100,96,52,12
  CheckBox 16,20,24,12, "N-1", .CheckBox_1
  CheckBox 48,20,28,12, "N-2", .CheckBox_2
  Text 84,8,60,16, "Report file name"
  TextBox 84,20,84,16, .EditBox_1
  TextBox 144,52,20,12, .EditBox_2
  TextBox 144,68,20,12, .EditBox_3
  Text 108,52,24,12, "High ="
  Text 112,72,24,8, "Low ="
  Text 12,72,48,8, "Threshold %="
  TextBox 60,68,16,12, .EditBox_4
End Dialog

  Dim dlg As InputDlg

  dlg.CheckBox_1 = 1
  dlg.Editbox_1  = "outage.rep"
  dlg.Editbox_2  = 0.95
  dlg.Editbox_3  = 1.05
  dlg.Editbox_4  = 85
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


  ' Report Title page
  Print #1, "                            CONTINGENCY ANALYSIS REPORT"
  Print #1, "              GENERATED FROM ASPEN POWERFLOW BY A POWERSCRIPT PROGRAM"
  Print #1, ""
  Print #1, "                                 Date: ", Date()
  Print #1, ""
  If DoN1 = 1 Then
    Print #1, "N-1 contingency           [X]"
  Else
    Print #1, "N-1 contingency           [ ]"
  End If
  If DoN2 = 1 Then
    Print #1, "N-2 contingency           [X]"
  Else
    Print #1, "N-2 contingency           [ ]"
  End If
  Print #1, "Upper voltage limit(PU) = ", PUhigh
  Print #1, "Lower voltage limit(PU) = ", PUlow
  Print #1, "Use line rating         = ", nLineRating
  Print #1, "Current threshold %     = ", Threshold
  Print #1, "Slack bus               = ", FullBusName( SlackHnd )
  Print #1, ""
  Print #1, ""


  CaseCount = 0

  BusHnd1&  = 0
  While NextBusByName( BusHnd1 ) > 0
    If GetData( BusHnd1, BUS_sName, BusName1$ ) = 0 Then GoTo HasError
    BranchHnd& = 0
    While GetBusEquipment( BusHnd1, TC_BRANCH, BranchHnd ) > 0
      If GetData( BranchHnd, BR_nType, BrType& ) = 0 Then GoTo HasError
      If BrType <> TC_LINE Then GoTo Skip1 ' Look at lines only
      If GetData( BranchHnd, BR_nInService, BrFlag& ) = 0 Then GoTo HasError
      If BrFlag <> 1 Then GoTo Skip1 ' Must be in-service
      ' Check opposite bus name to see if it has been covered
      If GetData( BranchHnd, BR_nBus1Hnd, BrBus2& ) = 0 Then GoTo HasError
      If BrBus2 = BusHnd1 Then ' Must get the other bus
        If GetData( BranchHnd, BR_nBus2Hnd, BrBus2& ) = 0 Then GoTo HasError
      End If
      If GetData( BrBus2, BUS_sName, BrBus2Name$ ) = 0 Then GoTo HasError
      If BrBus2Name < BusName1 Then GoTo Skip1
      ' Outage this line
      If GetData( BranchHnd, BR_nHandle, LineHnd1& ) = 0 Then GoTo HasError
      If SetData( LineHnd1, LN_nInService, 0 ) = 0 Then GoTo HasError
      If PostData( LineHnd1 ) = 0 Then GoTo HasError
      If DoN1 = 1 Then
        Print #1, ""
        CaseCount = CaseCount + 1
        Print #1, "======Case #", CaseCount, _
                  "=============================================================================="
        Print #1, "Outage: ", FullBusName( BusHnd1 ), " - ", FullBusName( BusHnd2 ), " ", LineID1$
        Print #1, ""
        ' Do the power flow
        If DoPF( SlackHnd, vdPFCriteria, vnPFOption, nMethod ) = 0 Then 
          Print #1, "PowerFlow failed"
        Else
          Call PFReport( PUlow, PUhigh, LineRating, Threshold )
        End If

      Skip1:
    Wend
  Wend

  Close
  Exit Sub

  LineHnd1& = 0
  While GetEquipment( TC_LINE, LineHnd1 ) > 0
    If GetData( LineHnd1, LN_nInService, nActiveFlag& ) = 0 Then GoTo HasError
    If nActiveFlag = 1 Then
      ' Take this line out of service
      If SetData( LineHnd1, LN_nInService, 0 ) = 0 Then GoTo HasError
      If PostData( LineHnd1 ) = 0 Then GoTo HasError
      If DoN1 = 1 Then
        ' Print case title
        If GetData( LineHnd1, LN_nBus1Hnd, BusHnd1& ) = 0 Then GoTo HasError
        If GetData( LineHnd1, LN_nBus2Hnd, BusHnd2& ) = 0 Then GoTo HasError
        If GetData( LineHnd1, LN_sID,      LineID1$  ) = 0 Then GoTo HasError
        Print #1, ""
        CaseCount = CaseCount + 1
        Print #1, "======Case #", CaseCount, _
                  "=============================================================================="
        Print #1, "Outage: ", FullBusName( BusHnd1 ), " - ", FullBusName( BusHnd2 ), " ", LineID1$
        Print #1, ""
        ' Do the power flow
        If DoPF( SlackHnd, vdPFCriteria, vnPFOption, nMethod ) = 0 Then 
          Print #1, "PowerFlow failed"
        Else
          Call PFReport( PUlow, PUhigh, LineRating, Threshold )
        End If
      End If
      If DoN2 = 1 Then
        LineHnd2& = LineHnd1
        While GetEquipment( TC_LINE, LineHnd2 ) > 0
          If GetData( LineHnd2, LN_nInService, nActiveFlag& ) = 0 Then GoTo HasError
          If nActiveFlag = 1 Then
            ' Take this line out of service
            If SetData( LineHnd2, LN_nInService, 0 ) = 0 Then GoTo HasError
            If PostData( LineHnd2 ) = 0 Then GoTo HasError
            ' Print case title
            If GetData( LineHnd2, LN_nBus1Hnd, BusHnd3& ) = 0 Then GoTo HasError
            If GetData( LineHnd2, LN_nBus2Hnd, BusHnd4& ) = 0 Then GoTo HasError
            If GetData( LineHnd2, LN_sID,      LineID2$  ) = 0 Then GoTo HasError
            ' Print case title
            Print #1, ""
            CaseCount = CaseCount + 1
            Print #1, "======Case #", CaseCount, _
                      "=============================================================================="
            Print #1, "Outage: ", FullBusName( BusHnd1 ), " - ", FullBusName( BusHnd2 ), " ", LineID1$
            Print #1, "        ", FullBusName( BusHnd3 ), " - ", FullBusName( BusHnd4 ), " ", LineID2$
            Print #1, ""
            ' Do the power flow
            If DoPF( SlackHnd, vdPFCriteria, vnPFOption, nMethod ) = 0 Then 
              Print #1, "PowerFlow failed"
            Else
              Call PFReport( PUlow, PUhigh, LineRating, Threshold )
            End If
             ' Put this line back in service
            If SetData( LineHnd2, LN_nInService, 1 ) = 0 Then GoTo HasError
            If PostData( LineHnd2 ) = 0 Then GoTo HasError
          End If
        Wend
      End If
      ' Put this line back in service
      If SetData( LineHnd1, LN_nInService, 1 ) = 0 Then GoTo HasError
      If PostData( LineHnd1 ) = 0 Then GoTo HasError
    End If
  Wend

  Close
  Print "Analyzed ", CaseCount, " contigencies. Report is in ", OutFile$
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub

Sub PFReport( ByVal PUlow#, ByVal PUhigh#, ByVal nLineRating&, ByVal Threshold# )
  Dim Mag(16) As Double, Ang(16) As Double
  Dim Rating(4) As Double

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
    Aline$ = Aline$ & Format( Mag(1), "##.###" )
    Aline$ = Aline$ & Space( 45 - Len(Aline$) )
    Aline$ = Aline$ & FFlag$
    Print #1, Aline$
  Wend

  ' Current report
  Print #1, ""
  Print #1, "__Line____________________________________________Current(A)_____Rating(A)______Flag______"
  LineHnd& = 0
  While GetEquipment( TC_LINE, LineHnd ) > 0
    If GetData( LineHnd, LN_nInService, nFlag& ) = 0 Then GoTo HasError
    If nFlag = 1 Then
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
    End If
  Wend
  Exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub