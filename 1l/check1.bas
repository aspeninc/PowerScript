' ASPEN PowerScript sample program
'
' CHECK1.BAS
'
' Plot graph of relay time vs. distance
' 
' Version 1.0
' Category: OneLiner
'
' PowerScript functions called:
'   DoFault()
'   PickFault()
'   EquipmentType()
'   GetSCCurrent()
'   FaultDescription()
'
' Main program code
Sub main()
   ' Variable declaration
   Dim MagArray(16) As Double
   Dim AngArray(16) As Double
   Dim FltConnection(4) As Long
   Dim FltOption(14) As Double
   Dim OutageType(3) As Long
   Dim OutageList(15) As Long
   Dim Br2List( 50 ) As Long
   Dim Bus3List( 50 ) As Long
   Dim BusNameList(50) As String
   Dim RlyGrp2List( 50 ) As Long
   Dim StepSize As Double
   Dim PickedHnd As Long, BranchHnd As Long, DoneFlag As Long, ShowFaultFlag As Long
   Dim FltConnStr(4) As String
Begin Dialog FAULTDLG 48,46, 125, 191, "Relay time report"
  OptionGroup .GROUP_1
    OptionButton 8,16,24,8, "3PH"
    OptionButton 36,16,24,8, "2LG"
    OptionButton 64,16,24,8, "1LG"
    OptionButton 92,16,20,8, "LL"
  OptionGroup .GROUP_2
    OptionButton 90,132,28,12, ".TXT"
    OptionButton 90,139,28,12, ".CSV"
  GroupBox 4,4,115,24, "Select fault connection"
  Text 4,32,84,12, "Sliding fault step size ="
  TextBox 92,32,12,12, .EditBox_2
  Text 108,32,8,12, "%"
  OKButton 13,168,48,12
  CancelButton 64,168,48,12
  ListBox 4,56,115,68, BusNameList$(), .ListBox_1
  Text 4,48,68,8, "Select remote bus"
  Text 4,124,80,12, "Enter output file name"
  TextBox 4,135,86,12, .EditBox_1
  CheckBox 5,154,88,8, "Output relay currents", .CheckBox_1
End Dialog
   Dim dlg As FAULTDLG

   FltConnStr(1) = "3PH"
   FltConnStr(2) = "2LG"
   FltConnStr(3) = "1LG"
   FltConnStr(4) = "LL"

   For ii = 1 To 4 
     FltConnection(ii) = 0
   Next 
   For ii = 1 To 14
     FltOption(ii) = 0.0
   Next
   For ii = 1 To 3
     OutageType(ii) = 0
   Next
   dFltR     = 0.0   '
   dFltX     = 0.0
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 Then GoTo hasError
   ' Must be a relay group
   If EquipmentType( PickedHnd ) <> TC_RLYGROUP Then
     Print "Must select a relay group"
     Exit Sub
   End If

'   If InputDialog( FltConnection, BusNameList ) = 0 Then Stop

   ' Determine starting branch BrHnd1
   If GetData( PickedHnd, RG_nBranchHnd, BrHnd1& ) = 0 Then GoTo HasError

   ' Find branch buses
   If GetData( BrHnd1, BR_nBus1Hnd, BusHnd1& ) = 0 Then GoTo HasError
   If GetData( BrHnd1, BR_nBus2Hnd, BusHnd2& ) = 0 Then GoTo HasError

   ' Find remote branches
   BrHnd2&   = 0
   CountBr2& = 0
   While GetBusEquipment( BusHnd2, TC_BRANCH, BrHnd2 ) > 0
     ' Must skip the picked branch
     If GetData( BrHnd2, BR_nBus2Hnd, BusHnd3& ) = 0 Then GoTo HasError
     If BusHnd3 = BusHnd1 Then GoTo EndLoop
     ' Must consider line only
     If GetData( BrHnd2, BR_nType, BrType& ) = 0 Then GoTo HasError
     If BrType <> TC_LINE Then GoTo EndLoop
     ' Line must have a relay group
     If GetData( BrHnd2, BR_nRlyGrp1Hnd, RlyGrp2& ) <= 0 Then GoTo EndLoop
     ' Record this branch for processing
     CountBr2 = CountBr2 + 1
     Br2List( CountBr2 )     = Br2Hnd
     RlyGrp2List( CountBr2 ) = RlyGrp2
     Bus3List( CountBr2 )    = BusHnd3
     BusNameList( CountBr2 ) = FullBusName( BusHnd3 )
   EndLoop:
   Wend

   ' Check all fault connection
   dlg.Group_1 = 2
   dlg.GROUP_2 = 1
   dlg.EditBox_1 = "c:\temp\timedist"
   dlg.EditBox_2 = 5
   If Dialog( dlg ) = 0 Then ' Canceled
     Exit Sub
   End If

   CsvOut = dlg.GROUP_2
   If CsvOut Then
     OFname = dlg.EditBox_1 + ".csv"
   Else
     OFname = dlg.EditBox_1 + ".txt"
   End If
   FltConnection(dlg.Group_1+1) = 1
   RemoteIdx& = 1 + dlg.ListBox_1
   OutRlyCur  = dlg.CheckBox_1
   StepSize#  = dlg.EditBox_2

   ClearPrev = 1	' Clear previous fault result
   ' Close in fault
   FltOption(1)  = 1
   If 0 = DoFault( PickedHnd, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev ) Then GoTo HasError

   ClearPrev = 0	' Start to keep all result

   ' Simulate intermediate faults on branch 1
   FltOption(1)  = 0
   FltOption(9)  = StepSize  ' Every 5%
   FltOption(13) = 0  ' Start from 0% 
   FltOption(14) = 100  ' To 100% 
   If 0 = DoFault( PickedHnd, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev ) Then GoTo HasError

   ' Simulate remote bus fault on branch 1
   FltOption(5)  = 1
   FltOption(9)  = 0
   If 0 = DoFault( PickedHnd, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev ) Then GoTo HasError

   ' Simulate intermediate faults on branch 2
   RlyGrpHnd2    = RlyGrp2List(RemoteIdx)
   StepSize      = 5
   FltOption(1)  = 0
   FltOption(5)  = 0
   FltOption(9)  = StepSize  ' Every 5%
   FltOption(13) = 0  ' Start from 0% 
   FltOption(14) = 100  ' To 100% 
   If 0 = DoFault( RlyGrpHnd2, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev ) Then GoTo HasError

   ' Simulate remote bus fault on branch 2
   FltOption(5)  = 1
   FltOption(9)  = 0
   If 0 = DoFault( RlyGrpHnd2, FltConnection, FltOption, OutageType, OutageList, _
                   dFltR, dFltX, ClearPrev ) Then GoTo HasError

   ' Write to output file
   Open OFname For Output As 1

   Print #1, "L1 = " & FullBusName(BusHnd1) & "-" & FullBusName(BusHnd2)
   Print #1, "L2 = " & FullBusName(BusHnd2) & "-" & FullBusName(Bus3List(RemoteIdx))
   Print #1, "Fault connection = " & FltConnStr(dlg.Group_1+1)
   Print #1, ""

   Call RelayOut( PickedHnd, StepSize, OutRlyCur, CsvOut )
   Call RelayOut( RlyGrpHnd2, StepSize, OutRlyCur, CsvOut )

   Close 1

   Message$ = "Output has been written successfully to " & FileName
   If CsvOut = 1 Then
     Message = Message & Chr(10) & "Do you want to open it in Excel"
     If 6 <> MsgBox( Message, 4, "DistTime" ) Then Exit Sub
     Set xlApp = CreateObject("excel.application")
     xlApp.Workbooks.Open Filename:=OFname
     xlApp.Visible = True
   Else
     MsgBox( Message$ )
   End If
Exit Sub
HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()
' ===================== End of Main() =========================================

' ===================== RelayOut() ============================================
' Purpose:
'   Print table of relay current and time for a range of fault
'
Sub RelayOut( ByVal RlyGrpHnd As Long, ByVal StepSize As Double, _
              ByVal printCurr As Long, ByVal printCSV As Long)
   Dim MagArray(16) As Double
   Dim AngArray(16) As Double
   Dim ShowFlagRly(4) As Long

   ' Initialize 
   For ii = 1 To 4 
     ShowFlagRly(ii) = 1
   Next 

   If GetData( RlyGrpHnd, RG_nBranchHnd, BrHnd1& ) = 0 Then GoTo HasError

   Delim$ = Chr(34) & "," & Chr(34)

   If printCurr = 1 Then 
     If printCSV = 1 Then
       HeaderLine$ = Chr(34) & "Distance" & Delim & "Ia Mag" & Delim & "Ia Ang" & Delim & _
                 "Ib Mag" & Delim & "Ib Ang" & Delim & "Ic Mag" & Delim & "Ic Ang" & Chr(34)
     Else
       HeaderLine$ = "Distance         Ia         Ib         Ic"
     End If
   Else
     If printCSV = 1 Then
       HeaderLine$ = Chr(34) & "Distance" & Chr(34)
     Else
       HeaderLine$ = "Distance"
     End If
   End If

   LineNO$  = "L1"
   Distance = 0
   ShowFaultFlag = 1 ' Starting from the first one
   ' Pick fault does not update single line diagram screen like ShowFault
'   While PickFault( ShowFaultFlag ) > 0
   While ShowFault( ShowFaultFlag, 1, 1, 0, ShowFlagRly ) > 0
    aLine$ = Chr(34) & LineNO$ & "@" & Distance & "%" & Chr(34)
    If printCurr Then
      ' Get branch current
      If GetSCCurrent( BrHnd1, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
      ' Output it
      If printCSV = 1 Then
        aLine$ = aLine$ & ","  & _
               MagArray(1) & "," & AngArray(1) & "," & _
               MagArray(2) & "," & AngArray(2) & "," & _
               MagArray(3) & "," & AngArray(3)
      Else
        aLine$ = LineNO$ & "@" & Distance & "%   " & _
               Format( "######.#", MagArray(1) ) & "@" & Format( "###.#", AngArray(1) ) & _
               Format( "######.#", MagArray(2) ) & "@" & Format( "###.#", AngArray(2) ) & _
               Format( "######.#", MagArray(3) ) & "@" & Format( "###.#", AngArray(3) )
      End If
    End If

    ' Get relay times
    ' Loop through all relays and find their operating times
    RelayHnd   = 0
    While GetRelay( RlyGrpHnd, RelayHnd ) > 0
      If HeaderLine$ <> "" Then
        TypeCode = EquipmentType( RelayHnd )
        Select Case TypeCode
          Case TC_RLYOCG 
            ParamID = OG_sID
            sType   = "OCG"
          Case TC_RLYOCP 
            ParamID = OP_sID
            sType   = "OCP"
          Case TC_RLYDSG
            ParamID = DG_sID
            sType   = "DSG"
          Case TC_RLYDSP
            ParamID = DP_sID
            sType   = "DSP"
          Case TC_FUSE
            ParamID = FS_sID
            sType   = "FUSE"
        End Select
        If GetData( RelayHnd, ParamID, sID$ ) > 0 Then
          sID$ = sType & " " & sID$
          HeaderLine$ = HeaderLine$ & "," & Chr(34) & sID & Chr(34)
        End If
      End If
      If GetRelayTime( RelayHnd, 1.0, OpTime ) > 0 Then 
        aLine$ = aLine$ & "," & OpTime
      End If
    Wend  'Each local relay
    
    ' Printout
    If HeaderLine <> "" Then
      Print #1, HeaderLine$
      HeaderLine$ = ""
    End If

    Print #1, aLine$

    Distance = Distance + StepSize
    If Distance > 100 Then
      Distance = StepSize
      LineNO$  = "L2"
    End If
    ShowFaultFlag = SF_NEXT   ' Show next fault
   Wend ' Each fault
   Print #1, ""

   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
End Sub
