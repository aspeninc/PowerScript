' ASPEN PowerScript sample program
' ATCXLS.BAS
'
' Determine max transfer capability of a tie line
'
' Demonstrate how to:
'   Update network data 
'   Perform PF
'   Control other window program via Automation
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
  'Variable declarations
  Dim vdP(200) As Double, vdV(200) As Double
  Dim vdVal1(12) As Double, vdVal2(12) As Double
  Dim vdLoadMW(3) As Double, vdLoadMVAR(3) As Double
  Dim vnPFOption(10) As Long
  Dim vdPFCriteria(5) As Double

  If 1 <> DoInput( sSlackBus$, dSlackKV#, sGenBus$,  dGenKV#, _
                   sLoadBus$,  dLoadKV#,  sTieBus1$, dTie1KV#, _
                   sTieBus2$,  dTie2KV#,  dDemandFrom#, dDemandTo#, _
                   dDemandStep#, nDemandType&, sFilename$ ) Then Exit Sub

  ' Prepare output file
  Open sFilename$ For Output As 1

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

  If FindBusByName( sSlackBus$, dSlackKV#, nSlackBus& ) = 0 Then GoTo HasError
  If FindBusByName( sTieBus1$, dTie1KV#, nTieBus1& ) = 0 Then GoTo HasError
  If FindBusByName( sTieBus2$, dTie2KV#, nTieBus2& ) = 0 Then GoTo HasError

  ' Find the tie line
  nDevHnd& = 0
  nFound&  = 0
  While GetEquipment( TC_LINE, nLineHnd& ) > 0 And nFound = 0
      ' Print line info
      If GetData( nLineHnd&, LN_nBus1Hnd, nBus1Hnd& ) = 0 Then GoTo HasError
      If GetData( nLineHnd&, LN_nBus2Hnd, nBus2Hnd& ) = 0 Then GoTo HasError
      If (nBus1Hnd=nTieBus1 And nBus2Hnd=nTieBus2) Or _
         (nBus1Hnd=nTieBus2 And nBus2Hnd=nTieBus1)  Then nFound = 1
  Wend
  If nFound = 0 Then 
    Print "Tie line not found"
    Close
    Exit Sub
  End If
  sVal2$ = FullBusname( nBus1Hnd )
  sVal3$ = FullBusName( nBus2Hnd )

  ' Find generator unit handle
  If FindBusByName( sGenBus$, dGenKV#, nSendingBus& ) = 0 Then GoTo HasError
  nGenUnitHnd& = 0
  nFound&  = 0
  While GetBusEquipment( nSendingBus, TC_GENUNIT, nGenUnitHnd& ) > 0 And nFound = 0
      ' Print line info
      If GetData( nGenUnitHnd&, GU_sID, sID$ ) = 0 Then GoTo HasError
      If sID = "1" Then nFound = 1
  Wend
  If nFound = 0 Then 
    Print "Sending genunit not found"
    Close
    Stop
  End If
  
  ' Find the load unit handle
  If FindBusByName( sLoadBus$, dLoadKV#, nRecevingBus& ) = 0 Then GoTo HasError
  nLoadUnitHnd& = 0
  nFound&  = 0
  While GetBusEquipment( nRecevingBus, TC_LOADUNIT, nLoadUnitHnd& ) > 0 And nFound = 0
      ' Print line info
      If GetData( nLoadUnitHnd&, LU_sID, sID$ ) = 0 Then GoTo HasError
      If sID = "1" Then nFound = 1
  Wend
  If nFound = 0 Then 
    Print "Receiving load unit not found"
    Close
    Stop
  End If
  nNum = 0
  For dGen# = dDemandFrom# To dDemandTo# Step dDemandStep#
    ' Modify Generation data
    If SetData( nGenUnitHnd&, GU_dSchedP, dGen ) = 0 Then GoTo HasError
    If PostData( nGenUnitHnd& ) = 0 Then GoTo HasError
    ' Get load data
    vdLoadMW(nDemandType&) = dGen
    If SetData( nLoadUnitHnd&, LU_vdMW, vdLoadMW ) = 0 Then GoTo HasError
    If PostData( nLoadUnitHnd& ) = 0 Then GoTo HasError
    ' Do the power flow
    If DoPF( nSlackBus, vdPFCriteria, vnPFOption, nMethod ) = 0 Then 
      Close 1
      Print "PowerFlow failed at: "; dGen; " MW. " & _
             "Output written to" & sFileName$ & " Click OK to see Excel plot"
      Call PlotExcelGraph( vdP(), vdV(), nNum )
      Exit Sub
    End If

   'Get voltagge at the end bus
    If GetPFVoltage( nLineHnd&, vdVal1, vdVal2, ST_PU ) = 0 Then GoTo HasError
    dV1# = vdVal1(1)
    dV2# = vdVal1(2)
    ' Get power into end buses
    If GetFlow( nLineHnd&, vdVal1, vdVal2 ) = 0 Then GoTo HasError
    dP1# = vdVal1(1)
    dP2# = vdVal1(2)
    dQ1# = vdVal2(1)
    dQ2# = vdVal2(2)
    ' Store data for ploting
    nNum = nNum + 1
    vdP(nNum) = dP1
    vdV(nNum) = dV1
    ' Print it to file
    Print #1, _
              Format( dGen, "#0.0");   " , "; _
              Format( dQ1,  "#0.0");   " , "; _
              Format( dP1,  "#0.0");   " , "; _
              Format( dV1,  "#0.##0"); " , "; _
              Format( dQ2,  "#0.0");   " , "; _
              Format( dP2,  "#0.0");   " , "; _
              Format( dV2,  "#0.##0")
  Next	' every 2MW
  Close 1
  Print "Output written to" & sFileName$ & " Click OK to see Excel plot"
  Call PlotExcelGraph( vdP(), vdV(), nNum )
  Exit Sub
  HasError:
  Print "Error: ", ErrorString( )
  Exit Sub 
End Sub

Sub PlotExcelGraph( ByRef vdX() As Double, ByRef vdY() As Double, ByVal nSize )
  Dim xlApp As Object    ' Declare variable to hold the reference.
  Dim wkbook As Object    ' Declare variable to hold the reference.
  Dim rRange As Object
  Dim dataSheet As Object
  Dim cChart As Object

  ' Get Pointer to Excel application
  Set xlApp = CreateObject("excel.application")
  xlApp.Workbooks.Open Filename:=CurDir() + "\PlotingExample.xls"
  Set dataSheet = xlApp.Sheets("1")
  dataSheet.Range("$A1").Value = vdX(1)
  dataSheet.Range("$A2").Value = vdX(nSize)
  dataSheet.Range("$B1").Value = vdY(1)
  dataSheet.Range("$B2").Value = vdY(nSize)
  For ii = 2 To nSize-1
    sRange$ = "$A$" & ii & ":$B$" & ii
    Set rRange = xlApp.Sheets("1").Range(sRange)
    rRange.Insert (3)
    dataSheet.Range("$A" & ii).Value = vdX(ii)
    dataSheet.Range("$B" & ii).Value = vdY(ii)
  Next
  ' Now show excel graph
  xlApp.Visible = True
End Sub

Function DoInput( ByRef sSlackBus$, ByRef dSlackKV#, _
             ByRef sGenBus$,   ByRef dGenKV#, _
             ByRef sLoadBus$,  ByRef dLoadKV#, _
             ByRef sTieBus1$,  ByRef dTie1KV#, _
             ByRef sTieBus2$,  ByRef dTie2KV#, _
             ByRef dDemandFrom#, ByRef dDemandTo#, _
             ByRef dDemandStep#, ByRef nDemandType&, _
             ByRef sFilename$ ) As Long
Begin Dialog INPUTDLG 44,22, 289, 154, "Available Transfer Capacity Calculation"
  OptionGroup .GROUP_1
    OptionButton 28,104,52,8, "Constant P"
    OptionButton 84,104,52,8, "Constant I"
    OptionButton 136,104,52,8, "Constant Z"
  Text 8,12,64,8, "System slack bus"
  TextBox 76,8,60,12, .EditBox_1
  TextBox 140,8,16,12, .EditBox_2
  Text 160,12,16,8, "kV"
  Text 8,32,60,8, "Generator bus"
  TextBox 76,28,60,12, .EditBox_3
  TextBox 140,28,16,12, .EditBox_4
  Text 160,32,16,8, "kV"
  Text 8,52,60,8, "Load bus"
  TextBox 76,48,60,12, .EditBox_5
  TextBox 140,48,16,12, .EditBox_6
  Text 160,52,16,8, "kV"
  Text 8,68,32,8, "Tie line"
  TextBox 76,64,60,12, .EditBox_10
  TextBox 140,64,16,12, .EditBox_11
  Text 160,64,20,12, "kV  - "
  TextBox 184,64,60,12, .EditBox_12
  TextBox 248,64,16,12, .EditBox_13
  Text 268,68,16,8, "kV"
  Text 8,84,88,12, "Vary power demand from"
  TextBox 100,84,20,12, .EditBox_7
  Text 124,84,12,12, "To"
  TextBox 140,84,20,12, .EditBox_8
  Text 164,84,32,12, "MW, step"
  TextBox 200,84,20,12, .EditBox_14
  Text 8,124,60,8, "Output file name"
  TextBox 76,120,68,12, .EditBox_9
  OKButton 88,136,52,12
  CancelButton 172,136,40,12
End Dialog
  Dim dlg As INPUTDLG

  ' Initialize
  dlg.EditBox_1  = "Slack"
  dlg.EditBox_2  = 20.0
  dlg.EditBox_3  = "North"
  dlg.EditBox_4  = 500.0
  dlg.EditBox_5  = "Station S"
  dlg.EditBox_6  = 500.0
  dlg.EditBox_10 = "StationNN"
  dlg.EditBox_11 = 500.0
  dlg.EditBox_12 = "Station S"
  dlg.EditBox_13 = 500.0
  dlg.EditBox_7  = 1000
  dlg.EditBox_8  = 5000
  dlg.EditBox_14 = 50
  dlg.GROUP_1    = 1
  dlg.EditBox_9  = "atc.csv"

  ' Show the dialog
  button = Dialog( Dlg )
  If button = 0 Then
    DoInput = 0 
    Exit Function
  End If

  ' Record data
  sSlackBus$ = dlg.EditBox_1
  dSlackKV#  = dlg.EditBox_2
  sGenBus$   = dlg.EditBox_3
  dGenKV#    = dlg.EditBox_4
  sLoadBus$  = dlg.EditBox_5
  dLoadKV#   = dlg.EditBox_6
  sTieBus1$  = dlg.EditBox_10
  dTie1KV    = dlg.EditBox_11
  sTieBus2$  = dlg.EditBox_12
  dTie2KV#   = dlg.EditBox_13
  dDemandFrom# = dlg.EditBox_7
  dDemandTo#   = dlg.EditBox_8
  dDemandStep# = dlg.EditBox_14
  nDemandType& = dlg.GROUP_1 + 1
  sFilename$   = "atc.csv"
  DoInput = 1
End Function