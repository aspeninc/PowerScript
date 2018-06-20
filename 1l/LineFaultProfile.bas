' ASPEN PowerScrip sample program
'
' LINEFLT.BAS
'
' Run intemmediate fault simulation on transmission line
' Read list of lines from an input CSV file
' Report fault currents in CSV file.
'
' Version 1.0
' Category: OneLiner
'
' Details input file format is in LineFault_Profile.pdf 
'
'******************************************************************************************************
'TODO: Modify parameters in this section if needed
 Const DataPath$        = "c:\000tmp\"
 Const FileNameIn$      = "linelist.txt"     ' Input file name
 Const PercentStepSize  = 40	             ' intermediate fault percent step
'
'******************************************************************************************************
'
'
Const ES_LEFT             = &h0000&  ' Try these different styles or-ed together
Const ES_CENTER           = &h0001&  ' as the last parameter of Textbox the change
Const ES_RIGHT            = &h0002&  ' the text box style.
Const ES_MULTILINE        = &h0004&  ' A 1 in the last parameter position defaults to 
Const ES_UPPERCASE        = &h0008&  ' A multiline, Wantreturn, AutoVScroll testbox.
Const ES_LOWERCASE        = &h0010&
Const ES_PASSWORD         = &h0020&
Const ES_AUTOVSCROLL      = &h0040&
Const ES_AUTOHSCROLL      = &h0080&
Const ES_NOHIDESEL        = &h0100&
Const ES_OEMCONVERT       = &h0400&
Const ES_READONLY         = &h0800&
Const ES_WANTRETURN       = &h1000&
Const ES_NUMBER           = &h2000&
Const WS_VSCROLL          = &h200000&

Const ES_MEMO = ES_MULTILINE Or ES_AUTOHSCROLL Or ES_WANTRETURN Or WS_VSCROLL  ' Indicates multiline TextBox


'*****************Dialog Spec***********************
Begin Dialog LINEFLT 97,14,273,217, "Run Line Fault"
  Text 7,8,41,8,"Data path ="
  Text 8,21,38,8,"Input file ="
  TextBox 48,8,158,11,.EditPath
  TextBox 48,20,158,11,.EditName
  TextBox 6,49,259,137,.EditList, ES_MEMO
  PushButton 206,193,40,13,"Done", .Button2
  PushButton 81,34,88,13,"Load data to list below", .Button1
  PushButton 87,194,73,13,"Run Line Faults", .Button3
End Dialog

'****************************************************

'Global variables

Dim StepSize As Double

Sub main()
  dim nChecked As long
  dim dlg As LINEFLT
  dim sFileNameOut As String
 
  sDataPath$        = DataPath
  sFileNameIn$      = FileNameIn  
  StepSize          = PercentStepSize
  sList$            = ""
     
  Do 
    dlg.EditPath = sDataPath
    dlg.EditName = sFileNameIn
    dlg.EditList = sList
    nButton = Dialog( dlg )
    If nButton <= 1 Then Stop

    sDataPath   = dlg.EditPath 
    sFileNameIn = dlg.EditName

    If nButton = 2 Or nButton = 3 Then
      Open sDataPath & sFileNameIn For Input As 2
      sList$ = ""
      Do While Not EOF(2)
       Line Input #2, aLine$ ' Read a line of data.
       If Len(sList) > 0 Then sList = sList & Chr(13) & Chr(10)
       sList = sList & aLine
      Loop
      If nButton = 2 Then GoTo ContinueDo
    End If

    nChecked = 0

    sData$ = sList$
    Do While Len(sData$) > 0 
      Call parseALine( sData, Chr(13) & Chr(10), aLine$, sData )
      ' Bus1, kV1, Bus2, kV2, cktID
      If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sBus1$,  aLine$ )
      If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sKV1$,   aLine$ )
      If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sBus2$,  aLine$ )
      If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sKV2$,   aLine$ )
      If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sCktID$, aLine$ )
      If Len(aLine$) > 0 Then Call parseALine( aLine$, ",", sFileNameOut$, aLine$ )

      If Len(sFileNameOut) = 0 Then sFileNameOut = sBus1 & "_" & sBus2 & "_" & sKV1 & "_" & sCktID & ".CSV"
      '       Print "Line = ", sBus1$, sKV1$, sBus2$, sKV2$, sCktID$, sFileNameOut      ' Construct message.
        
      BranchHnd = branchSearch( sBus1$, Val(sKV1$), sBus2$, Val(sKV2$), sCktID$ )
      If BranchHnd > 0 Then
        ' Get branch type
        Call GetData( BranchHnd, BR_nType, TypeCode )
        If TypeCode = TC_LINE Then
         If SimulateLineFaults( BranchHnd, sDataPath & sFileNameOut ) > 0 Then nChecked = nChecked + 1 Else GoTo HasError
        End If
      End If
    Loop
    If nChecked > 0 Then  Print "Simulated", nChecked, " lines. Output is in folder " & sDataPath
  continueDo:
  Loop

 
 If nChecked > 0 Then
'  sMsg$ = "Checked " + Str(nChecked) + " relays. Report is in " + sFile _
'           + Chr(13) + "Do you want to open this file in Excel?"
'  If 6 = MsgBox( sMsg, 4, "Check DS Zone" ) Then
'   Set xlApp = CreateObject("excel.application")
'   xlApp.Workbooks.Open Filename := sFile
'   xlApp.Visible = True
 End If
 
 exit Sub
HasError:
  Print "Error: ", ErrorString( )
End Sub


Function SimulateLineFaults( ByVal BranchHnd&, ByVal sPathOut ) As long
  Dim FltConnection(4) As Long
  Dim FltOption(14) As Double
  Dim OutageType(3) As Long
  Dim OutageList(15) As Long
  Dim FltConnStr(4) As String
  dim vdMag(12) As double
  dim vdAng(12) As double
  dim vdMagN(12) As double
  dim vdAngN(12) As double
  Dim DummyArray(6) As Long   '

  Open sPathOut For output As 1 
  
  For ii = 1 To 14
  FltOption(ii) = 0.0
  Next
  For ii = 1 To 4
  FltConnection(ii) = 1
  Next
  For ii = 1 To 3
   OutageType(ii) = 0
  Next

  FltOption(13) = 0			'Intermediate percent from
  FltOption(14) = 0   		'Intermediate percent to

  dFltR     = 0
  dFltX     = 0

  Percent# = StepSize
  Do While Percent < 100
    sOut$ = ""
    'Simulate faults
    FltOption(9)  = Percent  	'Intermediate percent
    If 0 = DoFault( BranchHnd, FltConnection, FltOption, OutageType, OutageList, dFltR, dFltX, 1 ) Then
      SimulateLineFaults = 0
      exit Function
    End If

    PickFault( SF_FIRST )
    If ShowFault( SF_FIRST, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError
    Do
      ' Total fault current
      If 0 = GetSCCurrent( HND_SC, vdMag, vdAng, 4 ) Then GoTo HasError
      ' Current from near end
      If 0 = GetSCCurrent( BranchHnd, vdMagN, vdAngN, 4 ) Then GoTo HasError
      FaultIMag# = 0
      LineIMag#  = 0
      For ii = 1 to 3
       ' Compute contribution from remote end by substracting near end contribution from total
       LineIMagF = cSub( vdMag(ii), vdAng(ii), vdMagN(ii), vdAngN(ii) )
       ' Highest value
       If FaultIMag < vdMag(ii)  Then FaultIMag# = vdMag(ii)
       If LineIMag  < vdMagN(ii) Then LineIMag#  = vdMagN(ii)
       If LineIMag  < LineIMagF  Then LineIMag#  = LineIMagF
      Next

      Call GetData( HND_SC, FT_dXR, dXR# )

      sFltDesc$ = FaultDescription()
      If InStr( 1, sFltDesc, " 3LG " ) > 0 Then _
        sFltConn = "3LG" _
      Else If InStr( 1, sFltDesc, " LL " ) > 0 Then _
        sFltConn = "LL" _
      Else If InStr( 1, sFltDesc, " 1LG " ) > 0 Then _
        sFltConn = "1LG" _
      Else If InStr( 1, sFltDesc, " 2LG " ) > 0 Then _
        sFltConn = "2LG"
'      sOut$ = Str(Percent) & "," & sFltConn & "," _
'               & Chr(13) & Chr(10) & " FaultI=" & polarFormatI( FaultIMag, FaultIAng ) _
'               & Chr(13) & Chr(10) & " LineI=" & polarFormatI( LineIMag, LineIAng ) _
'               & Chr(13) & Chr(10) & " X/R=" & Format(dXR, "#0.00" )
      If Len(sOut$) = 0 Then sOut$ = Format(Percent,"#0.00")
      sOut$ = sOut$ & "," & sFltConn & "," & Format(FaultIMag,"#0") _
               & "," & Format(dXR,"#0.00") & "," & Format(LineIMag,"#0")
    Loop While PickFault( SF_NEXT ) > 0
    Print #1, sOut$  
    Percent = Percent + StepSize
  Loop

  Close 1
  
  SimulateLineFaults = 1
  exit Function
HasError:
  Print "Error: ", ErrorString( )
  stop
End Function	'SimulateLineFaults

Function polarFormatI( Mag#, Ang# ) As String
  polarFormatI = Format( Mag, "#0.0") & "@" & Format( Ang, "#0.00")
End Function

Function cSub( Mag1#, Ang1#, Mag2#, Ang2# ) As double
  Pi# = 3.14156
  Ang1 = Ang1/180*Pi
  Ang2 = Ang2/180*Pi
  dReal# = Mag1 * Cos(Ang1) - Mag2 * Cos(Ang2)
  dImag# = Mag1 * Sin(Ang1) -  Mag2 * Sin(Ang2)
  cSub = Sqr(dReal*dReal + dImag*dImag)
End Function

Sub parseALine( ByVal aLine$, ByVal Delim$, ByRef sLeft$,  ByRef sRight$ )
  nPos = InStr( 1, aLine$, Delim$ )
  If nPos = 0 Then
    sLeft = aLine$
    sRight = ""
  Else
    sLeft = Left(aLine$, nPos-1)
    sRight = Mid(aLine$, nPos+Len(Delim), 9999 )
  End If
  sLeft  = Trim(sLeft)
  sRight = Trim(sRight)
End Sub

Function  branchSearch( sBus1$, KV1#, sBus2$, KV2#, sCktID$ ) As long
  branchSearch = 0
  If 0 = FindBus( sBus1, KV1, nHndBus1& ) Then exit Function
  If 0 = FindBus( sBus2, KV2, nHndBus2& ) Then exit Function
  BranchHnd = 0
  While GetBusEquipment( nHndBus1, TC_BRANCH, BranchHnd ) > 0
    Call GetData( BranchHnd, BR_nBus2Hnd, nHndFarBus& )
    If nHndFarBus = nHndBus2 Then
      Call GetData( BranchHnd, BR_nType, nBrType& )
      select case nBrType
        case TC_LINE
          nCode = LN_sID
        case TC_XFMR
          nCode = XF_sID
        case TC_XFMR3
          nCode = X3_sID
        case TC_PS
          nCode = PS_sID
      End select
      Call GetData( BranchHnd, BR_nHandle, nItemHnd& )
      Call GetData( nItemHnd, nCode, sID$ )
      If sID = sCktID Then
        branchSearch = BranchHnd
        exit Function
      End If
    End If
  Wend
  branchSearchReturn:
End Function

Function FindBus( sName$, dKV#, ByRef BusHnd& ) As long
  BusHnd& = 0
  Do While GetEquipment( TC_BUS, BusHnd ) = 1
    Call GetData( BusHnd, BUS_sName, thisName$ )
    If thisName = sName Then
      Call GetData( BusHnd, BUS_dKVnominal, thisKV# )
      If Abs(dKV-thisKV) < 0.000001 Then
        FindBus = 1
        exit Function
      End If
    End If
  Loop
  FindBus = 0
End Function
