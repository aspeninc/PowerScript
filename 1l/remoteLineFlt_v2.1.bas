' ASPEN PowerScript Sample Program
'
' REMOTELINEFLT.BAS
'
' Simulate fault on remote line(s).
' Lines with tap buses are handled correctly
'
' Version 2.1
' Category: OneLiner
'
' Global vars
Const thisFile = "remoteLineFlt"
dim FarBrHnd(50) As long
dim FarLnZ(50) As double
dim FarFarBsName(50) As String
dim RemoteBs(1) As String
dim CountFarBr As long
dim ThisLnZ As double
dim PickedHnd As long
Begin Dialog REMOTELINEFLT 65,17,227,160, "Outage"
  Text 5,4,43,8,"Remote bus:"
  Text 163,71,10,8,"+j"
  Text 5,16,43,8,"Remote line:"
  Text 15,33,34,8,"Percent ="
  Text 172,33,16,8,"to"
  Text 11,124,79,8,"Save results in CSV file"
  GroupBox 116,91,103,28,"Double outage"
  GroupBox 7,91,103,28,"Single outage"
  GroupBox 121,60,87,28,"Fault Z (ohm)"
  GroupBox 6,60,110,28,"Phase connections"
  ListBox 47,3,164,11,RemoteBs(), .ListBox1
  DropListBox 47,15,164,120,FarFarBsName(), .ComboBox1
  TextBox 47,31,18,11,.Edit3_Pcnt
  CheckBox 72,32,79,11,"Auto Increment from", .CheckBox5_Ainc
  TextBox 150,31,18,11,.Edit4_From
  TextBox 181,31,18,11,.Edit5_To
  CheckBox 15,47,185,11,"Percentage is based on the relay line impedance", .CheckBox6_Zbase
  CheckBox 11,71,24,11,"3LG", .CheckBox1
  CheckBox 36,71,24,11,"2LG", .CheckBox2
  CheckBox 62,71,25,11,"1LG", .CheckBox3
  CheckBox 88,71,23,11,"L-L", .CheckBox4
  TextBox 128,70,27,11,.Edit1_R
  TextBox 172,70,27,11,.Edit2_X
  CheckBox 11,101,46,11,"Local bus", .CheckBox_outageNear
  CheckBox 57,101,52,11,"Remote bus", .CheckBox_OutageFar
  CheckBox 120,101,45,11,"Local bus", .CheckBox_DoutageNear
  CheckBox 166,101,50,11,"Remote bus", .CheckBox_DoutageFar
  TextBox 93,123,117,11,.Edit_csv
  CheckBox 10,138,86,11,"Clear previous results", .CheckBox5_clear
  PushButton 106,136,54,13,"Simulate", .Button1
  CancelButton 167,136,38,13
End Dialog



dim dlg As REMOTELINEFLT



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
   
   Call GetData( RlyBrHnd, BR_nBus1Hnd, NearBsHnd& )
   NearBsName = FullBusName(NearBsHnd)
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
           FarLnZ(CountFarBr)   = compuLineBrZ(BranchHnd, nDummy&, Bus2Hnd& )
           If NearBsName <> FullBusName(Bus2Hnd) Then
             FarFarBsName(CountFarBr) = "To " + FullBusName(Bus2Hnd) + " " + cktID$ + "L" 
             Call GetData( BranchHnd, BR_nHandle, LnHnd& )
             Call GetData( LnHnd, LN_sID, cktID$ )
             FarBrHnd(CountFarBr) = BranchHnd
             CountFarBr = CountFarBr + 1
           End If
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
     FarFarBsName(CountFarBr+4) = "All lines at the remote bus"
   End If
   
   RunFault

End Sub

Sub initDlg()
  dlg.Edit1_R = 0
  dlg.Edit2_X = 0
  dlg.Edit3_Pcnt = 10
  dlg.CheckBox1 = true
  dlg.CheckBox3 = true
  dlg.Edit4_From = 0
  dlg.Edit5_To   = 100
  dlg.Edit_csv   = "None"
  
  ' Attempt to read data from file
  On Error GoTo cont 
  Open thisFile & ".dlg" For Input As 1
  While Not EOF(1)
    Line Input #1, aLine
    If 1=InStr(1,aLine,"Edit1_R=") Then
      dlg.Edit1_R = Mid(aLine,Len("Edit1_R=")+1,99)
    elseif 1=InStr(1,aLine,"Edit2_X=") Then
      dlg.Edit2_X = Mid(aLine,Len("Edit2_X=")+1,99)
    elseif 1=InStr(1,aLine,"Edit3_Pcnt=") Then
      dlg.Edit3_Pcnt = Mid(aLine,Len("Edit3_Pcnt=")+1,99)
    elseif 1=InStr(1,aLine,"Edit4_From=") Then
      dlg.Edit4_From = Mid(aLine,Len("Edit4_From=")+1,99)
    elseif 1=InStr(1,aLine,"Edit5_To=") Then
      dlg.Edit5_To = Mid(aLine,Len("Edit5_To=")+1,99)
    elseif 1=InStr(1,aLine,"Edit_csv=") Then
      dlg.Edit_csv = Mid(aLine,Len("Edit_csv=")+1,99)
    elseif 1=InStr(1,aLine,"CheckBox1=") Then
      If Mid(aLine,Len("CheckBox1=")+1,99) = "1" Then dlg.CheckBox1 = true Else dlg.CheckBox1 = false 
    elseif 1=InStr(1,aLine,"CheckBox2=") Then
      If Mid(aLine,Len("CheckBox2=")+1,99) = "1" Then dlg.CheckBox2 = true Else dlg.CheckBox2 = false 
    elseif 1=InStr(1,aLine,"CheckBox3=") Then
      If Mid(aLine,Len("CheckBox3=")+1,99) = "1" Then dlg.CheckBox3 = true Else dlg.CheckBox3 = false 
    elseif 1=InStr(1,aLine,"CheckBox4=") Then
      If Mid(aLine,Len("CheckBox4=")+1,99) = "1" Then dlg.CheckBox4 = true Else dlg.CheckBox4 = false 
    elseif 1=InStr(1,aLine,"CheckBox5_Ainc=") Then
      If Mid(aLine,Len("CheckBox5_Ainc=")+1,99) = "1" Then dlg.CheckBox5_Ainc = true Else dlg.CheckBox5_Ainc = fase 
    elseif 1=InStr(1,aLine,"CheckBox6_Zbase=") Then
      If Mid(aLine,Len("CheckBox6_Zbase=")+1,99) = "1" Then dlg.CheckBox6_Zbase = true Else CheckBox6_Zbase = false 
    elseif 1=InStr(1,aLine,"CheckBox5_Clear=") Then
      If Mid(aLine,Len("CheckBox5_Clear=")+1,99) = "1" Then dlg.CheckBox5_Clear = true Else CheckBox5_Clear = false 
    elseif 1=InStr(1,aLine,"CheckBox_outageNear=") Then
      If Mid(aLine,Len("CheckBox_outageNear=")+1,99) = "1" Then dlg.CheckBox_outageNear = true Else dlg.CheckBox_outageNear = false 
    elseif 1=InStr(1,aLine,"CheckBox_outageFar=") Then
      If Mid(aLine,Len("CheckBox_outageFar=")+1,99) = "1" Then dlg.CheckBox_outageFar = true Else dlg.CheckBox_outageFar = false 
    elseif 1=InStr(1,aLine,"CheckBox_DoutageNear=") Then
      If Mid(aLine,Len("CheckBox_DoutageNear=")+1,99) = "1" Then dlg.CheckBox_DoutageNear = true Else dlg.CheckBox_DoutageNear = false 
    elseif 1=InStr(1,aLine,"CheckBox_DoutageFar=") Then
      If Mid(aLine,Len("CheckBox_DoutageFar=")+1,99) = "1" Then dlg.CheckBox_DoutageFar = true Else dlg.CheckBox_DoutageFar = false 
    End If
  Wend
  Close #1
  cont:

End Sub

Sub saveDlg()
  On Error GoTo cont 
  Open thisFile & ".dlg" For output As 1
  Print #1, "Edit1_R=" & dlg.Edit1_R
  Print #1, "Edit2_X=" & dlg.Edit2_X
  Print #1, "Edit3_Pcnt=" & dlg.Edit3_Pcnt
  Print #1, "Edit4_From=" & dlg.Edit4_From
  Print #1, "Edit5_To=" & dlg.Edit5_To
  Print #1, "Edit_csv=" & dlg.Edit_csv
  If dlg.CheckBox1 Then Print #1, "CheckBox1=1" Else Print #1, "CheckBox1=0"
  If dlg.CheckBox2 Then Print #1, "CheckBox2=1" Else Print #1, "CheckBox2=0"
  If dlg.CheckBox3 Then Print #1, "CheckBox3=1" Else Print #1, "CheckBox3=0"
  If dlg.CheckBox4 Then Print #1, "CheckBox4=1" Else Print #1, "CheckBox4=0"
  If dlg.CheckBox5_Ainc Then Print #1, "CheckBox5_Ainc=1" Else Print #1, "CheckBox5_Ainc=0"
  If dlg.CheckBox6_Zbase Then Print #1, "CheckBox6_Zbase=1" Else Print #1, "CheckBox6_Zbase=0"
  If dlg.CheckBox5_Clear Then Print #1, "CheckBox5_Clear=1" Else Print #1, "CheckBox5_Clear=0"
  If dlg.CheckBox_outageNear Then Print #1, "CheckBox_outageNear=1" Else Print #1, "CheckBox_outageNear=0"
  If dlg.CheckBox_outageFar Then Print #1, "CheckBox_outageFar=1" Else Print #1, "CheckBox_outageFar=0"
  If dlg.CheckBox_DoutageNear Then Print #1, "CheckBox_DoutageNear=1" Else Print #1, "CheckBox_DoutageNear=0"
  If dlg.CheckBox_DoutageFar Then Print #1, "CheckBox_DoutageFar=1" Else Print #1, "CheckBox_DoutageFar=0"
  Close #1
  cont:

End Sub

Sub printCSV( ByVal HndRly, ByRef csvName$)
 Dim MagArray(12) As Double
 Dim AngArray(12) As Double
 Dim DummyArray(6) As Long   '

 Open csvName For Output As 1 


 Call GetData( HndRly, RG_nBranchHnd, HndBranch& )
 Call GetData( HndBranch, BR_nBus1Hnd, Bus1Hnd& )
 
 StringVal$ = FullBusName( Bus1Hnd )
 
 nShow = SF_FIRST
 ' Must alway show fault before getting V and I
 If ShowFault( 1, 1, 4, 0, DummyArray ) = 0 Then GoTo HasError
 Do
  ' Get bus voltage
  If GetSCVoltage( Bus1Hnd, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Print it
  Print #1, "Va = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Vb = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Vc = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
  ' Print it

  ' Get branch current
  If GetSCCurrent( HndBranch, MagArray, AngArray, 4 ) = 0 Then GoTo HasError
  ' Print it
  Print #1, _
      "Ia = "; Format( MagArray(1), "#0.0"); "@"; Format( AngArray(1), "#0.0"); _
      "; Ib = "; Format( MagArray(2), "#0.0"); "@"; Format( AngArray(2), "#0.0"); _
      "; Ic = "; Format( MagArray(3), "#0.0"); "@"; Format( AngArray(3), "#0.0")
 Loop While ShowFault( SF_NEXT, 1, 4, 0, DummyArray ) > 0
 Close #1   'Close output file
 Exit Sub
HasError:
 Print "Error: "; ErrorString( ) 
End Sub

Sub runFault()
  Dim FltConn(5) As Long
  Dim FltOpt(15) As double
  Dim OutageOpt(5) As Long
  Dim OutageLst(30) As Long
  Dim vnShowRelay(4)
  
  
  For ii=0 to 14
    FltOpt(ii) = 0
  Next
  
  Call initDlg()
  
  While 0 <> Dialog(dlg)
    Call SaveDlg()
    nIdx& = dlg.ComboBox1
    doAll = false
    If nIdx >= CountFarBr Then
      If RemoteIdx = CountFarBr-1 Then   ' Shortes line
        zZ = 9999
        For ii=0 to CountFarBr-1
          If FarLnZ(ii) < zZ Then
            zZ = FarLnZ(ii) < zZ 
            nIdx = ii
          End If 
        Next
      ElseIf RemoteIdx = CountFarBr Then   ' Longest line
        zZ = 0
        For ii=0 to CountFarBr-1
          If FarLnZ(ii) > zZ Then
            zZ = FarLnZ(ii) < zZ 
            nIdx = ii
          End If 
        Next
      Else
        nIdx  = 0
        doAll = true
      End If
    End If
    RemoteIdx = nIdx
    
    If dlg.CheckBox5_Ainc Then
      FltOpt(13) = dlg.Edit4_From
      FltOpt(14) = dlg.Edit5_To
    End If
    
    FltConn(1) = dlg.CheckBox1   ' Do 3PH
    FltConn(2) = dlg.CheckBox2
    FltConn(3) = dlg.CheckBox3   ' Do 1LG
    FltConn(4) = dlg.CheckBox4
    OutageOpt(1) = 0 ' With one outage at a time
    OutageOpt(2) = 0 ' With two outage at a time
    OutageOpt(3) = 0 ' With two outage at a time
    Rflt#        = dlg.Edit1_R ' Fault R
    Xflt#        = dlg.Edit2_X ' Fault X
    ClearPrev    = dlg.CheckBox5_Clear ' keep previous result?

    For ii = RemoteIdx  to CountFarBr-1
      FltOpt(9)  = dlg.Edit3_Pcnt   ' Intermediate %
      If dlg.CheckBox6_Zbase <> 0 Then
        zZ# = FltOpt(9)*ThisLnZ/FarLnZ(ii)
        FltOpt(9) = zZ#
        If FltOpt(9) > 99.5 Then FltOpt(9) = 99.5
      End If
  
      ' Simulate the fault
      BrHnd& = FarBrHnd(ii)
      If DoFault( BrHnd, FltConn, FltOpt, OutageOpt, OutageLst, _
           Rflt, Xflt, ClearPrev ) = 0 Then
        Print "Error: ", ErrorString( )
        Stop
      End If
      If Not doAll Then exit For
      ClearPrev = 0
    Next
    Call ShowFault( SF_LAST, 0, 4, 0, vnShowRelay )
    If dlg.Edit_csv <> "" And dlg.Edit_csv <> "None" Then
      csvName$ = dlg.Edit_csv
      Call printCSV( PickedHnd, csvName$ )
     Print "Fault(s) simulated successfully. Results had been saved in " + csvName
    Else
     Print "Fault(s) simulated successfully."
    End If
    exit Do
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

