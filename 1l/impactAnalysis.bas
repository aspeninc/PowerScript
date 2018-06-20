' ASPEN PowerScript sample program
'
' IMPACTANALYSIS.BAS
'
' Determine sensitivity of generator outage on a bus fault current
'
' List of generator units are stored in GenFile in comma delimited format:
' AAAAA xxxKV, N
' Where: AAAA- Bus name; xxxx- nominal kV; N- Unit ID
'
' This file can be extracted from csv report from Generator browser table.
'
' Sensitivity output is in FileOut in comma delimited format
'
' Version 1.0
' Category: OneLiner
'
Const Mx_Gen=100
Const Mx_Bs=100
Const FileOut$ = "c:\impact.csv"
Const GenFile$ = "c:\gen.csv"
Const BsFile$ = "c:\bus.csv"
Const HVDCBusName$ = "Nevada"
Const HVDCBusKV# = 132

Dim nHVDCBusHnd As long
Dim OutageList(15) As Long
Dim OutageType(3) As Long

Sub main

 Dim GenHnd(Mx_Gen) As long
 Dim BusHnd(Mx_Bs)  As long
 Dim nGenCount As long

 If 0 = FindBusByName( HVDCBusName$, HVDCBusKV, nHVDCBusHnd& ) Then
   Print "Error finding HVDC bus"
   exit Sub
 End If
 
 'Read file with generator list
 nGenCount = 0
 Open GenFile For Input As #1
 Do While Not EOF(1)
  Line Input #1, TextLine
  'Extract bus name and kV and genunit ID
  Pos2 = InStr( 2, TextLine, "kV" )
  Pos1 = Pos2 - 1
  Do While " " <> Mid( TextLine, Pos1, 1 )
    Pos1 = Pos1 - 1
  Loop
  sKV$ = Trim(Mid( TextLine, Pos1, Pos2-Pos1 ))
  KV# = Val(sKV)
  BName$ = Trim(Mid( TextLine, 1, Pos1-1 ))
  If Left(BName,1)=Chr(34) Then BName = Mid(BName,2,99)
  Pos1 = InStr( 1, TextLine, "," )
  UID$ = Mid( TextLine, Pos1+1, 2 )
  If UID = Chr(34) Then UID$ = Mid( TextLine, Pos1+2, 2 )
  'Find bus handle
  If 0 <> FindBusByName( bName, KV, nBusHnd& ) Then
   'Find generator unit handle
   nGUHnd& = 0
   Do 
     If 1 <> GetBusEquipment( nBusHnd, TC_GENUNIT, nGUHnd ) Then Exit Do
     Call GetData( nGUHnd, GU_sID, sID$ )
     Call GetData( nGUHnd, GU_nOnLine, nFlag& )
     If (sID = UID) And (nFlag<>0)  Then
      nGenCount = nGenCount + 1 
      GenHnd(nGenCount) = nGUHnd
      BusHnd(nGenCount) = nBusHnd
      Exit Do
     End If
   Loop
  End If
 Loop
 Close #1
 
 If nGenCount = 0 Then
   Print "No generator found"
   exit Sub
 End If

 'Open output file
 Open FileOut For Output As #2

 Print #2, "Gen Bus,Gen HV Bus,GenOut:Isc@HVDC,GenOut:Vpu@HV,GenOut:Isc@HV,GenOut:Vpu@HVDC,GenIn:Isc@HVDC"

 ' Simulate with no outage
 OutageList(1) = -1
 For ii = 1 To 3
   OutageType(ii) = 0
 Next
 
 ' Process each generator in the list
 For ii=1 To nGenCount
   Call ProcessGen( GenHnd(ii), BusHnd(ii) )
 Next

 Close #2

 Print "Sensitivity study completed. Results are in: " + FileOut
End Sub

Function ProcessGen( ByVal nGUHnd As long, ByVal nBusHnd As long ) As long
'  Dim nFlag As long		    ' Generator original status
'  Dim nHVBusHnd As long  	'High side bus handle
  
   ProcessGen = 0
   
  ' Check gen status
  If 0 = GetData( nGUHnd, GU_nOnLine, nFlag& ) Then GoTo hasError
  If nFlag = 0 Then exit Function	' Generator is deleted
 
  ' Find handle of generator HV bus 
  nHVBusHnd& = nBusHnd
  BranchHnd& = 0
  If GetBusEquipment( nBusHnd, TC_BRANCH, BranchHnd ) > 0 Then
    ' Check branch type. 
    Call GetData( BranchHnd, BR_nType, TypeCode& )
    If TypeCode = TC_XFMR Then 
      Call GetData( BranchHnd, BR_nBus2Hnd, nHVBusHnd )
    End If
  End If
    
    
  'Take generator out of service
  If nFlag = 1 Then
    nNewFlag&  = 2
    If SetData( nGUHnd, GU_nOnLine, nNewFlag ) = 0 Then GoTo hasError
    If PostData( nGUHnd ) = 0 Then GoTo hasError
  End If 

  OutputLine$ = FullBusName( nBusHnd ) + "," + FullBusName( nHVBusHnd )
  
  'Simulate fault at HVDC bus
  If DoOneFault( nHVDCBusHnd, nHVBusHnd, OutputLine$ ) = 0 Then Exit Function  ' Error
  
  'Simulate fault at generator HV bus
  If DoOneFault( nHVBusHnd, nHVDCBusHnd, OutputLine$ ) = 0 Then Exit Function  ' Error
  
  'Put generator in service
  nNewFlag  = 1
  Call SetData( nGUHnd, GU_nOnLine, nNewFlag )
  Call PostData( nGUHnd )
  
  'Simulate fault at HVDC bus
  If DoOneFault( nHVDCBusHnd, 0, OutputLine$ ) = 0 Then Exit Function  ' Error

  Print #2, OutputLine$
  
  'Put generator back in original state
  If nFlag = 2 Then
    Call SetData( nGUHnd, GU_nOnLine, nFlag& )
    Call PostData( nGUHnd )
  End If
  
  ProcessGen = 1	' Success
  exit Function
hasError:
  Print "Error: ", ErrorString( )
End Function

'Simulate one fault. Output fault current
Function DoOneFault( ByVal nBus As Long, ByVal nBus2 As Long, ByRef OutputLine As String ) As Long
 Dim FltConnection(4) As Long
 Dim FltOption(14) As Double
 Dim ShowRelayFlag(4) As Long
 Dim Rflt As Double, Xflt As Double
 Dim vdMag(9) As Double, vdAng(9) As Double


 ' Initialize 
 DoOneFault = 0
 FltConnection(1) = 1  '3LG
 FltConnection(2) = 0  '2LG
 FltConnection(3) = 0  '1LG
 FltConnection(4) = 0  'LL


 For ii = 1 To 14
   FltOption(ii) = 0.0
 Next
 FltOption(1)  = 1   ' Bus fault no outage
 For ii = 1 To 4
   ShowRelayFlag(ii) = 0
 Next
 Rflt        = 0.0   ' No fault impedance
 Xflt        = 0.0
 ClearPrev   = 1     ' Don't keep previous result
 
 'Simulate fault at nBus
 If 0 = DoFault( nBus, FltConnection, FltOption, OutageType, OutageList, _
        Rflt, Xflt, ClearPrev ) Then Exit Function
        
 'Get fault simulation result
 Call PickFault(1)
 
 ' Fault current
 Call GetSCCurrent( HND_SC, vdMag, vdAng, 4 )
 OutputLine$ = OutputLine$ + "," + Str(vdMag(1))
 
 If nBus2 <> 0 Then
   Call GetSCVoltage( nBus2, vdMag, vdAng, 4 )
   Call GetData( nBus2, BUS_dKVnorminal, KVNom# )
   OutputLine$ = OutputLine$ + "," + Str(vdMag(1)/KVNom*Sqr(3))
 End If
 
 DoOneFault = 1  'Success
End Function
