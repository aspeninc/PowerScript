' SENSITIVITY.BAS
'
' Determine sensitivity of generator outage on bus fault current
'
' List of generator units are stored in GenFile in comma delimited format:
' AAAAA xxxKV, N
' Where: AAAA- Bus name; xxxx- nominal kV; N- Unit ID
'
' List of monitor buses are stored in BsFile in comma delimited format
' AAAA, xxxx
' Where: AAAA- Bus name; xxxx- nominal kV
'
' These files can be extracted from csv report from Generator and Bus data browser tables.
'
' Sensitivity output is in FileOut in comma delimited format
'
' Fault connection (3LG, 1LG) selection is done by setting the
' corresponding element of FltConnection() to 1
'
'
Const Mx_Gen=100
Const Mx_Bs=100
Const FileOut$ = "c:\sensitivity.csv"
Const GenFile$ = "c:\gen.csv"
Const BsFile$ = "c:\bus.csv"

Sub main

 Dim GenHnd(Mx_Gen)
 Dim BusHnd(Mx_Bs)

 'Open output file
 Open FileOut For Output As #2

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
  UID$ = Mid( TextLine, Pos1+1, 1 )
  If UID = Chr(34) Then UID$ = Mid( TextLine, Pos1+2, 1 )
  'Find bus handle
  If 0 <> FindBusByName( bName, KV, nBusHnd& ) Then
   'Find generator unit handle
   nGUHnd& = 0
   Do 
     nCode& = GetBusEquipment( nBusHnd, TC_GENUNIT, nGUHnd )
     If nCode <> 1 Then Exit Do
     Call GetParam( nGUHnd, GU_sID, sID$ )
     Call GetParam( nGUHnd, GU_nOnLine, nFlag& )
     If (sID = UID) And (nFlag=1)  Then
      nGenCount = nGenCount + 1 
      GenHnd(nGenCount) = nGUHnd
      LineOut$ = Chr(34) + "G" + Trim(Str(nGenCount)) + Chr(34)+","+Chr(34) + _
                 BName + " " + sKV + "kV Unit " + UID + Chr(34)
      Print #2, LineOut
      Exit Do
     End If
   Loop
  End If
 Loop
 Close #1

 'Read file from bus list
 nBSCount = 0
 Open BsFile For Input As #1
 Do While Not EOF(1)
  Line Input #1, TextLine
  'Extract bus name and kV
  Pos1 = InStr( 1, TextLine, "," )
  BName$ = Left( TextLine, Pos1-1 )
  If Left(BName,1)=Chr(34) Then BName=Mid(BName,2,99)  'Remove leading "
  If Right(BName,1)=Chr(34) Then BName=Left(BName,Len(BName)-1) 'Remove trailing "
  sKV = Mid( TextLine, Pos1+1, 99)
  If Left(sKV,1)=Chr(34) Then sKV=Mid(sKV,2,99)  'Remove leading "
  If Right(sKV,1)=Chr(34) Then sKV=Left(sKV,Len(sKV)-1) 'Remove trailing "
  KV = Val(sKV)
  'Find bus handle
  If 0 <> FindBusByName( bName, KV, nBusHnd& ) Then
   nBSCount = nBSCount + 1
   BusHnd(nBSCount) = nBusHnd
   LineOut$ = Chr(34) + "B" + Trim(Str(nBSCount)) + Chr(34)+","+Chr(34) + _
              BName + " " + sKV + "kV" + Chr(34)
   Print #2, LineOut
  End If
 Loop
 Close #1

 LineOut = Chr(34)+Chr(34)
 For ii=1 To nBSCount
   LineOut = LineOut + ","+Chr(34)+"B"+Trim(Str(ii))+Chr(34)
 Next
 Print #2, LineOut
 For ii=0 To nGenCount
  If ii > 0 Then
    'Take generator out of service
    nGUHnd = GenHnd(ii)
    nFlag  = 2
    Call SetData( nGUHnd, GU_nOnLine, nFlag& )
    Call PostData( nGUHnd )
  End If 

  'Simulate fault at buses in the list. Output result
  LineOut = Chr(34)+"G"+Trim(Str(ii))+Chr(34)
  For jj=1 To nBSCount
   nBus& = BusHnd(jj)
   Call DoOneFault( nBus, LineOut )
  Next
  Print #2, LineOut

  If ii > 0 Then
    'Put generator back in service
    nGUHnd = GenHnd(ii)
    nFlag  = 1
    Call SetData( nGUHnd, GU_nOnLine, nFlag& )
    Call PostData( nGUHnd )
  End If 
 Next

 Close #2

 Print "Sensitivity study completed. Results are in: " + FileOut
End Sub

'Simulate one fault. Output fault current
Function DoOneFault( ByVal nBus As Long, ByRef sLine As String) As Long
 Dim FltConnection(4) As Long
 Dim FltOption(14) As Double
 Dim OutageType(3) As Long
 Dim OutageList(15) As Long
 Dim ShowRelayFlag(4) As Long
 Dim Rflt As Double, Xflt As Double
 Dim vdMag(4) As Double, vdAng(4) As Double

' vdMag(1) = nBus
' GoTo Skip

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
 OutageList(1) = -1  ' Terminate the list for good measure
 For ii = 1 To 3
   OutageType(ii) = 0
 Next
 For ii = 1 To 4
   ShowRelayFlag(ii) = 0
 Next
 Rflt        = 0.0   ' No fault impedance
 Xflt        = 0.0
 ClearPrev   = 1     ' Don't keep previous result
 'Simulate fault at this point in the network
 If 0 = DoFault( nBus, FltConnection, FltOption, OutageType, OutageList, _
        Rflt, Xflt, ClearPrev ) Then Exit Function
 'Get fault current magnitude
 Call PickFault(1)
 Call GetSCCurrent( HND_SC, vdMag, vdAng, 4 )
 skip:
 sLine = sLine + "," + Str(vdMag(1))
 DoOneFault = 1  'Success
End Function