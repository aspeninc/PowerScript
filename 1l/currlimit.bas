' ASPEN PowerScript Sample Program
'
' CURRLIMIT.BAS
'
' Simulate current limiting fuse
' by applying fault impedance.
'
'
' Global vars and constant
'
'

Const FuseLimit = 10000

Begin Dialog FLTDLG 30,60,114,86, "Simulate Fault"
  OptionGroup .phase
    OptionButton 65,17,27,11, "1LG"
    OptionButton 24,17,30,11, "3LG"
  Text 8,43,59,8,"Current limit (A) ="
  TextBox 68,42,36,11,.Edit3
  PushButton 12,61,51,13,"Simulate", .Button1
  CancelButton 70,61,29,13
  GroupBox 7,7,97,31,"Phase connection"
End Dialog


Sub main()
   If GetEquipment( TC_PICKED, PickedHnd& ) = 0 Or _
      (EquipmentType( PickedHnd ) <> TC_BUS And EquipmentType( PickedHnd ) <> TC_RLYGROUP ) Then 
      Print "Please select a bus or relay group"
      Exit Sub
   End If

  dim dlg As FLTDLG
  Dim FltConn(5) As Long
  Dim FltOpt(15) As double
  Dim OutageOpt(5) As Long
  Dim OutageLst(30) As Long
  Dim vnShowRelay(4) As long
  dim vndMag(6) As double
  dim vndAng(6) As double
  
  ClearPrev     = 1 ' do not keep previous result
  For ii=0 to 14
    FltOpt(ii) = 0
  Next
  FltOpt(1)  = 1 ' close-in
  
  FltConn(1) = 0 ' Do 3PH
  FltConn(2) = 0
  FltConn(3) = 0   ' Do 1LG
  FltConn(4) = 0
  OutageOpt(1) = 0 ' With one outage at a time
  OutageOpt(2) = 0 ' With two outage at a time
  OutageOpt(3) = 0 ' With two outage at a time
  Rflt#        = 0 ' Fault R
  Xflt#        = 0 ' Fault X
  
  dlg.phase = 0
  dlg.edit3 = FuseLimit
  
  While 0 <> Dialog(dlg)
    bCont = false
    dLimit# = Val(dlg.Edit3)
    If dLimit <= 0 Then
      Print "Please enter current limit"
      bCont = true
    End If
    If Not bCont Then exit Do
  Wend
  
  If dlg.phase = 0 Then FltConn(3) = 1 Else FltConn(1) = 1
  
  While true
  
    ' Simulate the fault
    If DoFault( PickedHnd, FltConn, FltOpt, OutageOpt, OutageLst, _
           Rflt, Xflt, ClearPrev ) = 0 Then
      Print "Error: ", ErrorString( )
      Stop
    End If
    Call ShowFault( SF_LAST, 0, 4, 0, vnShowRelay )
    If 0 = GetSCCurrent( HND_SC, vndMag, vndAng, 4) Then GoTo hasError
    
    dCurr# = vndMag(1)
    If dCurr <= dLimit Then
      Print "Fault current (A)= " & Format(dCurr,"#.") & " is below specified limit of " & Str(dLimit)
      Stop
    End If
    aLine1$ = "Fault current (A)= " & Format(dCurr,"#.")
    If  FltConn(3) = 1 Then 
      If 0 = GetData( HND_SC, FT_dRPt, dRt# ) Then GoTo hasError
      If 0 = GetData( HND_SC, FT_dXPt, dXt# ) Then GoTo hasError
      If 0 = GetData( HND_SC, FT_dRNt, dR# ) Then GoTo hasError
      If 0 = GetData( HND_SC, FT_dXNt, dX# ) Then GoTo hasError
      dRt = dRt + dR
      dXt = dXt + dX
      If 0 = GetData( HND_SC, FT_dRZt, dR# ) Then GoTo hasError
      If 0 = GetData( HND_SC, FT_dXZt, dX# ) Then GoTo hasError
      dRt = dRt + dR
      dXt = dXt + dX
      Rflt = dRt*(dCurr/dLimit - 1)/3
      Xflt = dXt*(dCurr/dLimit - 1)/3
    Else
      If 0 = GetData( HND_SC, FT_dRPt, dRt# ) Then GoTo hasError
      If 0 = GetData( HND_SC, FT_dXPt, dXt# ) Then GoTo hasError
      Print "Zt=", dRt, dXt
      Rflt = dRt*(dCurr/dLimit - 1)
      Xflt = dXt*(dCurr/dLimit - 1)
    End If
    ' Simulate the fault
    If DoFault( PickedHnd, FltConn, FltOpt, OutageOpt, OutageLst, _
           Rflt, Xflt, ClearPrev ) = 0 Then
      Print "Error: ", ErrorString( )
      Stop
    End If
    Call ShowFault( SF_LAST, 0, 4, 0, vnShowRelay )
    If 0 = GetSCCurrent( HND_SC, vndMag, vndAng, 4) Then GoTo hasError
    dCurr# = vndMag(1)
    aLine2 = "Z (ohm)= " & Format(Rflt,"#0.00") & "+j" & Format(Xflt,"#0.00") & " Fault current (A)= " & Format(dCurr,"#.")
'    Print aLine1 & chr13 & Chr(10) & aLine2
    printTTY( aLine1 )
    printTTY( aLine2 )
    Stop
  Wend
hasError:
  Print ErrorString()
End Sub
