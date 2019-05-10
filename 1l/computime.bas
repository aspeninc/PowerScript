' ASPEN PowerScript Sample Program
'
' COMPUTIME.BAS
'
' Report relay operating time given voltage and current input
'
' Version 1.1
' Category: OneLiner
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
Const ES_MEMO = ES_MULTILINE Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL Or ES_WANTRETURN ' Indicates multiline TextBox

Sub main()
   ' Variable declaration
   Dim RelayList(50) As String
   Dim RlyHndList(50) As Long
   dim vdVmag(3) As double
   dim vdVang(3) As double
   dim vdImag(5) As double
   dim vdIang(5) As double
   dim dVpreMag As double, dVpreAng As double
   
   ' Make sure a relay group or is selected
   If GetEquipment( TC_PICKED, PickedHnd& ) = 0 Or _
       EquipmentType( PickedHnd& ) <> TC_RLYGROUP  Then
     Print "Please select a relay group"
     Exit Sub
   End If

  
   ' Find relay bus nominal kV
   Call GetData( PickedHnd, RG_nBranchHnd, RlyBranchHnd )
   Call GetData( RlyBranchHnd, BR_nBus1Hnd, RlyBusHnd )
   Call GetData( RlyBusHnd, BUS_dkVNominal, dRlyBuskV )

   vdVmag(0) = dRlyBuskV / Sqr(3)
   vdVang(0) = 0
   vdVmag(1) = vdVmag(0)
   vdVmag(2) = vdVmag(0)
   vdVang(1) = vdVang(0) + 120
   vdVang(2) = vdVang(0) - 120
   dVpreMag  = dRlyBuskV
   dVpreAng  = 0
   
   ' Inventory of all relay at the selected location
   ' Loop through all relays and find their operating times
   RelayCount = 0
   RelayHnd&   = 0
   While GetRelay( PickedHnd&, RelayHnd& ) > 0
     TypeCode = EquipmentType( RelayHnd )
     If TypeCode = TC_RLYOCG Then 
       ParamID = OG_sID
       sDev    = "OCG"
     ElseIf TypeCode = TC_RLYOCP Then 
       ParamID = OP_sID
       sDev    = "OCP"
     ElseIf TypeCode = TC_RLYDSG Then 
       ParamID = DG_sID
       sDev    = "DSG"
     ElseIf TypeCode = TC_RLYDSP Then 
       ParamID = DP_sID
       sDev    = "DSP"
     ElseIf TypeCode = TC_FUSE   Then 
       ParamID = FS_sID
       sDev    = "FUSE"
     Else
       GoTo HasError
     End If
     If GetData( RelayHnd&, ParamID, sID$ ) = 0 Then GoTo HasError
     BrList$ = "[" + sDev + "] " + sID$
     RelayList(RelayCount) = BrList$
     RlyHndList(RelayCount) = RelayHnd&
     RelayCount = RelayCount + 1
   Wend  'Each relay
   
   If RelayCount = 0 Then
     Print "No relay found in the relay group"
     Stop
   End If

'=============Dialog Spec=============
Begin Dialog DIALOG_1 80,37,458,157, "Relay Time Calculation"
  Text 8,8,84,8,"Relay"
  Text 91,7,355,8,"Enter pre-fault and phase voltages and currents phasors as magnitude@angle (primary kV L-N, A and degree)"
  ListBox 8,20,78,55,RelayList(), .ListBox_1
  PushButton 203,69,102,12,"Compute Relay Operation", .Button2
  PushButton 267,134,59,13,"Done", .Button3
  PushButton 107,136,40,10,"Clear", .Button4
  PushButton 8,134,74,13,"Apply and Compute", .Button1
  TextBox 90,20,363,44,.Edit1
  TextBox 90,86,364,44,.Edit2
  TextBox 9,95,77,35,.Edit3
  PushButton 19,78,51,13,"Get Settings", .Button5
End Dialog

'=====================================

   Dim Dlg As Dialog_1
   dim sInput As String
   dim sOutput As String
   
   sInput    = "IP=1000@0 VP=" & Format(vdVmag(0),"0.00") & "@" & Format(vdVang(0),"0.00 ") & _
               "IA=1000@0 VA=1000@0 " & _
               "IB=1000@0 VB=1000@0 " & _
               "IC=1000@0 VC=1000@0 " & _
               "IN1=1000@0 " & _
               "IN2=1000@0 "
   sOutput   = ""
   sSettings = ""

   Do
    ' show the dialog
    dlg.Edit1 = sInput
    dlg.Edit2 = sOutput
    dlg.Edit3 = sSettings
    Button = Dialog( dlg )
    If Button = 0 Then Stop
    nRlyHnd = RlyHndList(Dlg.ListBox_1)
    sSettings = dlg.Edit3
    sInput    = dlg.Edit1
    sOutput   = dlg.Edit2
    If Button = 2 Then 
      Exit Sub	' Done
    ElseIf Button = 3 Then 
      sOutput = ""
	  GoTo Continue
    elseif Button = 4 Then
      If Len(sSettings) = 0 Then 
        Print "Invalid Setting"
        GoTo Continue
      End If
    elseif Button = 5 Then
      Call PrintSettings(nRlyHnd,sSettings) 
      GoTo continue
    elseif Button = 1 Then
    Else
	  GoTo Continue
    End If
    
    If Len(sInput) = 0 Then 
      Print "Invalid VI input"
      GoTo Continue
    End If

    sD$ = Chr(13) + Chr(10)
    nTokS = 1
    While nTokS > 0
      If Button = 4 Then
        sLineS$ = sSettings
        If Len(sLineS) > 0 And 0 = ProcessSettingsText(nRlyHnd,sLineS) Then
          Print "Unable to apply settings: " & sLineS
          GoTo continue
        End If
        nTokS = 0
      Else
        nTokS = 0
      End If
      nTokVI& = 1
      While nTokVI& > 0
        sLine$ = strToken( sInput, sD$, nTokVI& )
        If 0 < ProcessVIText(nRlyHnd, sLine$, vdVmag, vdVang, vdImag, vdIang, dVpreMag, dVpreAng ) Then
          Call ComputeRelayTime(nRlyHnd, vdImag, vdIang, vdVmag, vdVang, dVpreMag, dVpreAng, _
                dTime#, sDevice$ ) 
          sOutput = sOutput & RelayList(Dlg.ListBox_1) & ":" & _
                  " T="   & Format(dTime,"0.00") & "(" & sDevice & ")" & _
                  ";Ia="  & Format(vdImag(1),"0.0") & "@" & Format(vdIang(1),"0.0") & _
                  ";Ib="  & Format(vdImag(2),"0.0") & "@" & Format(vdIang(2),"0.0") & _
                  ";Ic="  & Format(vdImag(3),"0.0") & "@" & Format(vdIang(3),"0.0") & _
                  ";IN1=" & Format(vdImag(4),"0.0") & "@" & Format(vdIang(4),"0.0") & _
                  ";IN2=" & Format(vdImag(5),"0.0") & "@" & Format(vdIang(5),"0.0") & _
                  ";Va="  & Format(vdVmag(1),"0.0") & "@" & Format(vdVang(1),"0.0") & _
                  ";Vb="  & Format(vdVmag(2),"0.0") & "@" & Format(vdVang(2),"0.0") & _
                  ";Vc="  & Format(vdVmag(3),"0.0") & "@" & Format(vdVang(3),"0.0") & _
                  Chr(13) & Chr(10)
        End If
      Wend 'While nTokVI& > 0
    Wend 'nTokS > 0
Continue:    
   Loop
   
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
   Close
End Sub

Function Enable( ByVal ControlID$, ByVal Action%, ByVal SuppValue%) As long

   ReturnValue = 0

   DlgMain = ReturnValue

End Function

Function ProcessSettingsText(ByVal nRlyHnd&,ByVal S$) As long
  ProcessSettingsText = 0
  nTok& = 1
  S$ = UCase(S$)
  While nTok > 0
    sTok$ = strToken( S$, " " & chr(13) & chr(10), nTok )
    If KeyValPair(sTok,sKey$,dMag#,dAng#) = 0 Then GoTo cont1
    nParamID = 0
    If sKey = "CT" Then
      If TC_RLYOCG = EquipmentType(nRlyHnd) Then
        nParamID = OG_dCT
      elseif TC_RLYOCP = EquipmentType(nRlyHnd) Then
        nParamID = OP_dCT
      elseif TC_RLYDSG = EquipmentType(nRlyHnd) Then
        nParamID = DG_dCT
      elseif TC_RLYDSP = EquipmentType(nRlyHnd) Then
        nParamID = DP_dCT
	  End If
    ElseIf sKey = "TAP" Then
      If TC_RLYOCG = EquipmentType(nRlyHnd) Then
        nParamID = OG_dTap
      elseif TC_RLYOCP = EquipmentType(nRlyHnd) Then
        nParamID = OP_dTap
	  End If
    ElseIf sKey = "TD" Then
      If TC_RLYOCG = EquipmentType(nRlyHnd) Then
        nParamID = OG_dTDial
      elseif TC_RLYOCP = EquipmentType(nRlyHnd) Then
        nParamID = OP_dTDial
	  End If
    End If 'If sKey = 
    If nParamID = 0 Then GoTo cont1
    If 0 = SetData(nRlyHnd,nParamID,dMag) Then GoTo cont1
    ProcessSettingsText = ProcessSettingsText + 1
    cont1:
  Wend
  If 0 = PostData(nRlyHnd) Then ProcessSettingsText = 0
End Function

Function PrintSettings(ByVal nRlyHnd,ByRef sSettings)
  dim vLabel(255) As variant
  dim vSetting(255) As variant
 
  sSettings = "" 
  nRlyType = EquipmentType(nRlyHnd)
  If nRlyType = TC_RLYDSP Then
    nIDS& = DP_vParams
    nIDP& = DP_vParamLabels
    nIDCT& = DP_dCT
    nIDPT& = DP_dVT
  ElseIf nRlyType = TC_RLYDSG Then
    nIDS& = DG_vParams
    nIDP& = DG_vParamLabels
    nIDCT& = DG_dCT
    nIDPT& = DG_dVT
  elseif nRlyType = TC_RLYOCP Then
    nIDCT&  = OP_dCT
    nIDTAP& = OP_dTap
    nIDTD&  = OP_dTDial
  elseif nRlyType = TC_RLYOCG Then
    nIDCT&  = OG_dCT
    nIDTAP& = OG_dTap
    nIDTD&  = OG_dTDial
  Else
    exit Function
  End If
  If nRlyType = TC_RLYOCP Or nRlyType = TC_RLYOCG Then
    Call GetData(nRlyHnd,nIDCT,dCT)
    Call GetData(nRlyHnd,nIDTD,dTD)
    Call GetData(nRlyHnd,nIDTap,dTap)
    sSettings = sSettings & "CT=" & Format(dCT,"0.0 ") & chr(13)&Chr(10) & _
                 "TD=" & Format(dTD,"0.00 ") & chr(13)&chr(10) & _
                 "TAP=" & Format(dTap,"0.00 ")
  Else
    Call GetData(nRlyHnd,nIDS,vSetting)
    Call GetData(nRlyHnd,nIDP,vLabel)
    Call GetData(nRlyHnd,nIDCT,dCT)
    Call GetData(nRlyHnd,nIDPT,dPT)
    sSettings = sSettings & "CT=" & Format(dCT,"0.0 ") & chr(13)&chr(10) & _
                 "PT=" & Format(dPT,"0.00 ") &chr(13)&chr(10)
    For ii=1 to 255
      If vLabel(ii) = "" Then exit For
      sSettings = sSettings & vLabel(ii) & "=" & vSetting(ii) & Chr(13)&Chr(10)
    Next
  End If
End Function

Function ProcessVIText( ByVal RlyHnd, ByVal Str$, ByRef vdVmag() As double, ByRef vdVang() As double, _
  ByRef vdImag() As double, ByRef vdIang() As double, ByRef dVpreMag#, ByRef dVpreAng# ) As long
   
  dim nTok As long
  
  ProcessVIText = 0
  nTok = 1
  nI = 0
  nV = 0
  Str$ = UCase(Str$)
  While nTok > 0
    sTok$ = strToken( Str$, " ", nTok )
    nPos = StrPos("=",sTok)
    If nPos = 0 Then GoTo cont1
    sKey = Mid(sTok,1,nPos-1)
    sTok = Mid(sTok,nPos+1,99)
    nPos = StrPos("@",sTok)
    If nPos = 0 Then
      dMag = Val(sTok)
      dAng = 0
    Else
      dMag = Val(Mid(sTok,1,nPos-1))
      dAng = Val(Mid(sTok,nPos+1,99))
    End If
    If sKey = "IMAG" Then
      vdImag(0) = dMag
      nI = nI + 1
    elseif sKey = "IANG" Then
      vdIang(0) = dVal
      nI = nI + 1
    elseif sKey = "I"Then
      vdImag(0) = dMag
      vdIang(0) = dAng
      vdImag(1) = vdImag(0)
      vdImag(2) = vdImag(0)
      vdImag(3) = vdImag(0)
      vdImag(4) = vdImag(0)
      vdIang(1) = vdIang(0) + 120
      vdIang(2) = vdIang(0) - 120
      vdIang(3) = vdIang(0)
      vdIang(4) = vdIang(0)
      nI = nI + 1
    elseif sKey = "V" Then
      vdVmag(0) = dMag
      vdVang(0) = dAng
      vdVmag(1) = vdVmag(0)
      vdVmag(2) = vdVmag(0)
      vdVang(1) = vdVang(0) + 120
      vdVang(2) = vdVang(0) - 120
      dVpreMag  = vdVmag(0)
      dVpreAng  = vdVang(0)
      nV = nV + 1
    elseif sKey = "IA" Then
      vdImag(1) = dMag
      vdIang(1) = dAng
      nI = nI + 1
    elseif sKey = "IB" Then
      vdImag(2) = dMag
      vdIang(2) = dAng
      nI = nI + 1
    elseif sKey = "IC" Then
      vdImag(3) = dMag
      vdIang(3) = dAng
      nI = nI + 1
    elseif sKey = "IN1" Then
      vdImag(4) = dMag
      vdIang(4) = dAng
      nI = nI + 1
    elseif sKey = "IN2" Then
      vdImag(5) = dMag
      vdIang(5) = dAng
      nI = nI + 1
    elseif sKey = "VA" Then
      vdVmag(1) = dMag
      vdVang(1) = dAng
      nV = nV + 1
    elseif sKey = "VB" Then
      vdVmag(2) = dMag
      vdVang(2) = dAng
      nV = nV + 1
    elseif sKey = "VC" Then
      vdVmag(3) = dMag
      vdVang(3) = dAng
      nV = nV + 1
    elseif sKey = "VP" Then
      dVpreMag = dMag
      dVpreAng = dAng
      nV = nV + 1
    End If
    cont1:
  Wend
  ProcessVIText = nV + nI
End Function

Function StrPos( ByVal Substr$, ByVal Str$ ) As long
  nLen = Len(Substr)
  For ii = 1 to Len(Str)-nLen+1
    If StrComp( Substr, Mid(Str,ii,nLen) ) = 0 Then 
      StrPos = ii
      exit Function
    End If
  Next
  StrPos = 0
End Function

Function strToken( ByVal Str$, ByVal Delim$, ByRef nBegin& ) As String
  dim ii As long, nLen As long
  nLen = Len(str$)
  For ii = nBegin& to nLen ' Skip leading delimiters
    If StrPos(Mid(Str,ii,1),Delim$) = 0 Then exit For
  Next
  Do
   If ii > nLen Then exit Do
   If StrPos(Mid(Str,ii,1),Delim$) > 0 Then exit Do
   ii = ii + 1
  Loop
  strToken = Mid(Str,nBegin&,ii-nBegin&)
  Do ' Skip traling delimiters
    If ii > nLen Then exit Do
    If StrPos(Mid(Str,ii,1),Delim$) = 0 Then exit Do
    ii = ii + 1
  Loop
  If ii > nLen Then 
    nBegin& = 0 
  Else
    nBegin& = ii  
  End If
End Function

Function KeyValPair(S$,ByRef Key$, ByRef Mag#, ByRef Ang#) As long
  nRet& = 0
  Mag   = 0
  Ang   = 0
  nPos = StrPos("=",S)
  If nPos = 0 Then exit Sub
  Key = Mid(S,1,nPos-1)
  sVal = Mid(S,nPos+1,99)
  nPos = StrPos("@",sVal)
  If nPos = 0 Then
    Mag  = Val(sVal)
    Ang  = 0
    nRet = 1
  Else
    Mag  = Val(Mid(sVal,1,nPos-1))
    Ang  = Val(Mid(sVal,nPos+1,99))
    nRet = 2
  End If
  KeyValPair = nRet
End Function


