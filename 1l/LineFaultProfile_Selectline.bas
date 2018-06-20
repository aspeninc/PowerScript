' ASPEN PowerScrip sample program
'
' SELECTLINE.BAS
'
' Run query on transmission line data to select desired record
' Create csv file with list of selected lines
'
' Details are in LineFault_Profile.pdf 
'
'
' Version 1.0
' Category: OneLiner
'
'********** Windows dialog constants
Const ES_LEFT             = &h0000& 
Const ES_CENTER           = &h0001& 
Const ES_RIGHT            = &h0002& 
Const ES_MULTILINE        = &h0004& 
Const ES_UPPERCASE        = &h0008& 
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

'=============Dialog Spec=============
Begin Dialog SELELINE 97,14,271,217, "Line Selector"
  Text 8,7,21,8,"KV ="
  Text 67,8,42,8,"Bus name ="
  Text 97,47,73,11,"Label"
  Text 107,19,94,8," (Use % as wildcat char)"
  TextBox 26,6,36,11,.EditKV
  TextBox 110,6,104,11,.EditName
  TextBox 6,98,258,93,.EditLineList, ES_MEMO
  TextBox 6,29,259,47,.EditQueryResult, ES_MEMO
  TextBox 70,197,135,11,.EditFilePath
  PushButton 219,196,40,13,"Done", .Button2
  PushButton 221,6,40,13,"Query", .Button1
  PushButton 99,80,73,13,"Add to List Below", .Button3
  PushButton 6,196,61,13,"Save List to File", .Button4
End Dialog
'=====================================
Sub main

  dim dlg As SELELINE
  
  sKV$   = "0-999"
  sName$ = ""
  sRes$  = ""
  sList$ = ""
  sPath$ = "c:\000tmp\linelist.txt"
  nButton = -2
  While true 
    dlg.EditKV = sKV
    dlg.EditName = sName
    dlg.EditQueryResult = sRes$
    dlg.EditLineList    = sList$
    dlg.EditFilePath    = sPath$
    nButton = Dialog(dlg)
    If nButton = 0 Or nButton = 1 Then Stop
    sKV    = dlg.EditKV
    sName  = dlg.EditName
    sList$ = dlg.EditLineList 
    sPath$ = dlg.EditFilePath    
    If nButton = 2 Then
      nLines = queryLines( sKV, Trim(sName), sRes$ )
    End If
    If nButton = 3 And Len(sRes) Then
      sTmp$ = sRes
      Do While( Len(sTmp$) > 0 )
        Call parseALine(sTmp,Chr(13)&Chr(10),aLine$,sTmp)
        If Len(sList) > 0 Then sList = sList + Chr(13) + Chr(10)
        sList = sList + aLine
        Call parseALine(aLine, ",", aLeft$, aLine$)
        aName$ = aLeft$
        Call parseALine(aLine$,",",aLeft$,aLine$)
        Call parseALine(aLine$,",",aLeft$,aLine$)
        aName$ = aName & "_" & aLeft$
        Call parseALine(aLine$,",",aLeft$,aLine$)
        aName$ = aName & "_" & aLeft$
        Call parseALine(aLine$,",",aLeft$,aLine$)
        aName$ = aName & "_" & aLeft$ & ".CSV"
        sList = sList & "," & aName
      Loop
    End If
    If nButton = 4 And Len(sList) > 0 Then
     Open sPath For output As 1 
     Print #1, sList
     Print "List saved in file " & sPath$
     Close #1
    End If
  Wend  
End Sub

Function queryLines( sKV$, sName$, ByRef sResult$ ) As long
  sResult = ""
  queryLines = 0
  Call parseALine( sKV, "-", sLow$, sHigh$ )
  If sHigh = "" Then sHigh = sLow
  kvLow  = Val(sLow)
  kvHigh = Val(sHigh)
  If kvHigh < kvLow Then
    kvTemp = kvHigh
    kvHigh = kvLow
    kvLow  = kvTemp
  End If
  If kvHigh <= 0 Then kvHigh = 999.0
  BusHnd& = 0
  Do While GetEquipment( TC_BUS, BusHnd ) = 1
    Call GetData( BusHnd, BUS_nTapBus, tapFlag )
    If tapFlag = 0 Then
      Call GetData( BusHnd, BUS_dKVnominal, thisKV# )
      If thisKV >= kvLow And thisKV <= kvHigh Then
        Call GetData( BusHnd, BUS_sName, thisName$ )
        If StrStrWildCard( UCase(thisName),UCase(sName) ) > 0 Then
          nCount = linesAtBus( BusHnd, sTmp$ )
          If nCount > 0 Then
            If queryLines > 0 Then sResult = sResult + Chr(13) + Chr(10)
            sResult = sResult + sTmp
            queryLines = queryLines + 1
          End If
        End If
      End If
    End If
  Loop
  FindBus = 0
  
End Function


Function StrStrWildCard( ByVal sStr$, ByVal sSubStr$ ) As long
  StrStrWildCard = 0
  If Len(sSubStr) = 0 Then 
    exit Function
  End If
  nPos = InStr(1, sSubStr, "%")
  If nPos = 0 Then
    nRes = InStr( 1, sStr, ssubStr )
    If nRes > 0 Then StrStrWildCard = nRes
    exit Function
  End If
  If Len(sSubStr) = 1 Then 
    StrStrWildCard = 1
    exit Function
  End If
  If nPos = 1 Then
    sSubStr1$ = Mid(sSubStr, 2, 99)
    sStr1$ = Right(sStr, Len(sSubStr1) )
    If sStr1$ = sSubStr1 Then StrStrWildCard = 1
    exit Function
  End If
  If nPos = Len(sSubStr) Then
    sSubStr1$ = Left(sSubStr, nPos-1)
    sStr1$    = Left(sStr, nPos-1 )
    If sStr1$ = sSubStr1 Then StrStrWildCard = 1
    exit Function
  End If
  sSubStr1$ = Left(sSubStr, nPos-1)
  sStr1$    = Left(sStr, nPos-1 )
  If sStr1$ = sSubStr1 Then 
    sSubStr1$ = Mid(sSubStr, nPos+1, 99)
    sStr1$    = Right(sStr, Len(sSubStr1) )
    If sStr1$ = sSubStr1 Then StrStrWildCard = 1
  End If
End Function

Function linesAtBus( BusHnd&, ByRef sResult$ ) As long
  sResult = ""
  linesAtBus = 0
  BranchHnd = 0
  While GetBusEquipment( BusHnd, TC_BRANCH, BranchHnd ) > 0
    Call GetData( BranchHnd, BR_nType, nBrType& )
    If nBrType = TC_LINE Then
      Call GetData( BranchHnd, BR_nBus2Hnd, nHndFarBus& )
      Call GetData( BranchHnd, BR_nHandle, nItemHnd& )
      Call GetData( nItemHnd, LN_sID, sID$ )
      Call GetData( BusHnd, BUS_sName, sBus1$ )
      Call GetData( BusHnd, BUS_dKVnominal, dKV )
      Call GetData( nHndFarBus, BUS_sName, sBus2$ )
      
      If linesAtBus > 0 Then sResult = sResult & Chr(13) & Chr(10)
      sResult = sResult + sBus1 & "," & Str(dKV) & "," & sBus2 & "," & Str(dKV) & "," & sID
      linesAtBus = linesAtBus + 1
    End If
  Wend
End function


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
