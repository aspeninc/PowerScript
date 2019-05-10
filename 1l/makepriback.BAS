' ASPEN PowerScript sample program
'
' MAKEPRIBACK.BAS
'
' Generate list of backups relay groups on transmission lines
' using following criteria:
' - Backup must be located on transmission line behind the primary
' - Close-in with end open fault at the primary group must be
'   seen as forward fault (MTA = 30) at backup location
' Output RAT file with coordination pairs which can be imported to OLR
' file using Relay | Import command
'
' Version: 1.0
'
' Remark:If a relaygroup is selected before script is launched
'        only this group will be processed
'
Const optMaxTiers = 9
Const optMaxKVLevels = 0
Const optExistingPairsOnly = 0
Const debug = 0  '=1: Print trace

' Global variables
dim crlf As String
dim countEx&         ' New pairs found by script
dim countOld&        ' Existing pairs
dim countNF&         ' Existing pairs not found by script
dim countPartial&    ' Partially backed up groups
dim countTotal&      ' Total groups checked
dim countNoBackup&   ' Groups with no pair found
dim countFull&       ' Fully backed up groups
dim kvLo#            ' Lower kV limit
dim kvHi#            ' Upper kV limit
dim optDoRAT&        ' Output RAT file
dim optDoTXT&        ' Output TXT report

Sub main
 crlf = chr(13)&chr(10)
 countGr& = 0
 countEx& = 0
 countNF& = 0
 countOld& = 0
 countPartial& = 0
 countTotal& = 0
 countNoBackup& = 0
 countFull& = 0
 kvLo# = 132
 kvHi# = 999
 optDoRAT& = 0
 optDoTXT& = 0
 fnTXT$ = GetOlrFileName() & ".txt"
 fnRAT$ = GetOlrFileName() & ".rat"
 
 Call doUIInput(kvLo,kvHi,fnTXT$,fnRAT$)
 
 If fnTXT <> "" Then
  Open fnTXT For output As 1
  optDoTXT = 1
 End If
 
 If fnRAT <> "" Then
  Open fnRAT For output As 2
  Print #2, "[ASPEN RELAY DATA]"
  Print #2, "delimiter='"
  Print #2, "app= 'ASPEN OneLiner and Power Flow'"
  Print #2, "ver= 2014 'A' 14 6"
'  Print #2, "Date= 'Sat Dec 29 14:16:32 2018'"
  Print #2, "" 
  Print #2, "COORDINATION PAIRS"
  optDoRAT = 1
 End If
 
 bDoOne = false
 If 1 = GetEquipment(TC_PICKED, hnd&) Then
  If TC_RLYGROUP = EquipmentType(hnd&) Then bDoOne = true
 End If
 If Not bDoOne Then _
  Call doAllGroups(kvLo,kvHi) _
 Else _
  Call doOneGroup(0)
 a$ = "Total groups processed = " & Str(countTotal) & _
       crlf & "Groups with full backup = " & Str(countFull) & _
       crlf & "(#)Groups with patial backup = " & Str(countPartial) & _
       crlf & "(##)Groups with no backup = " & Str(countNoBackup) & _
       crlf & "(*)New pairs found by script = " & Str(countEx) & _
       crlf & "(!)Existing pairs = " & Str(countOld) & _
       crlf & "(!!)Existing pairs missed by script = " & Str(countNF)
 If optDoTXT Then 
  Print #1, a
  Print a & crlf & "Report in: " & fnTXT
 Else
  Print a
 End If
End Sub

Function doUIInput(ByRef kvLo#, ByRef kvHi#, ByRef fnTXT$, ByRef fnRAT$ ) As long
 Begin Dialog UIInput 60,26,309,81, "Coordination Pair Generator"
  Text 4,6,44,8,"KV Range:"
  Text 4,19,45,8,"Report file:"
  Text 4,32,37,8,"RAT file:"
  Text 7,47,145,8,"(Leave edit box empty if file is not desired)"
  Text 73,6,7,8,"-"
  TextBox 45,31,260,11,.FileRAT
  TextBox 45,18,260,11,.FileReport
  TextBox 45,4,25,11,.KVLo
  TextBox 80,4,25,11,.KVHi
  OKButton 133,60,57,13
  CancelButton 197,60,30,13
 End Dialog

 Dim Dlg1 As UIInput
 Dlg1.KVLo = Str(kvLo)
 Dlg1.KVHi = Str(kvHi)
 Dlg1.FileReport = GetOlrFileName() & ".txt"
 Dlg1.FileRAT = GetOlrFileName() & ".rat"

 Button = Dialog (Dlg1)
 If button = 0 Then Stop
 kvLo# = Val(dlg1.KVLo)
 kvHi# = Val(dlg1.KVHi)
 fnTXT$ = Dlg1.FileReport
 fnRAT$ = Dlg1.FileRAT
 
End Function

Sub doAllGroups(doKVlo#,doKVhi#)
' Check all group with existig pairs
 hndBs& = 0
 ii = 0
 While 1 = GetEquipment(TC_BUS,hndBs)
  Call GetData(hndBs,BUS_dKVNominal,kV#)
  If kV >= doKVlo And kV <= doKVhi Then
   hndBr& = 0
   While 1 = GetBusEquipment(hndBs,TC_BRANCH,hndBr)
    Call GetData(hndBr,BR_nType,brType&)
    Call GetData(hndBr,BR_nInservice,nFlag&)
    If nFlag = 1 And brType = TC_LINE Then
     hndGr& = 0
     Call GetData(hndBr,BR_nRlyGrp1Hnd,hndGr&)
     If hndGr& > 0 Then
      If ii < -99 Then ii = 0 Else ii = ii - 2
      Button = ProgressDialog( 1, "Creating coordination pairs", fullBranchName(hndGr), ii )
      If Button = 2 Then 
       ProgressDialog( 0 )
       Stop
      End If
      Call doOneGroup(hndGr)
      If debug > 0 Then
       'If "y" <> InputBox("Continue?", "Confirm", "y" ) Then Stop
      End If
     End If
    End If
   Wend
  End If
 Wend
 ProgressDialog( 0 )
End Sub

Function doOneGroup(hnd&) As long
 If hnd = 0 Then
  bOK = false
  If 1 = GetEquipment(TC_PICKED, hnd&) Then
   If TC_RLYGROUP = EquipmentType(hnd&) Then bOK = true
  End If
  If Not bOK Then
   Print "Must select a relay group"
   exit Sub
  End If
 End If
 
 dim listB(20) As long,listP(20) As long  ' List of existing pairs
 
 Call listPribackPairs(hnd,listB,countB&,listP,countP&)
 If optExistingPairsOnly = 1 Then
  If countB = 0 Then exit Function
 End If
 
 countOld = countOld + countB
 
 dim listBacGr(100) As long, countBacGr As long
 countBacGr& = 0

 nRet = findBackupGroups(hnd, listBacGr, countBacGr&)
 countTotal = countTotal + 1
 
 ' Report
 a$ = "Relaygroup: " & fullBranchName(hnd)
 
 If countBacGr + countB = 0 Then
  a$ = a$ & crlf & "No backup (##)"
  If optDoTXT = 1 Then Print #1, a & crlf
  countNoBackup = countNoBackup + 1
  exit Function
 End If
 
 If nRet = 1 Then
  countFull = countFull + 1
  a$ = a$ & crlf & "Backups:"
 Else
  countPartial = countPartial + 1
  a$ = a$ & crlf & "Backups (#):"
 End If

 For i& = 1 to countBacGr
  bFound = false
  For j& = 1 to countB
   If listB(j) = listBacGr(i) Then
    bFound = true
    listB(j) = 0
    GoTo exitFor
   End If
  Next
  exitFor:
  a$ = a$ & crlf & fullBranchName(listBacGr(i))
  If Not bFound Then 
   a$ = a$ & " (*)"
   countEx = countEx + 1
   If optDoRAT = 1 Then Print #2, outRATCoordPair(hnd,listBacGr(i))
  Else
   a$ = a$ & " (!)"
  End If
 Next
 For i& = 1 to countB
  If listB(i) > 0 Then
   a$ = a$ & crlf & fullBranchName(listB(i)) & "(!!)"
   countNF = countNF + 1
  End If
 Next
 If optDoTXT = 1 Then Print #1, a & crlf
End Function

Const maxStack& = 200
dim listToDo(maxStack) As long, listTierLevel(maxStack) As long, listKVLevel(maxStack) As long, countToDo&
Function findBackupGroups(hndGr&, ByRef listGr() As long, ByRef countGr& ) As long
' Find all backup groups of hndGr, which
'  - is behind hndGr within tier and kvLevelNo limits. AND
'  - has in forward direction the close-in with EOP fault at hndGr
'
' Return:
'  0- Found no valid branch behind this group to traverse
'  1- Successfull traversal on every branch behind hndGr
'  2- traversal was interrupted on some branches behind hndGr
'
'
 Call getdata(hndGr,RG_nBranchHnd,hndBr&)
 
 dim listBs(100) As long  ' list of buses of primary line
 dim listBr(500) As long  ' list of branches already traversed 
 
 countGr& = 0
 countBr& = 0
 countBs& = 0
 findBackupGroups = 0
 
 Call listSegsAndBs(hndBr,listBr,countBr&,listBs,countBs&)
 
 If debug = 1 Then
  a$ = "Rly group: " & FullBranchName(hndGr) & crlf
  If countBr > 1 Then
   For i& = 1 to countBr
    a$ = a$ & crlf & fullBranchName(listBr(i))
   Next
  End If
  Print a
 End If
 
 ' Traverse backward from primary bus. 
 ' List all relay groups that sees the fault as forward
 countToDo& = 0
 Call getdata(hndBr,BR_nBus1Hnd,hndBsPri&)
 hndBr& = 0
 While getbusequipment(hndBsPri,TC_BRANCH,hndBr) = 1
  Call getdata(hndBr,BR_nInService,flag&)
  If flag = 1 And 0 = findBranchInList(listBr,countBr,hndBr,0) Then 
   countToDo = countToDo + 1
   listToDo(countToDo) = hndBr
   listTierLevel(countToDo) = 1
   Call getdata(hndBr,BR_nType,brType&)
   If brType <> TC_LINE Then listKVLevel(countToDo) = 1 Else listKVLevel(countToDo) = 0
  End If
 Wend
 If countToDo = 0 Then
  findBackupGroups = 0
  exit Function
 End If
 
 ' Run a close-in fault with end open 
 Call runCEOFault(hndGr)
 
 countLeafHit& = 0
 i& = 1
 While i <= countToDo
  ' Look for backup group on all ToDo branches.
  ' traverseOneBranch() may add more branch to listToDo
  t1& = listTierLevel(i)
  t2& = listKVLevel(i)
  h& = listToDo(i)
  nRet = traverseOneBranch(h,listBr,countBr&,listBs,countBs&,listGr,countGr&,t1,t2)
  If nRet = 3 Then
   countLeafHit = countLeafHit + 1
  End If
  i = i + 1
 Wend
 If debug = 2 Then
  a$ = "Primary group: " & FullBranchName(hndGr) & crlf
  PrintTTY(" ")
  PrintTTY(a)
  a$ = "Found " & Str(countGr&) & " backups"
  PrintTTY(a)
  For i = 1 to countGr&
   a = "  " & FullBranchName(listGr(i))
   PrintTTY(a)
  Next
 End If
 If countLeafHit = 0 Then findBackupGroups = 1 Else findBackupGroups = 2
End Function

Function traverseOneBranch(ByVal hndBr&, _
       ByRef listBr() As long,ByRef countBr&, _
       ByRef listBs() As long, ByRef countBs&, _
       ByRef listGr() As long ,ByRef countGr&, _
       ByVal thisTier&, ByVal thisKVLevelNo& ) As long
' Traverse a branch to look for backup relay group
' If none is found update todoList for further action up to optMaxTiers
' Return: 0-Did not traverse (branch had been encontered before); 
'         1-Backup Found; 
'         2-Not found. Todo list is updated;
'         3-Not found. Limit hit
'
 dim listBrX(100) As long, listBsX(100) As long
 countBrX& = 0
 countBsX& = 0
 Call listSegsAndBs(hndBr,listBrX,countBrX&,listBsX,countBsX&)
 ' Reject all loops
 For i& = 1 to countBrX
  If findBranchInList(listBr,countBr,listBrX(i),0) <> 0 Then 
   ' found loop. Add the rest of line segs to list of already traversed branches
   ' then quit
   For j& = i + 1 to countBrX 
    Call findBranchInList(listBr,countBr,listBrX(j),1)
   Next
   If debug = 1 Then Print "Rejecting traversed branch: " & FullBranchName(listBrX(i))
   traverseOneBranch = 0
   exit Function
  End If
 Next 
 Call getdata(hndBr,BR_nBus1Hnd,hndB1&)
 For i& = 1 to countBsX
  If hndB1 <> listBsx(i) Then 
   If findInList(listBs,countBs,listBsx(i),0) <> 0 Then
    ' Loop
    If debug = 1 Then Print "Rejecting bus " & FullBusName(listBsx(i))
    For j& = 1 to countBrX 
     Call findBranchInList(listBr,countBr,listBrX(j),1)
    Next
    traverseOneBranch = 0
    exit Function
   End If
  End If
 Next
 
 If debug = 1 Then
  Print "Traversing " & FullBranchName(hndBr)
 End If
 
 countGrS& = countGr
 For i& = 1 to countBrX
  hndBrX& = listBrX(i)
  For j& = 1 to 3
   hndGr& = 0
   If j = 1 Then Call getdata(hndBrX,BR_nRlyGrp1Hnd,hndGr&)
   If j = 2 Then Call getdata(hndBrX,BR_nRlyGrp2Hnd,hndGr&)
'   If j = 3 Then Call getdata(hndBrX,BR_nRlyGrp3Hnd,hndGr&)
   If hndGr <> 0 Then
    If 0 = findInList(listGr,countGr,hndGr,0) Then
     If isForwardFault(hndGr) <> 0 Then
      countGr = countGr + 1
      listGr(countGr) = hndGr
     End If
    End If
   End If
  Next
 Next
 ' Add all this tier branch(s) to the list of already traversed branches
 For i& = 1 to countBrX
  countBr = countBr + 1
  listBr(countBr) = listBrX(i)
 Next
 If countGr > countGrS Then
  If debug = 1 Then
   a$ = "Found:" & crlf
   For i& = countGrS+1 to countGr
    a$ = a$ & crlf & FullBranchName(listGr(i))
   Next
   Print a$
  End If
  traverseOneBranch = 1
  exit Function
 End If
 'Must traverse further if possible
 If thisTier = optMaxTiers Or countToDo = maxStack Then 
  traverseOneBranch = 3
  exit Function
 End If
 countToDoS = countToDo
 For i& = 1 to countBsX
  hndBs& = listBsX(i)
  If findInList(listBs,countBs,hndBs,0) = 0 Then
   hndBrX& = 0
   While 1 = GetBusEquipment(hndBs,TC_BRANCH,hndBrX)
    Call getdata(hndBrX,BR_nInService,flag&)
    If flag = 1 And findBranchInList(listBr,countBr,hndBrX,0) = 0 Then
     Call getdata(hndBrX,BR_nType,brType&)
     select case brType
      case TC_LINE
       nextKVLevelNo& = thisKVLevelNo
       nextTier& = thisTier + 1
      case TC_SWITCH
       nextKVLevelNo& = thisKVLevelNo
       nextTier& = thisTier
      case Else
       nextKVLevelNo& = thisKVLevelNo + 1
       nextTier& = thisTier + 1
     End select
     If nextTier <= optMaxTiers And nextKVLevelNo <= optMaxKVLevels Then
      countToDo = countToDo + 1
      listToDo(countToDo) = hndBrX
      listTierLevel(countToDo) = nextTier
      listKVLevel(countToDo) = nextKVLevelNo
      If countToDo = maxStack Then exit Function
     End If
    End If
   Wend
  End If
 Next
 If countToDoS = countToDo Then
  traverseOneBranch = 3
 Else
  traverseOneBranch = 2
 End If
End Function

Function listPribackPairs(hndGr&,ByRef listB() As long,ByRef countB&,ByRef listP() As long, ByRef countP&) As long
 ' Find all groups in coord. pairs that invole hndGr
 countB = 0
 countP = 0
 hnd& = 0
 While 1 = getdata(hndGr,RG_nBackupHnd,hnd)
   countB = countB + 1
   listB(countB) = hnd
 Wend
 hnd& = 0
 While 1 = getdata(hndGr,RG_nPrimaryHnd,hnd)
   countP = countP + 1
   listP(countP) = hnd
 Wend
 If debug = 1 Then
  a$ = "Selected group: " & fullBranchName(hndGr)
  a$ = a$ & crlf & "Backups:"
  For i& = 1 to countB
   a$ = a$ & crlf & fullBranchName(listB(i))
  Next
  a$ = a$ & crlf & "Backing up:"
  For i& = 1 to countP
   a$ = a$ & crlf & fullBranchName(listP(i))
  Next
  Print a
 End If
End Function

Function listSegsAndBs(hndBr&,listBr() As long, ByRef countBr&,listBs() As long, ByRef countBs&)
' Build list of all segments and buses of hndBr
 Call getdata(hndBr,BR_nType,brType&)
 If brType = TC_LINE Then
  ' Find all tap segments of this line
  Call getdata(hndBr,BR_nHandle,hndL&)
  Call findTapLineSegments(hndL,listBr,countBr&,listBs,countBs&)
 Else
  listBr(1) = hndBr
  countBr& = 1
  Call getdata(hndBr,BR_nBus1Hnd,hndB&)
  listBs(1) = hndB
  Call getdata(hndBr,BR_nBus2Hnd,hndB&)
  listBs(2) = hndB
  If 1= getdata(hndBr,BR_nBus3Hnd,hndB&) Then
   listBs(3) = hndB
   countBs = 3
  Else
   countBs = 2
  End If
 End If
End Function

Function findBranchInList(listBr() As long,ByRef countBr&, hndBr&, add&) As long
' determine if no is is list
 findBranchInList = 0
 For i& = 1 to countBr
  If isSameBranch(listBr(i), hndBr) Then
   findBranchInList = i
   exit Function
  End If
 Next
 If add <> 0 Then
  countBr = countBr+1
  listBr(countBr) = hndBr
 End If
End Function

Function findTapLineSegments(hndL&,ByRef listBr() As long,ByRef countBr&, _
           ByRef listBs() As long,ByRef countBs&) As long
' Build list of all branch segments on a line
' Taps are considered part of the line
 
 Call getdata(hndL,LN_nBus1Hnd,hndB&)
 hndBr&=0
 Do 
  If getbusequipment(hndB,TC_BRANCH,hndBr) <> 1 Then exit Do
  Call getdata(hndBr,BR_nHandle,hndX)
 Loop While hndX <> hndL

 If 1 <> GetData(hndBr, BR_nType, nT2&) Then Print ErrorString()
 ' The first segment
 listBr(1) = hndBr
 countBr = 1
 Call GetData(hndBr,BR_nBus1Hnd, hndB&)
 listBs(1) = hndB
 Call GetData(hndBr,BR_nBus2Hnd, hndB&)
 listBs(2) = hndB
 countBs& = 2
 Idx& = 0
 ' Check for tap ends
 While Idx  < countBs
  Idx = Idx + 1
  hndB = listBs(Idx)
  Call countTapEndBr(hndB,listBr,countBr,listBs,countBs)
 Wend
End Function

Function isTapBus(hndB&) As long
 Call getdata(hndB,BUS_nTapBus,tap&)
 If tap = 0 Then isTapBus = 0 Else isTapBus = 1
End Function

Function countTapEndBr(hndB&,ByRef listBr() As long, ByRef countBr&, _
           ByRef listBs() As long, ByRef countBs&) 
 'Process a tap end
 If isTapBus(hndB) = 0 Then exit Function
 hndBr& = 0
 While 1 = getBusEquipment(hndB,TC_BRANCH,hndBr)
  Call GetData(hndBr,BR_nInService,t&)
  If t <> 1 Then GoTo contW
  Call GetData(hndBr,BR_nType,t&)
  If t <> TC_LINE Then GoTo contW
  If 0 <> findBranchInList(listBr,countBr,hndBr,1) Then GoTo contW
  Call GetData(hndBr,BR_nBus2Hnd, hndB2&)
  Call findInList(listBs,countBs,hndB2,1)
  contW:
 Wend
End Function

Function isSameBranch(ByVal hndBr1&, ByVal hndBr2&) As long
 isSameBranch = 0
 Call GetData(hndBr1, BR_nType, nT1&)
 Call GetData(hndBr2, BR_nType, nT2&)
 If nT1 <> nT2 Then exit Function
 Call GetData(hndBr1, BR_nHandle, nT1&)
 Call GetData(hndBr2, BR_nHandle, nT2&)
 If nT1 <> nT2 Then exit Function
 isSameBranch = 1
End function

Function isForwardFault(ByVal hndGr&) As long
' Check if fault is in forward direction
 isForwardFault = 0
 MTA = 30
 dim MagArrayI(15) As double, AngArrayI(15) As double
 dim MagArrayV(15) As double, AngArrayV(15) As double
 If GetSCCurrent(hndGr, MagArrayI, AngArrayI, 2 ) = 0 Then GoTo HasError1
 'If debug = 1 Then Print FullBranchName(hndGr), MagArrayI(1),MagArrayI(2),MagArrayI(3) 
 If MagArrayI(2) < 10 Then exit Function
 If GetSCVoltage(grNearBus(hndGr), MagArrayV, AngArrayV, 2 ) = 0 Then GoTo HasError1
 If MagArrayV(2) < 1 Then exit Function
 diffA = AngArrayV(2) - AngArrayI(2) 
 If diffA >= MTA-90 And diffA <= MTA+90 Then bForwardFault = 1 Else bForwardFault = 0
 'If debug = 1 Then Print diffA, bForwardFault
 isForwardFault = bForwardFault
 exit Function
hasError1:
 Print FullBranchName(hndGr) & crlf & ErrorString()
 Stop
End Function

Function runCEOFault(ByVal hndGr) As long
' Run close-in fault with end open
 Dim vnFltConn(4) As Long
 Dim vdFltOpt(15) As Double
 Dim vnOutageOpt(4) As Long
 Dim vnOutageLst(30) As Long
 'fault connections
 vnFltConn(1) = 1   ' Do 3PH
 For i& = 1 to 15
  vdFltOpt(i) = 0.0
 Next
 vdFltOpt(3) = 1.0   ' close-in fault with end open
 If DoFault( hndGr, vnFltConn, vdFltOpt, vnOutageOpt, vnOutageLst, 0, 0, 1 ) = 0 Then GoTo HasError
 dim vnShowRelay(4) As long
 Call ShowFault(SF_LAST, 9, 2, 0, vnShowRelay )
 exit Function
 hasError:
 Print ErrorString()
 Stop
End Function

Function listLocalGroups(ByVal hndB&, ByRef list() As long, ByRef l& ) As long
' List all local relay groups
' a$ = "Bus " & Str(hndB) & ":" 
 hndBr& = 0
 While GetBusEquipment(hndB, TC_BRANCH, hndBr) = 1
  hndGr& = 0
  Call GetData(hndBr, BR_nRlyGrp1Hnd, hndGr& )
  If hndGr > 0 Then Call findInList(list, l, hndGr, 1)
'  If hndGr > 0 Then a$ = a$ & Str(hndGr) & " "
 Wend
' Print a$
End Function

Function grNearBus(hndG&) As long
 Call getdata(hndG, RG_nBranchHnd, hndBr&)
 Call getdata(hndBr, BR_nBus1Hnd, hndB&)
 grNearBus = hndB
End Function

Function findInList(ByRef list() As long, ByRef l&, no&, add&) As long
' determine if no is is list
 findInList = 0
 For i& = 1 to l
  t& = list(i)
  If list(i) = no Then
   findInList = i
   exit Function
  End If
 Next
 If add <> 0 Then
  l = l+1
  list(l) = no
 End If
End Function

Function outRATCoordPair(hndGrPri&,hndGrBack&) As String
 a$ = ""
 For i = 1 to 2
  If i = 1 Then 
   hndG& = hndGrPri 
  Else 
   hndG& = hndGrBack
   a$ = a$ & crlf 
  End If
  Call getdata(hndG, RG_nBranchHnd, hndBr&)
  For j = 1 to 2
   If j = 1 Then Call getdata(hndBr, BR_nBus1Hnd, hndB&) _
    Else Call getdata(hndBr, BR_nBus2Hnd, hndB&)
    Call getdata(hndB, BUS_nNumber, bsNO&)
    Call getdata(hndB, BUS_sName, bsName$)
    Call getdata(hndB, BUS_dKVNominal, bsKV#)
    a$ = a$ & Str(bsNO) & ";'" & bsName & "';" & Str(bsKV) & ";"
  Next
  Call getdata(hndBr, BR_nType, brType&)
  select case brType
   case TC_LINE
    typeCode = "1"
    idCode = LN_sID
   case TC_SWITCH
    typeCode = "7"
    idCode = SW_sID
   case TC_XFMR
    typeCode = "2"
    idCode = XR_sID
   case TC_XFMR3
    typeCode = "10"
    idCode = X3_sID
   case TC_PS
    typeCode = "3"
    idCode = PS_sID
   case Else
    Print "Internal logic error on " & fullBranchName(hndBr)
    Stop
  End select
  Call getdata(hndBr, BR_nHandle, hndItem&)
  Call getdata( hndItem,idCode,sID$)
  a$ = a$ & "'" & sID & "';" & typeCode
 Next 
 outRATCoordPair = a
End Function
