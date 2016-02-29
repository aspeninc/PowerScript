' ORPHANBUS.BAS
' Sample ASPEN PowerScript Program
'
' Print list of buses that have no connection to any other buses
'

Sub main
 dim BusList() As long
 dim MapList() As long

 NObus&  = 0
 BusHnd& = 0
 While 1 = GetEquipment(TC_BUS, BusHnd)
   NObus = NObus + 1
 Wend
' Print NObus
 ReDim BusList(NObus+1)
 NObus&  = 0
 BusHnd& = 0
 While 1 = GetEquipment(TC_BUS, BusHnd)
   BusList(NObus) = BusHnd
   NObus = NObus + 1
 Wend

' For ii = 0 to 10
'  sStr$ = BusList(ii)
'  PrintTTY(sStr)
' Next
 
 ReDim MapList(NObus+1)
 For ii = 0 to NObus-1
  MapList(ii) = 1
 Next 

 DevHnd& = 0 
 While 1 = GetEquipment(TC_LINE, DevHnd)
   Call GetData(DevHnd, LN_nBus1Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
   Call GetData(DevHnd, LN_nBus2Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
 Wend

 DevHnd& = 0 
 While 1 = GetEquipment(TC_SWITCH, DevHnd)
   Call GetData(DevHnd, SW_nBus1Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
   Call GetData(DevHnd, SW_nBus2Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
 Wend

 DevHnd& = 0 
 While 1 = GetEquipment(TC_PS, DevHnd)
   Call GetData(DevHnd, PS_nBus1Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
   Call GetData(DevHnd, PS_nBus2Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
 Wend

 DevHnd& = 0 
 While 1 = GetEquipment(TC_XFMR, DevHnd)
   Call GetData(DevHnd, XR_nBus1Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
   Call GetData(DevHnd, XR_nBus2Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
 Wend

 DevHnd& = 0 
 While 1 = GetEquipment(TC_XFMR3, DevHnd)
   Call GetData(DevHnd, X3_nBus1Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
   Call GetData(DevHnd, X3_nBus2Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
   Call GetData(DevHnd, X3_nBus3Hnd, BusHnd)
   nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
   If nIdx > -1 Then MapList(nIdx) = 0 Else Print "Stop"
 Wend

 nCount& = 0
 PrintTTY(" ")
 PrintTTY("File name: " & GetOLRFileName())
 For ii = 0 to NObus-1
  If MapList(ii) = 1 Then 
   sS$ = fullBusName(BusList(ii))
   If nCount = 0 Then PrintTTY("Following bus(es) do not have connection to any other bus")
   PrintTTY(sS)
   nCount = nCount + 1
  End If
 Next 
 If nCount = 0 Then
   Print "Found no orphan bus in this network"
   PrintTTY("No orphan bus found")
 Else
   Print "Found ", nCount, " no orphan bus(s) in this network. See TTY window for full list"
 End If

End Sub

Function binarySearch( Array() As long, nKey As long, nMin As long, nMax As long ) As long
 While nMax >= nMin
   nMid& = (nMax+nMin) / 2
   If Array(nMid) = nKey Then
     binarySearch = nMid
     exit Function
   Else 
     If Array(nMid) < nKey Then nMin = nMid + 1 Else nMax = nMid -1
   End If
 Wend
 binarySearch = -1
End Function
