' ASPEN Sample Script
' SolitaryGen.BAS
'
' Print report of solitary generators 
' (ones with no connections to the rest of the network)
'
' Version 1.0
'

Sub Main
  hndBus& = 0
  hndGen& = 0
  nCountGen = 0
  nCountGenSolitary = 0
  While GetEquipment( TC_BUS, hndBus ) = 1
    nRet = GetBusEquipment(hndBus, TC_GEN, hndGen)
    If nRet = 1 Then
      nCountGen = nCountGen + 1
      ' This bus has a generator. Count number of connected branches
      hndBranch& = 0
      hasBranch = False
      If GetBusEquipment(hndBus, TC_BRANCH, hndBranch) = 1 Then hasBranch = True
      aLine = FullBusName(hndBus)
      If Not hasBranch Then
'        Print "Solitary gen: " & aLine
        If nCountGenSolitary = 0 Then 
          printTTY(" ")
          printTTY(" ")
          printTTY("List of solitary generators in this file:")
        End If
        printTTY("  " & aLine)
        nCountGenSolitary = nCountGenSolitary + 1
      Else
'        Print "Gen: " & aLine
      End If
    End If
  Wend
  aLine = "There are " & nCountGen & " generators in this file." &Chr(13)&Chr(10)
  If nCountGenSolitary > 0 Then
    If nCountGenSolitary = 1 Then HasHave = " has " Else HasHave = " have "
    aLine = aLine & nCountGenSolitary & HasHave & "no connection to the rest of the network." &Chr(13) & Chr(10) & "Details are in TTY."
  Else
    aLine = aLine & "All are connected to the network."
  End If
  Print aLine
End Sub
