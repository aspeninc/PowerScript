' ASPEN PowerScript Sample Program
'
' UPDATEZ1D.BAS
'
' Update Distance relay setting in OLR file
'
' Version: 1.0
' Category: OneLiner
' 
'

Const ParamIdx_Z1D = "10400"

Sub main()
   
   sAns$ = InputBox( "This PowerScript program will update all SEL relay settings in this OLR file" &Chr(13)&Chr(10)& _
         "from Z1PD or Z1GD=OFF to Z1PD or Z1GD=0.0" &Chr(13)&Chr(10)& _
         "Enter R (Run check and update) or"&Chr(13)&Chr(10)& _
         "C (Check only) and click OK", "Confirmation", "C" )
   If sAns = "" Then Stop
   
   nCountOK = 0
   nCount = 0
   hndGroup& = 0
   While 1 = GetEquipment(TC_RLYGROUP, hndGroup)
     Call GetData(hndGroup, RG_nBranchHnd, hndBranh&)
     Call GetData(hndBranch, BR_nBus1Hnd, hndBus1&)
     Call GetData(hndBranch, BR_nBus2Hnd, hndBus2&)


     hndRelay& = 0
     While 1 = GetRelay(hndGroup, hndRelay)
       TypeCode = EquipmentType( hndRelay )
       If TypeCode = TC_RLYDSP Then
         nParamDSType& = DP_sDSType
         nParamID& = DP_sParam
         sParam$ = "Z1PD"
       Elseif TypeCode = TC_RLYDSG Then
         nParamDSType& = DG_sDSType
         nParamID& = DG_sParam
         sParam$ = "Z1GD"
       Else
         GoTo Cont1
       End If
       Call GetData( hndRelay, nParamDSType, sDSType$)
       If 1 = InStr(1, sDSType, "SEL") And 0 < InStr(1, sDSType, "__") Then
         sOldParam = sParam
         If 0 <> GetData(hndRelay, nParamID, sOldParam) Then
           If sOldParam = "OFF" Then
             nCount = nCount + 1
             If sAns = "R" Then
               sNewParam$ = ParamIdx_Z1D & Chr(9) & "0.0"
               sResult$ = "Failed"
               If 1 = SetData(hndRelay, nParamID, sNewParam) Then
                 If 1 = PostData(hndRelay) Then 
                   sResult$ = "OK"
                   nCountOK = nCountOK + 1
                 End If
               End If
               printTTY(FullRelayName(hndRelay) & ": Update " & sDSType & "." & sParam & "=" & sOldParam & " " & sResult)
              else
               printTTY(FullRelayName(hndRelay) & ": " & sDSType & "." & sParam & "=" & sOldParam)
              end if
           End If
         End If
       End If
cont1:
     Wend
   Wend
   If nCount = 0 Then 
     Print "This OLR file has no SEL relay with setting Z1PD or Z1GD=OFF"
   Else
     If sAns = "R" Then
       Print "Found " & nCount & " SEL relay with setting Z1PD or Z1GD=OFF" &Chr(13)&Chr(10)& _
           "Updated " & nCountOK & " successfully." &Chr(13)&Chr(10)& _
           "Details are in TTY window"
     else
       Print "Found " & nCount & " SEL relay with setting Z1PD or Z1GD=OFF" &Chr(13)&Chr(10)& _
           "Details are in TTY window"
     end if
   End If   
   
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()
