' ASPEN PowerScrip program
'
' FixEquivBr.BAS
'
' Update impedance of boundary equivalent phase shifters to 
' fix a bug in command Network | Boundary Equivalent command
' in OneLiner version 14.5 and earlier.
'
'   Z1new = Z1old * mult
'   Z2new = Z2old * mult
' where 
'   mult = cos(ShiftAngle)
'
'
' MAKE SURE YOU RUN THIS SCRIPT ONCE, AND ONLY ONCE, ON EACH OLR FILE.
'
' Version 1.0
' Category: OneLiner
'

Sub main()
'=============Dialog Spec=============
   Begin Dialog Dialog_1 45,59,224,116, "Enter ID prefix"
     Text 6,2,195,8,"Update impedance of boundary equivalent phase shifters to"
     Text 6,12,212,8,"fix a bug in command Network | Boundary Equivalent command"
     Text 6,22,194,8,"in OneLiner version 14.5 and earlier."
     Text 6,35,119,8,"  Z1new = Z1old * mult"
     Text 6,42,73,8,"  Z2new = Z2old * mult"
     Text 6,50,142,8,"  where mult = cos(ShiftAngle in radian)"
     Text 6,64,150,8,"Boundary equivalent equipment ID prefix = "
     Text 6,79,209,8,"MAKE SURE YOU RUN THIS SCRIPT ONCE, AND ONLY ONCE, "
     Text 6,87,268,8,"ON EACH OLR FILE."
     TextBox 150,62,41,11,.Edit1
     OKButton 133,98,40,13
     CancelButton 176,98,40,13
   End Dialog
'=====================================
   Dim dlg As Dialog_1
   Dlg.Edit1 = "N"         ' Default prefix
   Button = Dialog( dlg )
   If Button = 0 Then Exit Sub	' Canceled 
   sPrefix$ = Dlg.Edit1
   If sPrefix = "" Then Stop
   nLen = Len( sPrefix )
   nCount = 0
   nCountOK = 0
   PsHnd = 0
   While GetEquipment( TC_PS, PsHnd ) > 0
      If GetData( PsHnd, PS_nInservice, nFlag& ) = 0 Then GoTo HasError
      If nFlag = 0 Then GoTo cont1
      If GetData( PsHnd, PS_dAngle, dAngle# ) = 0 Then GoTo HasError
      If Abs(dAngle) < 1E-5 Then GoTo cont1   ' Shift angle is zero. No need to do anything
            
      If GetData( PsHnd, PS_sID, PsID ) = 0 Then GoTo HasError
      If StrComp( Left(PsID, nLen), sPrefix ) <> 0 Then GoTo cont1  
      
      If GetData( PsHnd, PS_nBus1Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus1ID = FullBusName( BusHnd )
      If GetData( PsHnd, PS_nBus2Hnd, BusHnd ) = 0 Then GoTo HasError
      Bus2ID = FullBusName( BusHnd )
      nCount = nCount + 1  
      sMsg = Bus1ID & "  - " & Bus2ID & "   ID = " & PsID       
      printTTY( sMsg ) 
      
      
      If GetData( PsHnd, PS_dR,  dR1# ) = 0 Then GoTo HasError
      If GetData( PsHnd, PS_dX,  dX1# ) = 0 Then GoTo HasError
      If GetData( PsHnd, PS_dR2, dR2# ) = 0 Then GoTo HasError
      If GetData( PsHnd, PS_dX2, dX2# ) = 0 Then GoTo HasError    
      sMsg = "   Original: R1 = " & Format(dR1, "####0.00000") & "  " & "X1 = " & Format(dX1, "####0.00000") & "  " & _
                       "R2 = " & Format(dR2, "####0.00000") & "  " & "X2 = " & Format(dX2, "####0.00000")
      printTTY( sMsg )              
      dMul# = Cos(dAngle*3.14159/180.0)
      dR1_new = dR1 * dMul
      dX1_new = dX1 * dMul
      dR2_new = dR2 * dMul
      dX2_new = dX2 * dMul
      If SetData( PsHnd, PS_dR,  dR1_new ) = 0 Then GoTo HasError
      If SetData( PsHnd, PS_dX,  dX1_new ) = 0 Then GoTo HasError 
      If SetData( PsHnd, PS_dR2, dR2_new ) = 0 Then GoTo HasError
      If SetData( PsHnd, PS_dX2, dX2_new ) = 0 Then GoTo HasError 
      sMsg = "   Updated:  R1 = " & Format(dR1_new, "####0.00000") & "  " & "X1 = " & Format(dX1_new, "####0.00000") & "  " & _
      	               "R2 = " & Format(dR2_new, "####0.00000") & "  " & "X2 = " & Format(dX2_new, "####0.00000")
      printTTY( sMsg ) 
      If PostData( PsHnd ) = 1 Then nCountOK = nCountOK + 1 
      printTTY( " " ) 
      cont1:
   Wend  
   If nCount = 0 Then 
     Print "This OLR file has no equivalent phase shifter"
   Else
     Print "Found " & nCount & " equivalent phase shifter" &Chr(13)&Chr(10)& _
           "Updated " & nCountOK & " successfully." &Chr(13)&Chr(10)& _
           "Details are in TTY window"
   End If   
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
   Print "Don't save this OLR file"
End Sub  



 