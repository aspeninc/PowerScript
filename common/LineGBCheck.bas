' ASPEN PowerScrip sample program
'
' LineGBCheck.BAS
'
' Update all lines with very small shunt adimittance value
'
' Version: 1.0
'

Sub main()
   dim MinValue#
   
   MinV$ = InputBox( "Set very small line G + jB to this minimum threshold value:", "Fix very small line G + jB", "0.00000001" ) 
   MinValue# = Val( MinV$ )
   If MinValue = 0 Then Stop
   UpdateG# = MinValue#
   UpdateB# = MinValue#
   
   ' Specify output file
'   sOutFile = "C:\Source\bas\dev\LineCheck\LineCheck.csv"
   sOutFile = GetOLRFileName() & ".csv"

   Open sOutFile For output As 1
   Print #1, "Line,Param,Old value,New value"   
  
   ' Loop thru all lines
   
   
   dim nCode(8) As long
   nCode(1) = LN_dG1
   nCode(2) = LN_dG10   
   nCode(3) = LN_dB1   
   nCode(4) = LN_dB10
   nCode(5) = LN_dG2
   nCode(6) = LN_dG20   
   nCode(7) = LN_dB2   
   nCode(8) = LN_dB20
   dim sCode(8) As String
   sCode(1) = "G1"
   sCode(2) = "G10"   
   sCode(3) = "B1"   
   sCode(4) = "B10"
   sCode(5) = "G2"
   sCode(6) = "G20"   
   sCode(7) = "B2"   
   sCode(8) = "B20"
   
   LineCount = 0
   DevHandle& = 0
   While GetEquipment( TC_LINE, DevHandle ) > 0
     LineCount = LineCount + 1
   Wend
   Counts&   = 0
   DevHandle& = 0
   jj = 0
   While GetEquipment( TC_LINE, DevHandle ) > 0
     jj = jj + 1
     Button =  ProgressDialog( 1, "Checking line G + jB", "Line " + Str(jj) +" of " + Str(LineCount), 100 * jj / LineCount )
     If Button = 2 Then 
       Print "Cancel button pressed"
       GoTo Done
     End If
    
     If GetData( DevHandle,  LN_nInService, nFlag ) = 0 Then GoTo HasError
     If nFlag <> 0 Then      
       nChanged& = 0
       sLine$ = PrintObj1LPF( DevHandle )
       sText$ = ""
       For ii = 1 to 8
         nThisCode& = nCode(ii)
         If GetData( DevHandle, nThisCode, dVal# ) = 0 Then GoTo HasError
         If dVal <> 0 And (MinValue - Abs(dVal)) > MinValue/1000.0 Then
           If SetData( DevHandle, nThisCode, MinValue ) = 0 Then GoTo HasError
           sText$ = sLine$ & "," & sCode(ii) & "," & Str(dVal) & "," & Str(MinValue)
           Print #1, sText$
           nChanged& = 1
         End If
       Next
       If nChanged = 1 Then
         If PostData( DevHandle ) = 0 Then GoTo HasError   
         Counts = Counts + 1
       End If
     End If
   Wend
   Done:
   Call ProgressDialog( 0, "", "", 0 )   
   If Counts > 0 Then
     Print "Updated G+jB in " & Str(Counts) & " lines. Details are in" & Chr(13) & Chr(10) & sOutFile
   Else
     Print "Found no line with G+jB lower than threshold"
   End If
   
   Close 1
   Exit Sub
HasError:
   Close 1
   Call ProgressDialog( 0, "", "", 0 )   
   Print "Error: ", ErrorString( )
End Sub

