' ASPEN PowerScript sample program
'
' XFMRZ.BAS
'
'ASPEN Transformer Impedance Calculator
'
' The transformer impedance is typically given in percent On transformer test sheets.
' One can break down the impedance into its resistance And reactance components using
' the transformer load-loss data, which are also give On test sheets.
' The formulas are As follows:
'
' R = S / ( 10 * T)
' X = sqrt( Z*Z - R*R)
' 
' where
' S is the load loss at rated load in watts
' T is the transformer rating in kVA
' Z is the impedance in percent
' R is the resistance in percent
' X is the reactance in percent
'
' Reference: IEEE, Inc., IEEE Standard Test Code For Dry-Type Distribution And Power 
' Transformers, ANSI/IEEE C57.12.91-1979, Section 15.4.4
'
' Additional Background Information
' Many engineers model their transformer With a resistance of zero And a reactance equal
' to the impedance given by the test sheet. This method is Not strictly correct; the
' resistance of power transformers can be As high As 5% to 10% of the reactance.  
' The assumption of zero resistance can lead to unrealistic X/R ratios.  We strongly
' recommend that you enter the correct transformer resistance And reactance
' If you plan to use the X/R ratios For breaker rating studies.  The short circuit
' currents are usually Not affected significantly by the zero resistance.
'
'
' Dialog data
Begin Dialog DIALOG_1 16,43, 169, 79, "ASPEN Transformer Impedance Calculator"
  Text 24,8,69,12, "Transformer rating="
  TextBox 92,8,37,12, .EditBox_1
  Text 132,8,21,12, "kVA"
  Text 48,24,44,12, "Impedance="
  TextBox 92,24,37,12, .EditBox_2
  Text 132,24,37,12, "percent"
  Text 12,40,84,12, "Load loss at rated load="
  TextBox 92,40,37,12, .EditBox_3
  Text 132,40,37,12, "watts"
  OKButton 44,60,53,12
  CancelButton 104,60,37,12
End Dialog
' End dialog data

Sub main()

  dim dlg As DIALOG_1
  
   ' Get picked device
  If GetEquipment( TC_PICKED, PickedHnd& ) > 0 Then 
   TypeCode = EquipmentType( PickedHnd )
   If TypeCode = TC_XFMR Then
    Call GetData( PickedHnd, XR_dX, dX# )
    Call GetData( PickedHnd, XR_dMVA, dMVA# )
    dlg.EditBox_1 = Format( dMVA * 1000, "#0.0" )
    dlg.EditBox_2 = Format(dX * 100, "#0.0" )
   End If
  End If

  Do
   button = Dialog( Dlg )
   If button = 0 Then exit Sub
   bOK = TRUE
   T# = dlg.EditBox_1
   If T < 0.001 Then
    Print "Rating must be positive"
    bOK = FALSE
   End If
   Z# = dlg.EditBox_2
   If Z < 0.001 Then
    Print "Impedance must be positive"
    bOK = FALSE
   End If
   S# = dlg.EditBox_3
   If S < 0.001 Then
    Print "Loss must be positive"
    bOK = FALSE
   End If
  Loop While Not bOK
  
  R# = S / ( 10 * T)
  X# = Sqr( Z*Z - R*R)
  
  Print "Impedance in percent: R=", Format(R, "0.0"); " X=", Format(X, "0.0")
  
  Exit Sub
 HasError:
  Print "Error: ", ErrorString( )
End Sub