' ASPEN PowerScript Sample Program
'
' TIELINES.BAS
'
' Produce tie line report
'
' PowerScript functions called:
'   GetData()
'   FullBusName()
'   GetEquipment()
'
Dim nArea As Long
Dim fileName As String

' ===================== Dialog box spec (generated by Dialog Editor) ==========
'
'
Begin Dialog Dialog_1 49,60, 187, 73, "Ties report by zones"
  OptionGroup .GROUP_1
    OptionButton 16,8,25,12, "All"
    OptionButton 60,8,48,12, "From zone:"
  GroupBox 8,0,168,24, "Scope"
  GroupBox 8,28,168,24, "Output to"
  TextBox 116,8,24,12, .EditBox_2
  TextBox 12,36,157,12, .EditBox_1
  OKButton 44,56,52,12
  CancelButton 108,56,48,12
End Dialog

'
' ===================== End of Dialog box spec ================================

Sub main()
   nArea = 1
   ' Open output file
   rCode = OpenOutput() 
   If rCode = 0 Then Exit Sub
   OutStr$ = "Area1,Area2,Branch Type,Branch ID"
   Print #1, OutStr$
   Print #1, ""
   
   Counts    = 0
   ' Loop thru all lines
   cType$    = "L"
   DevHandle = 0
   While GetEquipment( TC_LINE, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , LN_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , LN_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     If GetData( Bus1Handle, BUS_nZone, Area1 ) = 0 Then GoTo HasError
     If GetData( Bus2Handle, BUS_nZone, Area2 ) = 0 Then GoTo HasError
     Call GetData( DevHandle, LN_sID, BrID$ )
     If Area1 <> Area2 And (nArea = -1 Or nArea = Area1 Or nArea = Area2) Then
       Call PrintBranch( Bus1Handle, Bus2Handle, BrID, cType )
       Counts = Counts + 1
     End If
   Wend

   ' Loop thru all series cap
   cType$    = "P"
   DevHandle = 0
   While GetEquipment( TC_SCAP, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , SC_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , SC_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     If GetData( Bus1Handle, BUS_nZone, Area1 ) = 0 Then GoTo HasError
     If GetData( Bus2Handle, BUS_nZone, Area2 ) = 0 Then GoTo HasError
     Call GetData( DevHandle, SC_sID, BrID$ )
     If Area1 <> Area2 And (nArea = -1 Or nArea = Area1 Or nArea = Area2) Then
       Call PrintBranch( Bus1Handle, Bus2Handle, BrID, cType )
       Counts = Counts + 1
     End If
   Wend

   ' Loop thru all phase shifter
   cType$    = "S"
   DevHandle = 0
   While GetEquipment( TC_PS, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , PS_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , PS_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     If GetData( Bus1Handle, BUS_nZone, Area1 ) = 0 Then GoTo HasError
     If GetData( Bus2Handle, BUS_nZone, Area2 ) = 0 Then GoTo HasError
     Call GetData( DevHandle, PS_sID, BrID$ )
     If Area1 <> Area2 And (nArea = -1 Or nArea = Area1 Or nArea = Area2) Then
       Call PrintBranch( Bus1Handle, Bus2Handle, BrID, cType )
       Counts = Counts + 1
     End If
   Wend

   ' Loop thru all switches
   cType$    = "W"
   DevHandle = 0
   While GetEquipment( TC_SWITCH, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , SW_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , SW_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     If GetData( Bus1Handle, BUS_nZone, Area1 ) = 0 Then GoTo HasError
     If GetData( Bus2Handle, BUS_nZone, Area2 ) = 0 Then GoTo HasError
     BrID$ = ""
     If Area1 <> Area2 And (nArea = -1 Or nArea = Area1 Or nArea = Area2) Then
       Call PrintBranch( Bus1Handle, Bus2Handle, BrID, cType )
       Counts = Counts + 1
     End If
   Wend

   ' Loop thru all transformers
   cType$    = "T"
   DevHandle = 0
   While GetEquipment( TC_XFMR, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , XR_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , XR_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     If GetData( Bus1Handle, BUS_nZone, Area1 ) = 0 Then GoTo HasError
     If GetData( Bus2Handle, BUS_nZone, Area2 ) = 0 Then GoTo HasError
     Call GetData( DevHandle, XR_sID, BrID$ )
     If Area1 <> Area2 And (nArea = -1 Or nArea = Area1 Or nArea = Area2) Then
       Call PrintBranch( Bus1Handle, Bus2Handle, BrID, cType )
       Counts = Counts + 1
     End If
   Wend
   cType$    = "X"
   DevHandle = 0
   While GetEquipment( TC_XFMR3, DevHandle ) > 0
     ' Get Line ID and end bus names
     If GetData( DevHandle , X3_nBus1Hnd, Bus1Handle ) = 0 Then GoTo HasError
     If GetData( DevHandle , X3_nBus2Hnd, Bus2Handle ) = 0 Then GoTo HasError
     If GetData( Bus1Handle, BUS_nZone, Area1 ) = 0 Then GoTo HasError
     If GetData( Bus2Handle, BUS_nZone, Area2 ) = 0 Then GoTo HasError
     Call GetData( DevHandle, X3_sID, BrID$ )
     If Area1 <> Area2 And (nArea = -1 Or nArea = Area1 Or nArea = Area2) Then
       Call PrintBranch( Bus1Handle, Bus2Handle, BrID, cType )
       Counts = Counts + 1
     End If
   Wend

   Print Counts, " Ties found. Output has been written to: " & fileName
   Exit Sub
HasError:
   Print "Error: ", ErrorString( )
End Sub

   
Function PrintBranch( ByVal Bus1Handle&, ByVal Bus2Handle&, ByVal BrID$, ByVal cType$ ) As integer
  Call GetData( Bus1Handle, BUS_nZone, Area1& )
  Call GetData( Bus2Handle, BUS_nZone, Area2& )
  sBus1$ = FullBusName(Bus1Handle)
  sBus2$ = FullBusName(Bus2Handle)
  OutStr$ = Format( Area1, "      ###" ) & "," & _
            Format( Area2, "      ###" ) & "," & cType & "," & _
            sBus1 & " - " & sBus2 & " " & BrID
  Print #1, OutStr$
End Function

' ===================== OpenOutput() ==========================================
' Purpose:
'   Open file for output
Function OpenOutput() As Long
   Dim dlg As Dialog_1
   Dlg.EditBox_1 = "c:\ties.csv"         ' Default name
   ' Dialog returns -1 for OK, 0 for Cancel, button # for PushButtons
   dlg.editbox_2 = nArea
   button = Dialog( Dlg )
   If button = 0 Then 
     OpenOutput = 0
     Exit Function
   End If
   If dlg.group_1 = 0 Then nArea = -1 Else nArea = dlg.EditBox_2
   fileName = Dlg.EditBox_1
   Open fileName For Output As #1
   OpenOutput = 2
End Function

