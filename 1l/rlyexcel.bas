' ASPEN PowerScript Sample Program
'
' RLYEXCEL.BAS
'
' Version 0.1
'
' Print relay settings for copying/paste to Excel
'
' Version: 1.0
' Category: OneLiner
'
Const MXDEVCOUNT = 10  ' Max number of devices in a group
dim arPickup(MXDEVCOUNT) As String
dim arTD(MXDEVCOUNT) As String
dim arCurve(MXDEVCOUNT) As String
dim DeviceList(MXDEVCOUNT) As String
dim RelayCount As long

Sub main()
  
   ' Get picked device
   If GetEquipment( TC_PICKED, PickedHnd ) = 0 or _
      TC_RLYGROUP <> EquipmentType( PickedHnd ) Then
     Print "You must select a relay group to run this script program"
     Stop
   End If

   ' Loop through all OC relays to make selection list
   RelayCount = 0
   RelayHnd   = 0
   sRlyList$ = ""
   While GetRelay( PickedHnd, RelayHnd ) > 0
     TypeCode = EquipmentType( RelayHnd )
     Select Case TypeCode
       Case TC_RLYOCG
         RlyDevice$ = "OCG= "
         param_nInservice& = OG_nInService
         param_sID& = OG_sID
         param_dPickup& = OG_dTap
         param_dTD& = OG_dTDial
         param_sCurve& = OG_sType
       Case TC_RLYOCP
         RlyDevice$ = "OCP= "
         param_nInservice& = OP_nInService
         param_sID& = OP_sID
         param_dPickup& = OP_dTap
         param_dTD& = OP_dTDial
         param_sCurve& = OP_sType
       Case Else
         GoTo Cont
     End Select
     Call GetData( RelayHnd, param_nInservice, nInService& )
     If nInService <> 1 Then GoTo Cont
     RelayCount = RelayCount + 1
     Call GetData( RelayHnd, param_sID, sID$ )
     Call GetData( RelayHnd, param_dPickup, dPickup# )
     Call GetData( RelayHnd, param_dTD, dTD# )
     Call GetData( RelayHnd, param_sCurve, sCurve$ )
     arPickup(RelayCount) = Format( dPickup, "####0.000")
     arTD(RelayCount) = Format( dTD, "####0.0")
     arCurve(RelayCount) = LookupSELCurve(sCurve)
     DeviceList(RelayCount) = Str(RelayCount) & ". " & RlyDevice & sID & _
              " Curve= " & sCurve & " TD= " & arTD(RelayCount) & " Pickup= " & arPickup(RelayCount) 
   Cont:
     if MXDEVCOUNT <= RelayCount then
       print "RelayCount had reached max limit. Edit MXDEVCOUNT in the script to show more"
     end if
   Wend  'Each relay in group
   
   If RelayCount = 0 Then
     Print "Found no active device in selected relay group"
     Stop
   End If
   
   Call PrintDeviceSettings()
   
   Stop
   
   mySelection = SelDevice( )
   If mySelection > 0 Then 
     Call ExportToExcel()
   End If
   Exit Sub
   ' Error handling
   HasError:
   Print "Error: ", ErrorString( )
End Sub  ' End of Sub Main()

Function LookupSELCurve( sASPENCurve$ ) As String
 LookupSELCurve = ""
 If InStr(1, sASPENCurve, "U1" ) > 0 Then
   LookupSELCurve = "U1"
   exit Function
 End If
 If InStr(1, sASPENCurve, "U2" ) > 0 Then
   LookupSELCurve = "U2"
   exit Function
 End If
 If InStr(1, sASPENCurve, "U3" ) > 0 Then
   LookupSELCurve = "U3"
   exit Function
 End If
 If InStr(1, sASPENCurve, "U4" ) > 0 Then
   LookupSELCurve = "U4"
   exit Function
 End If
End Function

'Try these different styles or-ed together as the last parameter of Textbox
' to define the text box style.
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
Const WS_VSCROLL          = &h00200000&

Function PrintDeviceSettings( ) As long
Begin Dialog PRINTSETTINGS 52,10,236,203, "Export Settings To Excel"
  Text 7,5,73,8,"Devices in this group"
  Text 11,106,46,8,"Device No"
  Text 21,124,43,8,"Pickup ="
  Text 22,132,32,8,"Curve ="
  Text 22,139,27,8,"TD="
  TextBox 54,105,25,11,.Edit1
  TextBox 7,19,221,80,.NumberedList, ES_AUTOVSCROLL Or ES_AUTOHSCROLL Or ES_MULTILINE Or ES_READONLY
  TextBox 55,124,161,48,.Edit2, ES_MULTILINE
  PushButton 82,104,146,13,"Print this device settings to edit box below", .Button3
  PushButton 100,181,57,13,"Done", .Done
End Dialog


  Dim dlg As PRINTSETTINGS
  dlg.NumberedList = ""
  For ii = 1 to RelayCount
   If ii > 1 Then dlg.NumberedList = dlg.NumberedList & Chr(13) & Chr(10)
   dlg.NumberedList = dlg.NumberedList & DeviceList(ii) 
  Next
  SelNo = 0
  Do 
   If SelNo > 0 Then
    aStr$ = arPickup(SelNo) & Chr(13) & Chr(10) & _
            arCurve(SelNo) & Chr(13) & Chr(10) & _
            arTD(SelNo)  & Chr(13) & Chr(10) & _
            Chr(13) & Chr(10) & _
            "Select settings above and copy/paste to Excel"
    dlg.Edit2 = aStr
   End If
   button = Dialog( dlg )
   If button <> 1 Then Stop
   SelNo = Val(dlg.Edit1)
   If SelNo < 1 Or SelNo > RelayCount Then
     Print "Selection is out of range"
   End If
  Loop While button = 1
End Function



Function ExportToExcel() As long
  Const PCCSetting_Col = 10   ' Column in DG PCC sheet
  Const PCCSetting_PhaseTOC_Pickup = 67  ' Row in DG PCC sheet
  Const PCCSetting_PhaseTOC_Curve  = 68
  Const PCCSetting_PhaseTOC_TD     = 69
  Const PCCSetting_GroundTOC_Pickup = 85
  Const PCCSetting_GroundTOC_Curve  = 86
  Const PCCSetting_GroundTOC_TD     = 87
  
  sOlrPathName$ = GetOlrFileName()
  sOlrPath$ = ExtractFilePath(sOlrPathName)
  ExcelFile$ = FileOpenDialog( sOlrPath, "Excel File| *.xls;*.xlsx||", 1 )
  
  If Len(ExcelFile) = 0 Then 
    Stop
  End If
'  Print ExcelFile$
  
  ' Get Pointer to Excel application
  On Error GoTo excelErr  
  Set xlApp = CreateObject("excel.application")
  Set wkbook = xlApp.Workbooks.Open( ExcelFile, True, True) 
  On Error GoTo dataSheetErr
  Set dataSheet = xlApp.Worksheets(4)
  ' Read file header row
  aHeader$ = dataSheet.Cells(1,2).Value
  Print aHeader
  wkbook.Close
excelErr:
dataSheetErr:
  ' Free memory  
  Set dataSheet = Nothing
  Set wkbook    = Nothing
  Set xlApp     = Nothing
  If Err.Number > 0 Then Print "Excecution error: " & Err.Description
  Stop    
    
End Function

Function SelDevice( ) As long
Begin Dialog SELECTLIST 30,60,236,149, "Export Settings To Excel"
  Text 8,4,73,8,"Devices in this group"
  ListBox 6,17,225,104,DeviceList, .DevNo
  OKButton 77,128,65,13
  CancelButton 148,128,40,13
End Dialog

  SelDevice = 0
  Dim dlg As SELECTLIST
  button = Dialog( dlg )
  If button <> 0 Then ' Canceled
    SelDevice  = 1 + dlg.DevNo
  End If
End Function

Function ExtractFilePath( sFullPathName$ ) As String
 ExtractFilePath = ""
 nLen& = Len(sFullPathName)
 Do
   If nLen <= 1 Then GoTo breakWhile
   nLen = nLen - 1
   aChar$ = Mid(sFullPathName, nLen, 1)
 Loop While aChar <> "\"
 breakWhile:
 ExtractFilePath$ = Left(sFullPathName,nLen)
End Function

