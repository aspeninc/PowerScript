' ASPEN Sample Script 
' ChangeBkrData.BAS
'
' Make change to breaker data using Change file
'
' Version 1.0
' Category: OneLiner
'
' Usage:
' 1- Run OneLiner File | Export | Breaker data to save breaker data in  CHF file
' 2- Execute Tool | Scripting | Run script. Open ChangeBkrData.BAS
' 3- Enter full path name of the CHF file in #1 to generate new change file
'    with breaker data modifications.
' 4- Execute File | Read change file to apply the newly generated change file
'    to the OLR file.
'
Sub main
  
  DataFile$ = InputBox("Enter change file full path name")
' DataFile$ = "h:\data\00Support&Dev\00Todo\00Dev_PSChangeFile\421DSG.CHF"
  
 If Len(DataFile) = 0 Then 
  Print "Bye"
  Stop
 End If
  
 Open DataFile For Input As 1
 nLineCount& = 0  
 nReadingBreakerData = 0
 Do While Not EOF(1)
  Line Input #1, aLine$ ' Read a line of data.
  aRecord$ = aLine
  If InStr( aLine$, " /" ) > 0 Then
   Do While Not EOF(1)
    Line Input #1, aLine$
    aRecord$ = aRecord$ & Chr(13) & Chr(10) & aLine
    If InStr( aLine$, " /" ) <= 0 Then exit Do
   Loop
  End If
  If InStr( aRecord$, "[ADD BREAKER" ) = 1 Then 
   aRecord$ = "[MODIFY BREAKER]"
   nReadingBreakerData = 1
  ElseIf nReadingBreakerData = 1 Then
   nLineCount& = nLineCount + 1
  End If
 Loop
 Close 1
 If nLineCount = 0 Then 
   Print "No breaker data found in file"
   Stop
 End If
 
 
 Open DataFile For Input As 1
 
 DataFile = UCase(DataFile)
 nExtension = InStr( DataFile, ".CHF" )
 If nExtension > 0 Then
  ChangeFile$ = Mid(DataFile, 1, nExtension-1)
 Else
  ChangeFile$ = DataFile
 End If
 ChangeFile$ = ChangeFile & "_M.CHF"
 
 Open ChangeFile For output As 2

 ii& = 0  
 Button = 0
 nReadingBreakerData = 0
 Do While Not EOF(1)
  If Button = 2 Then 
   Print "Cancel button pressed"
   exit Do
  End If
  Line Input #1, aLine$ ' Read a line of data.
  aRecord$ = aLine
  If InStr( aLine$, " /" ) > 0 Then
   Do While Not EOF(1)
    Line Input #1, aLine$
    aRecord$ = aRecord$ & Chr(13) & Chr(10) & aLine
    If InStr( aLine$, " /" ) <= 0 Then exit Do
   Loop
  End If
  If InStr( aRecord$, "[ADD BREAKER" ) = 1 Then 
   aRecord$ = "[MODIFY BREAKER]"
   nReadingBreakerData = 1
  ElseIf nReadingBreakerData = 1 Then
   aRecord = updateBreakerData( aRecord )
  End If
  Print #2, aRecord$
  ii& = ii + 1
  Button = ProgressDialog( 1, "Process Breaker Data", "Record " & Str(ii) & " of " & Str(nLineCount ), 100*ii/nLineCount )
 Loop
 Call ProgressDialog( 0, "", "", 0 )
 Print "Modifed ", nLineCount, " records in: ", ChangeFile
End Sub

Function updateBreakerData( aRec$ ) As String
 dim myFields(100) As String
 
 nStart& = 1
 nFCount = 0
 
' aRec = "ABC 'DE FGH' IK /" & Chr(13) & Chr(10) & "1234 '' 56 "

 ' Read all data fields in array
 Do While true
  aField = CHFRecField( aRec, nStart )
  If aField = "" Then exit Do
'  Print aRec & "Field" & Str(nFCount) & "=" &  aField
  If aField <> "/" Then 
   nFCount = nFCount + 1
   myFields(nFCount) = aField
  End If
 Loop
 
 ' Make data modification
 ' Assign MaxDesign kV to operating kV
 myFields(19) = myFields(20)
  
 ' copy data to output 
 For ii = 1 to nFCount
  If ii = 1 Then 
   updateBreakerData = myFields(ii)
  Else
   updateBreakerData = updateBreakerData & " " & myFields(ii)
  End If
 Next
 
' Print aRec &Chr(13) & Chr(10) & updateBreakerData 
End Function

Function CHFRecField( aRec$, ByRef nStart& ) As String
 CHFRecField = ""
 nLen   = Len(aRec)
 If nStart >= nLen Then exit Function
 isString = 0
 aField$  = ""
 inField  = 0
 endField = 0
 Do While endField = 0
  aCh = Mid(aRec,nStart,1)
'  Print "char=" & aCh
  If aCh = " " Or aCh = Chr(13) Or aCh = Chr(10) Then
   If inField = 1 Then
    If isString = 1 Then aField = aField & aCh Else endField = 1
   End If
  Else
   inField = 1
   If aCh = "'" Then
    If isString = 0 Then
     isString = 1
    Else
     isString = 0
    End If
   End If
   aField = aField & aCh
  End If
  If nStart+nIndex = nLen Then endField = 1
  If endField <> 1 Then nStart = nStart + 1
 Loop 
 CHFRecField = aField
End Function
