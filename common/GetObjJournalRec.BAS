
' ASPEN PowerScript sample program
'
' GETOBJJOURNALREC.BAS
'
' Version: 1.0
' Category: Common
'
' Demo the GetObjJournalRecord() PowerScript function
'
Sub main
  hnd& = 0
  count = 0
  While 1 = GetEquipment(TC_RLYOCG,hnd)
    JRec$ = GetObjJournalRecord(hnd)
    Call parseALine( JRec$, Chr(10), dateCreated$, JRec$ )
    Call parseALine( JRec$, Chr(10), CreatedBy$, JRec$ )
    Call parseALine( JRec$, Chr(10), dateModified$, JRec$ )
    Call parseALine( JRec$, Chr(10), ModifiedBy$, JRec$ )
    Print PrintObj1LPF(hnd) & Chr(10) & _
      " Created: " & dateCreated & " by: " & CreatedBy & _
      Chr(10) & " Modified: " & dateModified & " by: " & ModifiedBy
    If count >= 5 Then Stop
    count = count + 1
  Wend
End Sub

Sub parseALine( ByVal aLine$, ByVal Delim$, ByRef sLeft$,  ByRef sRight$ )
  nPos = InStr( 1, aLine$, Delim$ )
  If nPos = 0 Then
    sLeft = aLine$
    sRight = ""
  Else
    sLeft = Left(aLine$, nPos-1)
    sRight = Mid(aLine$, nPos+Len(Delim), 9999 )
  End If
  sLeft  = Trim(sLeft)
  sRight = Trim(sRight)
End Sub
