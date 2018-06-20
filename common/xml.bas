' ASPEN PowerScrip sample program
'
' XML.BAS
'
' XML functions
'
' Version: 1.0
'
' PowerScript functions called:
'

' Global variables
Sub main()
   xml$ = xmlMakeNode ( "MyNode" )
   Call xmlSetAttribute( xml, "Att1", "value1" )
   Call xmlSetAttribute( xml, "Att1", "valueX" )
   Call xmlSetAttribute( xml, "Att1", "valX" )
   Call xmlSetAttribute( xml, "Att1", "" )
End Sub

Function xmlMakeNode( sName$ ) As String
  xmlMakeNode = "<" & sName & " />"
End Function

Function xmlSetAttribute( ByRef sNode$, sAttName$, sAttValue$ ) As long
  sAtt$ = sAttName & "="""
  nFind = InStr(1, sNode, sAtt)
  If nFind = 0 Then 
    ' New Attribute 
    nFind = InStr(1, sNode, "/>" )
'    Print "sAtt", sAtt, "sNode", sNode, "nFind", nFind, "Left", Left(sNode, nFind-1), "right", Mid(sNode, nFind, 99)
    sNode = Left(sNode, nFind-1) & sAtt$ & sAttValue & """" & Mid(sNode, nFind, Len(sNode))
  Else
    ' Existing Attribute 
    nFind2 = InStr(nFind+Len(sAtt), sNode, """" )
'    Print "sAtt", sAtt, "sNode", sNode, "nFind", nFind, "Left", Left(sNode, nFind-1), "nFind2", nFind2, "mid", Mid(sNode, nFind2, Len(sNode))
    If sAttValue = "" Then
	  sNode = Left(sNode, nFind-1) & Mid(sNode, nFind2+1, Len(sNode))
	Else
	  sNode = Left(sNode, nFind-1) & sAtt$ & sAttValue & Mid(sNode, nFind2, Len(sNode))
	End If
  End If
'  Print sNode
  xmlSetAttribute = 1
End Function
