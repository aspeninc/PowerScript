' ASPEN PowerScript Sample Program
'
' Progress.Bas
'
' Demonstrate the use of ProgressDialog function
'
Sub Main
  Print "Start"
  For ii = 1 to 100
    ' Show the dialog
    Button = ProgressDialog( 1, "My Dialog", "Progress =" + Str(ii) +"%", ii )
    ' Carry out desired logic
    For jj = 1 to 1000000
    Next
    ' Check for button pressed
    If Button = 2 Then 
      Print "Cancel button pressed"
      GoTo Done
    End If
  Next
  Print "done"
Done:
  ' Hide the dialog
  Call ProgressDialog( 0, "", "", 0 )
End Sub
