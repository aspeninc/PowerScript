'ASPEN Sample Script

Sub main
 If GetEquipment( TC_PICKED, PickedHnd& ) = 0 Or _
    EquipmentType( PickedHnd& ) <> TC_LINE Then
    Print "Please select a line"
    Stop
  End If

  dim vdRating(4) As double
  Call GetData( PickedHnd, LN_vdRating, vdRating )
  Print "Rating A=", vdRating(1), _
        "Rating B=", vdRating(2), _
        "Rating C=", vdRating(3), _
        "Rating D=", vdRating(4)

End Sub
