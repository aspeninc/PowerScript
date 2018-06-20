' ASPEN PowerScript Sample Program
'
' COMPLEX.BAS
'
' Do complex number addtion
'
' Version 1.0
'

Type COMPLEX
 real As double
 imag As Double
End Type

Sub main
 dim num1 As COMPLEX, num2 As COMPLEX, num3 As COMPLEX

 num1.real = 1
 num1.imag = 2
 num2.real = 3
 num2.imag = 4
 Call dadd(num1, num2, num3)
 Print num1.real, "+j", num1.imag
 Print num2.real, "+j", num2.imag
 Print num3.real, "+j", num3.imag
End Sub

Function dadd( ByRef num1 As COMPLEX, ByRef num2 As COMPLEX, ByRef sum As COMPLEX ) As long
  dim result As COMPLEX
  sum.real = num1.real + num2.real
  sum.imag = num1.imag + num2.imag
  dadd = 1
End Function
