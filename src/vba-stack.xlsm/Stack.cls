Option Explicit

Private ary() As Variant   'Use an array as a stack
Private size As Long       'A variable that remembers the size of the stack

Private Sub Class_Initialize()
  'constructor
  
  ReDim ary(0)
  size = 0
  
End Sub
Public Sub Push(v As Variant)
  'Add data to the stack
  
  size = UBound(ary) + 1
  ReDim Preserve ary(size)
  
  Dim vType As Long
  vType = VarType(v)
  
  Select Case vType
    'Use Set syntax for objects, data access objects, and user-defined types
    Case vbObject
      Set ary(size) = v
    Case vbDataObject
      Set ary(size) = v
    Case vbUserDefinedType
      Set ary(size) = v
    Case Else
      ary(size) = v
  End Select
  
End Sub
Public Function Pop() As Variant
  'Pop the data on the stack
  
  If size = 0 Then
    'Raise an error number 1001 when there is no data on the stack
    Err.Raise 1001, "Stack", "There is no data in the stack"
  End If
  
  Dim vType As Long
  vType = VarType(ary(size))
  
  Select Case vType
    'Use Set syntax for objects, data access objects, and user-defined types
    Case vbObject
      Set Pop = ary(size)
    Case vbDataObject
      Set Pop = ary(size)
    Case vbUserDefinedType
      Set Pop = ary(size)
    Case Else
      Pop = ary(size)
  End Select
  
  size = size - 1
  ReDim Preserve ary(size)
  
End Function
Public Function Count() As Long
  'Return the size of the stack
  
  Count = size

End Function
Public Function GetContents() As Variant
  'Returns the contents of the stack as an array
  'Array index 0 is always empty.
  'Array index 1 is the bottom data on the stack.
  'Array index 2 is the second data from the bottom of the stack.
  
  GetContents = ary

End Function

