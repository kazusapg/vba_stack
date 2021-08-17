# VBA Stack

[![license](https://img.shields.io/github/license/mashape/apistatus.svg)](LICENSE)

This cls file is to implement stack with vba dynamic array.

## Install
1. Download Stack.cls.
2. Add Stack.cls to class module of the VBA Project.

## Usage
1. Make stack object from Stack.cls.
2. Use push method to add data to the stack.
3. Use pop method to pop data from the stack.
4. If you use pop method when there is no data in the stack, an error 1001 will be raised.


## Code Example
```VBA
Public Sub TestStack()

  'Make stack object
  Dim s As Stack
  Set s = New Stack
  
  'Push different types of data to the stack
  s.Push 1
  Dim c As Collection
  Set c = New Collection
  c.Add "value_test", "key_test"
  s.Push c
  s.Push "a"

  'Show the number of data in the stack
  MsgBox "The number of data in the stack is " & s.Count & "."
  
  'Get all the data from the stack and store it in an array
  Dim ary() As Variant
  ary = s.GetContents
  
  'Pop the data in the stack
  MsgBox "The data popped from the stack is " & s.Pop & "."
  Dim returnC As Collection
  Set returnC = s.Pop
  MsgBox "The data popped from the stack is " & returnC("key_test") & "."
  MsgBox "The data popped from the stack is " & s.Pop & "."
  
On Error GoTo ErrorLabel
  'If you pop when the stack is empty, an error 1001 will be raised.
  MsgBox "The data popped from the stack is " & s.Pop & "."
  
  GoTo Finally
  
ErrorLabel:
  If Err.Number = 1001 Then
    MsgBox Err.Description, vbCritical
  End If

Finally:
  Set s = Nothing
  
End Sub
```
