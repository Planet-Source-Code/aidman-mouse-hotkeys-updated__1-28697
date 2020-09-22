<div align="center">

## Mouse HotKeys \(Updated\)


</div>

### Description

An easy, very fast and well explained program that let&#8217;s you control your mouse courser actions by custom keyboard keys and custom movements speed. Showing the easiest way of hotkeys and controlling your mouse cursor movements and button actions. It has been tested and works in all Windows platforms (95, 98, ME, NT, 2000, XP) and the most games like Half-life (Counter-Strike).

There are many functions including:

* Moving at all directions

* Left, right and middle buttons

* Single clicking

* Double clicking

* Drag and drop

* Several keys down and up at once

All in just one single small and very fast module sub, the main sub. Whit out anything else like form, dll or ocx. Read "Assumes" in the Source Code for updated using instructions and notes.
 
### More Info
 
See the new Numpad Keys! When using the program in games, increase the AddSpeed and MaxSpeed, remove also the Doevents function.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aidman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aidman.md)
**Level**          |Advanced
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aidman-mouse-hotkeys-updated__1-28697/archive/master.zip)

### API Declarations

```
Option Explicit ' Of course
' Position variable type for api
Private Type POINTAPI
 X As Long
 Y As Long
End Type
' Windows api declartions
Private Declare Sub Sleep Lib "Kernel32" (ByVal milliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
' Mouse event messages
Const MOUSEEVENTF_ABSOLUTE = 32768 ' not used in this program
Const MOUSEEVENTF_LEFTDOWN = 2
Const MOUSEEVENTF_LEFTUP = 4
Const MOUSEEVENTF_MIDDLEDOWN = 32
Const MOUSEEVENTF_MIDDLEUP = 64
Const MOUSEEVENTF_MOVE = 1
Const MOUSEEVENTF_RIGHTDOWN = 8
Const MOUSEEVENTF_RIGHTUP = 16
```


### Source Code

```
' Key variable numbers
Const MoveUp = 1
Const MoveDown = 2
Const MoveLeft = 3
Const MoveRight = 4
Const ButtonLeft = 5
Const ButtonMiddle = 6
Const ButtonRight = 7
Const EndProgram = 8
' Movement Speed settings
Const AddSpeed = 0.1 ' Pixel(s)
Const MaxSpeed = 100 ' Pixel(s)
Const SleepTime = 1 ' Millisecond(s)
Private Sub Main() ' Start sub
Dim KeyNumber(1 To 8) As Long ' Key numbers
Dim KeyValue(1 To 8) As Boolean ' Key press values
Dim OldValue(1 To 8) As Boolean ' Old key press values
Dim MoveSpeed(1 To 4) As Single ' Speed of move keys
Dim Position As POINTAPI ' Cursor position in api type
Dim MoveKeys As Boolean ' Value of any true move key(s)
Dim Count As Integer ' For-next counter
  KeyNumber(MoveUp) = vbKeyNumpad8 ' Set move up key
  KeyNumber(MoveDown) = vbKeyNumpad5 ' Set move down key
  KeyNumber(MoveLeft) = vbKeyNumpad4 ' Set move left key
  KeyNumber(MoveRight) = vbKeyNumpad6 ' Set move right key
  KeyNumber(ButtonLeft) = vbKeyNumpad7 ' Set button left key
  KeyNumber(ButtonMiddle) = vbKeyNumpad2 ' Set button middle key
  KeyNumber(ButtonRight) = vbKeyNumpad9 ' Set button right key
  KeyNumber(EndProgram) = vbKeyEscape ' Set end program key
  Do ' Start the loop
    Sleep SleepTime ' Loops works better with the sleep function
    DoEvents ' Check other events too
    MoveKeys = False ' Clear last move keys value
    For Count = 1 To 8 ' Get the all 8 key press values
      GetAsyncKeyState KeyNumber(Count) ' Clear last key press
      OldValue(Count) = KeyValue(Count) ' Save old value
      KeyValue(Count) = False ' Clear last key press value
      If GetAsyncKeyState(KeyNumber(Count)) Then ' Check if key press
        KeyValue(Count) = True ' Set key press as true
        If Count < 5 Then MoveKeys = True ' If move key then set move key(s) as true
      End If
    Next Count ' Get next key press value
    If KeyValue(EndProgram) Then End ' If end key is pressed then end program
    If MoveKeys Then ' If any move key(s) are pressed then
      GetCursorPos Position ' Get the current mouse cursor position
      For Count = 1 To 4 ' Do all 4 movement actions
        If KeyValue(Count) Then ' If move key is pressed then,
          If Not OldValue(Count) Then MoveSpeed(Count) = 0 ' If key has just been pressed then set movement speed to null
          If MoveSpeed(Count) < MaxSpeed Then ' If movement speed is lower then 100 then,
            MoveSpeed(Count) = MoveSpeed(Count) + AddSpeed ' Increase movement speed
          Else
            MoveSpeed(Count) = MaxSpeed ' Else, set movement speed as maximum speed limit
          End If
          Select Case Count ' Select movement direction
            Case MoveUp: Position.Y = Position.Y - MoveSpeed(MoveUp) ' Decrease "Y" position
            Case MoveDown: Position.Y = Position.Y + MoveSpeed(MoveDown) ' Increase "Y" position
            Case MoveLeft: Position.X = Position.X - MoveSpeed(MoveLeft) ' Decrease "X" position
            Case MoveRight: Position.X = Position.X + MoveSpeed(MoveRight) ' Increase "X" position
          End Select
        End If
      Next Count ' Next movement action
      SetCursorPos Position.X, Position.Y ' Set new mouse cursor position
      mouse_event MOUSEEVENTF_MOVE, 0, 0, 0, 0 ' Inform other programs that mouse has moved
    End If
    For Count = 5 To 7 ' Do all 3 click actions
      If KeyValue(Count) And Not OldValue(Count) Then ' If button key has just been pressed then,
        Select Case Count ' Select button down
          Case ButtonLeft: mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0 ' Send button left mouse down command
          Case ButtonMiddle: mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0 ' Send button middle mouse down command
          Case ButtonRight: mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0 ' Send button right mouse down command
        End Select
      ElseIf Not KeyValue(Count) And OldValue(Count) Then ' If button key has just been released then,
        Select Case Count ' Select button up
          Case ButtonLeft: mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0 ' Send button left mouse up command
          Case ButtonMiddle: mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0 ' Send button middle mouse up command
          Case ButtonRight: mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0 ' Send button right mouse up command
        End Select
      End If
    Next Count ' Next click action
  Loop ' Continue looping until end program key is true
End Sub
```

