<div align="center">

## Form bouncing against the Start Menu at exit\! 2\.0


</div>

### Description

This code is a must have! It's cooler than the "Cool Form Close" code , cooler than the "Implode and Explode" code! So what does it do? When you close you program a really cool effect will appear. Your form will shrink so just the Titlebar is being showned, then the titlebar accelerates and bounces againt the start menu, goes up in the air, bounces a couple of more times and then disappears behind the Start Menu! Way Cool! This code is a very advanced one but it's really easy to use, try it!!! Includes functions for getting the top position of your start menu and offcourse the bounce code! New for ver. 2 is that the form now can bounce sideways if you edit the code just a little tiny bit, now also supports maximized windows!!!
 
### More Info
 
None really, you might wanna change the speed property in the code if you can find it.

Paste the 'main code' into the form_unload section. The declare the variables (IMPORTANT!)

A really cool 'bouncy' effect. I use it always in my progs!

None, but their might be bugs..


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Johan Otterud](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/johan-otterud.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/johan-otterud-form-bouncing-against-the-start-menu-at-exit-2-0__1-1920/archive/master.zip)

### API Declarations

```
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Dim What As RECT
```


### Source Code

```
Private Sub Form_Unload(Cancel As Integer)
  If Me.WindowState <> 0 Then
  Me.WindowState = 0
  End If
Cancel = -1
Dim HeightOfStartMenu As Long
Dim Speed As Long
Dim StartAt As Long
For I = 1 To 999 '// The start menu never uses a HWND higher than 1000
 z$ = Space$(128)
    Y = GetClassName(I, z$, 128)
    X = Left$(z$, Y)
    If LCase(X) = "shell_traywnd" Then
    GoTo JumpOut:
    End If
Next I
JumpOut:
GetWindowRect I, What
'// Get the top pos of the Start Menu
HeightOfStartMenu = What.Top * 15
If HeightOfStartMenu <= 0 Then
HeightOfStartMenu = Screen.Height
'// If some smart guy moves the start-menu, to say
'// the top, left or right bounce at the bottom of
'// the screen
End If
'// Turn the value into twips (more commonly used)
StartAt = HeightOfStartMenu - 4000
If StartAt < Me.Top Then
StartAt = Me.Top
'// This code prevents the form from bouncing
'// higher than itself (not logical, the start menu isn't made
'// of rubber you now)
End If
'// How many "bounces?"
Speed = 100
'// How fast should this go?
Me.Height = 0
Me.Width = 4000
GoAgain:
Do Until Me.Top >= HeightOfStartMenu
DoEvents
Me.Top = Me.Top + Speed
Me.Left = Me.Left + 15 '<--- Remove the " ' " to make the window bounce sideways!
Loop
Do Until Me.Top <= StartAt
DoEvents
Me.Top = Me.Top - Speed
Me.Left = Me.Left + 15 '<--- Remove the " ' " to make the window bounce sideways!
Loop
If StartAt >= 10000 And Me.Top >= HeightOfStartMenu Then
  Do Until Me.Top >= HeightOfStartMenu + 15000
  Me.Top = Me.Top + Speed
  Loop
End
Exit Sub
End If
StartAt = StartAt + 1000
Speed = Speed - 5
'// Decrease speed with 5 after each "bounce",
'// You can change the value all ya want :)
If Speed <= 0 Then
Speed = 5
'// If the Speed value gets under zero i will
'// automatically turn into 5 (cause if it don't
'// It will stop or do something crazy
End If
GoTo GoAgain:
End Sub
```

