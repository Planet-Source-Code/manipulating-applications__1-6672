<div align="center">

## Manipulating Applications


</div>

### Description

This code will minimize, maximize and send keystrokes from within your program to any application.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/manipulating-applications__1-6672/archive/master.zip)





### Source Code

```
'There are three programs which:
'send any keystroke to any application.
'minimize any window(of any application).
'maximise any window(of any application).
'Send Keystroke'
'~~~~~~~~~~~~~~'
Dim ReturnValue, I
ReturnValue = Shell("App. Name", 1)  ' e.g. Shell ("Calc.exe",1)
AppActivate ReturnValue  ' Activate the Calculator.
For I = 1 To 100  ' Set up counting loop.
  SendKeys I & "{+}", True  ' Send keystrokes to Calculator
Next I  ' to add each value of I.
SendKeys "=", True  ' Get grand total.
SendKeys "%{F4}", True  ' Send ALT+F4 to close Calculator.
'Minimize'
'~~~~~~~~'
Private Declare Function CloseWindow Lib "user32" Alias "CloseWindow" _
(ByVal hwnd As Long) As Long
Shell "Calc.exe",1  'This will start calc. Any appl. can be opened
		   'like this
a = screen.activeform.hwnd 'will return the window handle of calc.
			  'to get handle of own app. don't use shell
			  'command and write the code as it is.
closewindow(a) 'will minimize calc
'Maximize'
'~~~~~~~~'
'this code assumes that the application is opened but minimized
'If application is not opened you may use /Shell "App Nm"/ to open it
Appactivate "Title",True 'Here Title is the one which is shown in the
			 'title bar of the application
			 'Shell command automatically gives title so
			 'in above example to send keystroke we did
			 'not mention title.
SendKeys "%( x)" '% stands for alt and then a blank and x is sent
		 'to maximize.
```

