VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Big Factorial"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Calculation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3270
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   270
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5280
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Percentage of completed calculation:"
      Height          =   195
      Left            =   510
      TabIndex        =   6
      Top             =   1620
      Width           =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number to find factorial: "
      Height          =   255
      Left            =   1170
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'By Isbat Sakib.
'E-mail: sakib039@hotmail.com

'So, this is the updated program. Use it as you like. Test it,
'experiment with it. This program should not do any harm to
'your system but if it does, I am in no way responsible for
'that. Sorry, but I have to say this.

'Now something about the algorithm. This algorithm is not, I
'repeat, is not invented by me. Actually it is quite a common
'algorithm for computer science students. I tried to make the
'same program in C++, but I don't know why it didn't work out,
'probably due to the limitations of DOS. It could only calculate
'upto 9274. And at present I don't know how to make windows
'programs in C/C++. However, I think, very expert C/C++ programmers
'can find the factorial of 1,000,000,000 with the same
'algorithm.

'Actually this program holds the whole result in a string.
'A string can hold maximum 2^31 characters. So, this program
'can actually calculate the factorial of those numbers whose
'result contains less than 2^31 digits. This means it should be
'able to calculate the factorial of more than 100,000,000.
'I don't have time to wait for hours and check whether it actually
'can do it. But theoritically it should be able to do it.
'This is the present limitation of this program. I hope you guys
'can overcome this limitation. And I have to say, if you like,
'you can vote for me. I appreciate votes. Anyway, thanks.

Dim stopcal As Boolean

Private Sub Form_Load()
'The following lines ensures that this program will not run
'in the IDE. You should compile and run the executable file
'for faster calculations. However, if you want to run it in
'the IDE, then just comment out the following lines.

If App.LogMode <> 1 Then
    MsgBox "Please, compile me and then run me. I will run faster in that way.", vbInformation, "Compile first"
    Unload Me
End If
End Sub

Private Sub Command1_Click()
'Well, this button actually starts the calculation by first
'checking the value. This value is quite arbitrarily taken
'because I didn't have time to check how big numbers this
'program can actually calculate. It cannot possibly calculate
'the factorial of 1,000,000,000 right now because I think the
'result will contain more than 2^31 digits as a result of
'which a string even cannot hold the result. So, if you can,
'then try to make this program more powerful. Please, send me
'the source code if you really can find the factorial of
'1,000,000,000 by this program. And please don't forget to send
'me the huge result also. My address is sakib039@hotmail.com


If Val(Text1) > 50000000 Then
    MsgBox "Unable to calculate such a large number.", vbCritical, "Too Big Number"
Else
    ShowResult
End If
End Sub

Private Sub ShowResult()
'These are just safety precautions taken.
stopcal = False
Command1.Enabled = False
Command2.Enabled = True
Text1.Locked = True
Text2.Text = ""
Text3.Text = ""

'This is the actual array that is going to hold all the digits.
Dim a() As Byte

'These are the numerous variables. They are all long with one
'double. You can possibly calculate larger numbers if you
'declare all these variables as doubles. But then, you probably
'have to change some other sides.
Dim num As Long
Dim p As Double
Dim numdigit As Long, i As Long, carry As Long, z As Long
Dim sum As Long, j As Long, flag As Long
Dim result As String

num = Val(Text1) 'So 'num' is the number you have input.
p = 0#

'This part is important. If you take the log base 10 from 2 to
'any special number and add them and then add one to that result and
'then round the result to get a whole number, then that number
'will indicate how many digits will the factorial of that special
'number contain.
For i = 2 To num
    p = p + (Log(num) / Log(10#))
Next i

numdigit = Round(p) + 1

'So we initialize the array
ReDim a(numdigit) As Byte

For i = 1 To numdigit
    a(i) = 0
Next i
a(0) = 1
p = 0#

'This is the actual algorithm. Try to understand it.
For i = 2 To num
    DoEvents
    carry = 0 'this is the remainder. It is needed to initialize it to zero on every iteration.
    
    p = p + (Log(i) / Log(10#))
    z = Round(p) + 1 'we get the number of digits in the factorial of every 'i' in the variable z
        
    If stopcal = True Then Exit Sub   'this line just checks whether the 'stop calculation' button has been pressed or not

    Text4.Text = CStr(Format(i / num * 90, "##0.000")) + "%"   'this line prints the percentage from 1 to 90%
    
    For j = 0 To CLng(z)
        sum = CLng(a(j) * i) + carry
        carry = sum \ 10
        a(j) = sum Mod 10
    Next j
Next i

flag = 0

For i = (numdigit - 1) To 0 Step -1
    DoEvents
    If a(i) <> 0 And flag = 0 Then flag = 1
    If flag = 1 Then
        
        'this line just checks whether the 'stop calculation' button has been pressed or not
        If stopcal = True Then Exit Sub
         
         'this line prints the percentage from 91 to 99.999%
        If num > 3 Then Text4.Text = CStr(Format((numdigit - 1 - i) / (numdigit - 1) * 10 + 89.999, "##0.000")) + "%"
        
        'this is the retrieval of the result in the string
        result = result + CStr(a(i))
    End If
Next i
Text4.Text = "100.000%"

Dim msg As String, length As Long
Dim nameoffile As String
Dim resp
length = Len(result)


'Hey...the windows textbox is limited to only 64k characters.
'So, those results exceeding that number of digits cannot be
'shown in the textbox. So, I added the capability to save the
'result in a file.

If length < 64000 Then
    Text2.Text = result
Else
    msg = "The whole result cannot be displayed here due to the limitation of the textbox. A textbox can show maximum 64k characters whereas the total result contains " + CStr(length) + " digits."
    Text2.Text = msg
End If

'This result is always shown. This is the scientific form
'and is shown in the lower textbox.
If Len(result) > 1 Then
    Text3.Text = Left$(result, 1) & "." & Mid$(result, 2, 26) & "+e" & CStr(Len(result) - 1)
Else
    Text3.Text = result
End If

'It just asks the user to save the result.
If length > 64000 Then
    msg = msg + " However, you can save the whole result in a file. Do you want to save the result in a disk file?"
    resp = MsgBox(msg, vbInformation + vbYesNo, "Do you want to save?")
Else
    resp = MsgBox("Do you want to save the total result in a disk file?", vbQuestion + vbYesNo, "Asking to save")
End If

'These are saving routines. Pretty straight forward.
If resp = vbYes Then
    cmdlg.DialogTitle = "Save the result"
    cmdlg.Filter = "Text files (*.txt)|*.txt"
    cmdlg.Flags = &H2
    cmdlg.ShowSave
    If cmdlg.FileName <> "" Then nameoffile = cmdlg.FileName
    If nameoffile <> "" Then
        Open nameoffile For Output As #1
        Print #1, result
        Close #1
    End If
End If

'And these are finishing touch-ups.
Text1.Locked = False
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command2_Click()
'This button just stops the calculation.

stopcal = True
Command1.Enabled = True
Command2.Enabled = False
Text1.Locked = False
Text4.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'This ensures that the user cannot input anything but numbers
'in the textbox.
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

