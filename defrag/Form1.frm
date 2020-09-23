VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Defrag"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   600
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   120
      Top             =   720
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1680
      TabIndex        =   2
      Text            =   "PM"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   840
      TabIndex        =   1
      Text            =   "00"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Text            =   "3"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example shows how to make your computer defragement at a
'certain time.  I took this from a much larger program I made.
'                                            -Michael Nipper

Private Sub Form_Load()

Do While a < 12          'Loads combo1 with numbers 1 to 12 so I
a = a + 1                'Don't have to manually type them in.
Combo1.AddItem (a)
Loop

Do While b < 60
b = b + 1
If b < 10 Then
Combo2.AddItem ("0" & b) 'Makes it so every number under 10 it adds
Else                     'a zero in front of it to make it look more
Combo2.AddItem (b)       'like a time.
End If
Loop

Combo3.AddItem ("AM")
Combo3.AddItem ("PM")

'Counts the number of hard drives the user has
Dim NumDrives, FSO, Drives, Count
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Drives = FSO.Drives
NumDrives = Drives.Count
NumDrives = NumDrives - NumDrives
Const DriveTypeFixed = 2
Const DriveTypeNetwork = 3
For Each Drive In Drives
If Drive.DriveType = DriveTypeFixed Then NumDrives = NumDrives + 1
Next
Set Drives = Nothing
Set FSO = Nothing
Label1 = NumDrives

End Sub

Private Sub Timer1_Timer()
'When the current time (Time) is equal to the comboboxes, then begin.
If Time = Combo1.Text + ":" + Combo2.Text + ":00 " + Combo3.Text Then
Set wshshell = CreateObject("WScript.Shell")    'Loads up the
wshshell.Run "dfrg.msc"                         'Defrag program.
Pause 5
Dim mmcmainframe As Long, mdiclient As Long, mmcchildfrm As Long
Dim mmcviewwindow As Long, mmcocxviewwindow As Long, atlaxwinex As Long
Dim atldac As Long, btn As Long
'First it finds the main window of the disk defragmenter.
mmcmainframe = FindWindow("mmcmainframe", vbNullString)
mdiclient = FindWindowEx(mmcmainframe, 0&, "mdiclient", vbNullString)
mmcchildfrm = FindWindowEx(mdiclient, 0&, "mmcchildfrm", vbNullString)
mmcviewwindow = FindWindowEx(mmcchildfrm, 0&, "mmcviewwindow", vbNullString)
mmcocxviewwindow = FindWindowEx(mmcviewwindow, 0&, "mmcocxviewwindow", vbNullString)
atlaxwinex = FindWindowEx(mmcocxviewwindow, 0&, "atlaxwinex", vbNullString)
atldac = FindWindowEx(atlaxwinex, 0&, "atl:6d3a99c8", vbNullString)
'Since there is two buttons, it must first find the first button and then find
'the second button according to the first.
btn = FindWindowEx(atldac, 0&, "button", vbNullString)
btn = FindWindowEx(atldac, btn, "button", vbNullString)
'Click the second button using the space bar.  First WM_KEYDOWN, VK_SPACE presses the
'space bar down and the second line of code lets it back up.  WM_KEYDOWN is a Constant
'in module1.bas because it is easier to remember than &H100 so you don't have to keep
'looking up or refering back to what the code for what keydown is.
Call PostMessage(btn, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(btn, WM_KEYUP, VK_SPACE, 0&)
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
'This searches for the window that pops up saying that it has finished defragementing
'the disk it had been working on.  If it does not find it then it keeps looping until
'it does.
Dim X As Long
X = FindWindow("#32770", "Disk Defragmenter")
    If X = 0 Then
    Timer2.Enabled = True
    Else
    Timer2.Enabled = False
    Label1 = Label1 - 1
        'It then makes sure that it is not the last disk so it does not keep going
        'after it has already defragemented all the disk.
        If Label1.Caption = "0" Then
        'If there is no more disk then it stops the timer
        Timer2.Enabled = False
        Else
        'If label1.caption does not = 0 then there are more hard drives to be
        'defragemented and it then selects the next disk and defragments it.
        Pause 5
        Dim Button As Long
        'This is where it takes care of the box that came up telling you it has
        'finished defragementing the current volume.
        X = FindWindow("#32770", "Disk Defragmenter")
        Button = FindWindowEx(X, 0&, "button", vbNullString)
        Button = FindWindowEx(X, Button, "button", vbNullString)
        Button = FindWindowEx(X, Button, "button", vbNullString)
        Button = FindWindowEx(X, Button, "button", vbNullString)
        Button = FindWindowEx(X, Button, "button", vbNullString)
        'Hit the ok button
        Call SendMessageLong(Button, &H201, 0&, 0&)
        Call SendMessageLong(Button, &H202, 0&, 0&)
        'Allow pleanty of time for the window to close
        Pause 5
        'This clicks on the list so we can just use a sendkeys to press down because
        'it is just quicker and just as effective in this instance as API
        Dim mmcmainframe As Long, mdiclient As Long, mmcchildfrm As Long
        Dim mmcviewwindow As Long, mmcocxviewwindow As Long, atlaxwinex As Long
        Dim atldac As Long, syslistview As Long
        mmcmainframe = FindWindow("mmcmainframe", vbNullString)
        mdiclient = FindWindowEx(mmcmainframe, 0&, "mdiclient", vbNullString)
        mmcchildfrm = FindWindowEx(mdiclient, 0&, "mmcchildfrm", vbNullString)
        mmcviewwindow = FindWindowEx(mmcchildfrm, 0&, "mmcviewwindow", vbNullString)
        mmcocxviewwindow = FindWindowEx(mmcviewwindow, 0&, "mmcocxviewwindow", vbNullString)
        atlaxwinex = FindWindowEx(mmcocxviewwindow, 0&, "atlaxwinex", vbNullString)
        atldac = FindWindowEx(atlaxwinex, 0&, "atl:6d3a99c8", vbNullString)
        syslistview = FindWindowEx(atldac, 0&, "syslistview32", vbNullString)
        Call SendMessageLong(syslistview, &H201, 0&, 0&)
        Call SendMessageLong(syslistview, &H202, 0&, 0&)
        'This is using AppActivate to send the key down.  Since this is a relatively
        'simple function in a relatively simple program it is ok in this instance.
        'Normally, however, it should be avoided and API should be used.
        AppActivate ("Disk Defragmenter")
        SendKeys "{DOWN}"
        Pause 3
        'Hits the defrag button again for the new disk volume.
        mmcmainframe = FindWindow("mmcmainframe", vbNullString)
        mdiclient = FindWindowEx(mmcmainframe, 0&, "mdiclient", vbNullString)
        mmcchildfrm = FindWindowEx(mdiclient, 0&, "mmcchildfrm", vbNullString)
        mmcviewwindow = FindWindowEx(mmcchildfrm, 0&, "mmcviewwindow", vbNullString)
        mmcocxviewwindow = FindWindowEx(mmcviewwindow, 0&, "mmcocxviewwindow", vbNullString)
        atlaxwinex = FindWindowEx(mmcocxviewwindow, 0&, "atlaxwinex", vbNullString)
        atldac = FindWindowEx(atlaxwinex, 0&, "atl:6d3a99c8", vbNullString)
        Button = FindWindowEx(atldac, 0&, "button", vbNullString)
        Button = FindWindowEx(atldac, Button, "button", vbNullString)
        Call PostMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button, WM_KEYUP, VK_SPACE, 0&)
        Pause 5
        'Start the process all over again.
        Timer2.Enabled = True
        End If
    End If
End Sub
