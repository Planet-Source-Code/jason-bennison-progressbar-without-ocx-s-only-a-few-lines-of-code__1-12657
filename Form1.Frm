VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1905
   ClientLeft      =   2640
   ClientTop       =   3225
   ClientWidth     =   6690
   Height          =   2310
   Left            =   2580
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   6690
   Top             =   2880
   Width           =   6810
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Progressbar Demo"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Height          =   275
      Left            =   630
      TabIndex        =   2
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Response
Dim J As Integer 'loop ctrl
Dim Percentage As Integer

Response = InputBox("Enter a value") 'get a value from the user

If Not IsNumeric(Response) Then Response = 5000
If Response <= 0 Or Response >= 32768 Then Response = 5000

Label2.Visible = True 'Show the blue bit

For J = 1 To Response
    'your event or function being recursed goes here.....
    'or you can put the code below into the recursed function itself.
    'If youre running a fast PC you might want to put something here to slow it down abit, - The progress bar may flicker.

DoEvents
    Percentage = J / Response * 100
    If Percentage <= 0 Then Percentage = 1
    Label2.Width = (Label1.Width / 100 * Percentage) - 30
    Label3.Caption = Str(Percentage) & "% Done"
Next

Label2.Visible = False 'hide the blue bit when done
Label3.Caption = "Progressbar Demo"
Label2.Width = 1
End Sub


Private Sub Form_Activate()
Label2.Width = 1
Label2.Visible = False
Label3.Caption = "Progressbar Demo"
End Sub

