VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "option exoplicit"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Command1_Click()
Dim lngStartTime As Long
Dim lngFinishTime As Long

lngStartTime = timeGetTime

' Here you usuallt put your sub calls that needs to be timed
' For example's sake I just put in a loop

Dim I As Integer
For I = 1 To 30000
Next I

lngFinishTime = timeGetTime

MsgBox ("This task took " & lngFinishTime - lngStartTime & " milliseconds!")

End Sub
