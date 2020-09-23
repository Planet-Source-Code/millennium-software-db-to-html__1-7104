VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#############################################
'In order for this to work correctly, you MUST
'make a reference to the Microsoft DAO Library
'
'Click on Project --> References
'Look for Microsoft DAO 3.0 Library (or higher)
'#############################################


Dim db As Database
Dim rs As Recordset


Private Sub Command1_Click()

'First you need to open the database
Set db = OpenDatabase(App.Path & "\booksales.mdb")

'Now open the table you want to use...
Set rs = db.OpenRecordset("Titles")

'Have VB create a file on the hard disk
'to write to
Open App.Path & "\Book_Titles.htm" For Output As #1

'Now print the initial HTML document structure
Print #1, "<HTML>"
Print #1, "<HEAD>"
Print #1, "<TITLE>"
Print #1, "Just Testing..."
Print #1, "</TITLE>"
Print #1, "</HEAD>"
Print #1, "<BODY>"

'Here the info from the database is written
'into the document
Do Until rs.EOF
Print #1, "<B>"; rs!Title; "</B><BR>"; "$"; rs!Price; ".00<P>"

'Go to the next record
rs.MoveNext

'Start over with the next record
Loop

'Print the closing HTML document tags
Print #1, "</BODY>"
Print #1, "</HEAD>"

'The file is finished,
'tell VB to stop writing
'this file
Close #1

'Let the user know we are done
Label1.Caption = "Finished!"
End Sub

