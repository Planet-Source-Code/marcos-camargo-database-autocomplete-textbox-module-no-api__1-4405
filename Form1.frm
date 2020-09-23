VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoComplete TextBox"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3915
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "mautheman@yahoo.com"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Type the name you're looking for :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Example of AutoComplete TextBox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This is an example of how to use the AutoCompleteTextBox
'just make sure the database "sampledatabase.mdb" and the
'program are in the same directory
'any problem with the program, e-mail me
'                    mautheman@yahoo.com
'**************************************************************
Public db As Database 'declare the database
Public ds As Recordset 'declare the table
Private Sub Command1_Click()
End
End Sub
Private Sub Form_Load()
On Error GoTo ErrLine:

Set db = OpenDatabase(App.Path & "\sampledatabase.mdb") 'open the database
Set ds = db.OpenRecordset("sampletable", dbOpenDynaset) 'open the table
    
Exit Sub

ErrLine:
MsgBox "There was a problem while opening the database or table, it's impossible to open this program (check if the database 'sampledatabase' and the programa are in the same directory )", vbCritical, "Problems"
End

End Sub
Private Sub txtName_Change()
AutoComplete txtName, db, "sampletable", "Author"
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
CheckIsDelOrBack (KeyCode)
End Sub
