VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBHelper As bcDatabase.DataHelper
Dim DB As bcDatabase.DataConnection

Private Sub Combo1_Click()
ShowTable
End Sub

Private Sub Form_Load()

'Create the object
Set DBHelper = New bcDatabase.DataHelper

'Give it the path to the database
DBHelper.DatabaseName = App.Path & "\test.mdb"

'Tell it to load the tables into a combo box
DBHelper.TablesToCombo Combo1

Combo1.ListIndex = 0
ShowTable

End Sub

Private Sub Form_Unload(Cancel As Integer)
'clean up
DBHelper.CloseDB DB
End Sub

Private Sub ShowTable()
'Tell it to select some records
List1.Clear
DBHelper.SelectRecords DB, Combo1.Text

If DBHelper.isRecordSetEmpty(DB) = False Then
Label1.Caption = DBHelper.GetRecordCount(DB) & " Records Found"
    Do Until DB.RS.EOF = True
        List1.AddItem DB.RS.Fields(0).Value
    DB.RS.MoveNext
    Loop
End If

End Sub
