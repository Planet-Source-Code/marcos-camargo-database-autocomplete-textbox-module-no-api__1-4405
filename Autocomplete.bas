Attribute VB_Name = "Module1"
'********************************************************************************************

'This module makes that autocomplete thing in textboxes, you start typing
'and the textbox is completed with the rest of the word, just like IE 4.0
'This module does the search in a table, not from a listbox or other things.
'you will need to do 3 things before using the functions from this module
'1st Choose the textbox to do the autocomplete thing
'2nd On the OnKeyDown event of the textbox, call the CheckIsDelOrBack function
'3rd On the OnChangeEvent of the textbox, call the AutoComplete function
'If you have more than one autocomplete textbox on the same form, set the property keypreview
'of the form to true and call the function CheckIsDelOrBack on the keydown event
'of the form instead of calling it from the textbox keydown event.
'KNOWN BUGS - it has problems when there is a " ' " (apostrophe) in the table, because of the SQL statement.
'The function will do the search correctly until it finds the apostrophe, but after that, it gets all screwed up.
'I couldn't figure how to fix that out, if you find a solution for that, I would like to
'be informed, so I can fix that in my projects too!!
'                     *####*      mautheman@yahoo.com       *####*

'********************************************************************************************
Public IsDelOrBack As Boolean
Public Function AutoComplete(TheText As TextBox, TheDB As Database, TheTable As String, TheField As String) As Boolean
On Error Resume Next
'****************************************************************************************
'TheText is the textbox that will do the autocomplete thing
'TheField is the field from the table that has the information that will fill the textbox
'TheTable is the table where you will search for the information to fill the textbox
'TheDB is the database with the TheTable
'****************************************************************************************

Dim OldLen As Integer
Dim dsTemp As Recordset

AutoComplete = False
If Not TheText.Text = "" And IsDelOrBack = False Then

OldLen = Len(TheText.Text)
    Set dsTemp = TheDB.OpenRecordset("Select * from " & TheTable & " where " & TheField & " like '" & TheText.Text & "*'", dbOpenDynaset)
      If Err = 3075 Then
        'here we got a bug!!
      End If
         If Not dsTemp.RecordCount = 0 Then
            TheText.Text = dsTemp(TheField)
                If TheText.SelText = "" Then
                    TheText.SelStart = OldLen
                Else
                    TheText.SelStart = InStr(TheText.Text, TheText.SelText)
                End If
                    TheText.SelLength = Len(TheText.Text)
                    AutoComplete = True
                    
        End If
        
End If

End Function

Public Function CheckIsDelOrBack(TheKey As Integer) As Boolean
'TheKey is the KeyCode - all you gotta do is write CheckIsDelOrBack(KeyCode) on the KeyDown event...
    If TheKey = vbKeyBack Or TheKey = vbKeyDelete Then
        IsDelOrBack = True
        CheckIsDelOrBack = True
    Else
        IsDelOrBack = False
        CheckIsDelOrBack = False
    End If
End Function
