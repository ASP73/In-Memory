Attribute VB_Name = "test"
Sub ProtectFormSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Form")

    ' Remove any existing protection
    ws.Unprotect

    ' Hide scroll bars
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False

    ' Remove any split Or frozen panes
    ActiveWindow.FreezePanes = False
    ActiveWindow.Split = False

    ' Set the view To the form area
    ws.Range("A1").Select
End Sub

Sub AddPlaceholders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Form")

    ' Add placeholder values
    ws.Range("F6").Value = "Deceased person's name"
    ws.Range("F8").Value = "Enter contact name"
    ws.Range("F10").Value = "Enter contact phone"
    ws.Range("F12").Value = "Enter contact ID"
    ws.Range("F16").Value = "Enter leaves cost"
    ws.Range("F18").Value = "Enter leaves price"
    ws.Range("F22").Value = "Enter trees cost"
    ws.Range("F24").Value = "Enter trees price"
    ws.Range("F28").Value = "Enter plaque cost"
    ws.Range("F30").Value = "Enter plaque price"
    ws.Range("F34").Value = "Enter sandwiches cost"
    ws.Range("F36").Value = "Enter sandwiches price"

    ' Set font color To light gray For placeholders
    ws.Range("F6,F8,F10,F12,F16,F18,F22,F24,F28,F30,F34,F36").Font.Color = RGB(192, 192, 192)
End Sub

Sub ChangeTextColorOnInput()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Form")

    Dim cell As Range
    For Each cell In ws.Range("F6,F8,F10,F12,F16,F18,F22,F24,F28,F30,F34,F36")
        If cell.Value <> "" And cell.Value <> cell.Comment.Text Then
            cell.Font.Color = RGB(0, 0, 0) ' Black color
        End If
    Next cell
End Sub

Sub AddDataToTable()
    Dim formSheet As Worksheet
    Dim tableSheet As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow

    ' Set the worksheet variables
    Set formSheet = ThisWorkbook.Sheets("Form") ' Change "Form" To your actual sheet name If different
    Set tableSheet = ThisWorkbook.Sheets("Table") ' Change "Table" To your actual sheet name If different

    ' Set the table
    Set tbl = tableSheet.ListObjects("Table7") ' Change "Table1" To your actual table name

    ' Check If required fields are filled And Not placeholders
    If IsEmpty(formSheet.Range("F6")) Or formSheet.Range("F6").Value = "Deceased person's name" Or _
        IsEmpty(formSheet.Range("F8")) Or formSheet.Range("F8").Value = "Enter contact name" Or _
        IsEmpty(formSheet.Range("F10")) Or formSheet.Range("F10").Value = "Enter contact phone" Or _
        IsEmpty(formSheet.Range("F12")) Or formSheet.Range("F12").Value = "Enter contact ID" Then
        MsgBox "Please fill in all required fields With valid data.", vbExclamation
     Exit Sub
    End If

    ' Check If at least one sell price is entered And Not a placeholder
    If (IsEmpty(formSheet.Range("F18")) Or formSheet.Range("F18").Value = "Enter leaves price") And _
        (IsEmpty(formSheet.Range("F24")) Or formSheet.Range("F24").Value = "Enter trees price") And _
        (IsEmpty(formSheet.Range("F30")) Or formSheet.Range("F30").Value = "Enter plaque price") And _
        (IsEmpty(formSheet.Range("F36")) Or formSheet.Range("F36").Value = "Enter sarnies price") Then
        MsgBox "Please enter at least one valid sell price.", vbExclamation
     Exit Sub
    End If
    ' Add a New row To the table
    Set newRow = tbl.ListRows.Add

    ' Transfer data from cell F6 on the Form sheet To column A in the New row of the table
    newRow.Range(1, 1).Value = formSheet.Range("F6").Value

    ' Transfer data from cell F8 on the Form sheet To column B in the New row of the table
    newRow.Range(1, 2).Value = formSheet.Range("F8").Value

    ' Transfer data from cell F10 on the Form sheet To column C in the New row of the table
    newRow.Range(1, 3).Value = formSheet.Range("F10").Value

    ' Transfer data from cell F12 on the Form sheet To column D in the New row of the table
    newRow.Range(1, 4).Value = formSheet.Range("F12").Value
    'Leaves
    ' Transfer data from cell F16 on the Form sheet To column E in the New row of the table
    newRow.Range(1, 5).Value = formSheet.Range("F16").Value

    ' Transfer data from cell F18 on the Form sheet To column F in the New row of the table
    newRow.Range(1, 6).Value = formSheet.Range("F18").Value
    'Trees
    ' Transfer data from cell F22 on the Form sheet To column I in the New row of the table
    newRow.Range(1, 9).Value = formSheet.Range("F22").Value

    ' Transfer data from cell F24 on the Form sheet To column J in the New row of the table
    newRow.Range(1, 10).Value = formSheet.Range("F24").Value
    'Plaque
    ' Transfer data from cell F28 on the Form sheet To column M in the New row of the table
    newRow.Range(1, 13).Value = formSheet.Range("F28").Value

    ' Transfer data from cell F30 on the Form sheet To column N in the New row of the table
    newRow.Range(1, 14).Value = formSheet.Range("F30").Value
    'Sarnies
    ' Transfer data from cell F34 on the Form sheet To column Q in the New row of the table
    newRow.Range(1, 17).Value = formSheet.Range("F34").Value

    ' Transfer data from cell F36 on the Form sheet To column R in the New row of the table
    newRow.Range(1, 18).Value = formSheet.Range("F36").Value

    Dim rowOffset As Long
    rowOffset = 5 ' Assuming data starts from row 6

    'Leaves profit in £
    ' Calculate And Set the profit in column G
    newRow.Range(1, 7).Formula = "=F" & (newRow.Index + rowOffset) & "-E" & (newRow.Index + rowOffset)
    'Trees profit in £
    ' Calculate And Set the profit in column K
    newRow.Range(1, 11).Formula = "=J" & (newRow.Index + rowOffset) & "-I" & (newRow.Index + rowOffset)
    'Plaque profit in £
    ' Calculate And Set the profit in column O
    newRow.Range(1, 15).Formula = "=N" & (newRow.Index + rowOffset) & "-M" & (newRow.Index + rowOffset)
    'Sarnies profit in £
    ' Calculate And Set the profit in column S
    newRow.Range(1, 19).Formula = "=R" & (newRow.Index + rowOffset) & "-Q" & (newRow.Index + rowOffset)

    'Leaves mark-up percentage
    newRow.Range(1, 8).Formula = "=If(E" & (newRow.Index + rowOffset) & "<>0, (F" & (newRow.Index + rowOffset) & "-E" & (newRow.Index + rowOffset) & ")/E" & (newRow.Index + rowOffset) & ", 0)"
    newRow.Range(1, 8).NumberFormat = "0%"

    'Trees mark-up percentage
    newRow.Range(1, 12).Formula = "=If(I" & (newRow.Index + rowOffset) & "<>0, (J" & (newRow.Index + rowOffset) & "-I" & (newRow.Index + rowOffset) & ")/I" & (newRow.Index + rowOffset) & ", 0)"
    newRow.Range(1, 12).NumberFormat = "0%"

    'Plaque mark-up percentage
    newRow.Range(1, 16).Formula = "=If(M" & (newRow.Index + rowOffset) & "<>0, (N" & (newRow.Index + rowOffset) & "-M" & (newRow.Index + rowOffset) & ")/M" & (newRow.Index + rowOffset) & ", 0)"
    newRow.Range(1, 16).NumberFormat = "0%"

    'Sarnies mark-up percentage
    newRow.Range(1, 20).Formula = "=If(Q" & (newRow.Index + rowOffset) & "<>0, (R" & (newRow.Index + rowOffset) & "-Q" & (newRow.Index + rowOffset) & ")/Q" & (newRow.Index + rowOffset) & ", 0)"
    newRow.Range(1, 20).NumberFormat = "0%"    

    'Sum of all cost columns in column U
    newRow.Range(1, 21).Formula = "=SUM(E" & (newRow.Index + rowOffset) & ",I" & (newRow.Index + rowOffset) & ",M" & (newRow.Index + rowOffset) & ",Q" & (newRow.Index + rowOffset) & ")"

    'Sum of all sell columns in column V
    newRow.Range(1, 22).Formula = "=SUM(F" & (newRow.Index + rowOffset) & ",J" & (newRow.Index + rowOffset) & ",N" & (newRow.Index + rowOffset) & ",R" & (newRow.Index + rowOffset) & ")"

    'Total profit in column W
    newRow.Range(1, 23).Formula = "=V" & (newRow.Index + rowOffset) & "-U" & (newRow.Index + rowOffset)

    'Overall mark-up percentage in column X
    newRow.Range(1, 24).Formula = "=If(U" & (newRow.Index + rowOffset) & "<>0, (V" & (newRow.Index + rowOffset) & "-U" & (newRow.Index + rowOffset) & ")/U" & (newRow.Index + rowOffset) & ", 0)"
    newRow.Range(1, 24).NumberFormat = "0%"

    'Add datestamp in column Y
    newRow.Range(1, 25).Value = Now()
    newRow.Range(1, 25).NumberFormat = "dd/mm/yyyy"

    ' Clear the form after successful entry
    formSheet.Range("F6,F8,F10,F12,F16,F18,F22,F24,F28,F30,F34,F36").ClearContents

    MsgBox "Data added To the table successfully!"
End Sub

