Attribute VB_Name = "Module1"
Function IsStringBetweenRows(StartRow As Integer, EndRow As Integer, ColNo As Integer, FindString As String) As Boolean

IsStringBetweenRows = False

For i = StartRow To EndRow
    If InStr(Worksheets(3).Cells(i, ColNo), "Journal") = 1 Then
        IsStringBetweenRows = True
    End If
Next i

End Function
Function IsContainedInArray(l As Integer, Arr As Variant) As Boolean

'Return True if Value contained in Array, return false if not

        IsContainedInArray = UBound(Filter(Arr, l)) > -1

End Function


Sub RingsJournalSort()

Dim ws As Worksheet
Application.DisplayAlerts = False
For Each ws In Worksheets
    If ws.Name <> "Original" Then ws.Delete
Next ws
Application.DisplayAlerts = True

'just for when i've run and want to run it again, save me going to excel window and deleting
'Application.DisplayAlerts = False 'Stop Excel from asking if I'm sure I want to delete
'Sheets("JnlList1").Delete
'Sheets("JnlList2").Delete
'Application.DisplayAlerts = True

Dim Sh1, Sh2, Sh3, Sh4 As Worksheet
Dim Str, Date1, msg, Jnl() As String
Dim CurrentJnl, FinalRow, JnlNo, z, y, x, a, b, c, d, i, j, k, m, n, newCode As Integer

'opening, naming and setting worksheets to variables in case sheets are moved or named etc
Set Sh1 = Sheets(1)
With ThisWorkbook
Set Sh2 = .Sheets.Add(after:=.Sheets(.Sheets.Count))
Sh2.Name = "JnlList1"
Set Sh3 = .Sheets.Add(after:=.Sheets(.Sheets.Count))
Sh3.Name = "JnlList2"
End With

'initialise variables
z = 1 'counter for rows in sheet 1 - Original
y = 0 'counter for rows in JnlList1
x = 0 'counter for rows in JnlList2
a = 1 'Column Counter - SOME ISSUES WITH THIS DON'T CHANGE

'Turn off screen updating to increase speed
Application.ScreenUpdating = False

'Loop checks number of rows for which first 3 digits are numbers
'Loop End Condition - checks to see if 3 consecutive rows are blank
Do While IsEmpty(Sh1.Cells(z, a)) = False And IsEmpty(Sh1.Cells(z + 1, a)) = False And IsEmpty(Sh1.Cells(z + 2, a)) = False
    
'extracts first 3 characters of a row and sets as variable - will check if these are numeric
    Str = Left(Sh1.Cells(z, a), 3)
    
'extracts first 8 characters of a row and sets as variable - will check if these are dates
    Date1 = Left(Sh1.Cells(z, a), 8)
            
'Find Cells which contain the word journal and extract the journal number which is set to JnlStr
    JnlStr = Mid(Sh1.Cells(z, a), InStr(Sh1.Cells(z, a), "Journal") + 6 + 4, 6)
 
    a = a + 1 'Shifts column printed out to the right by 1
    
'----Seperate the and print out "Journal No. XXX" to either sheet 1 and 2
'There happens to be a slightly different formatting for the sheets and I use this to differentiate
    If IsNumeric(JnlStr) = True And JnlStr > 100 Then
        x = x + 1
        Sh3.Cells(x, a) = "Journal No. " & JnlStr
    ElseIf IsNumeric(JnlStr) = False And IsNumeric(Replace$(JnlStr, ".", "")) = True Then
        y = y + 1
        Sh2.Cells(y, a) = "Journal No." & Replace$(JnlStr, ".", "")
    End If
    a = a - 1 'Shifts column printed out to back left by 1
            
'Filtering out rows that don't begin with a number and which contain Co. name
    If ((IsNumeric(Str) = True) Or (IsDate(Date1) = True)) And (InStr(Sh1.Cells(z, a), Left(Sh1.Cells(1, 1), 18)) = 0) Then
        
    'seperating remaining items between two worksheets using string length without spaces
    'NOTE the number beside the < should represent the longest string that isn't a Jnl with Dr and Cr
        If (Len(Replace$(Sh1.Cells(z, a), " ", "")) < 35) Or (IsDate(Date1) = True) Then
            y = y + 1

'----This If Statement copies over the a/c name and client a/c code and the total and moves the total into column 3
            If IsNumeric(Sh1.Cells(z, a)) = True Then
                a = a + 5 'just using this to shift which column output goes to
                Sh1.Cells(z, 1).Copy Sh2.Cells(y, a)
                a = a - 5 'returning general output to initial column
            Else
                Sh1.Cells(z, 1).Copy Sh2.Cells(y, a)
            End If
           
        Else
            x = x + 1
            Sh1.Cells(z, a).Copy Sh3.Cells(x, a)
        End If
                                             
    End If
    z = z + 1
Loop

c = 1
ReDim Preserve Jnl(c)
Jnl(0) = 100

'----Deleting Duplicate Journal No. Rows on Sheet 2
For b = 1 To y
    ReDim Preserve Jnl(c)
    
    If InStr(Sh2.Cells(b, a + 1), "Journal") Then
        Jnl(c) = Sh2.Cells(b, a + 1)
        If Jnl(c) = Jnl(c - 1) Then
            Sh2.Rows(b).Delete
        End If
        c = c + 1
    End If
Next

'----Deleting Duplicate Journal No. Rows on Sheet 3
For b = 1 To x
    ReDim Preserve Jnl(c)
    
    If InStr(Sh3.Cells(b, a + 1), "Journal") Then
        Jnl(c) = Sh3.Cells(b, a + 1)
        If Jnl(c) = Jnl(c - 1) Then
            Sh3.Rows(b).Delete
        End If
        c = c + 1
    End If
Next

'----Using text to columns in Sheet 2
For b = 1 To y
'Text to Columns for the longer lines of data
    If IsDate(Left(Sh2.Cells(b, a), 8)) Then
        Sh2.Select
        Cells(b, a).Select
        Selection.TextToColumns Destination:=Sh2.Cells(b, a), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 4), Array(8, 1), Array(23, 1), Array(71, 1), Array(85, 1)), _
        TrailingMinusNumbers:=True
'Text to columns for the Client Code and Name for Client Code so that it'll be easier to use just the Code as ref
   ElseIf IsEmpty(Sh2.Cells(b, a)) = False Then
        Sh2.Select
        Cells(b, a).Select
        Selection.TextToColumns Destination:=Sh2.Cells(b, a), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True '
    End If
Next

'----Formatting and Tidying of Sheet 2
Sh2.Columns(5).Delete 'Delete column of irrelevant info

'Autofit some columns
For b = 2 To 5
Sh2.Columns(b).AutoFit
Next

'Format a column of numbers
With Sh2.Columns(4)
.NumberFormat = "#,###;(#,###);0"
End With

'Format a column of numbers, made bold as they are totals
With Sh2.Columns(5)
.NumberFormat = "#,###;(#,###);0"
'.Font.FontStyle = "Bold" 'Decided not to have them bold
End With

'Borders on the column of numbers which are totals
For b = 1 To y
    If IsEmpty(Sh2.Cells(b, 5)) = False Then
        Sh2.Cells(b, 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Sh2.Cells(b, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
    End If
Next

'One of the columns goes very wide with autofit so just widening to sufficient size
Sh2.Columns(1).ColumnWidth = 9.29

'----Using text to columns in Sheet 3
For b = 1 To x
    If IsEmpty(Sh3.Cells(b, a)) = False Then
    Sh3.Select
    Cells(b, a).Select
    Selection.TextToColumns Destination:=Sh2.Cells(b, a), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(8, 1), Array(46, 1), Array(64, 1), Array(76, 1)), _
        TrailingMinusNumbers:=True
    End If
Next





'----Formatting and tidying of Sheet 3
'AutoFit the column width
For b = 1 To 5
Sh3.Columns(b).AutoFit
Next

'Format two columns of numbers
With Sh3.Columns(3)
    .NumberFormat = "#,###;(#,###);0"
    .Font.FontStyle = "Bold"
End With
With Sh3.Columns(4)
    .NumberFormat = "#,###;(#,###);0"
    .Font.FontStyle = "Bold"
End With


'Up to this point I've extracted and formatted each journal and the transactions within
'now I want to gather for each client code the journals making up the total in seperate sheets


Dim Sh As Worksheet
Dim Code() As Variant
Dim FindOutArray() As String

'----Create a new tab for each client account code from Sheet 2
'Loop to find the size of the array required
i = 0
For b = 1 To y
    If IsDate(Sh2.Cells(b, a)) = False And IsEmpty(Sh2.Cells(b, a)) = False Then
        i = i + 1
    End If
Next

ReDim Code(i)

d = 0
For b = 1 To y
    If IsDate(Sh2.Cells(b, a)) = False And IsEmpty(Sh2.Cells(b, a)) = False Then
        newCode = Sh2.Cells(b, a).Value
        If IsContainedInArray(newCode, Code) = False Then
            Code(d) = newCode
            With ThisWorkbook
            Set Sh = .Sheets.Add(after:=.Sheets(.Sheets.Count))
            Sh.Name = newCode
            End With
            d = d + 1
        End If
    End If
Next

'----Create a new tab for each client account code from Sheet 3
'Loop to find the size of the array required maintaining the same increment variable and array
For b = 1 To x
    If IsDate(Sh2.Cells(b, a)) = False And IsEmpty(Sh2.Cells(b, a)) = False Then
        i = i + 1
    End If
Next

'restating the new increased dimension of the array preserving it's contents
ReDim Preserve Code(i)

'the same loop as previously, keeping the original array to check for duplicates
For b = 1 To x
    If IsDate(Sh3.Cells(b, a)) = False And IsEmpty(Sh3.Cells(b, a)) = False And IsEmpty(Replace$(Sh3.Cells(b, a), " ", "")) Then
        newCode = Sh3.Cells(b, a).Value
        If IsContainedInArray(newCode, Code) = False Then
            Code(d) = newCode
            With ThisWorkbook
            Set Sh = .Sheets.Add(after:=.Sheets(.Sheets.Count))
            Sh.Name = newCode
            End With
            d = d + 1
        End If
    End If
Next

'----Tabs open for each client code
'----Next Code to fill, format and tidy these tabs from JnlList1 and JnlList2
'----The following is extracting from *JnlList1* 1)ClientCodes 2)Transactions 3)Jnl No.s

Dim JnlRow() As Variant
Dim CodeRow(), CodeRow2(), u, v(), w, LastRow As Integer
Dim ShName
ReDim CodeRow(i)

d = 0
'this creates an array containing the rows which have the client codes
For b = 1 To x + y
    If IsDate(Sh2.Cells(b, a)) = False And IsEmpty(Sh2.Cells(b, a)) = False Then
    ReDim Preserve CodeRow(d)
    CodeRow(d) = b
    d = d + 1
    End If
Next

'the Array v is for referencing to the sheets in this workbook
w = ThisWorkbook.Sheets.Count
ReDim v(w)

j = 1
'place row numbers for all "journal No." rows into an array
For b = 1 To y
    If InStr(Sh2.Cells(b, 2), "Journal") = 1 Then
        ReDim Preserve JnlRow(j)
        JnlRow(j) = b
        j = j + 1
    End If
Next b

'Cycle through the worksheets, match the sheet names to client codes, copy over the appropriate code rows and Journal No rows

For k = 1 To w

    v(k) = 1 'Array in order to retain last row printed on for each sheet
    
  'loop through the array holding the row numbers of the rows with the client codes
    For b = 0 To UBound(CodeRow)
        
    'excludes non numeric sheet names in order to ensure the correct sheets used
        If IsNumeric(Sheets(k).Name) Then
            ShName = CInt(Sheets(k).Name)
        Else
            ShName = 0
        End If
    'Check if the Client Code in a Row is equal to the current sheet looped onto
        If Sh2.Cells(CodeRow(b), a) = ShName Then
    
        'this deals with the final CodeRow(b+1) being outside of the array
            If b = UBound(CodeRow) Then
                With Sh2
                    LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
                End With
        'this finds the distance between the last row and the last code row to measure the amount of rows with transactions
        'add 1 for the row with the total -- this is in order to correctly place the total of the totals
                u = 1 + LastRow - CodeRow(UBound(CodeRow))
            Else
                u = CInt(CodeRow(b + 1)) - CInt(CodeRow(b))
            End If
         'This loop finds the row containing journal above a client code and copies it
         'pasting to the sheet with the same name, ignoring  journals after the code
         'and pastes them over each other until the last one which is just before the code
            For j = 1 To UBound(JnlRow)
                    If JnlRow(j) < CodeRow(b) Then
                        Sh2.Rows(JnlRow(j)).Copy
                        Sheets(k).Rows(v(k)).PasteSpecial xlPasteAll
                    End If
            Next j
            v(k) = v(k) + 2
        'this copies across the Client Code row plus the rows up to the next client code
        'this is in order to catch the transactions beneath the code row
        
            For j = 0 To u - 1
                If InStr(Sh2.Cells(CodeRow(b) + j, 2), "Journal") = 0 Then
                    Sh2.Rows(CodeRow(b) + j).Copy
                    Sheets(k).Rows(v(k)).PasteSpecial xlPasteAll
                    v(k) = v(k) + 1
                End If
            Next j
        v(k) = v(k) + 1
        End If
    Next b
    'autofit the columns in the client code sheets
    For b = 1 To 8
        Sheets(k).Columns(b).AutoFit
        If b = 5 Or b = 4 Then Sheets(k).Columns(b).ColumnWidth = 10
    Next b
    
'Putting a total of the totals moved over from Sheet 2
    v(k) = v(k) + 1
    For b = 1 To v(k) - 1
        Sheets(k).Cells(v(k), 5) = Sheets(k).Cells(v(k), 5) + Sheets(k).Cells(b, 5)
    Next b

'Formatting this one cell
    With Sheets(k).Cells(v(k), 5)
        .NumberFormat = "#,###;(#,###);0"
        .Font.FontStyle = "Bold"
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With

    v(k) = v(k) + 2
Next k

'----Successful extraction and formatting to each appropriate tab from JnlList1
'----Next Code to do same for *JnlList2*

'this creates an array containing the rows which have the client codes
d = 0
For b = 1 To x + y
    If IsEmpty(Sh3.Cells(b, a)) = False Then
    ReDim Preserve CodeRow2(d)
    CodeRow2(d) = b
    d = d + 1
    End If
Next

CurrentJnl = 1
j = 1
Dim RunningTotal, RunningTotal2 As Double
Dim CheckifAnythingPasted As Boolean

For k = 4 To w

    CheckifAnythingPasted = False
    RunningTotal = 0
    RunningTotal2 = 0
  'loop through the array holding the row numbers of the rows with the client codes
    For b = 1 To x
    'excludes non numeric sheet names in order to ensure the correct sheets used
        If IsNumeric(Sheets(k).Name) Then
            ShName = CInt(Sheets(k).Name)
        Else
            ShName = 0
        End If
        
        'Check if Cells are empty if so save row no. to a temp variable
        If IsEmpty(Sh3.Cells(b, a)) Then
            CurrentJnl = b
        End If
        
    'Check if the Client Code in a Row is equal to the current sheet looped onto
        If Sh3.Cells(b, a) = ShName Then
            CheckifAnythingPasted = True
    'if cell contains client code and is same as the current sheet then paste
    'first the Row containing the Journal No.
            If CInt(Sh3.Cells(b, a).Value) = ShName Then
                Sh3.Rows(CurrentJnl).Copy
                Sheets(k).Rows(v(k)).PasteSpecial xlPasteAll
                v(k) = v(k) + 2
            End If
    'second paste the row containing the client code and other information
            Sh3.Rows(b).Copy
            Sheets(k).Rows(v(k)).PasteSpecial xlPasteAll
            
            'These are just to allow me to print a total at the end of sheet
            RunningTotal = RunningTotal + Sheets(k).Cells(v(k), 3)
            RunningTotal2 = RunningTotal2 + Sheets(k).Cells(v(k), 4)

            v(k) = v(k) + 1
        End If
    Next b
    If CheckifAnythingPasted = True Then
        v(k) = v(k) + 1
        Sheets(k).Cells(v(k), 3) = RunningTotal
        Sheets(k).Cells(v(k), 4) = RunningTotal2
                    'formatting of the totals
            With Sheets(k).Cells(v(k), 3)
                .NumberFormat = "#,###;(#,###);0"
                .Font.FontStyle = "Bold"
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeTop).LineStyle = xlContinuous
            End With
            With Sheets(k).Cells(v(k), 4)
                .NumberFormat = "#,###;(#,###);0"
                .Font.FontStyle = "Bold"
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeTop).LineStyle = xlContinuous
            End With
    End If
    'autofit the columns in the client code sheets
    For b = 1 To 8
        Sheets(k).Columns(b).AutoFit
        If b = 5 Or b = 4 Then Sheets(k).Columns(b).ColumnWidth = 10
    Next b
    
Next k


'Loop to reorder the client code tabs in numeric order
For b = 4 To w
    For k = 4 To w - 1
        If CInt(Sheets(k).Name) > CInt(Sheets(k + 1).Name) Then Sheets(k).Move after:=Sheets(k + 1)
    Next k
Next b

'find the last row of the first sheet
With Sheets(1)
    FinalRow = .Cells(.Rows.Count, 1).End(xlUp).Row
End With
x = 1
'The idea of this whole loop is to first cycle through the sheets
'then search through each sheet for a cell containing the word "Journal"
'then set that entire cell's string as JnlStr
'Search through Sheet1/Original for Jnlno
'Loop up through the rows above until finding the cell which contains either sales or purchases ledger transfer
'blank row above original jnlno, set it equal to the row containing Sales or Purchases Ledger Transfer
For k = 4 To w
'k = 4

    With Sheets(k)
       LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
    End With

    For b = 1 To LastRow
        If InStr(Sheets(k).Cells(b, 2), "Journal") = 1 Then
            'JnlNo = CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", ""))
            
  ' i = 59 for every loop, this might mean that i is 59 coming into the this full loop and that it doesn't get through the if as
  ' it might be taking JnlNo to literally mean JnlNo instead of what it represents
            
            For j = 1 To FinalRow
                If InStr(Sheets(1).Cells(j, 1), "Journal") > 0 And InStr(Sheets(1).Cells(j, 1), CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", ""))) > 0 Then
                'If InStr(Sheets(1).Cells(j, 1), CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", ""))) > 0 Then
                'If InStr(Sheets(1).Cells(j, 1), Journal) = 1 Then
                    i = 1
                    i = j
                    'Debug.Print (i)
                    
                    For i = j To j - 70 Step -1
                        If i = 0 Then
                            Exit For
                        ElseIf InStr(Sheets(1).Cells(i, 1), "Purchase Ledger Transfer Report") > 0 Then
                            Sheets(k).Cells(b + 1, 1) = "Purchase Ledger Transfer Report"
                            'Debug.Print ("Purchase Ledger")
                            'If CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", "")) = 241 Then Debug.Print (j)
                            Exit For
                        ElseIf InStr(Sheets(1).Cells(i, 1), "Sales Ledger Transfer Report") > 0 Then
                            Sheets(k).Cells(b + 1, 1) = "Sales Ledger Transfer Report"
                            'Debug.Print ("Sales Ledger")
                            'If CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", "")) = 241 Then Debug.Print ("In Sales")
                            Exit For
                        ElseIf InStr(Sheets(1).Cells(i, 1), "Cash Book Transfer Report") > 0 Then
                            Sheets(k).Cells(b + 1, 1) = "Cash Book Transfer Report"
                            'Debug.Print ("Cash Book")
                            Exit For
                        End If
                    Next i
                    
                End If
            Next j
    
            'Sheets(k).Cells(b + 1, 1) = Sheets(1).Cells(i, 1)
            'Debug.Print (x)
            'Debug.Print (CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", "")))
            
        End If
    Next b
Next k


'Delete Working Sheets that may not be necessary but are useful to hang onto in case they are wanted or useful
Application.DisplayAlerts = False
'Sheets(2).Delete
'Sheets(2).Delete
Application.DisplayAlerts = True

'Reactive Screen Updating
Application.ScreenUpdating = True

'For b = 1 To FinalRow

'If InStr(Sheets(1).Cells(b, 1), "Purchase Ledger Transfer Report") > 0 Then Debug.Print (b)
'If InStr(Sheets(1).Cells(b, 1), "Sales Ledger Transfer Report") > 0 Then Debug.Print (b)
'Next b

MsgBox ("Completed :)")
'MsgBox (InStr(Sheets(1).Cells(671, 1), "Journal"))

End Sub



