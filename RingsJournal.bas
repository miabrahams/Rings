Dim PerfCounter As New PerformanceCounter


Function IsStringBetweenRows(StartRow As Integer, EndRow As Integer, ColNo As Integer, FindString As String) As Boolean
    IsStringBetweenRows = False
    For i = StartRow To EndRow
        If InStr(Worksheets(3).Cells(i, ColNo), "Journal") = 1 Then
            IsStringBetweenRows = True
        End If
    Next i
End Function

Function IsContainedInArray(l As Integer, Arr As Variant) As Boolean
    IsContainedInArray = UBound(Filter(Arr, l)) > -1
End Function

Function DBG_ElapsedTime(Optional additionalMessage As String = "")
    If Len(additionalMessage) > 0 Then
        Debug.Print (additionalMessage)
    End If
    Debug.Print ("Elapsed Time: " & Format(PerfCounter.TimeElapsed / 1000, "#.00") & " seconds" + vbCrLf)
    PerfCounter.StartCounter
End Function


Sub a_Test()
    PerfCounter.StartCounter
    MsgBox ("This is a test")
    DBG_ElapsedTime "PerfCounterTest"
End Sub






Sub RingsJournalSort()

    PerfCounter.StartCounter

    'Remove every sheet but the first
    Application.DisplayAlerts = False
    Do While Worksheets.Count > 1
        Worksheets(2).Delete
    Loop
    Application.DisplayAlerts = True


    Application.ScreenUpdating = False     'Faster

    
    ' Allocate variables
    Dim OrigSheet, JournalSheet1, JournalSheet2 As Worksheet
    Dim TempStr, TempDate, TempJnl() As String
    Dim CurrentJnl, JnlNo, z, numJournals1, numJournals2, currentCol, b, c, d, numClientCodes1, j, k, m, n, newCode As Integer
    
    
    'opening, naming and setting worksheets to variables in case sheets are moved or named etc
    Set OrigSheet = Sheets(1)
    With ThisWorkbook
        Set JournalSheet1 = .Sheets.Add(after:=.Sheets(.Sheets.Count))
        JournalSheet1.Name = "JnlList1"
        Set JournalSheet2 = .Sheets.Add(after:=.Sheets(.Sheets.Count))
        JournalSheet2.Name = "JnlList2"
    End With
    
    'initialise variables
    z = 1 'counter for rows in sheet 1 - Original
    numJournals1 = 0 'counter for rows in JnlList1
    numJournals2 = 0 'counter for rows in JnlList2
    currentCol = 1 'Column Counter - SOME ISSUES WITH THIS DON'T CHANGE
    
    
    '----Seperate the and print out "Journal No. XXX" to either sheet 1 and 2
    ' Sheet 1 is for "long" format journals, Sheet 2 for "short" format
    'Loop End Condition - see if 3 consecutive rows are blank
    Do While IsEmpty(OrigSheet.Cells(z, currentCol)) = False And IsEmpty(OrigSheet.Cells(z + 1, currentCol)) = False And IsEmpty(OrigSheet.Cells(z + 2, currentCol)) = False

        'Loop checks number of rows for which first 3 digits are numbers
        'extracts first 3 characters of a row and sets as variable - will check if these are numeric
        TempStr = Left(OrigSheet.Cells(z, currentCol), 3)
        
        'extracts first 8 characters of a row and sets as variable - will check if these are dates
        TempDate = Left(OrigSheet.Cells(z, currentCol), 8)
                
        'Find Cells which contain the word journal and extract the journal number which is set to JnlStr
        JnlStr = Mid(OrigSheet.Cells(z, currentCol), InStr(OrigSheet.Cells(z, currentCol), "Journal") + 6 + 4, 6)
     
        currentCol = currentCol + 1 'Shifts column printed out to the right by 1
        
        'There happens to be a slightly different formatting for the sheets and I use this to differentiate
        If IsNumeric(JnlStr) = True And JnlStr > 100 Then
            numJournals2 = numJournals2 + 1
            JournalSheet2.Cells(numJournals2, currentCol) = "Journal No. " & JnlStr
        ElseIf IsNumeric(JnlStr) = False And IsNumeric(Replace$(JnlStr, ".", "")) = True Then
            numJournals1 = numJournals1 + 1
            JournalSheet1.Cells(numJournals1, currentCol) = "Journal No." & Replace$(JnlStr, ".", "")
        End If
        currentCol = currentCol - 1 'Shifts column printed out to back left by 1
                
        'Filtering out rows that don't begin with a number and which contain Co. name
        If ((IsNumeric(TempStr) = True) Or (IsDate(TempDate) = True)) And (InStr(OrigSheet.Cells(z, currentCol), Left(OrigSheet.Cells(1, 1), 18)) = 0) Then
            
            'seperating remaining items between two worksheets using string length without spaces
            'NOTE the number beside the < should represent the longest string that isn't a Jnl with Dr and Cr
            If (Len(Replace$(OrigSheet.Cells(z, currentCol), " ", "")) < 35) Or (IsDate(TempDate) = True) Then
                numJournals1 = numJournals1 + 1

                '----This If Statement copies over the a/c name and client a/c code and the total and moves the total into column 3
                If IsNumeric(OrigSheet.Cells(z, currentCol)) = True Then
                    currentCol = currentCol + 5 'just using this to shift which column output goes to
                    OrigSheet.Cells(z, 1).Copy JournalSheet1.Cells(numJournals1, currentCol)
                    currentCol = currentCol - 5 'returning general output to initial column
                Else
                    OrigSheet.Cells(z, 1).Copy JournalSheet1.Cells(numJournals1, currentCol)
                End If

            Else
                numJournals2 = numJournals2 + 1
                OrigSheet.Cells(z, currentCol).Copy JournalSheet2.Cells(numJournals2, currentCol)
            End If
                                                 
        End If
        z = z + 1
    Loop
    
    DBG_ElapsedTime "Searched for journal entry numbers"
    
    
    c = 1
    ReDim Preserve TempJnl(c)
    TempJnl(0) = 100

    '----Deleting Duplicate Journal No. Rows on Sheet 2
    For b = 1 To numJournals1
        ReDim Preserve TempJnl(c)

        If InStr(JournalSheet1.Cells(b, currentCol + 1), "Journal") Then
            TempJnl(c) = JournalSheet1.Cells(b, currentCol + 1)
            If TempJnl(c) = TempJnl(c - 1) Then
                JournalSheet1.Rows(b).Delete
            End If
            c = c + 1
        End If
    Next
    
    '----Deleting Duplicate Journal No. Rows on Sheet 3
    For b = 1 To numJournals2
        ReDim Preserve TempJnl(c)
        
        If InStr(JournalSheet2.Cells(b, currentCol + 1), "Journal") Then
            TempJnl(c) = JournalSheet2.Cells(b, currentCol + 1)
            If TempJnl(c) = TempJnl(c - 1) Then
                JournalSheet2.Rows(b).Delete
            End If
            c = c + 1
        End If
    Next
    
    DBG_ElapsedTime "Deleted Duplicate Journals"
    
    '----Using text to columns in Sheet 2
    For b = 1 To numJournals1
        'Text to Columns for the longer lines of data
        If IsDate(Left(JournalSheet1.Cells(b, currentCol), 8)) Then
            JournalSheet1.Select
            Cells(b, currentCol).Select
            Selection.TextToColumns Destination:=JournalSheet1.Cells(b, currentCol), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 4), Array(8, 1), Array(23, 1), Array(71, 1), Array(85, 1)), _
            TrailingMinusNumbers:=True
        'Text to columns for the Client Code and Name for Client Code so that it'll be easier to use just the Code as ref
        ElseIf IsEmpty(JournalSheet1.Cells(b, currentCol)) = False Then
            JournalSheet1.Select
            Cells(b, currentCol).Select
            Selection.TextToColumns Destination:=JournalSheet1.Cells(b, currentCol), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(3, 1)), _
            TrailingMinusNumbers:=True '
        End If
    Next
    
    
    '----Formatting and Tidying of Sheet 2
    JournalSheet1.Columns(5).Delete 'Delete column of irrelevant info
    
    'Autofit some columns
    For b = 2 To 5
    JournalSheet1.Columns(b).AutoFit
    Next
    
    'Format a column of numbers
    With JournalSheet1.Columns(4)
    .NumberFormat = "#,###;(#,###);0"
    End With
    
    'Format a column of numbers, made bold as they are totals
    With JournalSheet1.Columns(5)
    .NumberFormat = "#,###;(#,###);0"
    '.Font.FontStyle = "Bold" 'Decided not to have them bold
    End With
    
    'Borders on the column of numbers which are totals
    For b = 1 To numJournals1
        If IsEmpty(JournalSheet1.Cells(b, 5)) = False Then
            JournalSheet1.Cells(b, 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
            JournalSheet1.Cells(b, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
    Next
    
    'One of the columns goes very wide with autofit so just widening to sufficient size
    JournalSheet1.Columns(1).ColumnWidth = 9.29
    
    '----Using text to columns in Sheet 3
    For b = 1 To numJournals2
        If IsEmpty(JournalSheet2.Cells(b, currentCol)) = False Then
        JournalSheet2.Select
        Cells(b, currentCol).Select
        Selection.TextToColumns Destination:=JournalSheet1.Cells(b, currentCol), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(8, 1), Array(46, 1), Array(64, 1), Array(76, 1)), _
            TrailingMinusNumbers:=True
        End If
    Next
    
    
    '----Formatting and tidying of Sheet 3
    'AutoFit the column width
    For b = 1 To 5
    JournalSheet2.Columns(b).AutoFit
    Next
    
    'Format two columns of numbers
    With JournalSheet2.Columns(3)
        .NumberFormat = "#,###;(#,###);0"
        .Font.FontStyle = "Bold"
    End With
    With JournalSheet2.Columns(4)
        .NumberFormat = "#,###;(#,###);0"
        .Font.FontStyle = "Bold"
    End With
    
    DBG_ElapsedTime "Text to Columns"
    
    
    'Up to this point I've extracted and formatted each journal and the transactions within
    'now I want to gather for each client code the journals making up the total in seperate sheets
    
    
    Dim Sh As Worksheet
    Dim Code() As Variant
    Dim FindOutArray() As String
    
    '----Create a new tab for each client account code from Sheet 2
    'Loop to find the size of the array required
    numClientCodes1 = 0
    For b = 1 To numJournals1
        If IsDate(JournalSheet1.Cells(b, currentCol)) = False And IsEmpty(JournalSheet1.Cells(b, currentCol)) = False Then
            numClientCodes1 = numClientCodes1 + 1
        End If
    Next
    
    ReDim Code(numClientCodes1)
    
    d = 0
    For b = 1 To numJournals1
        If IsDate(JournalSheet1.Cells(b, currentCol)) = False And IsEmpty(JournalSheet1.Cells(b, currentCol)) = False Then
            newCode = JournalSheet1.Cells(b, currentCol).Value
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
    For b = 1 To numJournals2
        If IsDate(JournalSheet1.Cells(b, currentCol)) = False And IsEmpty(JournalSheet1.Cells(b, currentCol)) = False Then
            numClientCodes1 = numClientCodes1 + 1
        End If
    Next
    
    'restating the new increased dimension of the array preserving it's contents
    ReDim Preserve Code(numClientCodes1)
    
    'the same loop as previously, keeping the original array to check for duplicates
    For b = 1 To numJournals2
        If IsDate(JournalSheet2.Cells(b, currentCol)) = False And IsEmpty(JournalSheet2.Cells(b, currentCol)) = False And IsEmpty(Replace$(JournalSheet2.Cells(b, currentCol), " ", "")) Then
            newCode = JournalSheet2.Cells(b, currentCol).Value
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
    
    DBG_ElapsedTime "Created Client Code Sheets"
    
    '----Tabs open for each client code
    '----Next Code to fill, format and tidy these tabs from JnlList1 and JnlList2
    '----The following is extracting from *JnlList1* 1)ClientCodes 2)Transactions 3)Jnl No.s
    
    Dim JnlRow() As Variant
    Dim CodeRow(), CodeRow2(), u, LastRow As Integer
    Dim ShName
    ReDim CodeRow(numClientCodes1)
    
    d = 0
    'this creates an array containing the rows which have the client codes
    For b = 1 To numJournals2 + numJournals1
        If IsDate(JournalSheet1.Cells(b, currentCol)) = False And IsEmpty(JournalSheet1.Cells(b, currentCol)) = False Then
        ReDim Preserve CodeRow(d)
        CodeRow(d) = b
        d = d + 1
        End If
    Next
    
    'the Array v is for referencing to the sheets in this workbook
    Dim SheetLength() As Integer
    ReDim SheetLength(ThisWorkbook.Sheets.Count)
    
    j = 1
    'place row numbers for all "journal No." rows into an array
    For b = 1 To numJournals1
        If InStr(JournalSheet1.Cells(b, 2), "Journal") = 1 Then
            ReDim Preserve JnlRow(j)
            JnlRow(j) = b
            j = j + 1
        End If
    Next b
    
    DBG_ElapsedTime "Cleaned Client Code Sheets"
    
    'Cycle through the worksheets, match the sheet names to client codes, copy over the appropriate code rows and Journal No rows
    
    For k = 4 To ThisWorkbook.Sheets.Count
        
        ShName = CInt(Sheets(k).Name)
        SheetLength(k) = 1 'Array in order to retain last row printed on for each sheet
        
        'loop through the array holding the row numbers of the rows with the client codes
        For b = 0 To UBound(CodeRow)

            'Check if the Client Code in a Row is equal to the current sheet looped onto
            If JournalSheet1.Cells(CodeRow(b), currentCol) = ShName Then
        
                'this deals with the final CodeRow(b+1) being outside of the array
                If b = UBound(CodeRow) Then
                    With JournalSheet1
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
                            JournalSheet1.Rows(JnlRow(j)).Copy
                            Sheets(k).Rows(SheetLength(k)).PasteSpecial xlPasteAll
                        End If
                Next j
                SheetLength(k) = SheetLength(k) + 2
                'this copies across the Client Code row plus the rows up to the next client code
                'this is in order to catch the transactions beneath the code row
            
                For j = 0 To u - 1
                    If InStr(JournalSheet1.Cells(CodeRow(b) + j, 2), "Journal") = 0 Then
                        JournalSheet1.Rows(CodeRow(b) + j).Copy
                        Sheets(k).Rows(SheetLength(k)).PasteSpecial xlPasteAll
                        SheetLength(k) = SheetLength(k) + 1
                    End If
                Next j
            SheetLength(k) = SheetLength(k) + 1
            End If
        Next b
        'autofit the columns in the client code sheets
        For b = 1 To 8
            Sheets(k).Columns(b).AutoFit
            If b = 5 Or b = 4 Then Sheets(k).Columns(b).ColumnWidth = 10
        Next b
        
        'Putting a total of the totals moved over from Sheet 2
        SheetLength(k) = SheetLength(k) + 1
        For b = 1 To SheetLength(k) - 1
            Sheets(k).Cells(SheetLength(k), 5) = Sheets(k).Cells(SheetLength(k), 5) + Sheets(k).Cells(b, 5)
        Next b
    
        'Formatting this one cell
        With Sheets(k).Cells(SheetLength(k), 5)
            .NumberFormat = "#,###;(#,###);0"
            .Font.FontStyle = "Bold"
            .Borders(xlEdgeBottom).LineStyle = xlDouble
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
    
        SheetLength(k) = SheetLength(k) + 2
    Next k
    
    DBG_ElapsedTime "Matched Journal 1 Client Data to Sheets"
    
    '----Successful extraction and formatting to each appropriate tab from JnlList1
    '----Next Code to do same for *JnlList2*
    
    'this creates an array containing the rows which have the client codes
    d = 0
    For b = 1 To numJournals2 + numJournals1
        If IsEmpty(JournalSheet2.Cells(b, currentCol)) = False Then
        ReDim Preserve CodeRow2(d)
        CodeRow2(d) = b
        d = d + 1
        End If
    Next
    
    CurrentJnl = 1
    j = 1
    Dim RunningTotal, RunningTotal2 As Double
    Dim CheckifAnythingPasted As Boolean
    
    For k = 4 To ThisWorkbook.Sheets.Count
    
        CheckifAnythingPasted = False
        RunningTotal = 0
        RunningTotal2 = 0
        'loop through the array holding the row numbers of the rows with the client codes
        For b = 4 To numJournals2
            ShName = CInt(Sheets(k).Name)

            'Check if Cells are empty if so save row no. to a temp variable
            If IsEmpty(JournalSheet2.Cells(b, currentCol)) Then
                CurrentJnl = b
            End If
            
            'Check if the Client Code in a Row is equal to the current sheet looped onto
            If JournalSheet2.Cells(b, currentCol) = ShName Then
                CheckifAnythingPasted = True
                'if cell contains client code and is same as the current sheet then paste
                'first the Row containing the Journal No.
                If CInt(JournalSheet2.Cells(b, currentCol).Value) = ShName Then
                    JournalSheet2.Rows(CurrentJnl).Copy
                    Sheets(k).Rows(SheetLength(k)).PasteSpecial xlPasteAll
                    SheetLength(k) = SheetLength(k) + 2
                End If
                'second paste the row containing the client code and other information
                JournalSheet2.Rows(b).Copy
                Sheets(k).Rows(SheetLength(k)).PasteSpecial xlPasteAll
                
                'These are just to allow me to print a total at the end of sheet
                RunningTotal = RunningTotal + Sheets(k).Cells(SheetLength(k), 3)
                RunningTotal2 = RunningTotal2 + Sheets(k).Cells(SheetLength(k), 4)
    
                SheetLength(k) = SheetLength(k) + 1
            End If
        Next b
        If CheckifAnythingPasted = True Then
            SheetLength(k) = SheetLength(k) + 1
            Sheets(k).Cells(SheetLength(k), 3) = RunningTotal
            Sheets(k).Cells(SheetLength(k), 4) = RunningTotal2
            'formatting of the totals
            With Sheets(k).Cells(SheetLength(k), 3)
                .NumberFormat = "#,###;(#,###);0"
                .Font.FontStyle = "Bold"
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeTop).LineStyle = xlContinuous
            End With
            With Sheets(k).Cells(SheetLength(k), 4)
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
    
    DBG_ElapsedTime "Matched Journal 2 Client Data to Sheets"
    
    'Loop to reorder the client code tabs in numeric order
    For b = 4 To ThisWorkbook.Sheets.Count
        For k = 4 To ThisWorkbook.Sheets.Count - 1
            If CInt(Sheets(k).Name) > CInt(Sheets(k + 1).Name) Then Sheets(k).Move after:=Sheets(k + 1)
        Next k
    Next b
    
    DBG_ElapsedTime "Reordered Client Code Tabs"
    
    'find the last row of the first sheet
    Dim FinalRow As Integer
    With Sheets(1)
        FinalRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With

    numJournals2 = 1
    'The idea of this whole loop is to first cycle through the sheets
    'then search through each sheet for a cell containing the word "Journal"
    'then set that entire cell's string as JnlStr
    'Search through Sheet1/Original for Jnlno
    'Loop up through the rows above until finding the cell which contains either sales or purchases ledger transfer
    'blank row above original jnlno, set it equal to the row containing Sales or Purchases Ledger Transfer
    For k = 4 To ThisWorkbook.Sheets.Count
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
                        numClientCodes1 = 1
                        numClientCodes1 = j
                        'Debug.Print (i)
                        
                        For numClientCodes1 = j To j - 70 Step -1
                            If numClientCodes1 = 0 Then
                                Exit For
                            ElseIf InStr(Sheets(1).Cells(numClientCodes1, 1), "Purchase Ledger Transfer Report") > 0 Then
                                Sheets(k).Cells(b + 1, 1) = "Purchase Ledger Transfer Report"
                                'Debug.Print ("Purchase Ledger")
                                'If CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", "")) = 241 Then Debug.Print (j)
                                Exit For
                            ElseIf InStr(Sheets(1).Cells(numClientCodes1, 1), "Sales Ledger Transfer Report") > 0 Then
                                Sheets(k).Cells(b + 1, 1) = "Sales Ledger Transfer Report"
                                'Debug.Print ("Sales Ledger")
                                'If CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", "")) = 241 Then Debug.Print ("In Sales")
                                Exit For
                            ElseIf InStr(Sheets(1).Cells(numClientCodes1, 1), "Cash Book Transfer Report") > 0 Then
                                Sheets(k).Cells(b + 1, 1) = "Cash Book Transfer Report"
                                'Debug.Print ("Cash Book")
                                Exit For
                            End If
                        Next numClientCodes1
                        
                    End If
                Next j
        
                'Sheets(k).Cells(b + 1, 1) = Sheets(1).Cells(i, 1)
                'Debug.Print (x)
                'Debug.Print (CInt(Replace$(Replace$(Right(Sheets(k).Cells(b, 2), 5), " ", ""), ".", "")))
                
            End If
        Next b
    Next k
    
    
    DBG_ElapsedTime "Completed Final Reverse Lookup"

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


