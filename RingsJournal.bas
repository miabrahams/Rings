Public PerfCounter As New PerformanceCounter


Function IsStringBetweenRows(StartRow As Integer, EndRow As Integer, ColNo As Integer, FindString As String) As Boolean
    IsStringBetweenRows = False
    For i = StartRow To EndRow
        If InStr(Worksheets(3).Cells(i, ColNo), "Journal") = 1 Then
            IsStringBetweenRows = True
        End If
    Next i
End Function

Function IsContainedInArray(target As Variant, Arr As Variant) As Boolean
    IsContainedInArray = False
    For Each a In Arr
        IsContainedInArray = IsContainedInArray Or (a = target)
    Next
End Function

Function pprintMS(ByRef milliseconds As Double) As String
    pprintMS = Format(milliseconds / 1000, "#0.00")
End Function

Function stripSpaces(ByVal S As String)
    stripSpaces = Replace$(S, " ", "")
End Function

Function strippedToNum(ByVal S As String) As Long
    strippedToNum = CLng(Replace$(stripSpaces(S), ".", ""))
End Function

Function ActuallyIsDate(S As String) As Boolean
    ActuallyIsDate = IsDate(Left(S, 8)) And Mid(S, 3, 1) = "/"
End Function

Function DBG_ElapsedTime(Optional additionalMessage As String = "")
    If Len(additionalMessage) > 0 Then
        Debug.Print (additionalMessage)
    End If
    Debug.Print ("Elapsed Time: " & (pprintMS(PerfCounter.TimeElapsed)) & " seconds" + vbCrLf)
    PerfCounter.StartCounter
End Function

' Returns a range starting on row n of length u - i.e. "An:H(n+u-1)"
Function rowBlock(ByVal n As Integer, ByVal u As Integer) As String
    rowBlock = "A" & n & ":Z" & (n + u - 1)
End Function

Function IsClientCode(cellValue As Variant) As Boolean
    IsClientCode = (TypeName(cellValue) <> "Date") And (IsEmpty(cellValue) = False)
End Function

'Returns index of last item in arr less than x. Assumes array is sorted.
Function lastItemLessThanX(ByRef Arr() As Integer, x As Integer) As Integer
    idx = -1
    For Each a In Arr
        If a < x Then
            idx = idx + 1
        Else
            Exit For
        End If
    Next
    lastItemLessThanX = idx
End Function


'Remove every sheet but the first
Sub deleteExtraSheets()
    Application.DisplayAlerts = False
    Do While Worksheets.Count > 1
        Worksheets(2).Delete
    Loop
    Application.DisplayAlerts = True
End Sub

Sub xxx_Test()
    Dim arr_test(10) As Integer
    For i = 0 To 9
        arr_test(i) = i
    Next
    Debug.Print ("Last item less than 5 - " & (lastItemLessThanX(arr_test, 5)))
End Sub

Function enterPerfMode()
    Application.ScreenUpdating = False     'Faster
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Function

Sub exitPerfMode()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub











Sub RingsJournalSort()

    PerfCounter.StartCounter

    Call deleteExtraSheets
    Call enterPerfMode

    ' Allocate variables
    Dim SheetLength() As Integer
    Dim OrigSheet, JournalSheet1, JournalSheet2 As Worksheet
    Dim TempJnl() As String
    Dim startsWithID, startsWithDate As Boolean
    Dim CurrentJnl, JnlNo, z, numJournals1, numJournals2, b, c, d, numClientCodes1, j, k, m, n As Integer

    Const lastColumn = 6
    Const lastColumnLetter = "F"

    Dim ReportTypes As Variant
    ReportTypes = Split("Sales Ledger Transfer Report;" & _
                        "Purchase Ledger Transfer Report;" & _
                        "Cash Book Transfer Report;" & _
                        "Nominal Journal Posting", ";")


    'XXX: Goal is to get all necessary information in one pass through the input.
    ' Dict type structure for VBA:
    ' Dim dict As New Scripting.Dictionary
    
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
    
    ' Separate and print out journal entries. Text is "Journal No. N"
    ' Journals come in two formats. Sheet 1 is for "long" format journals, Sheet 2 for "short" format.
    ' XXX: There is more logic to the way these journals are laid out that we're not taking advantage of.
    ' It's like we have no idea whether to expect a customer code or a date or whatever, but it's actually not that crazy.
    Do While IsEmpty(OrigSheet.Cells(z, 1)) = False And IsEmpty(OrigSheet.Cells(z + 1, 1)) = False And IsEmpty(OrigSheet.Cells(z + 2, 1)) = False
        'Loop End Condition - see if 3 consecutive rows are blank
        Dim currentCell As Range
        Set currentCell = OrigSheet.Cells(z, 1)

        'Check whether the row starts with a numeric value ID or a date
        startsWithID = IsNumeric(Left(currentCell, 3))
        startsWithDate = IsDate(Left(currentCell, 8))
                
        'Find Cells which contain the word journal and extract the journal number which is set to JnlStr
        JnlStr = Mid(OrigSheet.Cells(z, 1), InStr(currentCell, "Journal") + 6 + 4, 6)

        ' Construct a new journal string since journal types 1 and 2 have slightly different formatting
        If IsNumeric(JnlStr) = True And JnlStr > 100 Then
            numJournals2 = numJournals2 + 1
            JournalSheet2.Cells(numJournals2, 2) = "Journal No. " & JnlStr
        ElseIf IsNumeric(JnlStr) = False And IsNumeric(Replace$(JnlStr, ".", "")) = True Then
            numJournals1 = numJournals1 + 1
            JournalSheet1.Cells(numJournals1, 2) = "Journal No." & Replace$(JnlStr, ".", "")
        End If

        'Filtering out rows that don't begin with a number and which contain Co. name
        If (startsWithID Or startsWithDate) And (InStr(currentCell, Left(OrigSheet.Cells(1, 1), 18)) = 0) Then
            
            'seperating remaining items between two worksheets using string length without spaces
            'NOTE the number beside the < should represent the longest string that isn't a Jnl with Dr and Cr
            If (Len(stripSpaces(currentCell)) < 35) Or (startsWithDate = True) Then
                numJournals1 = numJournals1 + 1

                'XXX : CHECK
                '----This If Statement copies over the a/c name and client a/c code and the total and moves the total into column 3
                If IsNumeric(currentCell) = True Then
                    JournalSheet1.Cells(numJournals1, lastColumn).Value = currentCell.Value
                Else
                    JournalSheet1.Cells(numJournals1, 1).Value = currentCell.Value
                End If

            Else
                numJournals2 = numJournals2 + 1
                JournalSheet2.Cells(numJournals2, 1).Value = currentCell.Value
            End If
                                                 
        End If
        z = z + 1
    Loop
    
    Debug.Print ("Found numJournals1: " & numJournals1)
    Debug.Print ("Found numJournals2: " & numJournals2)
    
    
    c = 1
    ReDim TempJnl(c)
    TempJnl(0) = 100

    '----Deleting Duplicate Journal No. Rows on Sheet 2
    For b = 1 To numJournals1
        ReDim Preserve TempJnl(c)

        If InStr(JournalSheet1.Cells(b, 2), "Journal") Then
            TempJnl(c) = JournalSheet1.Cells(b, 2)
            If TempJnl(c) = TempJnl(c - 1) Then
                JournalSheet1.Rows(b).Delete
            End If
            c = c + 1
        End If
    Next
    
    '----Deleting Duplicate Journal No. Rows on Sheet 3
    For b = 1 To numJournals2
        ReDim Preserve TempJnl(c)
        
        If InStr(JournalSheet2.Cells(b, 2), "Journal") Then
            TempJnl(c) = JournalSheet2.Cells(b, 2)
            If TempJnl(c) = TempJnl(c - 1) Then
                JournalSheet2.Rows(b).Delete
            End If
            c = c + 1
        End If
    Next
    

    '----Text to columns for Journal List 1.
    For b = 1 To numJournals1
        JournalSheet1.Select
        Cells(b, 1).Select
        Set currentCell = JournalSheet1.Cells(b, 1)
        ' Ordinary log entries. Identify these if they start with a date.
        ' Excel will happily convert "152 004" to a date so make sure there is a forward slash
        If ActuallyIsDate(currentCell.Value) Then
            Selection.TextToColumns Destination:=currentCell, DataType:=xlFixedWidth, _
                FieldInfo:=Array(Array(0, 4), Array(8, 2), Array(28, 2), Array(40, 1), Array(72, 1), _
                Array(84, xlSkipColumn)), DataType:=standardDataType, TrailingMinusNumbers:=True
            Cells(b, 2).Value = Trim(Cells(b, 2).Value)
        'Columns with Client Code. These are not fixed width and may depend on the
        ElseIf IsEmpty(currentCell) = False Then
            ' Check for second set of three numbers like "123 001"
            If IsNumeric(Mid(currentCell, 5, 3)) Then
                Selection.TextToColumns Destination:=currentCell, DataType:=xlFixedWidth, _
                    FieldInfo:=Array(Array(0, 1), Array(8, 2)), TrailingMinusNumbers:=True
            Else
                Selection.TextToColumns Destination:=currentCell, DataType:=xlFixedWidth, _
                    FieldInfo:=Array(Array(0, 1), Array(4, 2)), TrailingMinusNumbers:=True
            End If
        End If
    Next
    
    
    ' Format Journal Sheet 1
    JournalSheet1.Columns(lastColumn - 1).NumberFormat = "#,###;(#,###);0" ' Numerical entries

    'Totals
    JournalSheet1.Columns(lastColumn).NumberFormat = "#,###;(#,###);0"
    For b = 1 To numJournals1
        If IsEmpty(JournalSheet1.Cells(b, lastColumn - 1)) = False Then
            JournalSheet1.Cells(b, lastColumn - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
            JournalSheet1.Cells(b, lastColumn - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        End If
    Next

    JournalSheet1.Columns("A:" & lastColumnLetter).AutoFit
    JournalSheet1.Columns(1).ColumnWidth = 9.29
    JournalSheet1.Columns(2).ColumnWidth = 50
    


    '----Text to columns for Journal List 2.
    'Note these contain different columns, including debit and credit entries.
    For b = 1 To numJournals2
        If IsEmpty(JournalSheet2.Cells(b, 1)) = False Then
        JournalSheet2.Select
        Cells(b, 1).Select
        Selection.TextToColumns Destination:=JournalSheet1.Cells(b, 1), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(14, 2), Array(44, 1), Array(63, 1), Array(76, 2)), _
            TrailingMinusNumbers:=True
        End If
    Next
    
    
    ' Format Journal Sheet 2
    
    'Format two columns of numbers
    With JournalSheet2.Columns(3)
        .NumberFormat = "#,###;(#,###);0"
        .Font.FontStyle = "Bold"
    End With
    With JournalSheet2.Columns(4)
        .NumberFormat = "#,###;(#,###);0"
        .Font.FontStyle = "Bold"
    End With
    
    JournalSheet2.Columns("B:" & lastColumnLetter).AutoFit
    JournalSheet2.Columns(1).ColumnWidth = 9.29
    
    'Up to this point I've extracted and formatted each journal and the transactions within
    'now I want to gather for each client code the journals making up the total in seperate sheets
    
    
    Dim Sh As Worksheet
    Dim AllCodes() As Long
    Dim FindOutArray() As String
    Dim newCode As Long
    ReDim AllCodes(0)
    d = 0

    '---- Sort client codes from Journal Sheet 1 and find unique values
    numClientCodes1 = 0
    For b = 1 To numJournals1
        Dim t As Boolean
        If IsClientCode(JournalSheet1.Cells(b, 1).Value) Then
            numClientCodes1 = numClientCodes1 + 1
            newCode = strippedToNum(JournalSheet1.Cells(b, 1).Value)
            If IsContainedInArray(newCode, AllCodes) = False Then
                d = d + 1
                ReDim Preserve AllCodes(d)
                AllCodes(d) = newCode
                With ThisWorkbook
                    Set Sh = .Sheets.Add(after:=.Sheets(.Sheets.Count))
                    Sh.Name = newCode
                End With
            End If
        End If
    Next
    
    '---- Sort client codes from Journal Sheet 2 and find unique values not already seen in Journal Sheet 1
    numClientCodes2 = 0
    For b = 1 To numJournals2
        If IsClientCode(JournalSheet2.Cells(b, 1).Value) Then
            numClientCodes2 = numClientCodes2 + 1
            newCode = strippedToNum(JournalSheet2.Cells(b, 1))
            If IsContainedInArray(newCode, AllCodes) = False Then
                d = d + 1
                ReDim Preserve AllCodes(d)
                AllCodes(d) = newCode
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
    
    Dim JnlRow() As Integer
    Dim CodeRow(), CodeRow2(), u, LastRow As Integer
    Dim ShName As Double
    ReDim CodeRow(numClientCodes1)
    
    'Reverse lookup rows with client codes (XXX: Probably should move this up?)
    d = 0
    For b = 1 To numJournals1
        If IsClientCode(JournalSheet1.Cells(b, 1).Value) Then
            ReDim Preserve CodeRow(d)
            CodeRow(d) = b
            d = d + 1
        End If
    Next
    
    'the Array v is for referencing to the sheets in this workbook
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

    
    'Search the list of Type 1 Journals for client codes.
    Dim codeHasEntries As Boolean
    JournalSheet1.Select
    For k = 4 To ThisWorkbook.Sheets.Count
        ShName = strippedToNum(Sheets(k).Name)
        SheetLength(k) = 1 'Array in order to retain last row printed on for each sheet
        codeHasEntries = False

        'loop through the array holding the row numbers of the rows with the client codes
        For b = 0 To UBound(CodeRow)

            'Check if the Client Code in a Row is equal to the current sheet looped onto
            If strippedToNum(JournalSheet1.Cells(CodeRow(b), 1)) = ShName Then
                
                codeHasEntries = True

                ' Find the distance between the last row in the current set of transactions to the beginning of the next set
                If b = UBound(CodeRow) Then
                    With JournalSheet1
                        LastRow = .Cells(.Rows.Count, 5).End(xlUp).Row
                    End With
                    'add 1 for the row with the total -- this is in order to correctly place the total of the totals
                    u = 1 + LastRow - CodeRow(UBound(CodeRow))
                Else
                    u = CInt(CodeRow(b + 1)) - CInt(CodeRow(b))
                End If

                'Copy the journal entry found closest directly above this client code
                Dim idx As Integer
                idx = lastItemLessThanX(JnlRow, (CodeRow(b)))
                Sheets(k).Cells(SheetLength(k), 2).Value = Trim(JournalSheet1.Cells(JnlRow(idx), 2).Value)
                SheetLength(k) = SheetLength(k) + 2

                'this copies across the Client Code row plus the rows up to the next client code
                'this is in order to catch the transactions beneath the code row
                If CodeRow(b) + u - 1 = JnlRow(idx) Then
                    u = u - 1
                End If
                Sheets(k).Range(rowBlock((SheetLength(k)), u)).Value = JournalSheet1.Range(rowBlock((CodeRow(b)), u)).Value
                SheetLength(k) = SheetLength(k) + u
            SheetLength(k) = SheetLength(k) + 1
            End If
        Next b
        
        If codeHasEntries Then
            'Putting a total of the totals moved over from Sheet 2
            ' Could vectorize this
            SheetLength(k) = SheetLength(k) + 1
            For b = 1 To SheetLength(k) - 1
                Sheets(k).Cells(SheetLength(k), 5) = Sheets(k).Cells(SheetLength(k), 5) + Sheets(k).Cells(b, 5)
            Next b
        
            'Format total
            'XXX: Should plug in an Excel formula for clarity
            With Sheets(k).Cells(SheetLength(k), 5)
                .NumberFormat = "#,###;(#,###);0"
                .Font.FontStyle = "Bold"
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeTop).LineStyle = xlContinuous
            End With
            SheetLength(k) = SheetLength(k) + 5
        End If

        'autofit the columns in the client code sheets
        Sheets(k).Columns("A:" & lastColumnLetter).AutoFit
        Sheets(k).Columns(5).ColumnWidth = 10
        Sheets(k).Columns(4).ColumnWidth = 10
    Next k
    
    PerfCounter.StartCounter

    '----Successful extraction and formatting to each appropriate tab from JnlList1
    '----Next Code to do same for *JnlList2*
    
    'this creates an array containing the rows which have the client codes
    d = 0
    For b = 1 To numJournals2
        If IsEmpty(JournalSheet2.Cells(b, 1)) = False Then
            ReDim Preserve CodeRow2(d)
            CodeRow2(d) = b
            d = d + 1
        End If
    Next
    

    CurrentJnl = 1
    j = 1
    Dim RunningTotal, RunningTotal2 As Double
    JournalSheet2.Select
    For k = 4 To ThisWorkbook.Sheets.Count
    
        ShName = strippedToNum(Sheets(k).Name)
        codeHasEntries = False
        RunningTotal = 0
        RunningTotal2 = 0
        'loop through the array holding the row numbers of the rows with the client codes
        For b = 1 To numJournals2
            
            'Check if Cells are empty if so save row no. to a temp variable
            If IsEmpty(JournalSheet2.Cells(b, 1)) Then
                CurrentJnl = b
            'Check if the Client Code in a Row is equal to the current sheet looped onto
            ElseIf strippedToNum(JournalSheet2.Cells(b, 1)) = ShName Then
                codeHasEntries = True
                'if cell contains client code and is same as the current sheet then paste
                'Copy the journal number
                Sheets(k).Cells(SheetLength(k), 2).Value = Trim(JournalSheet2.Cells(CurrentJnl, 2).Value)
                SheetLength(k) = SheetLength(k) + 2 ' Leave some room for the label!
                'second paste the row containing the client code and other information
                Sheets(k).Range(rowBlock(SheetLength(k), 1)).Value = JournalSheet2.Range(rowBlock(b, 1)).Value
                'These are just to allow me to print a total at the end of sheet
                RunningTotal = RunningTotal + Sheets(k).Cells(SheetLength(k), 3)
                RunningTotal2 = RunningTotal2 + Sheets(k).Cells(SheetLength(k), 4)
                SheetLength(k) = SheetLength(k) + 1
            End If
        Next b
        If codeHasEntries Then
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
        Sheets(k).Columns("A:" & lastColumnLetter).AutoFit
        Sheets(k).Columns(5).ColumnWidth = 10
        Sheets(k).Columns(4).ColumnWidth = 10
        
    Next k
    
    
    'Loop to reorder the client code tabs in numeric order
    ' XXX: Rewrite this
    For b = 4 To ThisWorkbook.Sheets.Count
        For k = 4 To ThisWorkbook.Sheets.Count - 1
            If strippedToNum(Sheets(k).Name) > strippedToNum(Sheets(k + 1).Name) Then Sheets(k).Move after:=Sheets(k + 1)
        Next k
    Next b
    
    
    'find the last row of the first sheet
    Dim FinalRow As Integer
    With Sheets(1)
        FinalRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With

    'Match journals with "Sales/Purchase Ledger Transfer Report"
    ' We go through the original source document. Unfortunate and unnecessary!
    For k = 4 To ThisWorkbook.Sheets.Count
    
        With Sheets(k)
           LastRow = .Cells(.Rows.Count, 3).End(xlUp).Row
        End With
    
        For b = 1 To LastRow
            If InStr(Sheets(k).Cells(b, 2), "Journal") = 1 Then
                ' Loop through sheets. When we locate a journal, begin search for category.
                targetType = "Other"  ' Could set this to "" to disable
                journalNum = strippedToNum(Right(Sheets(k).Cells(b, 2), 5))
                For j = 1 To FinalRow
                    currentCell = Sheets(1).Cells(j, 1)
                    If InStr(currentCell.Value, "Journal") And InStr(currentCell.Value, journalNum) Then
                        For i = j To Application.Max(j - 70, 1) Step -1
                            ' Each of the "Sales Ledger" type strings is in here
                            For Each rt In ReportTypes
                                If InStr(Sheets(1).Cells(i, 1).Value, rt) Then
                                    targetType = rt
                                    GoTo PostLabel ' Goto is horrible but better than the alternative
                                End If
                            Next
                        Next
                    End If
                Next
PostLabel:
                Sheets(k).Cells(b + 1, 1) = targetType
            End If
        Next b
    Next k
    
    

    'Delete Working Sheets that may not be necessary but are useful to hang onto in case they are wanted or useful
    Application.DisplayAlerts = True

    exitPerfMode
    
    DBG_ElapsedTime "Done!"
End Sub







