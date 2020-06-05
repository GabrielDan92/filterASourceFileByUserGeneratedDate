Sub master(ByVal Day As String, ByVal Month As String, ByVal Year As String)

    Dim lastRow As Long
    Dim rng As Range

'check the date and convert it if it's valid
    dateConcat = Day & " " & Month & ", " & Year
    On Error Resume Next
    datetoexcel = DateValue(dateConcat)
    On Error GoTo 0
    
    If datetoexcel = "" Then
        MsgBox "Wrong date!"
        Exit Sub
    End If
    
    'clear the previous data from Sheet2
    Set rng = ThisWorkbook.Sheets(2).Range("A1").CurrentRegion
    rng.ClearContents
    
    'open the source file
        Dim wb As Workbook
        Set wb = Workbooks.Open(ThisWorkbook.Sheets(3).Range("C2").Value, 3)
        Set rng = wb.Sheets(1).Range("A1").CurrentRegion
        lastRow = rng.Rows.Count

    'retrieve the running Date
        datetoexcel = CLng(datetoexcel)
     
     With wb.Sheets(1)
        .Rows("1:1").AutoFilter                                                                                     'add the auto Filter
        .Columns("S:S").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove                                'insert a new column in 'S'
        .Range("S1").Value = "Date Code"
        .Range("S2:S" & lastRow).FormulaR1C1 = "=DATEVALUE(LEFT(RC[-1],9))"                                         'convert the date to code
    .Range("A1:A" & lastRow).AutoFilter Field:=19, Criteria1:=">=" & datetoexcel, Operator:=xlAnd                   'filter by date equal or higher than the user selected date
        .Columns("U:U").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove                                'insert new column in 'U'
        .Range("U1").Value = "Duree du sejour"
        .Range("U2:U" & lastRow).FormulaR1C1 = "=DAYS(LEFT(RC[-1],9),LEFT(RC[-2],9))+1"                             'add the formula that counts the days difference
        .Range("A1:A" & lastRow).AutoFilter Field:=21, Criteria1:=">=10", Operator:=xlAnd                           'filter the days above or equal to 10 in column 'T'
        .Range("A1:A" & lastRow).AutoFilter Field:=12, Criteria1:="=F*", Operator:=xlAnd                            'filter the 'VISITED_ID' equal to 'F'
        .Range("A1:A" & lastRow).AutoFilter Field:=3, Criteria1:="<>F*", Operator:=xlAnd                            'filter the 'emp_ID' different from 'F'
        .Range("A1:A" & lastRow).AutoFilter Field:=8, Criteria1:="Submitted"                                        'filter by 'Submitted' status
     End With


    'copy the filtered results back to the original workbook
      Set rng = wb.Sheets(1).Range("A1").CurrentRegion
      rng.Copy (ThisWorkbook.Sheets(2).Range("A1"))
    
    'close the source workbook
      wb.Close (False)
      Set wb = Nothing
    
    'add the 'contact' column
      ThisWorkbook.Sheets(2).Range("AM1").Value = "Contact"
    
    Set rng = ThisWorkbook.Sheets(2).Range("A1").CurrentRegion
    lastRow = rng.Rows.Count
    
    'add the formula for the contact
      With ThisWorkbook.Sheets(2)
          .Range("AM2:AM" & lastRow).FormulaR1C1 = "=VLOOKUP(LEFT(RC[-36],1),PARAM!C1:C2,2,0)"
          .Activate
      End With

End Sub
