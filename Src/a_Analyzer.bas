Attribute VB_Name = "a_Analyzer"
Function Analyzer(writefile As Boolean)
    Application.ScreenUpdating = False
    Dim Analysis
    Dim Domo
    Dim FuelData
    Dim InvData
    Dim InvDataTemp
    Dim hptimer As PerformanceMonitor
    Dim datedict As Scripting.Dictionary
    Dim storedict As Scripting.Dictionary
    Dim invdict As Scripting.Dictionary
    Dim lithiadict As Scripting.Dictionary
    
    Dim xrow
    Dim storeindex
    Dim dateindex
    Dim domorow
    
    'get data from fuel processor
    'break out the gallon totals by month and store
    Sheet4.Cells.Delete
    Sheet8.Cells.Delete
    
    Set hptimer = New PerformanceMonitor
    Set datedict = New Scripting.Dictionary
    Set storedict = New Scripting.Dictionary
    Set invdict = New Scripting.Dictionary
    Set lithiadict = New Scripting.Dictionary
    hptimer.StartCounter
    
    
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    storelist = UniqueVals(Sheet2.Range("K2:K" & lastrow))
    Sheet2.Columns("A:N").Sort key1:=Sheet2.Range("A2"), _
      order1:=xlAscending, Header:=xlYes
    monthlist = UniqueVals(Sheet2.Range("M2:M" & lastrow))
    daylist = UniqueVals(Sheet2.Range("N2:N" & lastrow))
    datelist = UniqueVals(Sheet2.Range("A2:A" & lastrow))
    storerow = 2
    fcrow = 2
        
    FuelData = Sheet2.Range("A2:N" & lastrow).Value
    
    Debug.Print ("UniqueVal: " & hptimer.TimeElapsed)
    hptimer.StartCounter
    
    
    'Build Inventory Array
    InvDataTemp = Sheet3.Range("A1:Y" & Sheet3.Cells(Sheet3.Rows.Count, "A").End(xlUp).Row).Value
    svccars = Sheet3.Range("Z1:AA" & Sheet3.Cells(Sheet3.Rows.Count, "Z").End(xlUp).Row).Value
    ReDim InvData(1 To UBound(InvDataTemp, 1) - 1, 0 To 12)
        For x = 2 To UBound(InvDataTemp, 1)
            svcinv = 0
            For y = 2 To UBound(svccars, 1)
                If svccars(y, 1) = InvDataTemp(x, 1) Then
                    svcinv = svccars(y, 2)
                    Exit For
                End If
            Next y
            InvData(x - 1, 0) = InvDataTemp(x, 1)
            For y = 2 To UBound(InvDataTemp, 2)
                If InvDataTemp(1, y) <> "" Then
                    If invcol <> Split(InvDataTemp(1, y), ";")(0) Then mnthinx = 0
                    
                    invcol = Split(InvDataTemp(1, y), ";")(0)

                    InvData(x - 1, invcol) = InvData(x - 1, invcol) + InvDataTemp(x, y)
                    
                    If mnthinx = 0 And InvData(x - 1, invcol) <> 0 Then
                        InvData(x - 1, invcol) = InvData(x - 1, invcol) + svcinv
                    End If
                    
                    mnthinx = mnthinx + 1
                End If
            Next y
        Next x
    
    'resize inventory array
    For y = 0 To 11
        mnthtotal = 0
        For x = 2 To UBound(InvData, 1)
            mnthtotal = mnthtotal + InvData(x, 12 - y)
        Next x
        If mnthtotal <> 0 Then
            lastmnth = 12 - y
            Exit For
        End If
    Next y
    
    
    
    Debug.Print ("Inventory: " & hptimer.TimeElapsed)
    hptimer.StartCounter
    
    
    'Set up arrays
    ReDim Preserve InvData(1 To UBound(InvData, 1), 0 To lastmnth)
    ReDim Analysis(1 To 1 + Application.CountA(storelist), 1 To 4 + Application.CountA(datelist) * 2)
    ReDim Domo(1 To (1 + Application.CountA(datelist) * Application.CountA(storelist)), 1 To 4)
        
    Domo(1, 1) = "Store#"
    Domo(1, 2) = "F/C"
    Domo(1, 3) = "Transaction Date"
    Domo(1, 4) = "Account Name"
    
    Analysis(1, 1) = "Store#"
    
    domorow = 2
    
    'create index dictionaries
    For irow = 2 To UBound(storelist) + 1
        Analysis(irow, 1) = storelist(irow - 1)
        storedict(storelist(irow - 1)) = irow
    Next irow
    
    For irow = 2 To Sheet6.Cells(Sheet6.Rows.Count, "D").End(xlUp).Row
        lithiadict(Sheet6.Range("D" & irow).Value) = Sheet6.Range("E" & irow).Value
    Next irow
    
    For irow = 1 To UBound(InvData)
        invdict(InvData(irow, 0)) = irow
    Next irow
    
    For jcol = 2 To UBound(datelist) + 1
        Analysis(1, jcol) = datelist(jcol - 1) & " Fuel"
        Analysis(1, jcol + UBound(datelist)) = datelist(jcol - 1) & " F/C"
        
        datedict(Analysis(1, jcol)) = jcol
        datedict(Analysis(1, jcol + UBound(datelist))) = jcol + UBound(datelist)
        
    Next jcol
    
    For xrow = 1 To UBound(FuelData, 1)
        storeindex = storedict(FuelData(xrow, 11))
        dateindex = datedict(FuelData(xrow, 1) & " Fuel")
    
        Analysis(storeindex, dateindex) = Analysis(storeindex, dateindex) + Val(FuelData(xrow, 3)) + 0
        
    Next xrow
    
    
    'Calculate F/C and fill out Analysis array
    For Each transdate In datelist
        For Each store In storelist
            storeindex = storedict(store)
            invindex = invdict(store)
            fuelindex = datedict(transdate & " Fuel")
            fcindex = datedict(transdate & " F/C")
        
            If Not IsEmpty(invindex) And Not IsEmpty(storeindex) And Not IsEmpty(fuelindex) And Not IsEmpty(fcindex) Then
                If InvData(invindex, Month(transdate)) <> 0 And Analysis(storeindex, fuelindex) <> 0 Then
                    Analysis(storeindex, fcindex) = Analysis(storeindex, fuelindex) / InvData(invdict(store), Month(transdate))
                    
                    'fill out Domo array
                    Domo(domorow, 1) = store
                    Domo(domorow, 2) = Analysis(storeindex, fcindex)
                    Domo(domorow, 3) = transdate
                    Domo(domorow, 4) = lithiadict(store)
                    domorow = domorow + 1
                End If
            End If
        Next store
    Next transdate
    
    

    'Get average, stdev, and cov for F/C
    fcstart = datedict("1/1/2018 F/C")
    fcend = UBound(Analysis, 2) - 3
        
    Analysis(1, fcend + 1) = "Average F/C"
    Analysis(1, fcend + 2) = "Day Over Day"
    Analysis(1, fcend + 3) = "% Change"
        
    For storerow = 2 To UBound(Analysis, 1)
        avgtotal = 0
        avgcount = 0
        For n = fcstart To fcend
            If Analysis(storerow, n) <> 0 Then
                avgcount = avgcount + 1
                avgtotal = Analysis(storerow, n) + avgtotal
            End If
        Next n
            
        If avgtotal <> 0 And avgcount <> 0 Then
            Analysis(storerow, fcend + 1) = avgtotal / avgcount
        End If
        
        Analysis(storerow, fcend + 2) = Analysis(storerow, fcend) - Analysis(storerow, fcend - 1)
        
        If Analysis(storerow, fcend) <> 0 Then
            Analysis(storerow, fcend + 3) = Analysis(storerow, fcend + 2) / Analysis(storerow, fcend)
        Else
            Analysis(storerow, fcend + 3) = 0
        End If
        
        Debug.Print ("Store Loop: " & hptimer.TimeElapsed)
        hptimer.StartCounter
    Next storerow

    
    Sheet4.Range("A1:" & Split(Cells(1, UBound(Analysis, 2)).Address, "$")(1) & UBound(Analysis, 1)).Value = Analysis
    Sheet8.Range("A1:D" & (UBound(Domo, 1))).Value = Domo
    
    lastcol = Sheet4.Cells(1, Sheet4.Columns.Count).End(xlToLeft).Column
    
    Sheet4.Range(Sheet4.Columns(1), Sheet4.Columns(lastcol)).Sort key1:=Sheet4.Cells(storerow, lastcol - 2), _
        order1:=xlDescending, Header:=xlYes
    
    'overage highlight
    lastrow = Sheet4.Cells(Sheet4.Rows.Count, 2).End(xlUp).Row
    For i = 2 To lastrow
    Set stcell = Sheet4.Cells(i, lastcol - 3)
        If IsNumeric(stcell.Value) Then
            If stcell.Value > (Application.WorksheetFunction.VLookup(Sheet4.Cells(i, 1).Value, Sheet6.Range("J:L"), 3, 0) * 3) + _
                (Application.WorksheetFunction.VLookup(Sheet4.Cells(i, 1).Value, Sheet6.Range("J:L"), 2, 0)) Then
                stcell.EntireRow.Interior.Color = RGB(255, 102, 102)
                varstore = Sheet4.Cells(i, 1).Value & ", " & varstore
            End If
        End If
    Next i
    
    If writefile = True Then: Call FileWriter
    Analyzer = MsgBox("All Done!" & vbNewLine & "The following stores have an unusual variance:" & vbNewLine & varstore)
        
        
        Debug.Print ("Total time: " & hptimer.TimeElapsed)
          
    Application.ScreenUpdating = True
End Function

Function UniqueVals(rangein As Range) As Variant

    Dim tmp As String
    Dim cell
    Dim arr() As String
    Dim arr2() As String

    For Each cell In rangein
        If (cell.Value <> "") And (InStr(1, tmp, cell.Value & "|", vbTextCompare) = 0) Then
            tmp = tmp & cell.Value & "|"
        End If
    Next cell

    If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - 1)
    
    arr = Split(tmp, "|")
    
    ReDim arr2(1 To UBound(arr, 1) + 1)
    
    For i = 1 To (UBound(arr, 1) + 1)
        arr2(i) = arr(i - 1)
    Next i
    
    UniqueVals = arr2

End Function

Function FileWriter()
    Dim w As Workbook
    Dim reportwb As Workbook
    Set reportwb = ActiveWorkbook
    Set w = Application.Workbooks.Add


    reportwb.Sheets("Finished Analysis").Copy _
        Before:=w.Sheets(1)
    
    reportwb.Sheets("Domo-Ready").Copy _
        Before:=w.Sheets(1)
    
    reportwb.Sheets("Compiled Fuel Data").Copy _
        Before:=w.Sheets(1)
    
    Application.DisplayAlerts = False
    For s = 1 To 3
        If SheetExists("Sheet" & s, ActiveWorkbook) Then
            Sheets("Sheet" & s).Delete
        End If
    Next s
    Application.DisplayAlerts = True


    flnm = "Fuel Report (" & Month(Date) & "-" & Day(Date) & "-" & Year(Date) & ")"

    filepath = "\\ntoscar\T-Drive Accounts Payable\Procurement\Fuel Analysis\Daily Summary\"

    On Error Resume Next
    Application.DisplayAlerts = False
    w.SaveAs Filename:=filepath & flnm, FileFormat:=51
    Application.DisplayAlerts = True
    w.Saved = True
    On Error GoTo 0

    w.Close
    
End Function

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ActiveWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
