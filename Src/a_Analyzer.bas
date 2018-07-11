Attribute VB_Name = "a_Analyzer"
Function Analyzer()
    Application.ScreenUpdating = False
    Dim Analysis
    Dim Domo
                     

    'get data from fuel processor
    'break out the gallon totals by month and store
    Sheet4.Cells.Delete
    Sheet8.Cells.Delete
    Sheet2.ShowAllData
    
    lastrow = Sheet2.Cells(Sheet2.Rows.Count, "A").End(xlUp).Row
    storelist = UniqueVals(Sheet2.Range("K2:K" & lastrow))
    Sheet2.Columns("A:N").Sort key1:=Sheet2.Range("A2"), _
      order1:=xlAscending, Header:=xlYes
    monthlist = UniqueVals(Sheet2.Range("M2:M" & lastrow))
    daylist = UniqueVals(Sheet2.Range("N2:N" & lastrow))
    datelist = UniqueVals(Sheet2.Range("A2:A" & lastrow))
    storerow = 2
    fcrow = storerow
        
    ReDim Analysis(1 To Application.CountA(storelist), 1 To (4 + Application.CountA(datelist) * 3))
    
    Analysis(1, 1) = "Store#"

    For Each store In storelist
        mnthcol = 2
        Analysis(storerow, 1) = store
        On Error GoTo 0
        For Each transdate In datelist
        
    'Write the headers
            If storerow = 2 Then
                Analysis(1, mnthcol) = transdate & " Fuel"
                Analysis(1, (mnthcol) + Application.CountA(datelist)) = transdate & " Cars"
                Analysis(1, (mnthcol) + (Application.CountA(datelist) * 2)) = transdate & " F/C"
                Analysis(1, 2 + Application.CountA(datelist) * 3) = "Average F/C"
                Analysis(1, 3 + Application.CountA(datelist) * 3) = "Day over Day"
                Analysis(1, 4 + Application.CountA(datelist) * 3) = "% Change"
                'Sheet8.Range("A1:D1") = Split("Store#,F/C,Transaction Date,Account Name", ",")
            End If
            
    'Get the units of fuel for each month
            Sheet6.Range("A2").Value = "=" & """" & "=" & store & """"
            Sheet6.Range("B2").Value = transdate
            Sheet2.Range("A:M").AdvancedFilter _
                Action:=xlFilterInPlace, _
                criteriarange:=Sheet6.Range("A1:B2")
                
            totalunits = Application.WorksheetFunction.Sum(Sheet2.Columns("C:C").SpecialCells(xlVisible))
            
            If totalunits <> 0 Then
                Analysis(storerow, mnthcol) = totalunits
            End If

    'Get the inventory for each month

        On Error GoTo inverror
        
            invmnthcoln = Application.Match(Month(transdate) & ";n", Sheet3.Range("1:1"), 0)
            invmnthcolu = Application.Match(Month(transdate) & ";u", Sheet3.Range("1:1"), 0)
            invstorerow = Application.Match(store, Sheet3.Range("A:A"), 0)
            
            serviceval = 0
            
            If Not IsError(Application.Match("serv", Sheet3.Range("1:1"), 0)) Then
                servcol = Application.Match("serv", Sheet3.Range("1:1"), 0)
                If Not IsError(Application.Match(store, Sheet3.Columns(servcol - 1), 0)) Then
                    servrow = Application.Match(store, Sheet3.Columns(servcol - 1), 0)
                    serviceval = Sheet3.Cells(servrow, servcol)
                End If
            End If
                
            invval = Sheet3.Cells(invstorerow, invmnthcoln).Value + Sheet3.Cells(invstorerow, invmnthcolu).Value + serviceval
            
            fcvalue = totalunits / invval
            
            Analysis(storerow, (mnthcol) + Application.CountA(datelist)) = invval
            
            If fcvalue <> 0 Then
                Analysis(storerow, (mnthcol) + (Application.CountA(datelist) * 2)) = fcvalue
            End If
            
'            If fcvalue <> 0 Then
'                Sheet8.Cells(fcrow, 1).Value = store
'                Sheet8.Cells(fcrow, 2).Value = fcvalue
'                Sheet8.Cells(fcrow, 3).Value = transdate
'                Sheet8.Cells(fcrow, 4).Value = Application.WorksheetFunction.VLookup(store, Sheet6.Range("D:E"), 2, 0)
'                fcrow = fcrow + 1
'            End If
nxtmnth:
            mnthcol = mnthcol + 1
        Next transdate
        
        On Error GoTo 0
        
'    'Get average, stdev, and cov for F/C
'        Set fcrange = Range(Sheet4.Cells(storerow, 2 + (Application.CountA(datelist) * 2)), Sheet4.Cells(storerow, mnthcol + (Application.CountA(datelist) * 3) + 1))
'        If Application.WorksheetFunction.Count(fcrange) > 0 Then
'        Analysis(storerow, 2 + Application.CountA(datelist) * 3) = Application.WorksheetFunction.Average(fcrange.Value)
'            Analysis(storerow, 3 + Application.CountA(datelist) * 3) = Analysis(storerow, 1 + Application.CountA(datelist) * 3) - Analysis(storerow, Application.CountA(datelist) * 3)
'            If Analysis(storerow, Application.CountA(datelist) * 3) <> 0 Then
'                Analysis(storerow, 4 + Application.CountA(datelist) * 3) = Analysis(storerow, 3 + Application.CountA(datelist) * 3) / Analysis(storerow, Application.CountA(datelist) * 3)
'            Else
'                Analysis(storerow, 4 + Application.CountA(datelist) * 3) = 0
'            End If
'        End If
        
    storerow = storerow + 1
        
    Next store
    Sheet2.ShowAllData
    '.Range("A1:" & Split(Cells(1, UBound(fueldata, 2)).Address, "$")(1) & UBound(fueldata, 1)).Value = fueldata
    
    Sheet4.Range("A2:" & Split(Cells(1, UBound(Analysis, 2)).Address, "$")(1) & UBound(Analysis, 1)).Value = Analysis
    
    lastcol = Sheet4.Cells(1, Sheet4.Columns.Count).End(xlToLeft).Column
    
    Sheet4.Range(Sheet4.Columns(1), Sheet4.Columns(lastcol)).Sort key1:=Sheet4.Cells(storerow, lastcol - 2), _
        order1:=xlDescending, Header:=xlYes
    
    lastrow = Sheet4.Cells(Sheet4.Rows.Count, 2).End(xlUp).Row
    For i = 2 To lastrow
    Set stcell = Sheet4.Cells(i, 1 + Application.CountA(datelist) * 3)
        If IsNumeric(stcell.Value) Then

            If stcell.Value > (Application.WorksheetFunction.VLookup(Sheet4.Cells(i, 1).Value, Sheet6.Range("J:L"), 3, 0) * 3) + _
                (Application.WorksheetFunction.VLookup(Sheet4.Cells(i, 1).Value, Sheet6.Range("J:L"), 2, 0)) Then
                stcell.EntireRow.Interior.Color = RGB(255, 102, 102)
                varstore = Sheet4.Cells(i, 1).Value & ", " & varstore
            End If
        End If
    Next i

    Call FileWriter
    Analyzer = MsgBox("All Done!" & vbNewLine & "The following stores have an unusual variance:" & vbNewLine & varstore)
    
    Exit Function

inverror:
    Resume nxtmnth



    Application.ScreenUpdating = True
End Function

Function UniqueVals(rangein As Range) As Variant

    Dim tmp As String
    Dim cell

    For Each cell In rangein
        If (cell.Value <> "") And (InStr(1, tmp, cell.Value & "|", vbTextCompare) = 0) Then
            tmp = tmp & cell.Value & "|"
        End If
    Next cell

    If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - 1)

    UniqueVals = Split(tmp, "|")

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
