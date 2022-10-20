'Unhides all sheets, even if they're very hidden.
Sub UnhideAll()
    Dim WS As Worksheet
    For Each WS In Worksheets
        WS.Visible = True
    Next
End Sub

' Marks the current sheet as very hidden.
Sub HideSheet()
    Dim sheet As Worksheet
    Set sheet = ActiveSheet
    sheet.Visible = xlSheetVeryHidden
End Sub

' Is one of your PivotTables overlapping?
' Use this VBA to find out specifically which one it is.
' It will rename all of your pivot tables, then if you try to refresh, Excel will tell you which one's giving you the problem.
Sub renameAllPivotTables()
    Dim pvt As PivotTable
    Dim sh As Worksheet
    Dim num As Integer
    
    For Each sh In ThisWorkbook.Worksheets
        If sh.PivotTables.Count > 0 Then
            
            For Each pvt In sh.PivotTables
                ' Put num first because the fucking pivot table name textbox is like 20 pixels long....
                pvt.Name = num & sh.Name
                num = num + 1
                Debug.Print pvt.Name & pvt.SourceData
            Next pvt
        End If
    Next sh
End Sub

' Are you unable to figure out where your PivotTable is?
' Use this VBA to find it.
Sub findSpecificPivot()
    Dim pvt As PivotTable
    Dim sh As Worksheet
    Dim num As Integer
    
    For Each sh In ThisWorkbook.Worksheets
        If sh.PivotTables.Count > 0 Then
            
            For Each pvt In sh.PivotTables
                If pvt.Name = "28NATIONAL BROADCAST SUMMARY" Then
                    Debug.Print pvt.Name & pvt.TableRange1.Address
                End If
            Next pvt
        End If
    Next sh
End Sub


' If I have a thousand pivot tables and I want to do something to a specific field, here's a good start.
' In the below example, any pivot table on the current worksheet that has the fields Track Wk, Media Type, or Highlight to mark those fields as show items with no data.
Sub showAllItems()
    Dim pt As PivotTable
    On Error Resume Next
    For Each pt In ActiveSheet.PivotTables
        With pt.PivotFields("Track Wk")
            .showAllItems = True
        End With
        With pt.PivotFields("Media Type")
            .showAllItems = True
        End With
        With pt.PivotFields("Highlight")
            .showAllItems = True
        End With
    Next
End Sub

' Need to figure out if a field in your list of fields is actually being used in a pivot anywhere?
' This VBA will iterate through every pivot table in the workbook and list every field being used in it.
Sub fieldBeingUsed()
    Dim pvt As PivotTable
    Dim sh As Worksheet
    Dim num As Integer
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "USED FIELDS"
    For Each sh In ThisWorkbook.Worksheets
        If sh.PivotTables.Count > 0 Then
            For Each pvt In sh.PivotTables
            
                'column fields
                For Each cf In pvt.ColumnFields
                    num = num + 1
                    Cells(num, 1) = cf.SourceName
                    Cells(num, 2) = sh.Name
                Next cf
                
                'row fields
                For Each rf In pvt.RowFields
                    num = num + 1
                    Cells(num, 1) = rf.SourceName
                    Cells(num, 2) = sh.Name
                Next rf
                
                'filter fields
                For Each ff In pvt.ActiveFilters
                    num = num + 1
                    Cells(num, 1) = ff.SourceName
                    Cells(num, 2) = sh.Name
                Next ff
                
                'data fields
                For Each df In pvt.DataFields
                    num = num + 1
                    Cells(num, 1) = df.SourceName
                    Cells(num, 2) = sh.Name
                Next df
                
            Next pvt
        End If
    Next sh
    
    Worksheets("USED FIELDS").Select
    Columns("A:A").Select
    Selection.Replace What:="#N/A", Replacement:="S Values?", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub


