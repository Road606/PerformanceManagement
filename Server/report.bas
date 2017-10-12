Attribute VB_Name = "report"
Option Explicit


Public Function update_all()
    Dim ws As Worksheet
    Dim ws_data As Worksheet
    Dim pt As PivotTable
    Dim pt_cache As PivotCache
        
    Set ws_data = ThisWorkbook.Worksheets("data")
        
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            Debug.Print pt.Name
            Set pt_cache = ThisWorkbook.PivotCaches.create( _
                xlDatabase, _
                ws_data.Range("A1").CurrentRegion)
            pt.ChangePivotCache pt_cache
        Next
    Next
    
    Debug.Print ThisWorkbook.PivotCaches.Count
End Function
