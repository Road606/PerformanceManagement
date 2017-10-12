Attribute VB_Name = "hndl_performance_output"
Option Explicit

Public STR_WS_NAME As String
Public STR_DAILY_WS_NAME_KPI As String
Public STR_DAILY_WS_NAME_ADDITIONAL As String

Public Const STR_FIRST_ROW_TBL As String = "A1:U1"
Public Const STR_DATA_START_RG = "A2"

' outbound
Public STR_OUTBOUND_PATH As String
Public STR_OUTBOUND_FILE As String
Public STR_OUTBOUND_TMPL_PATH As String

Public str_file_prefix As String
Public str_file_name_separator As String
Public str_file_appendix As String
Public STR_TMPL_FILE_APPENDIX As String


Public str_passwd As String

Public Const STR_SAVE_MODE_DAILY As String = "daily"
Public Const STR_SAVE_MODE_WEEKLY As String = "weekly"
Public Const STR_SAVE_MODE_MONTHLY As String = "monthly"
Public str_save_mode As String

Public INT_APP_WORK_COL_DATE_OFFSET As Integer
Public STR_APP_WORK_FIRST_RG As String

Public STR_APP_MODULE_NAME As String

Public Function init()
    STR_WS_NAME = "data"

    'STR_OUTBOUND_PATH = ThisWorkbook.Path & "\data\outbound\"
    'STR_OUTBOUND_TMPL_PATH = ThisWorkbook.Path & "\tmpl\"
    str_file_prefix = "performance-inbound"
    str_file_name_separator = "-"
    str_file_appendix = ".xlsm"
    STR_TMPL_FILE_APPENDIX = "tmpl"
    str_passwd = "db_history"
    
    STR_APP_MODULE_NAME = "hndl_performance_output"
    
    INT_APP_WORK_COL_DATE_OFFSET = 0
    STR_APP_WORK_FIRST_RG = "A2"
    
    'STR_SAVE_MODE = STR_SAVE_MODE_WEEKLY
    'STR_SAVE_MODE = STR_SAVE_MODE_MONTHLY
End Function

Public Function save()
    Debug.Print "hndl_performance_output.save"
    Select Case str_save_mode
        Case STR_SAVE_MODE_DAILY
            save_daily
        Case STR_SAVE_MODE_WEEKLY
            'save_weekly
        Case STR_SAVE_MODE_MONTHLY
            'save_monthly
        Case Else
            Debug.Print STR_APP_MODULE_NAME & "->save->unknown case"
    End Select
End Function

Public Function save_daily()
    save_daily_kpi
End Function

Public Function save_daily_kpi()
    ' dal se podivat v souboru tmpl na formatovani napr BINu - ikdyz to vypada ted dobre

    Dim wb_history As Workbook
    Dim ws_history As Worksheet
    Dim rg_history As Range
    Dim rg_data As Range
'
'    Dim str_file_path_current As String
'    Dim str_file_path_last As String
'    Dim obj_history_dates As Collection
'    Dim var_history_date As Variant
'    Dim str_history_date As String
    
    Set wb_history = open_wb(get_file_name_daily())
    Set ws_history = wb_history.Worksheets(STR_DAILY_WS_NAME_KPI)
    
    Set rg_data = hndl_performance.get_data_daily_kpi()
    
    If Not rg_data Is Nothing Then
        rg_data.Copy
        Set rg_history = wb_history.Worksheets(STR_DAILY_WS_NAME_KPI).Cells(ws_history.Range("A:A").CountLarge, 1).End(xlUp).Offset(1)
        rg_history.PasteSpecial xlPasteAll
        Application.CutCopyMode = False
            ' get rid of duplicate records
'            wb_history.Worksheets(STR_WS_NAME).Range(STR_FIRST_ROW_TBL).CurrentRegion.RemoveDuplicates _
'                Columns:=Array(1, 2, 3, 4, 5, 6, _
'                7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20), Header:=xlYes

            'sort wb_history
        ' run update
        Application.run "'" & wb_history.Name & "'!app.update"

    End If
    
    'Set obj_history_dates = retrieve_dates
    
'    For Each var_history_date In obj_history_dates
'        str_file_path_current = get_file_name_weekly(CStr(var_history_date), app.int_week_beginning)
'        If str_file_path_current <> str_file_path_last Then
'            If Not (wb_history Is Nothing) Then
'                wb_history.Close SaveChanges:=True
'            End If
'            Set wb_history = open_wb(str_file_path_current)
'            Set ws_history = wb_history.Worksheets(STR_WS_NAME)
'        End If
'        'Set wb_history = open_wb(str_file_path_current)
'
'        Set rg_data = hndl_performance.get_data(CStr(var_history_date))
'
'        If Not rg_data Is Nothing Then
'            rg_data.Copy
'            Set rg_history = wb_history.Worksheets(STR_WS_NAME).Cells(ws_history.Range("A:A").CountLarge, 1).End(xlUp).Offset(1)
'            rg_history.PasteSpecial xlPasteAll
'            ' get rid of duplicate records
'            wb_history.Worksheets(STR_WS_NAME).Range(STR_FIRST_ROW_TBL).CurrentRegion.RemoveDuplicates _
'                Columns:=Array(1, 2, 3, 4, 5, 6, _
'                7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20), Header:=xlYes
'
'            sort wb_history
'
'        End If
'
'        str_file_path_last = str_file_path_current
'    Next
'
    If Not (wb_history Is Nothing) Then
        wb_history.Close SaveChanges:=True
    End If
End Function

Public Function save_weekly()
    ' dal se podivat v souboru tmpl na formatovani napr BINu - ikdyz to vypada ted dobre

    Dim wb_history As Workbook
    Dim ws_history As Worksheet
    Dim rg_history As Range
    Dim rg_data As Range
    
    Dim str_file_path_current As String
    Dim str_file_path_last As String
    Dim obj_history_dates As Collection
    Dim var_history_date As Variant
    Dim str_history_date As String
    
    Set obj_history_dates = retrieve_dates
    
    For Each var_history_date In obj_history_dates
        str_file_path_current = get_file_name_weekly(CStr(var_history_date), app.int_week_beginning)
        If str_file_path_current <> str_file_path_last Then
            If Not (wb_history Is Nothing) Then
                wb_history.Close SaveChanges:=True
            End If
            Set wb_history = open_wb(str_file_path_current)
            Set ws_history = wb_history.Worksheets(STR_WS_NAME)
        End If
        'Set wb_history = open_wb(str_file_path_current)
        
        Set rg_data = hndl_performance.get_data(CStr(var_history_date))
        
        If Not rg_data Is Nothing Then
            rg_data.Copy
            Set rg_history = wb_history.Worksheets(STR_WS_NAME).Cells(ws_history.Range("A:A").CountLarge, 1).End(xlUp).Offset(1)
            rg_history.PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            ' get rid of duplicate records
            wb_history.Worksheets(STR_WS_NAME).Range(STR_FIRST_ROW_TBL).CurrentRegion.RemoveDuplicates _
                Columns:=Array(1, 2, 3, 4, 5, 6, _
                7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20), Header:=xlYes
            
            'sort wb_history
            
        End If
                
        str_file_path_last = str_file_path_current
    Next
    
    If Not (wb_history Is Nothing) Then
        wb_history.Close SaveChanges:=True
    End If
End Function

Public Function save_monthly()
    ' dal se podivat v souboru tmpl na formatovani napr BINu - ikdyz to vypada ted dobre

    Dim wb_history As Workbook
    Dim ws_history As Worksheet
    Dim rg_history As Range
    Dim rg_data As Range
    
    Dim str_file_path_current As String
    Dim str_file_path_last As String
    Dim obj_history_dates As Collection
    Dim var_history_date As Variant
    Dim str_history_date As String
    
    Set obj_history_dates = retrieve_dates
    
    For Each var_history_date In obj_history_dates
        str_file_path_current = get_file_name_monthly(CStr(var_history_date))
        If str_file_path_current <> str_file_path_last Then
            If Not (wb_history Is Nothing) Then
                wb_history.Close SaveChanges:=True
            End If
            Set wb_history = open_wb(str_file_path_current)
            Set ws_history = wb_history.Worksheets(STR_WS_NAME)
        End If
        'Set wb_history = open_wb(str_file_path_current)
        
        Set rg_data = hndl_performance.get_data(CStr(var_history_date))
        
        If Not rg_data Is Nothing Then
            rg_data.Copy
            Set rg_history = wb_history.Worksheets(STR_WS_NAME).Cells(ws_history.Range("A:A").CountLarge, 1).End(xlUp).Offset(1)
            rg_history.PasteSpecial xlPasteAll
            ' get rid of duplicate records
            wb_history.Worksheets(STR_WS_NAME).Range(STR_FIRST_ROW_TBL).CurrentRegion.RemoveDuplicates _
                Columns:=Array(1, 2, 3, 4, 5, 6, _
                7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20), Header:=xlYes
            
            'sort wb_history
            update_data wb_history
        End If
                
        str_file_path_last = str_file_path_current
    Next
    
    If Not (wb_history Is Nothing) Then
        wb_history.Close SaveChanges:=True
    End If
End Function

Public Function get_file_name_daily() As String
    get_file_name_daily = STR_OUTBOUND_PATH & STR_OUTBOUND_FILE
End Function

Public Function get_file_name_weekly(str_date As String, int_week_beginning As Integer)
    get_file_name_weekly = _
        STR_OUTBOUND_PATH & _
        str_file_prefix & _
        str_file_name_separator & _
        Year(DateValue(str_date)) & _
        WorksheetFunction.WeekNum(DateValue(str_date), int_week_beginning) & _
        str_file_appendix
End Function

Public Function get_file_name_monthly(str_date As String)
    get_file_name_monthly = _
        STR_OUTBOUND_PATH & _
        str_file_prefix & _
        str_file_name_separator & _
        "monthly" & _
        str_file_name_separator & _
        Year(str_date) & _
        Format(Month(str_date), "00") & _
        str_file_appendix
End Function

Public Function get_tmpl_file_name()
    get_tmpl_file_name = _
        STR_OUTBOUND_TMPL_PATH & _
        str_file_prefix & _
        str_file_name_separator & _
        STR_TMPL_FILE_APPENDIX & _
        str_file_appendix
End Function

Public Function retrieve_dates() As Collection
    Dim rg_app_work As Range
    Dim rg_transaction_date As Range
    
    app_work.clear
    
    Set rg_transaction_date = _
        ThisWorkbook.Worksheets(hndl_performance.STR_WS_NAME). _
        Range(hndl_performance.STR_DATA_FIRST_CELL). _
        Offset(0, db_performance.INT_OFFSET_TRANSACTION_DATE)
    
    Set rg_app_work = _
            ThisWorkbook.Worksheets(app_work.STR_WS_NAME). _
            Range(app_work.STR_DATA_START_RG)
    
    If rg_transaction_date.Value <> "" Then
        Set rg_transaction_date = Range(rg_transaction_date, rg_transaction_date.End(xlDown))
        rg_transaction_date.Copy
                                
        rg_app_work.PasteSpecial xlPasteValues
        rg_app_work.RemoveDuplicates Columns:=Array(INT_APP_WORK_COL_DATE_OFFSET + 1)
    End If
    
    Set retrieve_dates = New Collection
    
    Do While rg_app_work.Value <> ""
        retrieve_dates.add rg_app_work.Value
        Set rg_app_work = rg_app_work.Offset(1)
    Loop
End Function

Public Function open_wb(str_file_path As String) As Workbook
    On Error GoTo ERR_FILE_NOT_FOUND
    Set open_wb = Application.Workbooks.Open(Filename:=str_file_path, readonly:=False, WriteResPassword:=str_passwd)
    On Error GoTo 0
    Exit Function
ERR_FILE_NOT_FOUND:
    Set open_wb = Application.Workbooks.Open(Filename:=get_tmpl_file_name, readonly:=False, WriteResPassword:=str_passwd)
    open_wb.SaveAs str_file_path, WriteResPassword:=str_passwd
End Function

Public Function sort_daily_kpi(wb As Workbook)
    Dim ws As Worksheet
    Dim rg_tbl As Range
    Dim rg_col As Range
    
    Set ws = wb.Worksheets(STR_DAILY_WS_NAME_KPI)
    Set rg_tbl = ws.Range(STR_FIRST_ROW_TBL)
    Set rg_tbl = ws.Range(rg_tbl, rg_tbl.End(xlDown))
    
    Set rg_col = ws.Range(STR_DATA_START_RG)
    'Dim rg_sort As Range
    
    ws.Activate
    rg_tbl.AutoFilter
    On Error GoTo ERR_AUTOFILTER_NOT_SET
    ws.AutoFilter.sort.SortFields.clear
    On Error GoTo 0
    ws.AutoFilter.sort.SortFields.add Key:= _
        Range( _
            rg_col.Offset(0, db_performance.INT_OFFSET_TRANSACTION_DATE), _
            rg_col.Offset(0, db_performance.INT_OFFSET_TRANSACTION_DATE).End(xlDown) _
            ), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ws.AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Exit Function
ERR_AUTOFILTER_NOT_SET:
    rg_tbl.AutoFilter
    Resume Next
End Function

Public Function update_data(wb As Workbook)
    Dim ws As Worksheet
    Dim ws_data As Worksheet
    Dim pt As PivotTable
    Dim pt_cache As PivotCache
        
    Set ws_data = wb.Worksheets("data")
        
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            Debug.Print pt.Name
            Set pt_cache = wb.PivotCaches.create( _
                xlDatabase, _
                ws_data.Range("A1").CurrentRegion)
            pt.ChangePivotCache pt_cache
        Next
    Next
    
    Debug.Print wb.PivotCaches.Count
End Function

