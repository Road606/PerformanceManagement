Attribute VB_Name = "util_file"
Option Explicit

Public Const module_name As String = "util_file"
Public Const str_path_separator As String = "\"

Function open_wb(file_path, Optional is_readonly = True, Optional is_visible As Boolean = True, Optional str_password As String = "") As Workbook
    On Error GoTo ERR_FILE_PATH
    If str_password = "" Then
        Set open_wb = Workbooks.Open(file_path, readonly:=is_readonly)
    Else
        Set open_wb = Workbooks.Open(file_path, readonly:=is_readonly, WriteResPassword:=str_password)
    End If
    ' set visibility of the file
    Windows(open_wb.Name).Visible = is_visible
    
    'If Not is_visible Then
    '    ' ActiveWindow.visible = False
    '    Windows(open_wb.Name).visible = is_visible
    'End If
    On Error GoTo 0
    Exit Function
ERR_FILE_PATH:
    app_error.raise app_error.LEVEL_ERR, module_name & ">open_wb", _
        "Opening file: " & file_path & " has failed"
End Function

Function file_exists(str_file_path As String) As Boolean
    file_exists = Dir(str_file_path) = ""
End Function

Function path_exists(str_dir_path As String) As Boolean
    On Error Resume Next
    path_exists = (GetAttr(str_dir_path) And vbDirectory) = vbDirectory
End Function

Function retrieve_folders(str_dir_path As String) As Collection
    Dim file_name As String
    
    If Not path_exists(str_dir_path) Then
        app_error.raise app_error.LEVEL_ERR, module_name & ">retrieve_files", _
            "Directory path: " & str_dir_path & " doesn't exist."
    End If
    
    If Right(str_dir_path, 1) <> str_path_separator Then
        str_dir_path = str_dir_path & str_path_separator
    End If
    
    Set retrieve_folders = New Collection
    file_name = Dir(str_dir_path, vbDirectory)
    
    Do While file_name <> ""
        If file_name <> "." And file_name <> ".." And file_name <> "queue" Then
            retrieve_folders.add file_name, file_name
        End If
        file_name = Dir()
    Loop
End Function

Function retrieve_files(str_dir_path As String, Optional str_specific As String = "*.*") As Collection
    Dim file_name As String
    
    If Not path_exists(str_dir_path) Then
        app_error.raise app_error.LEVEL_ERR, module_name & ">retrieve_files", _
            "Directory path: " & str_dir_path & " doesn't exist."
    End If
    
    If Right(str_dir_path, 1) <> str_path_separator Then
        str_dir_path = str_dir_path & str_path_separator
    End If
    
    Set retrieve_files = New Collection
    file_name = Dir(str_dir_path & str_specific)
    
    Do While file_name <> ""
        retrieve_files.add file_name, file_name
        file_name = Dir()
    Loop
End Function
