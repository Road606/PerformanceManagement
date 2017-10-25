Attribute VB_Name = "log4VBA"
Option Explicit

Public Const TRC As Integer = 1
Public Const DBG As Integer = 2
Public Const INF As Integer = 3
Public Const WRN As Integer = 4
Public Const ERR As Integer = 5

Public logging_is_enabled As Boolean

Public loggers As Collection

Public Function init()
    Set loggers = New Collection
    logging_is_enabled = True
End Function

Public Function trace(destination As String, message As msg)
    log TRC, destination, message
End Function

Public Function debg(destination As String, message As msg)
    log DBG, destination, message
End Function

Public Function info(destination As String, message As msg)
    log INF, destination, message
End Function

Public Function warn(destination As String, message As msg)
    log WRN, destination, message
End Function

Public Function error(destination As String, message As msg)
    log ERR, destination, message
End Function

Private Function log(lvl As Integer, destination As String, message As msg)
    Dim obj_logger As Object
    
    If logging_is_enabled Then
        For Each obj_logger In loggers
            obj_logger.log lvl, destination, message
        Next obj_logger
    End If
End Function

Public Function add_logger(obj_logger As Object)
    loggers.add obj_logger, obj_logger.name
End Function

Public Function remove_logger(logger_name As String)
    loggers.Remove logger_name
End Function
