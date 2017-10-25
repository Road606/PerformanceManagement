Attribute VB_Name = "test_logger"
Option Explicit

Public Function test()

Dim obj_logger As LoggerDebug
Dim message As msg


Set obj_logger = New LoggerDebug

obj_logger.init "test", log4VBA.DBG, "test"

log4VBA.init
log4VBA.add_logger obj_logger


Set message = New msg
log4VBA.debg "test", message.source("testing_module").text("test")
Set message = Nothing


End Function

Public Function test_mail()
    Dim oApp As Object
    Dim oMail As Object
    
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
        .to = "rostislav.jirasek@lego.com"
        .Subject = "testicek"
        .Body = "test"
        .Send
    End With
    
    
End Function
