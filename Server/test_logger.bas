Attribute VB_Name = "test_logger"
Option Explicit

Public Function test()

Dim obj_logger As LoggerMail
Dim message As MSG


Set obj_logger = New LoggerMail

obj_logger.init "test", log4VBA.ERRO, "test_destination"
obj_logger.mailAddress = "rostislav.jirasek@lego.com"

log4VBA.init
log4VBA.add_logger obj_logger



Set message = New MSG
log4VBA.error "test_destination", message.source("testing_module").text("Toto je testovací zpráva, která nemá žádný význam.")
Set message = Nothing


End Function

Public Function test_mail()
    Dim oApp As Object
    Dim oMail As Object
    
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
        .To = "rostislav.jirasek@lego.com"
        .Subject = "testicek"
        .Body = "test"
        .Send
    End With
    
    
End Function
