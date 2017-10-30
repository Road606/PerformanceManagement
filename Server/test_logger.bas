Attribute VB_Name = "test_logger"
Option Explicit

Public Function testMailLogger()

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

Public Function testFileLogger()

Dim obj_logger As LoggerFile
Dim message As MSG


Set obj_logger = New LoggerFile

obj_logger.init "test", log4VBA.ERRO, "test_destination"
obj_logger.logFilePath = "C:\Users\czJiRost\Desktop\log-format-example.xlsx"



log4VBA.init
log4VBA.add_logger obj_logger



Set message = New MSG
log4VBA.error "test_destination", message.source("testing_module").text("Problém je mezi židlí a klávesnicí")
Set message = Nothing


End Function

