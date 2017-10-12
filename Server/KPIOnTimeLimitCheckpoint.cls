VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeLimitCheckpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public str_type As String
Public str_next_checkpoint_id As String
Public obj_limit As Date

Public obj_previous_limit As KPIOnTimeLimitCheckpoint
Public obj_next_limit As KPIOnTimeLimitCheckpoint
