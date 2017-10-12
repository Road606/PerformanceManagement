VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AreaStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public str_place_from As String
Public str_place_to As String

Public private_obj_area As Area

Public Property Get obj_area() As Area
    Set obj_area = private_obj_area
End Property

Public Property Let obj_area(obj_area As Area)
    Set private_obj_area = obj_area
    obj_area.add_step Me
End Property
