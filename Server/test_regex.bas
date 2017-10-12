Attribute VB_Name = "test_regex"
Option Explicit

Private Sub simpleRegex()
    'Dim strPattern As String: strPattern = "^[0-9]{1,2}"
    Dim strPattern As String: strPattern = "^$"
    Dim strReplace As String: strReplace = "ahoj"
    Dim regEx As New RegExp
    Dim strInput As String
    Dim Myrange As Range

    Set Myrange = ActiveSheet.Range("M1")
    
    If strPattern <> "" Then
        strInput = Myrange.Value

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            MsgBox (regEx.Replace(strInput, strReplace))
        Else
            MsgBox ("Not matched")
        End If
    End If
End Sub
