Attribute VB_Name = "test_regex"
Option Explicit

Private Sub simpleRegex()
    'Dim strPattern As String: strPattern = "^[0-9]{1,2}"
    Dim strPattern As String: strPattern = "history-process-(\d{8})-(opened|closed)"
    Dim strReplace As String: strReplace = "ahoj"
    Dim regex As New RegExp
    Dim strInput As String
    Dim Myrange As Range
    Dim match As Object

    Set Myrange = ActiveSheet.Range("M1")
    
    If strPattern <> "" Then
        strInput = Myrange.Value

        With regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        'If regEx.test(strInput) Then
        Set match = regex.Execute("history-process-17082223-opened")
        Debug.Print match(0).SubMatches(0)
            'MsgBox (regEx.Replace(strInput, strReplace))
'        Else
'            MsgBox ("Not matched")
'        End If
    End If
End Sub
