Attribute VB_Name = "Module2"
Sub clearAll()

For Each ws In Worksheets

ws.Range("I:P").Value = Null
ws.Range("I:P").Interior.ColorIndex = 0

Next ws

MsgBox "Results Cleared"

End Sub
