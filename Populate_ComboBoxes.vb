Sub Populate()
Dim i As Integer, n As Integer, s As String, today() As String, exact() As String
Application.ScreenUpdating = False
today = Split(Now())
exact = Split(today(0), "/")

CurrencyConverter.Days.Text = exact(1) & "/" & exact(0) & "/" & exact(2)

Sheets("Datos").Visible = True

Sheets("Datos").Select
n = WorksheetFunction.CountA(Columns("A:A"))

For i = 1 To n
    s = Range("A" & i) & " - " & Range("B" & i)
    CurrencyConverter.From.AddItem s
Next i
CurrencyConverter.From.Text = Range("A1") & " - " & Range("B1")

For i = 1 To n
    s = Range("A" & i) & " - " & Range("B" & i)
    CurrencyConverter.EqualTo.AddItem s
Next i
CurrencyConverter.EqualTo.Text = Range("A2") & " - " & Range("B2")

Sheets("Datos").Visible = False
End Sub