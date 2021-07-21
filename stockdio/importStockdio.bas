Attribute VB_Name = "importStockdio"
'You can find JSON data from Stockdio.com by clicking on their "Data" tab.
'I used the "Get Latest Prices Ex" json data.

Sub importStockdioM()
Dim JsonConverter As New JsonConverter
Const rgCol As String = "A1:N13"
Const rgRow As String = "A2:N13"
Dim JSON As Object
Dim data As Dictionary
Dim prices As Dictionary
Dim columns As Collection
Dim column As Variant
Dim values(1 To 14, 1 To 14) As Variant
Dim JsonText As String
Dim Value As Variant
Dim wb As Workbook
Set wb = ThisWorkbook
Dim ws As Worksheet
Dim i As Integer
Dim m As Integer
Dim o As Integer

'enter your URL here:
JsonText = "http://localhost/dataStockdio2.json"

Set JSON = JsonConverter.ParseJson(fnGetHTTPXML(JsonText))

Let i = 1
For i = 1 To 14
With wb.Sheets("jsonSht").Range(rgCol).EntireColumn
    .AutoFit
    .Cells(i) = printColumn1(i, JSON)
End With
Next i

wb.Sheets("jsonSht").Range("A1:N1").Font.Bold = True

Let o = 0
For o = 0 To 11
    Let m = 1
    For m = 1 To 14
    wb.Sheets("jsonSht").Range(rgRow).Cells(m).Offset(o) = printRow1(m, o + 1, JSON)
    Next m
Next o

wb.Sheets("jsonSht").Range(rgRow).EntireColumn.AutoFit


End Sub

Function printColumn1(iter As Integer, jsn As Object) As String

Dim col As Variant
Dim result As Variant

col = iter

result = jsn("data")("prices")("columns")(col)
printColumn1 = result

End Function

Function printRow1(iter As Integer, z As Integer, jsn As Object) As String

Dim y As Integer
Dim ro As Variant

Dim valu As Variant

ro = iter
y = z

valu = jsn("data")("prices")("values")(y)(ro)

printRow1 = valu
End Function

'Following function from link https://github.com/VBA-tools/VBA-JSON/issues/112
Function fnGetHTTPXML(strURL As String) As String

Dim xmlHttp As Object

Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

xmlHttp.Open "GET", strURL
xmlHttp.setRequestHeader "Content-Type", "text/xml"
xmlHttp.send

fnGetHTTPXML = xmlHttp.responseText
End Function
