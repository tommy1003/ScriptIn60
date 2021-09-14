Option Explicit
'>>>>>> Tiktok.com/@scriptin60
'>>>>>> Follow to get more tutotials
'Reformat wrong date type in Excel with VBA Macro

'Fix Wrong Date format in Excel
'Let's start with this one
'Basic, suitable for small dataset, not recommended
Function ScriptIn60Basic()
    Dim vCell As Range, DateRg As Range
    With Sheets(1) 'Sheet1
        Set DateRg = .Range("B2:B7") ' Set the range
        For Each vCell In DateRg ' loop through each cell
            vCell.Value = Format(vCell.Value, "mm/dd/yyyy") 'Reformat
        Next
    End With
End Function

' Advanced, more efficient in large dataset
' And try this one
Function ScriptIn60Advanced()
   Dim vArr As Variant, i As Integer
   With Sheets(2) ' Sheet2
        vArr = .Range("B2:B7").Value ' Put values in an array
        For i = 1 To UBound(vArr) ' Loop through array
            vArr(i, 1) = Format(vArr(i, 1), "mm/dd/yyyy") ' Reformat
        Next
        .Range("B2:B7").Value = vArr ' Paste the values back
    End With
End Function
