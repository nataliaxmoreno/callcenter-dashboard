Attribute VB_Name = "datacleaning"
Sub getDataFromWbs()

Dim wb As Workbook, ws As Worksheet
Set fso = CreateObject("Scripting.FileSystemObject")
Set fldr = fso.GetFolder("C:\data-projects\call center project\")
y = ThisWorkbook.Sheets("sheet1").Cells(Rows.Count, 1).End(xlUp).Row + 1

'Loop through each file in that folder
For Each wbFile In fldr.Files
    If fso.GetExtensionName(wbFile.Name) = "xlsx" Then
      Set wb = Workbooks.Open(wbFile.Path)
      For Each ws In wb.Sheets
          wsLR = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
          'Loop through each record (row 2 through last row)
          For x = 2 To wsLR
            ThisWorkbook.Sheets("sheet1").Cells(y, 1) = ws.Cells(x, 1)
            ThisWorkbook.Sheets("sheet1").Cells(y, 2) = ws.Cells(x, 2)
            ThisWorkbook.Sheets("sheet1").Cells(y, 3) = ws.Cells(x, 3)
            ThisWorkbook.Sheets("sheet1").Cells(y, 4) = ws.Cells(x, 4)
            ThisWorkbook.Sheets("sheet1").Cells(y, 5) = ws.Cells(x, 5)
            ThisWorkbook.Sheets("sheet1").Cells(y, 6) = ws.Cells(x, 6)
            ThisWorkbook.Sheets("sheet1").Cells(y, 7) = ws.Cells(x, 7)
            ThisWorkbook.Sheets("sheet1").Cells(y, 8) = CDate(ws.Cells(x, 8))
            ThisWorkbook.Sheets("sheet1").Cells(y, 9) = ws.Cells(x, 9)
            ThisWorkbook.Sheets("sheet1").Cells(y, 10) = ws.Cells(x, 10)
            ThisWorkbook.Sheets("sheet1").Cells(y, 11) = ws.Cells(x, 11)
            ThisWorkbook.Sheets("sheet1").Cells(y, 12) = ws.Cells(x, 12)
            ThisWorkbook.Sheets("sheet1").Cells(y, 13) = ws.Cells(x, 13)
            ThisWorkbook.Sheets("sheet1").Cells(y, 14) = ws.Cells(x, 14)
            ThisWorkbook.Sheets("sheet1").Cells(y, 15) = ws.Cells(x, 15)
            ThisWorkbook.Sheets("sheet1").Cells(y, 16) = ws.Cells(x, 16)
            
            y = y + 1
          Next x
          
          
      Next ws
      
      wb.Close
    End If

Next wbFile

End Sub

Sub deletenullrows()
lastrow = ThisWorkbook.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
     If IsEmpty(Cells(i, 10)) = True Then
     Rows(i).EntireRow.Delete
     End If
Next i

End Sub
Sub genderdata()
lastrow = ThisWorkbook.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
     If Cells(i, 3).Value = "Agender" Or Cells(i, 3).Value = "Bigender" Or Cells(i, 3).Value = "Genderfluid" Or Cells(i, 3).Value = "Polygender" Or Cells(i, 3).Value = "Genderqueer" Then
        Cells(i, 3).Value = "Non-binary"
     End If
Next i

End Sub

Sub changetime()
lastrow = ThisWorkbook.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
     Range("I" & i).Value = (Range("I" & i).Value / 1440)
     Range("I" & i).NumberFormat = "h:mm:ss AM/PM;@"
     
Next i

End Sub

Sub deleteextraid()
Columns("D").Delete

End Sub
Function Number_of_Words(Text_String As String) As Integer
'Function counts the number of words in a string
'by looking at each character and seeing whether it is a space or not
Number_of_Words = 0
Dim String_Length As Integer
Dim Current_Character As Integer

String_Length = Len(Text_String)

For Current_Character = 1 To String_Length

If (Mid(Text_String, Current_Character, 1)) = " " Then
    Number_of_Words = Number_of_Words + 1
End If

Next Current_Character
End Function
Sub countwords()
'Range("L1").EntireColumn.Insert
'Range("L1").Value = "word_count"
lastrow = ThisWorkbook.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
     Range("L" & i).Value = Number_of_Words(Range("K" & i).Value)
     
Next i
End Sub

Sub calldurationinminutes()
lastrow = ThisWorkbook.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 6).Value = "call_duration_min"
For i = 2 To lastrow
    Cells(i, 6).Value = Cells(i, 6).Value / 60
Next i
End Sub

Sub changing()
lastrow = ThisWorkbook.Sheets("data").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
If Cells(i, 12).Value = "resolved" Then Cells(i, 16).Value = "Very satisfied"

Next i

End Sub

