Option Compare Text 'get rid of case sensitive

Const mainSheetName As String = "Öýëëï1"
Const mesimeriana As String = "ÌÅÓÇÌÅÑÉÁÍÁ ÅÎÏÄÏ×ÁÑÔÁ"
Const prohna As String = "ÐÑÙÉÍÁ ÅÎÏÄÏ×ÁÑÔÁ"
Global sheetCurrentRow As Integer

Sub dd()
   lastDay = day(getLastDayOfMonth())
   
   sheetCurrentRow = 1
   If Not sheetExists("ÌÅÓÇÌÅÑÉÁÍÁ ÅÎÏÄÏ×ÁÑÔÁ") Then
        CreateSheet (mesimeriana)
        tmp = WriteToSheet(sheetCurrentRow, "ÅÐÉÈÅÔÏ", "ÏÍÏÌÁ", "ÁÐÏ", "ÌÅ×ÑÉ", "ÙÑÁ-ÅÎÏ", "ÙÑÁ-ÌÅÓÁ")
   Else
        tmp = EmptySheet(mesimeriana)
        tmp = WriteToSheet(sheetCurrentRow, "ÅÐÉÈÅÔÏ", "ÏÍÏÌÁ", "ÁÐÏ", "ÌÅ×ÑÉ", "ÙÑÁ-ÅÎÏ", "ÙÑÁ-ÌÅÓÁ")
   End If
   
   
   'get current day
   currentday = day(Date)
  
    Dim loopOffset As Integer 'apo posa cells arxizi proti tou minos
    Dim currentDayStatusColumn As Integer 'to column tis simerinis meras
    Dim currentDayStatus As String 'to ti exei simera o sminitis
    Dim loopStartFrom As Integer 'to cell pou prepei na arxisi to loop
    Dim maxTimesOfLoop As Integer 'poses meres mexri to telos tou mina
    Dim endOfLoopMax As Integer 'to cell pou teleiwni o minas
    Dim currentRow As Integer 'to index gia tin grammi
    Dim innerLoopIndex As Integer
    Dim firstNameColumn As Integer 'to column number pou vriskete to name
    Dim lastNameColumn As Integer 'to column number pou vriskte to epitheto
    Dim epistasia As String
    
    loopOffset = 8
    currentDayStatusColumn = loopOffset + currentday
    currentDayStatus = ""
    loopStartFrom = currentDayStatusColumn + 1
    maxTimesOfLoop = lastDay - currentday
    endOfLoopMax = loopStartFrom + maxTimesOfLoop
    currentRow = 2
    innerLoopIndex = loopStartFrom
    firstNameColumn = 4
    lastNameColumn = 3
    epistasia = ""
    
    Do While Cells(currentRow, 1).Value <> ""
        'tsekarisma gia to poia epistasia diavzei twra
        If Cells(currentRow, 1).Value = "_" Then
            epistasia = Cells(currentRow, 3)
        End If
        currentDayStatus = Trim(Cells(currentRow, currentDayStatusColumn).Value)
        
        ' to endOfLoopMax-1 einai to cell tis teleuteas meras tou mina
        ' 1.ean kapoios feugei pros to telos tou mina kai xana mpainei xana ton allo
        ' 2.tote apo auto to excel tha vgenei exodoxarto mexri to telos tou mina
        ' 3.kai meta thelei na mpeneis sto excel tou allou mina na vgazeis exodoxarta kai apo ekei gia na deis pote tha xana mpei
        If currentDayStatus = "ÄÉÅ" Then
            Do While innerLoopIndex < endOfLoopMax
                If HasDuty(currentRow, innerLoopIndex, endOfLoopMax - 1, innerLoopIndex) Then
                    sheetCurrentRow = sheetCurrentRow + 1
                    If IsWeekend(innerLoopIndex - loopOffset) Then
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(innerLoopIndex - loopOffset), "12:00", "08:00")
                    Else
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(innerLoopIndex - loopOffset), "12:00", "06:30")
                    End If
                    Exit Do
                End If
                innerLoopIndex = innerLoopIndex + 1
            Loop
        ElseIf currentDayStatus = "ÂÁÑ" Then
            Do While innerLoopIndex < endOfLoopMax
                If HasDuty(currentRow, innerLoopIndex, endOfLoopMax - 1, innerLoopIndex - loopOffset) Then
                    'Debug.Print currentDay & "  " & innerLoopIndex - loopOffset
                    sheetCurrentRow = sheetCurrentRow + 1
                    If IsWeekend(innerLoopIndex - loopOffset) Then
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(currentday), "14:00", "20:00")
                        sheetCurrentRow = sheetCurrentRow + 1
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(innerLoopIndex - loopOffset), "22:30", "08:00")
                    Else
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(currentday), "14:00", "20:00")
                        sheetCurrentRow = sheetCurrentRow + 1
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(innerLoopIndex - loopOffset), "22:30", "06:30")
                    End If
                    Exit Do
                End If
                innerLoopIndex = innerLoopIndex + 1
            Loop
        ElseIf currentDayStatus = "ÕÐ" And epistasia = "ÌÁÃÅÉÑÉÁ" Then
            Do While innerLoopIndex < endOfLoopMax
                If HasDuty(currentRow, innerLoopIndex, endOfLoopMax - 1, innerLoopIndex - loopOffset) Then
                    'Debug.Print currentDay & "  " & innerLoopIndex - loopOffset
                    sheetCurrentRow = sheetCurrentRow + 1
                    If IsWeekend(innerLoopIndex - loopOffset) Then
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(innerLoopIndex - loopOffset), "21:30", "08:00")
                    Else
                        tmp = WriteToSheet(sheetCurrentRow, Cells(currentRow, lastNameColumn).Value, Cells(currentRow, firstNameColumn).Value, CStr(currentday), CStr(innerLoopIndex - loopOffset), "21:30", "06:30")
                    End If
                    Exit Do
                End If
                innerLoopIndex = innerLoopIndex + 1
            Loop
        End If
        innerLoopIndex = loopStartFrom
        currentDayStatus = ""
        currentRow = currentRow + 1
     Loop
    
End Sub


' teleutea mera tou mina
Function getLastDayOfMonth() As Date
    dyear = Year(Now)
    dmonth = Month(Now)
    getDate = DateSerial(dyear, dmonth + 1, 0)
    
    getLastDayOfMonth = getDate
End Function

'check if sheet exists
Public Function sheetExists(sheetToFind As String, Optional InWorkBook As Workbook) As Boolean
    If InWorkBook Is Nothing Then Set InWorkBook = ThisWorkbook
    
    Dim Sheet As Object
    For Each Sheet In InWorkBook.Sheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
    sheetExists = False
End Function


Public Function CreateSheet(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
End Function

Public Function EmptySheet(sheetName As String) As Boolean
    Worksheets(sheetName).Activate
    Cells.Clear
    Worksheets(mainSheetName).Activate
    EmptySheet = True
End Function


Public Function WriteToSheet(currentRow As Integer, lastName As String, firstName As String, dayStart As String, dayEnd As String, timeOut As String, timeIn As String) As Boolean
    Worksheets(mesimeriana).Activate
    Range("a" & currentRow).Value = lastName
    Range("b" & currentRow).Value = firstName
    Range("c" & currentRow).Value = dayStart
    Range("d" & currentRow).Value = dayEnd
    Range("e" & currentRow).Value = timeOut
    Range("f" & currentRow).Value = timeIn
    Worksheets(mainSheetName).Activate
    WriteToSheet = True
End Function


Public Function IsWeekend(d As Integer) As Boolean
    Dim dDate As Date
    
    dDate = ConvertDaytToDate(d)
    
    Select Case Weekday(dDate)
        Case vbSaturday, vbSunday
            IsWeekend = True
        Case Else
            IsWeekend = False
    End Select
    
End Function

Public Function ConvertDaytToDate(d As Integer) As Date
    ConvertDaytToDate = DateSerial(Year(Date), Month(Date), d)

End Function


Public Function HasDuty(currentRow As Integer, innerLoopIndex As Integer, lastDayOfMonthCell As Integer, currentColumn As Integer) As Boolean
    If Trim(Cells(currentRow, innerLoopIndex).Value) = "ÄÉÅ" Or Trim(Cells(currentRow, innerLoopIndex).Value) = "ÓÊ" Or Trim(Cells(currentRow, innerLoopIndex).Value) = "ÕÐ" Or Trim(Cells(currentRow, innerLoopIndex).Value) = "ÂÁÑ" Then
        HasDuty = True
    Else
        If Trim(Cells(currentRow, lastDayOfMonthCell).Value) = "ÅÎÏ" And lastDayOfMonthCell = currentColumn Then
            HasDuty = True
        Else
            HasDuty = False
        End If
    End If
End Function



