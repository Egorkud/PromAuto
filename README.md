# PromAuto

### Project initialisation
1. pip install -r requirements.txt
2. Use init_project() func in main

### How to work
Uncomment necessary functions to begin the work

### Guide notes
1. Create models table (large_import_data_to_excel, key_generator(if necessary))
2. Get photo data (get_photo_data)
3. Make descriptions ru, ukr (if necessary)
4. Input necessary data to data_changes
5. Generate data import file (prom_autofill_generator_with_differences or no differences)
6. Check all generated data


## Useful visual basic scripts for excel
### Put ';' around years
<pre>
Sub AddSemicolonsToYears()
    Dim cell As Range
    Dim yearPattern As Object
    Dim matches As Object
    Dim regex As Object
    Dim str As String
    
    ' Створюємо регулярний вираз для пошуку років у форматі ####-####
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\(\d{4}-\d{4}\)"
    regex.Global = True

    For Each cell In Selection
        str = cell.Value
        
        ' Знаходимо всі збіги
        Set matches = regex.Execute(str)
        
        ' Для кожного знайденого збігу додаємо ";" перед і після
        For Each match In matches
            str = Replace(str, match.Value, ";" & match.Value & ";")
        Next match
        
        cell.Value = str
    Next cell
End Sub

</pre>

### Delete empty cells
<pre>
1. ctrl + a
2. ctrl + g
3. alt + s (or click special)
4. choose option "blanks"
5. delete

Video:
https://www.youtube.com/watch?v=dWKSN4qplV0
</pre>

### Delete empty cells with left step
<pre>
Sub SelectAlternateHorizontalRows()
    Dim i As Long
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim rng As Range

    ' Визначаємо останню колонку з даними
    lastColumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Визначаємо останній рядок із даними
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

    ' Початковий рядок для вибору
    Dim startRow As Long
    startRow = 2 ' Починаємо з рядка 2

    ' Кількість рядків у кожному блоці для вибору
    Dim rowsPerGroup As Long
    rowsPerGroup = 2 ' Наприклад, вибираємо по 2 рядки

    ' Крок між групами рядків
    Dim stepRow As Long
    stepRow = 3 ' Наприклад, пропускаємо 3 рядки після кожного блоку

    ' Проходимо через кожну групу рядків до кінця документу
    For i = startRow To lastRow Step stepRow
        ' Перевіряємо, чи наступна пара рядків не виходить за межі останнього рядка
        If i + rowsPerGroup - 1 <= lastRow Then
            ' Якщо rng ще не визначений, присвоюємо йому першу пару рядків
            If rng Is Nothing Then
                Set rng = Range(Cells(i, 1), Cells(i + rowsPerGroup - 1, lastColumn))
            Else
                ' Об’єднуємо rng з наступною парою рядків
                Set rng = Union(rng, Range(Cells(i, 1), Cells(i + rowsPerGroup - 1, lastColumn)))
            End If
        End If
    Next i
    
    ' Виділяємо об’єднаний діапазон, якщо він не порожній
    If Not rng Is Nothing Then rng.Select
End Sub
</pre>

### Delete line with selected word
<pre>
Sub DeleteRowsWithWord()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim wordToFind As String
    
    wordToFind = "7 мест"
    Set ws = ThisWorkbook.Sheets("Export Products Sheet")


    For Each cell In ws.UsedRange
        If InStr(cell.Value, wordToFind) > 0 Then
            cell.EntireRow.Delete
        End If
    Next cell
End Sub
</pre>