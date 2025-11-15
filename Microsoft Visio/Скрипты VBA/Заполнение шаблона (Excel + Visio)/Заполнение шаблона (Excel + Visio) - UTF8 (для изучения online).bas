Attribute VB_Name = "Module1"
Sub sub_report_visio()
    Dim VisioApp As Visio.Application
    Dim VisioDoc As Visio.Document
    Dim VisioDocNew As Visio.Document
    Dim XLSheet As Worksheet
    Dim PageObj As Visio.Page
    Dim ShapeObj As Visio.Shape
    Dim i As Long, j As Long
    Dim s_path_temp As String
    Dim s_files_word As String
    Dim s_search_pattern As String
    Dim i_cnt_row_in_one_file As Integer
    Dim i_cnt_row_in_one_file_current As Integer
    
    i_cnt_row_in_one_file = 4         ' Количество строк на один файл
    i_cnt_row_in_one_file_current = 0 ' Текущая партия
    
    ' Инициализация Visio
    On Error Resume Next
        Set VisioApp = GetObject(, "Visio.Application")
        If VisioApp Is Nothing Then
            Set VisioApp = CreateObject("Visio.Application")
        End If
    On Error GoTo 0
    
    Set XLSheet = ActiveSheet
    Application.ScreenUpdating = False
    VisioApp.ScreenUpdating = False
    
    ' Полный путь к шаблону
    s_path_temp = XLSheet.Parent.Path & "\Шаблон.vsd"
    
    ' Проверка существования файла шаблона
    If Dir(s_path_temp) = "" Then
        MsgBox "Файл не найден:" & vbCrLf & s_path_temp, vbCritical
        Exit Sub
    End If
    
    ' Находим страницу "Шаблон"
    On Error Resume Next
        Set VisioDoc = VisioApp.Documents.Open(s_path_temp)
        Set PageObj = VisioDoc.Pages("Шаблон")
        If PageObj Is Nothing Then
            MsgBox "Страница 'Шаблон' в файле Visio не найдена!", vbCritical
            VisioDoc.Close
            Exit Sub
        End If
        VisioDoc.Close
    On Error GoTo 0
    
    For i = 4 To 3 + WorksheetFunction.CountA(XLSheet.Range("A4:A10000")) ' Цикл по строкам
        If XLSheet.Cells(i, 1).Value <> "" Then
            ' Открываем шаблон только для первой строки в группе из 4
            If (i - 4) Mod i_cnt_row_in_one_file = 0 Then
                ' Закрываем предыдущий файл, если он есть
                If Not VisioDocNew Is Nothing Then
                    VisioDocNew.Save
                    VisioDocNew.Close
                    Set VisioDocNew = Nothing
                End If
                
                ' Открываем шаблон VSD
                Set VisioDoc = VisioApp.Documents.Open(s_path_temp)
                
                ' Сохраняем копию
                i_cnt_row_in_one_file_current = i_cnt_row_in_one_file_current + 1
                VisioDoc.SaveAs XLSheet.Parent.Path & "\" & "Группа_" & i_cnt_row_in_one_file_current & ".vsd"
                Set VisioDocNew = VisioApp.Documents("Группа_" & i_cnt_row_in_one_file_current & ".vsd")
            End If
            ' Замена текста во всех фигурах на странице "Шаблон"
            For Each ShapeObj In VisioDocNew.Pages("Шаблон").Shapes
                If ShapeObj.Type = visTypeShape Then
                    For j = 1 To XLSheet.UsedRange.Columns.Count
                        If InStr(ShapeObj.Text, XLSheet.Cells(1, j)) > 0 Then
                            s_search_pattern = Replace(XLSheet.Cells(3, j).Value, "}", (i - 4) Mod 4 + 1 & "}")
                            ShapeObj.Text = Replace(ShapeObj.Text, s_search_pattern, XLSheet.Cells(i, j))
                        End If
                    Next j
                End If
            Next ShapeObj
        End If
    Next i
    
    ' Закрываем последний файл
    If Not VisioDocNew Is Nothing Then
        VisioDocNew.Save
        VisioDocNew.Close
    End If
  
    VisioApp.Quit
    Set VisioApp = Nothing
    Application.ScreenUpdating = True
    
    ' Определяем правильное окончание
    Select Case i_cnt_row_in_one_file_current
        Case 1: s_files_word = "файл"
        Case 2, 3, 4: s_files_word = "файла"
        Case 5 To 20: s_files_word = "файлов"
        Case Else:
            Select Case i_cnt_row_in_one_file_current Mod 10
                Case 1: s_files_word = "файл"
                Case 2, 3, 4: s_files_word = "файла"
                Case Else: s_files_word = "файлов"
            End Select
    End Select

    MsgBox "Готово. Создано " & i_cnt_row_in_one_file_current & " " & s_files_word & " Visio.", vbInformation
End Sub

Sub sub_report_XLSM_to_DOCX()
        Dim WDApp As Word.Application
        Dim WDDoc As Document
        Dim WDDocNew As Document
        Dim XLSheet As Object

        Set WDApp = New Word.Application
        Set XLSheet = ActiveSheet

        Application.ScreenUpdating = False
        WDApp.ScreenUpdating = False

        For i = 2 To WorksheetFunction.CountA(XLSheet.Range("A:A")) ' цикл по строкам
                Set WDDoc = WDApp.Documents.Open(XLSheet.Parent.Path & "Шаблон.docx")
                WDDoc.SaveAs2 (XLSheet.Parent.Path & "\" & XLSheet.Cells(i, 1) & ".docx")
                Set WDDocNew = WDApp.Documents(XLSheet.Cells(i, 1) & ".docx")
                
                For j = 1 To XLSheet.UsedRange.Columns.Count ' цикл по столбцам
                        Set MyRange = WDDocNew.Content
                        MyRange.Find.Execute FindText:=XLSheet.Cells(1, j), ReplaceWith:=XLSheet.Cells(i, j), Replace:=wdReplaceAll
                Next j

                WDDoc.Close
        Next i

        WDApp.Quit
        Set WDApp = Nothing

        MsgBox ("Готово.")
End Sub

Function f_Число_из_строки(s_input As String, i_pos As Integer) As Variant
    Dim i As Integer
    Dim v_num As Variant
    Dim s_temp As String
    Dim s_char As String
    Dim s_list_num() As String
    Dim collect_num As Collection
    
    ' Инициализация коллекции для чисел
    Set collect_num = New Collection
    
    ' Заменяем все нечисловые символы (кроме точек и минусов) на пробелы
    s_temp = ""
    
    s_input = Replace(s_input, "..", " ")
    s_input = Replace(s_input, "...", " ")
    s_input = Replace(s_input, "…", " ")
    
    For i = 1 To Len(s_input)
        s_char = Mid(s_input, i, 1)
        If IsNumeric(s_char) Or s_char = "." Or s_char = "," Or s_char = "-" Then
            s_temp = s_temp & s_char
        Else
            s_temp = s_temp & " "
        End If
    Next i
    
    ' Разделяем строку по пробелам и удаляем пустые элементы
    s_list_num = Split(Application.WorksheetFunction.Trim(s_temp), " ")

    ' Собираем только валидные числа в коллекцию
    For Each v_num In s_list_num
        If v_num <> "" And IsNumeric(Replace(v_num, ",", ".")) Then
            collect_num.Add CDbl(Replace(v_num, ",", "."))
        End If
    Next v_num

    ' Проверяем, существует ли число с указанной позицией
    If i_pos > 0 And i_pos <= collect_num.Count Then
        f_Число_из_строки = collect_num(i_pos)
    Else
        f_Число_из_строки = CVErr(xlErrNA) ' Ошибка
    End If
End Function

Function f_Единица_измерения(s_input As String) As String
    Dim i As Integer
    Dim s_char As String
    Dim b_num_finished As Boolean
    Dim s_unit As String

    s_unit = ""
    b_num_finished = False
        
    ' Удаление избыточных символов
    With CreateObject("VBScript.RegExp")
        .Pattern = "[-.,…()]|\.{2,3}"
        .Global = True
        s_input = .Replace(s_input, "")
    End With
    
    ' Проходим по каждому символу строки
    For i = 1 To Len(s_input)
        s_char = Mid(s_input, i, 1)
        
        ' Если символ НЕ является частью числа
        If Not (IsNumeric(s_char)) Then
            b_num_finished = True ' Числовая часть закончилась
        End If
        
        ' Если числовая часть закончилась и символ не пробел - добавляем к единице измерения
        If b_num_finished And s_char <> " " Then
            s_unit = s_unit & s_char
        End If
    Next i
    
    f_Единица_измерения = Trim(s_unit)
End Function

