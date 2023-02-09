Attribute VB_Name = "Module2"
Sub Messages()
Dim m, k, head, deputy, all  As Integer
Dim default, loadDefault, gov As String
Dim cell_info, cell_gov, cell_message, cell_sphere As Range
Dim phrase As Variant
Dim dict_one, dict_two As Dictionary

Set dict_one = New Dictionary
dict_one.Add "*Количество обращений*", "F3"
dict_one.Add "*Повторные*", "F4"
dict_one.Add "*Коллективные*", "F5"
dict_one.Add "*Взято на контроль*", "F6"
'dict_one.Add "*Заказное письмо*", "F8"
'dict_one.Add "*Электронная*", "F9"
'dict_one.Add "*Устная*", "F10"
dict_one.Add "*Администрация Губернатора СПб*", "F12"
dict_one.Add "*Законодательное собрание СПб*", "F13"
dict_one.Add "*ИОГВ СПб*", "F14"
dict_one.Add "*МО СПб*", "F15"
dict_one.Add "*Федеральные ОВ*", "F16"
dict_one.Add "*Органы Прокуратуры*", "F17"
dict_one.Add "*Региональные ОВ*", "F18"
dict_one.Add "*Заявители*", "F19"
dict_one.Add "*Иные*", "F20"
dict_one.Add "*Заявление*", "F22"
dict_one.Add "*Предложение*", "F23"
dict_one.Add "*Жалоба*", "F24"
dict_one.Add "*Иное (запрос, необращение и т.п.)*", "F25"
dict_one.Add "*Количество лиц, привлеченных к ответственности по результатам рассмотрения обращений*", "F26"
dict_one.Add "*Количество вопросов*", "F27"
dict_one.Add "*Всего вопросов со сроком рассмотрения в отчетном периоде*", "F93"
'dict_one.Add "*Разъяснено*", "F94"
dict_one.Add "*Поддержано*", "F96"
dict_one.Add "*в том числе: меры приняты*", "F97"
dict_one.Add "*Не поддержано*", "F98"
dict_one.Add "*Дан ответ автору*", "F99"
dict_one.Add "*Оставлено без ответа*", "F100"
dict_one.Add "*Направлено по компетенции*", "F102"
dict_one.Add "*Рассмотрено с выездом на место*", "F103"
dict_one.Add "*Рассмотрено с нарушением срока*", "F104"
'dict_one.Add "*На рассмотрении*", "F105"

Set dict_two = New Dictionary
dict_two.Add "*Государство, общество, политика*", "F30"
dict_two.Add "*Конституционный строй*", "F31"
dict_two.Add "*Основы государственного управления*", "F32"
dict_two.Add "*Гражданское право*", "F33"
dict_two.Add "*Международные отношения. Международное право*", "F34"
dict_two.Add "*Индивидуальные правовые акты по кадровым вопросам, вопросам награждения, помилования, гражданства, присвоения почетных и иных званий*", "F35"
dict_two.Add "*Социальная сфера*", "F37"
dict_two.Add "*Семья*", "F38"
dict_two.Add "*Труд и занятость населения*", "F39"
dict_two.Add "*Социальное обеспечение и социальное страхование*", "F40"
dict_two.Add "*Образование. Наука. Культура*", "F42"
dict_two.Add "*Образование (за исключением международного сотрудничества)*", "F43"
dict_two.Add "*Наука (за исключением международного сотрудничества и военной науки)*", "F44"
dict_two.Add "*Культура (за исключением международного сотрудничества)*", "F45"
dict_two.Add "*Средства массовой информации (за исключением вопросов информатизации)*", "F46"
dict_two.Add "*Здравоохранение. Физическая культура и спорт. Туризм*", "F48"
dict_two.Add "*Здравоохранение (за исключением международного сотрудничества)*", "F49"
dict_two.Add "*Физическая культура и спорт (за исключением международного сотрудничества)*", "F50"
dict_two.Add "*Туризм. Экскурсии (за исключением международного сотрудничества)*", "F51"
dict_two.Add "*Экономика*", "F53"
dict_two.Add "*Финансы*", "F54"
dict_two.Add "*Хозяйственная деятельность*", "F56"
dict_two.Add "*Промышленность*", "F57"
dict_two.Add "*Геология. Геодезия и картография*", "F58"
dict_two.Add "*Использование атомной энергии. Захоронение радиоактивных отходов и материалов (за исключением вопросов безопасности)*", "F59"
dict_two.Add "*Строительство*", "F60"
dict_two.Add "*Градостроительство и архитектура*", "F61"
dict_two.Add "*Сельское хозяйство*", "F62"
dict_two.Add "*Транспорт*", "F63"
dict_two.Add "*Связь*", "F64"
dict_two.Add "*Космическая деятельность*", "F65"
dict_two.Add "*Торговля*", "F66"
dict_two.Add "*Общественное питание*", "F67"
dict_two.Add "*Бытовое обслуживание населения*", "F68"
dict_two.Add "*Внешнеэкономическая деятельность. Таможенное дело*", "F69"
dict_two.Add "*Природные ресурсы и охрана окружающей природной среды*", "F70"
dict_two.Add "*Информация и информатизация*", "F71"
dict_two.Add "*Оборона, безопасность, законность*", "F73"
dict_two.Add "*Оборона*", "F74"
dict_two.Add "*Безопасность и охрана правопорядка*", "F75"
dict_two.Add "*Уголовное право. Исполнение наказаний*", "F76"
dict_two.Add "*Правосудие*", "F77"
dict_two.Add "*Прокуратура. Органы юстиции. Адвокатура. Нотариат*", "F78"
dict_two.Add "*Жилище*", "F80"
dict_two.Add "*Общие положения жилищного законодательства*", "F81"
dict_two.Add "*Жилищный фонд*", "F82"
dict_two.Add "*Обеспечение граждан жилищем, пользование жилищным фондом, социальные гарантии в жилищной сфере (за исключением права собственности на жилище)*", "F83"
dict_two.Add "*Коммунальное хозяйство*", "F84"
dict_two.Add "*Оплата строительства, содержания и ремонта жилья (кредиты, компенсации, субсидии, льготы)*", "F85"
dict_two.Add "*Нежилые помещения. Административные здания (в жилищном фонде)*", "F86"
dict_two.Add "*Перевод помещений из жилых в нежилые*", "F87"
dict_two.Add "*Риэлторская деятельность (в жилищном фонде)*", "F88"
dict_two.Add "*Дачное хозяйство*", "F89"
dict_two.Add "*Гостиничное хозяйство*", "F90"
dict_two.Add "*Разрешение жилищных споров. Ответственность за нарушение жилищного законодательства*", "F91"

'*****************************************************************************CODE*******************************************************************************************

Call UserForm1.Show

MsgBox "Выберите файл с информацией о результатах рассмотрения обращений граждан, организаций и общественных объединений (выгрузка из ЕСЭДД)", 48
default = Application.GetOpenFilename(Title:="Выберите необходимый файл", FileFilter:="Excel Files(*.xls*), *xls*")

If default <> False Then
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    loadDefault = Dir(default)
    GetObject (loadDefault)
    
    With ThisWorkbook.Worksheets("загрузочный_файл")
    gov = Workbooks(loadDefault).Worksheets("Лист1").Range("A2").Value
    For Each cell_info In ThisWorkbook.Worksheets("справочник").Range("D1:D68")
        If cell_info Like gov & "*" Then
            For j = 3 To .Cells(Rows.Count, 1).End(xlUp).Row
                If .Range("E" & j).Offset(0, -1) <> "" Then
                    .Range("E" & j) = cell_info.Offset(0, -1).Value
                End If
            Next
            Exit For
        End If
    Next
    
    m = Workbooks(loadDefault).Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Row
    For Each cell_message In Workbooks(loadDefault).Worksheets(1).Range("B1:B" & m)
        For i = 0 To dict_one.Count - 1
            If cell_message Like "*Письменная*" Or cell_message Like "*Запись на личный прием*" Then
                .Range("F8") = .Range("F8").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like "*Электронная*" Or cell_message Like "*МЭДО*" Then
                .Range("F9") = .Range("F9").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like "*Устная*" Or cell_message Like "*Личный прием*" Then
                .Range("F10") = .Range("F10").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like "*Разъяснено*" Or cell_message Like "*На рассмотрении*" Then
                .Range("F94") = .Range("F94").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like dict_one.Keys()(i) Then
                .Range(dict_one.Items()(i)) = cell_message.Offset(0, 1).Value
                Exit For
            End If
        Next i
    Next cell_message
    
    k = Workbooks(loadDefault).Worksheets(2).Cells(Rows.Count, 2).End(xlUp).Row
    For Each cell_sphere In Workbooks(loadDefault).Worksheets(2).Range("B1:B" & k)
        For j = 0 To dict_two.Count - 1
            If cell_sphere Like dict_two.Keys()(j) Then
                .Range(dict_two.Items()(j)) = cell_sphere.Offset(0, 1).Value
                Exit For
            End If
        Next j
    Next cell_sphere
    
    If .Range("F26") = "" Then
        .Range("F26") = 0
    End If
         
    head = InputBox("Количество граждан, принятых на личных приемах РУКОВОДИТЕЛЕМ ИСПОЛНИТЕЛЬНОГО ОРГАНА", Title:="Введите целое число")
    .Range("F107") = head
    deputy = InputBox("Количество граждан, принятых на личных приемах ЗАМЕСТИТЕЛЯМИ РУКОВОДИТЕЛЯ ИСПОЛНИТЕЛЬНОГО ОРГАНА", Title:="Введите целое число")
    .Range("F108") = deputy
    all = InputBox("Количество граждан, принятых на личных приемах РУКОВОДИТЕЛЕМ и ЗАМЕСТИТЕЛЯМИ РУКОВОДИТЕЛЯ ИСПОЛНИТЕЛЬНОГО ОРГАНА", Title:="Введите целое число")
    .Range("F109") = all
    
    End With
    Workbooks(loadDefault).Close (False)
    
    Workbooks.Add
    ThisWorkbook.Worksheets("загрузочный_файл").Range("A1:F109").Copy
    Workbooks("Книга1").Worksheets(1).Range("A1:F109").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    ThisWorkbook.Worksheets("Лист1").Cells(1, 100).Clear
    ThisWorkbook.Worksheets("Лист1").Cells(1, 101).Clear
    ThisWorkbook.Close (False)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
End If

End Sub
