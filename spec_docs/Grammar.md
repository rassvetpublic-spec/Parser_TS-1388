# Grammar for TS-1388 Parser

Документ описывает грамматику строк заказа типа TS-1388, а также общие правила работы парсера TS-1388 Parser.

Парсер реализован на VBA (Excel 2016) русский интерфейс, использует массивы:
- ParamName(1..20)
- ParamValue(1..20)
- ParamStart(1..20)
- ParamEnd(1..20)
- ParamErrorCode(1..20)

И глобальную строку WorkString (нормализованный текст заказа).

Логика пошагового парсинга и порядок вызова функций строго определены ТЗ.  
Содержание параметров (типовые значения, диапазоны, смысл) — ориентировано на форму заказа TS-1388 и реальные примеры строк.

---

## 1. Общая структура строки заказа

Строка заказа представляет собой последовательность параметров, разделённых символом `/` (возможны пробелы вокруг разделителя).

Нормализация строки
Перед парсингом исходная строка должна быть нормализована функцией Normalize.
Шаги нормализации:
Удалить ведущие и хвостовые пробелы.
Заменить длинные тире (—, –) на обычный дефис -.
Заменить все последовательности из двух и более пробелов на один пробел.
Удалить пробелы вокруг разделителя /:
" /" > "/"
"/ " > "/"
Заменить русские буквы на латиницу.
Привести всю строку к верхнему регистру (UCase).
Результат записывается в WorkString.

Пример работы Функции "Parser":
NormalString: "TS-1388/ A EXD B V3/ 11-1М/ 2NU/ 1EX DB IIВ T4 GB X / PT100/ -50...+350/ 250/ 5 (M8X1)/ 1,5/ KMMFE/ B/ -/ MIT8/ №3/ GP/ TU 4211-012-13282997-2014/ COMENT"
ParamValue(№):
1.	TYPE: TS-1388
2.	ISPOLN: AEXDBV3
3.	MODEL: 11-1M
4.	AES: 2NU
5.	EX: 1EXDBIIBT4GBX
6.	SENSOR: PT100
7.	T_LOW: -50 
8.	T_HIGH: +350
9.	DLINA: 250
10.	DIAMETR: 5 
11.	SHTUCER: M8X1
12.	LKABEL: 1,5
13.	TYPEKABEL: KMMFE
14.	CLASS: B
15.	HEAD: —
16.	PLUG: MIT8
17.	SCHEMA: N3
18.	GP: GP
19.	TU: TU4211-012-13282997-2014
20.	EXTRA: COMENT
WorkString: "#1_TS-1388%#2_AEXDBV3%#3_11-1M%#4_2NU%#5_1EXDBIIBT4GBX%#6_PT100%#7_-50%#8_+350%#9_250%#10_5%#11_M8X1%#12_1,5%#13_KMMFE%#14_B%#15_-%#16_MIT8%#17_N3%#18_GP%#19_TU4211-012-13282997-2014%#20_COMENT%"

Важно:
Фактический порядок параметров в строке должен следовать форме заказа.
Параметры строки по позиции (по форме заказа)
TYPE/ ISPOLN/ MODEL/ AES/ EX/ SENSOR/ T_LOW/ T_HIGH/ DLINA/ DIAMETR/ SHTUCER/ LKABEL/ TYPEKABEL/ CLASS/ HEAD/ PLUG/ SCHEMA/ GP/ TU/ EXTRA

Но, для однозначного нахождения при большой вариабельности функции парсера вызываются в строгом порядке:
SENSOR, TYPE, ISPOLN, EX, AES, MODEL, T_LOW и T_HIGH, SCHEMA, DLINA, DIAMETR и SHTUCER, GP, LKABEL, CLASS, TYPEKABEL, HEAD, PLUG, TU, EXTRA


Каждая функция:
работает с WorkString в её текущем состоянии,
при успехе:
находит свой параметр,
заполняет ParamValue/ParamStart/ParamEnd,
вызывает замену фрагмента на токен, 




TYPE
Допустимые формы до нормализации:
TS-1388
TS1388
TS 1388

варианты с лишними пробелами и тире.

После нормализации:

единый шаблон TS-1388.

Правило:

Если TYPE отсутствует, парсер должен добавить тип по умолчанию (TS-1388) в начало WorkString.

3.2 ISPOLN (исполнение)

Это первый значимый параметр в строке после TYPE (или сразу в начале, если TYPE нет).

Примеры из реальных строк:
- (нет исполнения)
B F3
B G2
B V3
BC
EX
EXB F3
EXB G2
EXB V3
EXBC
АВ V3
EXВ V3
EXВС
EXD
EXD B F3
EXD B G2
EXD B V3
EXD BC
AEXD
AEXD B F3
AEXD B G2
AEXD B V3
AEXD BC

4.2 MODEL (конструктивный номер исполнения)

Примеры значений:
1
3
5
5ShM
8-1
11
11PLT164
12
13M
21
25-1
3TKP
1-1M
2-2
2-1
2-3
1-1
Особенности:
Могут быть чистые числа (1, 3, 5, 11, 21).
Могут быть буквенно-цифровые коды (3TKP, 5SHM, 13M, 11PLT164, 11-1M, 2-2, 2-1M, 2-3, 8-1).

4.3 AES — класс безопасности для АЭС
По примерам:
-
2N
3N
2NU
3NU
2
3
4


4.4 EX (взрывозащита)
Примеры:
-
0ЕХ IA IIС T5 GA X
1ЕХ DB IIВ T3 GB X
0ЕХ IA IIВ T4 GA X
1ЕХ DB IIС T5 GB X

Особенности:
Может быть - (нет взрывозащиты).
Может содержать сложный текст с типом, группой, температурным классом.
Парсер должен:
найти фрагмент 0EX…;
склеить все пробелы;
сохранить нормализованное значение в ParamValue.

4.5 SENSOR / НСХ первичного преобразователя
Примеры:
100P
100M
50M
50P
PT100
PT500
PT1000

4.6 температурный диапазон (P_T_LOW / P_T_HIGH)
Формат:
-5,5...+200,7
-196...+500
-60...160
-50...+0
-196...0
Особенности:
Три точки могут быть как ..., так и символ многоточия после нормализации.
Нижняя и верхняя границы включают знак (-50, +200 и т.п.).
Парсер должен выделить отдельно:
T_LOW
T_HIGH

Контроль диапазона — по инженерным правилам (например, T_LOW < T_HIGH, значения в допустимых пределах).



5. 

при ошибке:
устанавливает ParamErrorCode(i),
возвращает код ошибки (см. ErrorCodes.md).

6. Примеры корректных строк

Ниже приведены реальные примеры строк, соответствующих форме заказа и поддерживаемых грамматикой (до нормализации пробелов/букв и т.п.):

-/ 1/ -/ -/ 100П/ -50...+200/ 20/ 5(М8Х1)/ 1,5/ КММФЭ/ C/ -/ -/ №2/ ГП/ -/ -
-/ 11/ -/ -/ PT100/ -196...+500/ 120/ 4/ 6/ КМНЭ/ B/ -/ МИТ8/ №3/ ГП/ -/ -
B V3/ 13М/ -/ -/ PT100/ -60...+160/ 30Х10Х3/ -/ 1/ МС-16-13/ B/ -/ -/ №2/ ГП/ -/ -
ТС-1388/B V3/ 2-2/ -/ 100П/ -50...+200/ 30/ 8/ 7/ КММФЭ/ B/ -/ -/ №3/ ГП/ ТУ 4211-012-13282997-2014/ -
EXВ V3/ 13М/ -/ 0ЕХ IA IIВ T4 GA X/ PT100/ -60...+160/ 190Х9Х2/ -/ 0,5/ МС-16-13/ B/ -/ -/ №2/ ГП/ -/ -
АВ V3/ 1-1М/ 2/ -/ PT100/ -60...+160/ 20/ 5/ 1,2/ КММФЭ/ B/ -/ -/ №3/ ГП/ ТУ 4211-012-13282997-2014/ -


Парсер TS-1388 обязан корректно разбирать такие строки, даже если в исходных данных присутствуют:
лишние пробелы, различные варианты написания тире и многоточий.

Стратегия разбора без Split:
При нахождении каждого параметра в строке WorkString запоминаем точку начала ParamStart(№) и конца параметра ParamEnd(№), ParamErrorCode(№)=255 (найден но не очищен). Всегда ищем самое длинное вхождение маски в строке.

Выполняй в строгой последовательности (Порядок обработки: №параметра ParamName(№): условия… NextParam – перейти на обработку следующего параметра):
•	№6 SENSOR: Если нашли в WorkString нормализованное значение ячеек из столбца ($M$12:$M$19 на листе с именем «1», уникальный точный список). ReplaceParam №6. NextParam. Если не нашли ошибка «Не определен НСХ». СТОП.
•	№1 TYPE: Если не нашли по шаблону "^TS[\-\s]*1388" в WorkString то TYPE=«TS-1388». WorkString=TYPE+WorkString. Ищем в WorkString шаблон "^TS[\-\s]*1388" ReplaceParam №1.
•	№2 ISPOLN: ищем самое длинное совпадение в фрагменте от WorkString от ParamEnd(1) до следующего разделителя «/» или до ParamStart(6) Regex.Pattern = "((A\s*)?(EXD?|EX)\s*(B|BC)?\s*(V3|N3|F3|G2))|(B\s*V3)|(N3)|(F3)|(G2)|-". ReplaceParam №2. NextParam. Если не нашли ошибка «Не определен вид исполнения». СТОП. . (примеры «-» «B V3» «A» «AB G2» «AEx» «A» «BC» «Exd» «AExd» «AExB V5» «ExdB G2»)
•	№5 EX:
Если в ISPOLN есть "EX" или «EXD», искать, двигаясь влево от ParamStart(6) до ParamEnd(2) самое длинное совпадение "(0|1)\s*([EeЕе][XxХх]\s*IA|EXD\s*DB)\s*II\s*([ABCАВС])\s*T[1-6]\s*([Gg][AaBb])\s*X", Если нашли (очистить EX от пробелов. ReplaceParam №5. NextParam) иначе ошибка «Не указана маркировка взрывозащиты для исп. Ex». СТОП. (примеры «0Ех ia IIВ T4 Ga X» «1Ехd db IIC T6 GB X»).
Если в ISPOLN нет "EXD|EX" искать, двигаясь влево от ParamStart(6) до ParamEnd(1) "-". Если нашли (ISPOLN= «-», ReplaceParam №5. NextParam), иначе выдавать ошибку «Параметры взрывозащиты указаны для не Ex исполнения» СТОП. (Примеры «/-/»)
•	№4 AES: 
o	Если в ISPOLN есть «A» то ищем двигаясь влево в WorkString с ParamStart(5) до ParamEnd(2) самое длинное совпадение "(2|3|4)\s*(НУ|HU|HУ|H|Н)?" если не нашли выдать ошибку «Не найден класс безопасности для исполнения АЭС» СТОП. Если нашли очистить AES от пробелов. ReplaceParam №4. NextParam. (примеры «/2/» «/3H/»).
o	Если в ISPOLN нет «A» то ищем двигаясь влево в WorkString с ParamStart(5) до ParamEnd(1) «-». ReplaceParam №4. NextParam. Если не нашли ошибка «Класс безопасности не определен». СТОП. (примеры «/-/»).
•	№3 MODEL: всё между ParamEnd(2) и ParamStart(4) самое длинное совпадение "^/?(\d+)(?:-(\d+))?/?(.+)?/?$", ReplaceParam №3. NextParam. Если шаблон не сработал – «Не определен конструктив» СТОП. (примеры «1» «1-1М» «3TKP» «13M»).
•	•№7 T_LOW и №8 T_HIGH:
(ищем только между ParamEnd(6)+1 до конца строки). В этом фрагменте ищем самый длинный match по регулярному выражению: "([+-]?\d+(?:,\d+)?[\s\.…+-]+[+-]?\d+(?:,\d+)?)" (поддержка чисел с запятой, знаками ±, разделённых пробелами, точками, многоточием … или их комбинациями).
Если match найден и содержит ≥2 числа — извлекаем первое как T_LOW (№7), второе — как T_HIGH (№8). ReplaceParam 8: WorkString = Left(WorkString, ParamStart(8) - 2) & Mid(WorkString, ParamStart(8)): ParamEnd(7) = ParamStart(8) – 2: ReplaceParam 7. NextParam. Если не нашли – ошибка! СТОП. (примеры «-65,3…500» «0...+200» «-60..+350».
•	№17 SCHEMA: Ищем с ParamEnd(8) (?:№|N)([1-6]) ReplaceParam №17. NextParam. (примеры «N2» «N6»).
•	№9 DLINA: Ищем с ParamEnd(8) до следующего разделителя «/» или до ParamStart(17) ((число без знака) или строку вида «Число (не обязательный пробел) «X» (не обязательный пробел) число (не обязательный пробел) «X» (не обязательный пробел) число») (Все это варианты одного DLINA). Если нашли ReplaceParam №9. NextParam. Если не нашли ошибка «Длина не определена». СТОП. (примеры «30» «30х10х3» «50,5».
•	№10 DIAMETR и №11 SHTUCER:
Если DLINA содержит «X» то {DIAMETR = «-» ReplaceParam №10, SHTUCER = «-» ParamStart(11)=ParamEnd(10)+1 ParamEnd(11)=ParamStart(11)}
Если DLINA не содержит «X» то {ищем с ParamEnd(8) до ParamStart(17) или (число «--->» число) или (число DIAMETR) (необязательный пробел) "(" (необязательный пробел) (строка SHTUCER) (необязательный пробел) ")". ReplaceParam #10. Если SHTUCER не найден, то (SHTUCER= «-» ParamStart(11)=ParamEnd(10)+1 ParamEnd(11)=ParamStart(11))} ReplaceParam №11. NextParam. Если не нашли ошибка «Длина не определена». СТОП. (примеры «5» «5(M6)» «10--->9»).
•	№18 GP: Ищем с ParamEnd(17) (?:-|GP)$ ReplaceParam №18. NextParam. (примеры «GP»).
•	№12 LKABEL: Ищем с ParamEnd(11) до ParamStart(17) (число без знака) или (число без знака с десятичной частью отделённой «,») или "-" ReplaceParam №12. NextParam. (примеры «2» «4,5»).
•	№14 KLASS: Ищем с ParamEnd(12) до ParamStart(17) ((пробел или «/»)(AA|A|B|C) (пробел или «/»)) ReplaceParam №14. NextParam. (примеры «AA» «A» «B» «C»).
•	№13 TYPEKABEL: ищем от ParamEnd(12) до ParamStart(14), ReplaceParam №13. NextParam.
•	№15 HEAD: Ищем с ParamEnd(14) до ParamStart(17) «(\s*\/?-\s*\/?\s*), ReplaceParam №15. NextParam. (примеры «AG-10» «ADXD»).
•	№16 PLUG: Ищем с ParamEnd(15) до ParamStart(17) (([A-ZА-Я0-9-]+|^-)) ReplaceParam №16. NextParam. (примеры «MIT8» «PLT164»).
•	№19 TU: Ищем с ParamEnd(18) «-», или (TU\s\d{4}-\d{3}-\d{8}-\d{4}) ReplaceParam №19. NextParam. (примеры «-» «TU 4211-012-13282997-2014»).
•	№20 EXTRA: всё после ParamEnd(19), ReplaceParam. End Sub Parser.

СТОП – Записать в ParamErrorCode(№), Вывести вообще всё что можно в debug.print и stopmsgerror. NextParam.


пример работающего кода: 

 ' Порядок строго по ТЗ:

    ' ----- 6. P_HCX (#6) -----
    Debug.Print LOG_PREFIX & "Шаг #6: Поиск P_HCX..."
    If Not Find_HCX_FromSheet(6) Then
        ParamErrorCode(6) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не определен НСХ (P_HCX).") Then Parser = 106: Exit Function
    Else
        Debug.Print LOG_PREFIX & "#6 найдено: " & ParamValue(6)
        ReplaceParam 6
    End If

    ' ----- 1. P_TYPE (#1) -----
    Debug.Print LOG_PREFIX & "Шаг #1: P_TYPE..."
    If Not RegexTest("^TS[\-\s]*1388", WorkString, False) Then
        ' добавить дефолт в начало
        ParamValue(1) = "TS-1388"
        ParamStart(1) = 1
        ParamEnd(1) = Len(ParamValue(1))
        ParamErrorCode(1) = ERR_FOUND_NOT_CLEANED
        WorkString = ParamValue(1) & "/" & WorkString
        Debug.Print LOG_PREFIX & "Добавлен дефолт P_TYPE в начало: " & ParamValue(1)
    End If
    If FindRegexInWork("^TS[\-\s]*1388", 1) Then
        ParamValue(1) = "TS-1388"
        Debug.Print LOG_PREFIX & "#1 найден: " & ParamValue(1)
        ReplaceParam 1
    Else
        ParamErrorCode(1) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден шаблон P_TYPE (TS-1388).") Then Parser = 101: Exit Function
    End If

    ' ----- 2. P_ISPOLN (#2) -----
    Debug.Print LOG_PREFIX & "Шаг #2: P_ISPOLN..."
    If Find_ISPOLN(2) Then
        Debug.Print LOG_PREFIX & "#2 найден: " & ParamValue(2)
        ReplaceParam 2
    Else
        ParamErrorCode(2) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не определен вид исполнения (P_ISPOLN).") Then Parser = 102: Exit Function
    End If

    ' ----- 5. P_EX (#5) -----
    Debug.Print LOG_PREFIX & "Шаг #5: P_EX..."
    If InStr(1, ParamValue(2), "EX", vbTextCompare) > 0 Or InStr(1, ParamValue(2), "EXD", vbTextCompare) > 0 Then
        If Find_EX_Label(5, 6, 2) Then
            Debug.Print LOG_PREFIX & "#5 найден(Ex): " & ParamValue(5)
            ReplaceParam 5
        Else
            ParamErrorCode(5) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Не указана маркировка взрывозащиты для исп. Ex (P_EX).") Then Parser = 105: Exit Function
        End If
    Else
        If FindLiteralDashInRange(5, ParamStart(6), ParamEnd(1)) Then
            Debug.Print LOG_PREFIX & "#5 найден (dash): " & ParamValue(5)
            ReplaceParam 5
        Else
            ParamErrorCode(5) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Параметры взрывозащиты указаны для не Ex исполнения (P_EX).") Then Parser = 150: Exit Function
        End If
    End If

    ' ----- 4. P_KL_AES (#4) -----
    Debug.Print LOG_PREFIX & "Шаг #4: P_KL_AES..."
    If InStr(1, ParamValue(2), "A", vbTextCompare) > 0 Then
        If FindRegexInRange_Longest("(2|3|4)\s*(NU|N)?", 4, ParamEnd(2), ParamStart(5)) Then
            Debug.Print LOG_PREFIX & "#4 найден (A): " & ParamValue(4)
            ReplaceParam 4
        Else
            ParamErrorCode(4) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Не найден класс безопасности для исполнения А (P_KL_AES).") Then Parser = 104: Exit Function
        End If
    Else
        If FindLiteralDashInRange(4, ParamStart(5), ParamEnd(1)) Then
            Debug.Print LOG_PREFIX & "#4 найден (dash): " & ParamValue(4)
            ReplaceParam 4
        Else
            ParamErrorCode(4) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Класс безопасности не определен (P_KL_AES).") Then Parser = 140: Exit Function
        End If
    End If

    ' ----- 3. P_MODEL (#3) -----
    Debug.Print LOG_PREFIX & "Шаг #3: P_MODEL..."
    If FindModelBetween(3, ParamEnd(2), ParamStart(4)) Then
        Debug.Print LOG_PREFIX & "#3 найден: " & ParamValue(3)
        ReplaceParam 3
    Else
        ParamErrorCode(3) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не определен конструктив (P_MODEL).") Then Parser = 103: Exit Function
    End If

    ' ----- 7 & 8. P_T_LOW (#7) и P_T_HIGH (#8) -----
    Debug.Print LOG_PREFIX & "Шаги #7/#8: Темп. диапазон..."
    If FindTemperatureRange(7, 8, ParamEnd(6) + 1, NextSeparatorPos(ParamEnd(6) + 2) - 1) Then
        ' По ТЗ: сначала ReplaceParam(8), затем ReplaceParam(7)
        Debug.Print LOG_PREFIX & "Найден температурный диапазон: " & ParamValue(7) & " .. " & ParamValue(8)
        ReplaceParam 8
        WorkString = Left(WorkString, ParamStart(8) - 2) & "/" & Mid(WorkString, ParamStart(8))
        ParamEnd(7) = ParamStart(8) - 2
        ReplaceParam 7
    Else
        ParamErrorCode(7) = ERR_NOT_DEFINED
        ParamErrorCode(8) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден температурный диапазон (P_T_LOW / P_T_HIGH).") Then Parser = 107: Exit Function
    End If
    
    ' ----- 17. P_CXEMA (#17) -----
    Debug.Print LOG_PREFIX & "Шаг #17: P_CXEMA..."
    If FindRegexInRange_Longest("(?:№|N)([1-6])", 17, ParamEnd(8) + 1, Len(WorkString)) Then
        Debug.Print LOG_PREFIX & "#17 найден: " & ParamValue(17)
        ReplaceParam 17
    Else
        ParamErrorCode(17) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найдена схема (P_CXEMA).") Then Parser = 117: Exit Function
    End If

    ' ----- 9. P_DLINA (#9) -----
    Debug.Print LOG_PREFIX & "Шаг #9: P_DLINA..."
    If FindDLINA(9, ParamEnd(8) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#9 найден: " & ParamValue(9)
        ReplaceParam 9
    Else
        ParamErrorCode(9) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Длина не определена (P_DLINA).") Then Parser = 109: Exit Function
    End If

    ' ----- 10 & 11. P_DIAMETR (#10) и P_SHTUCER (#11) -----
    Debug.Print LOG_PREFIX & "Шаги #10/#11: Диаметр / Штуцер..."
    If InStr(1, ParamValue(9), "X", vbTextCompare) > 0 Then
        ParamValue(10) = "-"
        ParamStart(10) = ParamEnd(9) + 1
        ParamEnd(10) = ParamStart(10)
        ParamErrorCode(10) = ERR_OK
        ParamValue(11) = "-"
        ParamStart(11) = ParamEnd(10) + 1
        ParamEnd(11) = ParamStart(11)
        ParamErrorCode(11) = ERR_OK
        Debug.Print LOG_PREFIX & "#10/#11: поставлены '-' (P_DLINA содержит X)"
    Else
        If FindDiameterAndShtucer(10, 11, ParamEnd(8) + 1, ParamStart(17) - 1) Then
            Debug.Print LOG_PREFIX & "#10 найден: " & ParamValue(10)
            ReplaceParam 10
            If ParamValue(11) <> "" Then ReplaceParam 11
        Else
            ParamErrorCode(10) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Диаметр или штуцер не определены (P_DIAMETR / P_SHTUCER).") Then Parser = 110: Exit Function
        End If
    End If

    ' ----- 18. P_GP (#18) -----
    Debug.Print LOG_PREFIX & "Шаг #18: P_GP..."
    If FindRegexInRange_Longest("(?:-|GP)", 18, ParamEnd(17) + 1, Len(WorkString)) Then
        Debug.Print LOG_PREFIX & "#18 найден: " & ParamValue(18)
        ReplaceParam 18
    Else
        ParamErrorCode(18) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найдено обозначение GP (P_GP).") Then Parser = 118: Exit Function
    End If

    ' ----- 12. P_L_KAB (#12) -----
    Debug.Print LOG_PREFIX & "Шаг #12: P_L_KAB..."
    If FindRegexInRange_Longest("^\-?\d+(?:,\d+)?$|^-$", 12, ParamEnd(11) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#12 найден: " & ParamValue(12)
        ReplaceParam 12
    Else
        ParamErrorCode(12) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден L_KAB (P_L_KAB).") Then Parser = 112: Exit Function
    End If

    ' ----- 14. P_KLASS (#14) -----
    Debug.Print LOG_PREFIX & "Шаг #14: P_KLASS..."
    If FindRegexInRange_Longest("(AA|A|B|C)", 14, ParamEnd(12) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#14 найден: " & ParamValue(14)
        ReplaceParam 14
    Else
        ParamErrorCode(14) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден класс (P_KLASS).") Then Parser = 114: Exit Function
    End If

    ' ----- 13. P_KABEL (#13) -----
    Debug.Print LOG_PREFIX & "Шаг #13: P_KABEL..."
    If FindBetweenAndReplace(13, ParamEnd(12) + 1, ParamStart(14) - 1) Then
        Debug.Print LOG_PREFIX & "#13 найден: " & ParamValue(13)
        ReplaceParam 13
    Else
        ParamErrorCode(13) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден кабель (P_KABEL).") Then Parser = 113: Exit Function
    End If

    ' ----- 15. P_HEAD (#15) -----
    Debug.Print LOG_PREFIX & "Шаг #15: P_HEAD..."
    If FindRegexInRange_Longest("(\s*\/?-\s*\/?\s*)", 15, ParamEnd(14) + 1, ParamStart(17) - 1) Then
        ' заменяем на '-'
        ParamValue(15) = "-"
        ReplaceParam 15
        Debug.Print LOG_PREFIX & "#15 заменён на '-'"
    Else
        ParamErrorCode(15) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден P_HEAD.") Then Parser = 115: Exit Function
    End If

    ' ----- 16. P_PLUG (#16) -----
    Debug.Print LOG_PREFIX & "Шаг #16: P_PLUG..."
    If FindRegexInRange_Longest("([A-ZА-Я0-9-]+|^-)", 16, ParamEnd(15) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#16 найден: " & ParamValue(16)
        ReplaceParam 16
    Else
        ParamErrorCode(16) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден P_PLUG.") Then Parser = 116: Exit Function
    End If

    ' ----- 19. P_TU (#19) -----
    Debug.Print LOG_PREFIX & "Шаг #19: P_TU..."
    If FindRegexInRange_Longest("TU\s*\d{4}-\d{3}-\d{8}-\d{4}|^-", 19, ParamEnd(18) + 1, Len(WorkString)) Then
        If Trim$(ParamValue(19)) = "-" Then ParamValue(19) = "-"
        ReplaceParam 19
        Debug.Print LOG_PREFIX & "#19 найден: " & ParamValue(19)
    Else
        ParamErrorCode(19) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден P_TU.") Then Parser = 119: Exit Function
    End If

    ' ----- 20. P_H3 (#20) -----
    Debug.Print LOG_PREFIX & "Шаг #20: P_H3..."
    If ParamEnd(19) > 0 And ParamEnd(19) < Len(WorkString) Then
        Dim tailStart As Long
        tailStart = ParamEnd(19) + 1
        ParamStart(20) = tailStart
        ParamEnd(20) = Len(WorkString)
        ParamValue(20) = Trim$(Mid$(WorkString, tailStart))
        ParamErrorCode(20) = ERR_FOUND_NOT_CLEANED
        ReplaceParam 20
        Debug.Print LOG_PREFIX & "#20 найден: " & ParamValue(20)
    Else
        ParamValue(20) = ""
        ParamStart(20) = 0
        ParamEnd(20) = 0
        ParamErrorCode(20) = ERR_NOT_DEFINED
    End If
