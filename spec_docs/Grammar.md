# Grammar for TS/TC-1388 Parser (Excel VBA 2016)

Документ описывает грамматику и алгоритм разбора строк заказа типа TS-1388 / ТС-1388 (после нормализации — латиница).
Парсер реализован на VBA (Excel 2016, RU), работает без Split: поиск параметров выполняется в глобальной строке WorkString
с фиксацией диапазонов ParamStart/ParamEnd и последующей токенизацией ReplaceParam.

## 0. Термины и структуры данных

Глобальные переменные/массивы (индексы 1..20):

- InputString — исходная строка (как есть).
- WorkString — нормализованный текст заказа, внутри которого параметры по мере нахождения заменяются токенами.
- ParamName(i) — имя параметра (строка).
- ParamValue(i) — значение параметра.
- ParamStart(i), ParamEnd(i) — позиции (1-based) в WorkString. 0 если не найден.
- ParamErrorCode(i) — код состояния параметра.

Требование: массивы 1..20 всегда заполнены по индексам (если параметр не найден — пустое значение, позиции 0, код ошибки).

## 1. Коды состояния параметров (ParamErrorCode)

Нормативные значения:

- 0  — OK: параметр найден и очищен (токенизирован).
- 255 — FOUND_NOT_CLEANED: параметр найден, ParamStart/End заданы, но фрагмент ещё не заменён токеном (ожидает ReplaceParam).
- 127 — NOT_DEFINED: параметр не определён (обязательный не найден или нарушена логика).

Примечание: дополнительные «детальные» коды ошибок (MALFORMED_VALUE и т.п.) допускаются как расширение,
но базовые три значения обязательны и используются для жизненного цикла параметра.

## 2. Нормализация строки (NormalizeGOST)

Вход: ToNormGOST (строка). Выход: NormalizeGOST (строка).

Шаги нормализации:

1) Заменить все не-ANSI пробелы на обычный пробел:
   ChrW(160), Tab, и набор unicode-пробелов (в т.ч. 8194/8195/8201/8202/8239/8287/12288).

2) Заменить длинные тире на дефис:
   EN DASH (8211) → "-"
   EM DASH (8212) → "-"

3) Удалить пробелы вокруг разделителя "/":
   " /" → "/"
   "/ " → "/"
   Повторять, пока такие сочетания встречаются.

4) Сжать множественные пробелы:
   "  " → " " (в цикле до исчезновения).

5) Trim по краям.

6) Посимвольная замена похожих символов (ГОСТ-совместимая «визуальная» латинизация):
   К→K, Е→E, Н→H, Х→X, В→B, А→A, Р→P, О→O, С→C, М→M, Т→T и т.п. (включая нижний регистр).

7) Транслитерация оставшихся русских букв по ГОСТ (например, Ж→ZH, Ц→TS, Ш→SH и т.п.).

8) Приведение к верхнему регистру (UCase).

Результат записывается в WorkString.

Важно:
- После NormalizeGOST допускается, что в строке останется символ "№" (для схемы), знаки "+/-", запятая в дробях.
- Десятичный разделитель — запятая (",") по исходным данным.

## 3. Разделители и общая структура строки

Логическая структура заказа — параметры, разделённые символом "/".
Позиционный (формальный) порядок по форме заказа:

1  P_TYPE
2  P_ISPOLN
3  P_MODEL
4  P_KL_AES
5  P_EX
6  P_HCX
7  P_T_LOW
8  P_T_HIGH
9  P_DLINA
10 P_DIAMETR
11 P_SHTUCER
12 P_L_KAB
13 P_KABEL
14 P_KLASS
15 P_HEAD
16 P_PLUG
17 P_CXEMA
18 P_GP
19 P_TU
20 P_H3

Примечание: фактически входные строки могут быть «грязными» (лишние пробелы, тире, смешение кириллицы/латиницы),
поэтому парсер ищет параметры не Split-ом, а по маскам и диапазонам.

## 4. Токенизация (ReplaceParam) и формат токена

После нахождения параметра i:
- ParamValue(i), ParamStart(i), ParamEnd(i) заполнены
- ParamErrorCode(i) устанавливается в 255 (FOUND_NOT_CLEANED)

ReplaceParam(i) обязан:
1) Очистить ParamValue(i) от служебных символов:
   - пробелы
   - "/"
   - "(" и ")"
   (и при необходимости Trim)
2) Заменить в WorkString фрагмент [ParamStart(i)..ParamEnd(i)] на токен вида:
   #<i>_<CLEANED_VALUE>%
3) Обновить ParamEnd(i) под новую длину токена.
4) Пересчитать позиции ParamStart/ParamEnd у параметров, которые находятся ПРАВЕЕ заменённого участка и уже были найдены.
5) Установить ParamErrorCode(i)=0.

Пример:
"PT100" → "#6_PT100%"

## 5. Строгий порядок разбора (алгоритм Parser)

Парсер обязан вызывать шаги строго в следующей последовательности:

(6)  P_HCX
(1)  P_TYPE
(2)  P_ISPOLN
(5)  P_EX
(4)  P_KL_AES
(3)  P_MODEL
(7)  P_T_LOW и (8) P_T_HIGH (в одном поиске)
(17) P_CXEMA
(9)  P_DLINA
(10) P_DIAMETR и (11) P_SHTUCER
(18) P_GP
(12) P_L_KAB
(14) P_KLASS
(13) P_KABEL
(15) P_HEAD
(16) P_PLUG
(19) P_TU
(20) P_H3

При критической ошибке обязательного параметра — STOP (возврат кода Parser, заполнение ParamErrorCode).

## 6. Правила распознавания параметров (по шагам)

### 6.1. P_HCX (#6) — НСХ (обязательный)
Источник допустимых значений: лист "1", диапазон M12:M19 (уникальный список).
Поиск: найти в WorkString одно из значений (без учёта регистра).
Если не найдено: ParamErrorCode(6)=127, STOP.

### 6.2. P_TYPE (#1) — тип изделия
Шаблон в начале строки: ^TC[\-\s]*1388
(после нормализации кириллическое "ТС-1388" также должно стать "TC-1388" из-за замены С→C).
Если шаблон не найден:
- вставить "TC-1388/" в начало WorkString
- считать тип найденным и токенизировать.
Итоговое значение ParamValue(1) нормализовать к "TC-1388".

### 6.3. P_ISPOLN (#2) — исполнение
Диапазон поиска:
- от (ParamEnd(1)+1) до (ParamStart(6)-1), если HCX найден,
- иначе — до ближайшего "/" после типа.

Регулярное выражение (самое длинное совпадение):
((A\s*)?(EXD?|EX)\s*(B|BC)?\s*(V3|N3|F3|G2))|(B\s*V3)|(N3)|(F3)|(G2)|-

Если не найдено: STOP.

### 6.4. P_EX (#5) — маркировка взрывозащиты
Условие:
- если в P_ISPOLN есть "EX" или "EXD" → P_EX обязателен и должен быть найден как маркировка.
- если "EX/EXD" нет → P_EX должен быть "-" (иначе ошибка «взрывозащита для не-Ex исполнения»).

Маркировка Ex (самое длинное совпадение), поиск в диапазоне между концом исполнения и началом HCX:
(0|1)\s*([EeЕе][XxХх]\s*IA|EXD\s*DB)\s*II\s*([ABCАВС])\s*T[1-6]\s*G[AaBb]\s*X

При успехе: удалить пробелы (через ReplaceParam).

### 6.5. P_KL_AES (#4) — класс безопасности (АЭС)
Условие:
- если в P_ISPOLN присутствует "A" → параметр обязателен.
  Поиск (самое длинное совпадение) в диапазоне между концом P_ISPOLN и началом P_EX:
  (2|3|4)\s*(НУ|HU|HУ|H|Н)?
- если "A" нет → значение должно быть "-" (dash) в допустимом диапазоне слева.

Если не найдено в обязательном случае: STOP.

### 6.6. P_MODEL (#3) — конструктив
Диапазон: между окончанием P_ISPOLN и началом P_KL_AES.
Применяется шаблон к фрагменту:
^/?(\d+)(?:-(\d+))?/?(.+)?/?$
Если не распознано: STOP.

### 6.7. P_T_LOW (#7) и P_T_HIGH (#8) — температурный диапазон
Диапазон поиска: после HCX (обычно до появления следующих ключевых параметров).
Выражение для извлечения двух чисел (LOW и HIGH):
([+\-]?\d+(?:,\d+)?)\D+([+\-]?\d+(?:,\d+)?)

Если не найдено: STOP.
Токенизация по ТЗ: сначала заменяется #8, затем #7 (с корректировкой разделителя).

### 6.8. P_CXEMA (#17) — схема
Поиск после диапазона температур:
(?:№|N)([1-6])
Если не найдено: STOP.

### 6.9. P_DLINA (#9) — длина или габарит
Поиск между концом температур и началом схемы (или в окне, заданном Parser):
\d+(?:,\d+)?(?:\s*[Xxх]\s*\d+(?:,\d+)?(?:\s*[Xxх]\s*\d+(?:,\d+)?)?)?
Поддерживает:
- "250"
- "190X9X2" / "190Х9Х2"
- "50,5"

Если не найдено: STOP.

### 6.10. P_DIAMETR (#10) и P_SHTUCER (#11)
Если P_DLINA содержит "X"/"Х" (габарит) → установить:
- P_DIAMETR = "-"
- P_SHTUCER = "-"
и считать параметр обработанным.

Иначе: распознать диаметр и опционально штуцер:
(\d+(?:,\d+)?)(?:\s*\(\s*([A-Za-z0-9\-хХ]+)\s*\))?
Примеры:
- "5"
- "5(M8X1)"

Примечание (важно): формат "10--->9" текущей маской не описан.

### 6.11. P_GP (#18)
Поиск:
(?:-|GP)
(после нормализации "ГП" обычно становится "GP")
Если не найдено: STOP.

### 6.12. P_L_KAB (#12) — длина кабеля
В допустимом диапазоне:
^\-?\d+(?:,\d+)?$|^-$
Если не найдено: STOP.

### 6.13. P_KLASS (#14) — класс точности
Поиск:
(AA|A|B|C)
Если не найдено: STOP.

### 6.14. P_KABEL (#13) — тип кабеля
Значение = всё между концом P_L_KAB и началом P_KLASS (Trim).
Если пусто: STOP.

### 6.15. P_HEAD (#15)
Поиск:
(\s*\/?-\s*\/?\s*)
Нормализовать значение к "-".
Если не найдено: STOP.

### 6.16. P_PLUG (#16)
Поиск:
([A-ZА-Я0-9-]+|^-)
Если не найдено: STOP.

### 6.17. P_TU (#19)
Поиск:
TU\s*\d{4}-\d{3}-\d{8}-\d{4}|^-
Если не найдено: STOP.

### 6.18. P_H3 (#20) — хвост / примечание
Значение = всё после конца P_TU (Trim).
Если хвоста нет — допускается пусто.

## 7. Возвращаемые коды Parser (рекомендация)

Parser возвращает 0 при успехе.
При критической ошибке — код 100+№параметра (например, 106 для P_HCX, 117 для P_CXEMA),
а для логических конфликтов допускаются отдельные коды (например, 140/150) и 999 как внутренняя ошибка.

## 8. Примеры входных строк (до нормализации)

(Список примеров должен проходить нормализацию и разбор согласно правилам выше.)

- "/-/ 1/ -/ -/ 100П/ -50...+200/ 20/ 5(М8х1)/ 1,5/ КММФЭ/ C/ -/ -/ №2/ ГП/ -/ -"
- "ТС-1388/ B V3/ 11PLT164/ -/ -/ PT100/ -50...+350/ 250/ 3/ -/ -/ B/ -/ PLT164/ №3/ ГП/ -/ -"
- "ExВ V3/ 1-1/ -/ 0Ех ia IIВ T4 Ga X/ Pt100/ -50...+200/ 20/ 5(М8х1)/ 5/ КММФЭ/ B/ -/ -/ №5/ ГП/ -/ -"


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
