Attribute VB_Name = "Main"
'--------------------------------------------------------------------------------------------------------------------------------
'Last Update: 2013-12-10
'Author:      Ruslan Salikhov
'Description:
'            - время накопления считается равным продолжительности простоя с причиной "Остановка в накопление" в отчетном месяце
'              (предварительно заполняется в TOTALS_MTH.ACC_TIME)
'            - из общего времени простоя отнимается продолжительность простоев с причиной "Остановка в накопление"
'            - при расчете уплотненных дебитов нефти/жидкости используется время работы с учетом продолжительности простоев
'              с причиной "Остановка в накопление" (сумма времени работы и времени накопления)
'--------------------------------------------------------------------------------------------------------------------------------
'Last Update: 2014-01-30
'Author:      Siraziev Ruslan RNInform
'Description:
'            - добавленно новое поле в отчет "Завод изготовитель", данные берутся из RSP_WELL_PUMP по полю ESP_MFR_TEXT
'--------------------------------------------------------------------------------------------------------------------------------
'Last Update: YYYY-MM-DD
'Author:      Krotova Anastasia RNInform
'Description:
'            - по просьбе В.Борисова реализован следующий алгоритм вычисления данных из 52-й колонки "Уплотн.скв.часы за месяц простоя"
'            с использованием данных из базы:
'            Если ((время_раб + время_накоп + время_прост) > кол_часов_в_месяце)
'            то (кол_часов_в_месяце - время_раб - время_накоп) иначе (время_прост).
'            Это сделано для того, чтобы в сумме колонки 50,51 и 52 давали календарное количество дней в месяце.
'--------------------------------------------------------------------------------------------------------------------------------
'Last Update: YYYY-MM-DD
'Author:      First_Name Last_Name
'Description: ...
'--------------------------------------------------------------------------------------------------------------------------------

Option Explicit


Function GetUnitNameByFieldID(FieldID As String) As String
'Название предприятия в зависимости от ID месторождения (для шапки отчета)
    Select Case FieldID
    Case "de84968b50f14b1c817c5a57436afc4e" 'Русское
        GetUnitNameByFieldID = "ОАО " & Chr(34) & "Тюменнефтегаз" & Chr(34)
    Case "b064f77791fd4fb89c0e5b13a4a5044f" 'Русско - Реченское
        GetUnitNameByFieldID = "ОАО " & Chr(34) & "Русско-Реченское" & Chr(34)
    Case "c40b67bb738344629ac18db421271717" 'Сузунское
        GetUnitNameByFieldID = "ОАО " & Chr(34) & "Сузун" & Chr(34)
    Case "2e376c62fe0e44028e0856e840d8db5d" 'Тагульское
        GetUnitNameByFieldID = "ООО " & Chr(34) & "Тагульское" & Chr(34)
    Case Else
        GetUnitNameByFieldID = "АО " & Chr(34) & "РОСПАН ИНТЕРНЕШНЛ" & Chr(34)
    End Select
End Function

Sub MainExec(rawParams() As Variant)
On Error GoTo LabelErr
Application.Cursor = xlIBeam

Dim cn               As ADODB.Connection ' Connection object
Dim RS               As ADODB.Recordset  ' Recordset object
Dim RS_Layer         As ADODB.Recordset  ' Recordset object

Dim StartDay         As Date             ' Дата отчета
Dim RowCounter       As Integer          ' Счетчик для движения по строкам отчета
Dim i                As Integer          ' Счетчик для движения по строкам рекордсета
Dim j                As Integer          ' Счетчик для движения по строкам рекордсета
Dim well_num         As Integer          ' Счетчик порядкового номера скважины

Dim Sql_main         As String           ' Строка sql-запроса
Dim Sql              As String           ' Строка sql-запроса
Dim SQL_MHP          As String           ' Часть строки sql-запроса

Dim CurrFIELD_name   As String           ' Текущее месторождение
Dim CurrLAYER_name   As String           ' Текущий пласт
Dim CurrWELL_name    As String           ' Текущая скважина
Dim ParField_Item_Id    As String        ' Id месторождения, выбранного пользователем
Dim ParField_Item_Name  As String        ' Name месторождения, выбранного пользователем
Dim OldFileName         As String

Dim HoursInMon       As Integer          'Количество часов в месяце
Dim HoursProd        As Double           'Часы работы
Dim HoursNakop       As Double           'Часы накопления
Dim HoursProst       As Double           'Часы простоя

Dim TitleDate        As Date             ' Дата заголовка (+1 день)

Application.ThisWorkbook.Application.Visible = False
Application.ThisWorkbook.Application.ScreenUpdating = False

Dim params As New Collection
For i = 0 To UBound(rawParams) - 1
    params.Add rawParams(i + 1), rawParams(i)
    
'    Sheets("Sheet1").Cells(100 + i, 1) = rawParams(i + 1)
'    Sheets("Sheet1").Cells(100 + i, 2) = rawParams(i)
  
    i = i + 1
Next i

Set cn = New ADODB.Connection
Dim strConn As String
strConn = "Provider=OraOLEDB.Oracle.1;" + params("CONNECTION")
cn.Open strConn

'Первая строчка после шапки
RowCounter = 14

StartDay = CDate(params("AS_OF_DATE"))
TitleDate = DateAdd("d", 1, StartDay)

If Len(params("ITEM")) > 0 Then
    'Id месторождения, выбранного пользователем
    ParField_Item_Id = Mid(params("ITEM"), InStr(params("ITEM"), ",") + 1, 32)
    'Name месторождения, выбранного пользователем
    ParField_Item_Name = Right(params("ITEM"), Len(params("ITEM")) - InStrRev(params("ITEM"), ","))
End If

With Sheets("Sheet1")

.Cells(2, 1) = "За :                   " & Choose(Month(StartDay), "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь") & " " & Year(StartDay) & " г."
.Cells(3, 1) = "Предприятие : " & GetUnitNameByFieldID(ParField_Item_Id)
.Cells(4, 1) = "Дата выдачи:  " & Day(Date) & " " & Choose(Month(Date), "Января", "Февраля", "Марта", "Апреля", "Мая", "Июня", "Июля", "Августа", "Сентября", "Октября", "Ноября", "Декабря") & " " & Year(Date) & " г."

' Зачистим хвосты, которые затем сами сгенерим
Cells(19, 6).ClearContents
Cells(20, 6).ClearContents
Cells(20, 7).ClearContents
Cells(22, 6).ClearContents
Cells(23, 6).ClearContents
Cells(23, 7).ClearContents

Sql_main = _
"select distinct WELL2UNITS.ORG_UNIT4_NAME FIELD_NAME, WELL2UNITS.ORG_UNIT4_LEGACY_ID FIELD_ID, FORMATION.ITEM_NAME LAYER_NAME, FORMATION.ITEM_ID LAYER_ID " & Chr(10) & _
"FROM ( select ITEM_ID, START_DATETIME from ITEM_EVENT where START_DATETIME=last_day(to_date(':DATAOTCHETA','DD.MM.YYYY')) and  EVENT_TYPE='RSP_STOCK_HIST') STOCK_HIST " & Chr(10) & _
"  inner join ( select w.item_id, w.ORG_UNIT4_NAME, w.ORG_UNIT4_LEGACY_ID from VCUSTOM_WELL2UNITS w inner join VI_ORG_UNIT4_RU_RU u on w.ORG_UNIT4_LEGACY_ID=u.LEGACY_ID " & Chr(10) & _
IIf(ParField_Item_Id <> "", " and u.ITEM_ID='" & ParField_Item_Id & "' ", "") & Chr(10) & _
"             ) WELL2UNITS on WELL2UNITS.ITEM_ID=STOCK_HIST.ITEM_ID " & Chr(10) & _
"  inner JOIN ( select ITEM_ID from ITEM_EVENT_EXT where START_DATETIME=trunc(to_date(':DATAOTCHETA','DD.MM.YYYY'),'mm') and EVENT_TYPE='CUM_VOL_DET' " & Chr(10) & _
"                 and VAL24/*LTD_OIL_MASS*/>0 and VAL29/*LTD_PROD_HOURS*/>0 ) MONTH_HIST ON MONTH_HIST.ITEM_ID=WELL2UNITS.ITEM_ID " & Chr(10) & _
"  LEFT JOIN VL_WELL_ZONE_RU_RU      WELL_ZONE_LINK      ON WELL2UNITS.ITEM_ID =WELL_ZONE_LINK.FROM_ITEM_ID AND WELL_ZONE_LINK.TO_ITEM_TYPE='ZONE' AND WELL_ZONE_LINK.FROM_ITEM_TYPE='COMPLETION' " & Chr(10) & _
"  LEFT JOIN VL_FORMATION_ZONE_RU_RU FORMATION_ZONE_LINK ON WELL_ZONE_LINK.TO_ITEM_ID = FORMATION_ZONE_LINK.TO_ITEM_ID AND FORMATION_ZONE_LINK.TO_ITEM_TYPE='ZONE' AND FORMATION_ZONE_LINK.FROM_ITEM_TYPE='FORMATION' " & Chr(10) & _
"  LEFT JOIN ( select i.item_id, p0.PROPERTY_STRING ITEM_NAME from ITEM i inner JOIN ITEM_PROPERTY p0 ON i.ITEM_ID=p0.ITEM_ID and i.ITEM_TYPE='FORMATION' AND p0.PROPERTY_TYPE='NAME' " & Chr(10) & _
"                AND p0.START_DATETIME<=trunc(to_date(':DATAOTCHETA','DD.MM.YYYY'),'mm') AND p0.END_DATETIME>trunc(to_date(':DATAOTCHETA','DD.MM.YYYY'),'mm') " & Chr(10) & _
"            ) FORMATION ON FORMATION.ITEM_ID=FORMATION_ZONE_LINK.FROM_ITEM_ID " & Chr(10) & _
"WHERE FORMATION.ITEM_ID is not null " & Chr(10) & _
"order by WELL2UNITS.ORG_UNIT4_NAME "
Sql_main = Replace(Sql_main, ":DATAOTCHETA", "01." & Month(StartDay) & "." & Year(StartDay))

Set RS = New ADODB.Recordset
RS.ActiveConnection = cn ' Assign the Connection object.
RS.CursorType = adOpenStatic
'.Cells(1, 8) = Sql_main
RS.Open Sql_main ' Extract the required records.

' Цикл перебора наборов месторождений-пластов
If Not (RS.BOF And RS.EOF And RS.RecordCount = 0) Then
    For i = 0 To RS.RecordCount - 1 ' бежим по рекордсету
    
    .Cells(RowCounter, 1).Font.Bold = 1
    .Cells(RowCounter + 1, 1).Font.Bold = 1
    .Cells(RowCounter + 2, 1).Font.Bold = 1
    .Cells(RowCounter, 1) = "Месторождение : " & RS.Fields("FIELD_NAME")
    .Cells(RowCounter + 1, 1) = "Объект : "
    .Cells(RowCounter + 2, 1) = "Пласт : " & RS.Fields("LAYER_NAME")
    RowCounter = RowCounter + 3
    
    CurrFIELD_name = RS.Fields("FIELD_ID")
    CurrLAYER_name = RS.Fields("LAYER_ID")
      
SQL_MHP = _
"       ( SELECT FORMATION_ZONE_LINK.TO_ITEM_ID ZONE_ID, " & Chr(10) & _
"           WELL2UNITS.ITEM_ID,                      WELL2UNITS.COMPLETION_NAME      WELL_NAME, " & Chr(10) & _
"           WELL2UNITS.ORG_UNIT4_LEGACY_ID FIELD_ID, WELL2UNITS.ORG_UNIT4_NAME       FIELD_NAME, " & Chr(10) & _
"           FORMATION.ITEM_ID              LAYER_ID, FORMATION.ITEM_NAME             LAYER_NAME, " & Chr(10) & _
"           MH.LIFT_TYPE,                            MH.LIFT_TYPE_TEXT, " & Chr(10) & _
"           MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN FORMATION.OIL_FVF END) AS OIL_FVF, " & Chr(10) & _
"           MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN FORMATION.WATER_DENSITY_1 END) AS WATER_DENSITY, " & Chr(10) & _
"           TO_DATE(':DATAOTCHETA','DD.MM.YYYY') AS START_DATETIME, " & Chr(10)
SQL_MHP = SQL_MHP & _
"           SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS MTD_OIL_MASS, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS YTD_OIL_MASS, /*MH.YTD_OIL_MASS,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS LTD_OIL_MASS, /*MH.LTD_OIL_MASS,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_VOL*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS MTD_OIL_VOL, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_VOL*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS YTD_OIL_VOL, /*MH.YTD_OIL_VOL,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_VOL*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS LTD_OIL_VOL, /*MH.LTD_OIL_VOL,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_MASS*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS MTD_WATER_MASS, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_MASS*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS YTD_WATER_MASS, /*MH.YTD_WATER_MASS,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_MASS*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS LTD_WATER_MASS, /*MH.LTD_WATER_MASS,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_VOL*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS MTD_WATER_VOL, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_VOL*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS YTD_WATER_VOL, /*MH.YTD_WATER_VOL,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_VOL*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS LTD_WATER_VOL, /*MH.LTD_WATER_VOL,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1) END) AS MTD_GAS_VOL, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1) END) AS YTD_GAS_VOL, /*MH.YTD_GAS_VOL,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1) END) AS LTD_GAS_VOL, /*MH.LTD_GAS_VOL,*/ " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1)*WELL_ZONE_LINK.RNI_WGHOR,MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1)) END) AS MTD_GAS_RG_VOL, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1)*WELL_ZONE_LINK.RNI_WGHOR,MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1)) END) AS YTD_GAS_RG_VOL, " & Chr(10) & _
"           SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1)*WELL_ZONE_LINK.RNI_WGHOR,MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1)) END) AS LTD_GAS_RG_VOL " & Chr(10) & _
"          ,/*MAX(NVL((SELECT LEAST(LAST_DAY(MH.START_DATETIME)+1,p.end_DATETIME) FROM ITEM_PROPERTY p WHERE p.PROPERTY_TYPE='LIFT_TYPE' AND p.ITEM_ID=WELL2UNITS.ITEM_ID and PROPERTY_STRING=MH.LIFT_TYPE AND (trunc(p.START_DATETIME,'mm')=trunc(MH.START_DATETIME,'mm') or trunc(p.end_DATETIME,'mm')=trunc(MH.START_DATETIME,'mm')) " & Chr(10) & _
"                      AND MH.START_DATETIME=to_date(':DATAOTCHETA','dd.mm.yyyy')),DECODE(MH.START_DATETIME,to_date(':DATAOTCHETA','dd.mm.yyyy'),LAST_DAY(MH.START_DATETIME)+1,NULL)))*/ LAST_DAY(to_date(':DATAOTCHETA','dd.mm.yyyy'))+1 LAST_DAY_PROD " & Chr(10)
SQL_MHP = SQL_MHP & _
"         FROM VCUSTOM_WELL2UNITS WELL2UNITS " & Chr(10) & _
"           LEFT OUTER JOIN vt_CUM_VOL_DET_ru_ru MH ON MH.ITEM_ID=WELL2UNITS.ITEM_ID " & Chr(10) & _
"           LEFT OUTER JOIN ( SELECT WZ.LINK_ID, WZ.TO_ITEM_ID, WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p0.PROPERTY_VALUE/100 AS REL_GAS_VOL, p1.PROPERTY_VALUE/100 AS REL_OIL_VOL, p2.PROPERTY_VALUE/100 AS REL_WATER_VOL, " & Chr(10) & _
"                               p9.PROPERTY_STRING AS RNI_ZONE_AGENT, DECODE(p9.PROPERTY_STRING,'OIL+WATER+GAS',p10.PROPERTY_VALUE,NULL)/1000 AS RNI_WGHOR FROM ITEM_LINK  WZ " & Chr(10) & _
"                               LEFT OUTER JOIN ITEM_LINK_PROPERTY p0 ON p0.LINK_TYPE='WELL_ZONE' AND p0.LINK_ID=WZ.LINK_ID AND p0.START_DATETIME<=WZ.START_DATETIME AND p0.END_DATETIME>WZ.START_DATETIME AND p0.PROPERTY_TYPE='REL_GAS_VOL' " & Chr(10) & _
"                               LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
"                               LEFT OUTER JOIN ITEM_LINK_PROPERTY p2 ON p2.LINK_TYPE='WELL_ZONE' AND p2.LINK_ID=WZ.LINK_ID AND p2.START_DATETIME<=WZ.START_DATETIME AND p2.END_DATETIME>WZ.START_DATETIME AND p2.PROPERTY_TYPE='REL_WATER_VOL' " & Chr(10) & _
"                               LEFT OUTER JOIN ITEM_LINK_PROPERTY p9 ON p9.LINK_TYPE='WELL_ZONE' AND p9.LINK_ID=WZ.LINK_ID AND p9.START_DATETIME<=WZ.START_DATETIME AND p9.END_DATETIME>WZ.START_DATETIME AND p9.PROPERTY_TYPE='RNI_ZONE_AGENT' " & Chr(10) & _
"                               LEFT OUTER JOIN ITEM_LINK_PROPERTY p10 ON p10.LINK_TYPE='WELL_ZONE' AND p10.LINK_ID=WZ.LINK_ID AND p10.START_DATETIME<=WZ.START_DATETIME AND p10.END_DATETIME>WZ.START_DATETIME AND p10.PROPERTY_TYPE='RNI_WGHOR' " & Chr(10) & _
"                             WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON WELL2UNITS.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND MH.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND MH.START_DATETIME<WELL_ZONE_LINK.END_DATETIME  " & Chr(10) & _
"           LEFT OUTER JOIN VL_FORMATION_ZONE_RU_RU FORMATION_ZONE_LINK ON WELL_ZONE_LINK.TO_ITEM_ID=FORMATION_ZONE_LINK.TO_ITEM_ID  " & Chr(10) & _
"           LEFT OUTER JOIN VI_FORMATION_ALL_RU_RU FORMATION ON FORMATION.ITEM_ID=FORMATION_ZONE_LINK.FROM_ITEM_ID AND MH.START_DATETIME>=FORMATION.START_DATETIME AND MH.START_DATETIME<FORMATION.END_DATETIME " & Chr(10) & _
"         WHERE MH.ITEM_ID = WELL2UNITS.ITEM_ID AND MH.LTD_OIL_MASS >0 AND MH.LTD_PROD_HOURS >0 " & Chr(10) & _
"           AND WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & CurrFIELD_name & "' " & Chr(10) & _
"           AND FORMATION.ITEM_ID = '" & CurrLAYER_name & "' " & Chr(10) & _
"           /*AND NOT EXISTS(SELECT RNI_ZONE_AGENT FROM VL_WELL_ZONE_RU_RU WHERE WELL2UNITS.ITEM_ID=FROM_ITEM_ID AND FORMATION_ZONE_LINK.TO_ITEM_ID=TO_ITEM_ID AND RNI_ZONE_AGENT='OIL+WATER+GAS') *//*FOR VNZ*/ " & Chr(10) & _
"         GROUP BY WELL2UNITS.ITEM_ID, WELL2UNITS.COMPLETION_NAME, WELL2UNITS.ORG_UNIT4_LEGACY_ID, WELL2UNITS.ORG_UNIT4_NAME, FORMATION_ZONE_LINK.TO_ITEM_ID, FORMATION.ITEM_ID, FORMATION.ITEM_NAME, MH.LIFT_TYPE, MH.LIFT_TYPE_TEXT ) "
Sql = _
"  SELECT FIELD_NAME, WELL_NAME, WELLPAD, LIFT_TYPE, LAYER_NAME, INITIAL_TYPE, WELL_STATUS, PROST_TEXT, WATER_DENSITY, CHOKE_DIAM, DATE_ON_PROD, PRESS_TUB, PRESS_BUF, " & Chr(10) & _
"       (SELECT TYPE_TEXT FROM vt_rsp_well_pump_ru_ru WHERE item_id=Q.ITEM_ID AND START_DATETIME=Q.PUMP_LST_DAY) PUMP_TYPE, " & Chr(10) & _
"       (SELECT ESP_MFR_TEXT FROM vt_rsp_well_pump_ru_ru WHERE item_id=Q.ITEM_ID AND START_DATETIME=Q.PUMP_LST_DAY) PUMP_MFR, " & Chr(10) & _
"       OIL_MASS, OIL_MASS_CUM_YEAR, OIL_MASS_CUM_FULL, " & Chr(10) & _
"       WATER_MASS, WATER_MASS_CUM_YEAR, WATER_MASS_CUM_FULL, " & Chr(10) & _
"       WATER_VOL, WATER_VOL_CUM_YEAR, WATER_VOL_CUM_FULL, " & Chr(10) & _
"       NVL((NVL(OIL_MASS,0)+NVL(WATER_MASS,0)),0)                    LIQ_MASS, " & Chr(10) & _
"       NVL((NVL(OIL_MASS_CUM_YEAR,0)+NVL(WATER_MASS_CUM_YEAR,0)),0)  LIQ_MASS_CUM_YEAR, " & Chr(10) & _
"       NVL((NVL(OIL_MASS_CUM_FULL,0)+NVL(WATER_MASS_CUM_FULL,0)),0)  LIQ_MASS_CUM_FULL, " & Chr(10) & _
"       NVL((NVL(OIL_VOL,0)+NVL(WATER_VOL,0)),0)                      LIQ_VOL, " & Chr(10) & _
"       NVL((NVL(OIL_VOL_CUM_YEAR,0)+NVL(WATER_VOL_CUM_YEAR,0)),0)    LIQ_VOL_CUM_YEAR, " & Chr(10) & _
"       NVL((NVL(OIL_VOL_CUM_FULL,0)+NVL(WATER_VOL_CUM_FULL,0)),0)    LIQ_VOL_CUM_FULL, " & Chr(10) & _
"       NVL(( NVL(OIL_VOL_PL,0)           +NVL(WATER_VOL,0)),0)         LIQ_VOL_PL, " & Chr(10) & _
"       NVL(( NVL(OIL_VOL_CUM_YEAR_PL,0)  +NVL(WATER_VOL_CUM_YEAR,0)),0) LIQ_VOL_CUM_YEAR_PL, " & Chr(10) & _
"       NVL(( NVL(OIL_VOL_CUM_FULL_PL,0)  +NVL(WATER_VOL_CUM_FULL,0)),0) LIQ_VOL_CUM_FULL_PL, " & Chr(10) & _
"       GAS_VOL, GAS_VOL_CUM_YEAR, GAS_VOL_CUM_FULL, GAS_RG_VOL, GAS_RG_VOL_CUM_YEAR, GAS_RG_VOL_CUM_FULL, " & Chr(10) & _
"       (GAS_VOL-GAS_RG_VOL) AS GAS_GH_VOL, (GAS_VOL_CUM_YEAR-GAS_RG_VOL_CUM_YEAR) AS GAS_GH_VOL_CUM_YEAR, (GAS_VOL_CUM_FULL-GAS_RG_VOL_CUM_FULL) AS GAS_GH_VOL_CUM_FULL, " & Chr(10)
Sql = Sql & _
"       NVL((GAS_VOL*1000/zvl(OIL_MASS)),0)                               GAS_FACTOR, " & Chr(10) & _
"       NVL((WATER_MASS*100/zvl(NVL(OIL_MASS,0)+NVL(WATER_MASS,0))),0)    WC_MASS, " & Chr(10) & _
"       NVL((WATER_VOL*100/zvl(NVL(OIL_VOL,0)+NVL(WATER_VOL,0))),0)       WC_VOL, " & Chr(10) & _
"       PROD_HOURS, PROD_DAYS, PROST_HOURS, NVL(((PROST_HOURS)/24),0) PROST_DAYS, PROST_DAYS_CUM_YEAR, PROD_DAYS_CUM_YEAR, PROD_DAYS_CUM_FULL, " & Chr(10) & _
"       NVL((OIL_MASS/zvl((PROD_HOURS+ACC_HOURS_MTH)/24)),0) OIL_RATE_MASS, " & Chr(10) & _
"       NVL((OIL_VOL/zvl((PROD_HOURS+ACC_HOURS_MTH)/24)),0) OIL_RATE_VOL, " & Chr(10) & _
"       NVL(((NVL(OIL_MASS,0)+NVL(WATER_MASS,0))/zvl((PROD_HOURS+ACC_HOURS_MTH)/24)),0) LIQ_RATE_MASS, " & Chr(10) & _
"       NVL(((NVL(OIL_VOL,0)+NVL(WATER_VOL,0))/zvl((PROD_HOURS+ACC_HOURS_MTH)/24)),0) LIQ_RATE_VOL, " & Chr(10) & _
"       NVL(( OIL_MASS/zvl((NVL(PROD_HOURS,0)+NVL(PROST_NAKOP,0))/24)),0) OIL_RATE_MASS_AVERAGE, " & Chr(10) & _
"       NVL(( (NVL(OIL_MASS,0)+NVL(WATER_MASS,0))/zvl((NVL(PROD_HOURS,0)+NVL(PROST_NAKOP,0))/24)),0) LIQ_RATE_MASS_AVERAGE, " & Chr(10) & _
"       NVL(( (NVL(OIL_VOL,0)+NVL(WATER_VOL,0))/zvl((NVL(PROD_HOURS,0)+NVL(PROST_NAKOP,0))/24)),0) LIQ_RATE_VOL_AVERAGE " & Chr(10) & _
"       ,ACC_HOURS_MTH, ACC_HOURS_MTH/24 ACC_DAYS_MTH, ACC_HOURS_YEAR, (ACC_HOURS_YEAR/24) ACC_DAYS_YEAR, MONTH_HOURS " & Chr(10) & _
"       ,PROD_DAYS_CUM_FULL+NVL(ACC_HOURS_FULL/24,0) AS EKSPL_DAIS_CUM_FULL " & Chr(10)
Sql = Sql & _
"  FROM ( " & Chr(10) & _
"    SELECT distinct t.ITEM_ID, t.FIELD_NAME, t.WELL_NAME, t.WELLPAD,  t.LIFT_TYPE, t.LAYER_NAME, t.INITIAL_TYPE, t.WELL_STATUS, t.PROST_TEXT, " & Chr(10) & _
"      NVL(t.CHOKE_DIAM,0) CHOKE_DIAM, t.DATE_ON_PROD, NVL(t.PRESS_TUB,0) PRESS_TUB, NVL(t.PRESS_BUF,0) PRESS_BUF, NVL(t.WATER_DENSITY,0) WATER_DENSITY " & Chr(10) & _
"      ,(SELECT MAX(START_DATETIME) FROM VT_RSP_WELL_PUMP_RU_RU WHERE ITEM_ID=T.ITEM_ID AND INSTR(UPPER(TYPE_TEXT),SUBSTR(T.LIFT_TYPE,2,2))>0 AND START_DATETIME<t.PUMP_LST_DAY) AS PUMP_LST_DAY " & Chr(10) & _
"      ,NVL(SUM(t.MTD_OIL_MASS   ),0) OIL_MASS " & Chr(10) & _
"      ,NVL(SUM(t.YTD_OIL_MASS   ),0) OIL_MASS_CUM_YEAR " & Chr(10) & _
"      ,NVL(SUM(t.LTD_OIL_MASS   ),0) OIL_MASS_CUM_FULL " & Chr(10) & _
"      ,NVL(SUM(t.MTD_OIL_VOL    ),0) OIL_VOL " & Chr(10) & _
"      ,NVL(SUM(t.YTD_OIL_VOL    ),0) OIL_VOL_CUM_YEAR " & Chr(10) & _
"      ,NVL(SUM(t.LTD_OIL_VOL    ),0) OIL_VOL_CUM_FULL " & Chr(10) & _
"      ,NVL(SUM(t.MTD_OIL_VOL_PL ),0) OIL_VOL_PL " & Chr(10) & _
"      ,NVL(SUM(t.YTD_OIL_VOL_PL ),0) OIL_VOL_CUM_YEAR_PL " & Chr(10) & _
"      ,NVL(SUM(t.LTD_OIL_VOL_PL ),0) OIL_VOL_CUM_FULL_PL " & Chr(10) & _
"      ,NVL(SUM(t.MTD_WATER_MASS ),0) WATER_MASS " & Chr(10) & _
"      ,NVL(SUM(t.YTD_WATER_MASS ),0) WATER_MASS_CUM_YEAR " & Chr(10) & _
"      ,NVL(SUM(t.LTD_WATER_MASS ),0) WATER_MASS_CUM_FULL " & Chr(10) & _
"      ,NVL(SUM(t.MTD_WATER_VOL  ),0) WATER_VOL " & Chr(10) & _
"      ,NVL(SUM(t.YTD_WATER_VOL  ),0) WATER_VOL_CUM_YEAR " & Chr(10) & _
"      ,NVL(SUM(t.LTD_WATER_VOL  ),0) WATER_VOL_CUM_FULL " & Chr(10)
Sql = Sql & _
"      ,NVL(SUM(t.MTD_GAS_VOL    ),0) GAS_VOL " & Chr(10) & _
"      ,NVL(SUM(t.YTD_GAS_VOL    ),0) GAS_VOL_CUM_YEAR " & Chr(10) & _
"      ,NVL(SUM(t.LTD_GAS_VOL    ),0) GAS_VOL_CUM_FULL " & Chr(10) & _
"      ,NVL(SUM(t.MTD_GAS_RG_VOL    ),0) GAS_RG_VOL " & Chr(10) & _
"      ,NVL(SUM(t.YTD_GAS_RG_VOL    ),0) GAS_RG_VOL_CUM_YEAR " & Chr(10) & _
"      ,NVL(SUM(t.LTD_GAS_RG_VOL    ),0) GAS_RG_VOL_CUM_FULL " & Chr(10) & _
"      ,NVL(SUM(t.MTD_PROD_HOURS),0) PROD_HOURS " & Chr(10) & _
"      ,NVL((SUM(t.MTD_PROD_HOURS)/24),0) PROD_DAYS " & Chr(10) & _
"      ,DECODE(SIGN(NVL(SUM(t.PROST_HOURS),0)         - NVL(ACC_HOURS_MTH,0))    ,-1,0,(NVL(SUM(t.PROST_HOURS),0)         - NVL(ACC_HOURS_MTH,0)))     PROST_HOURS " & Chr(10) & _
"      ,DECODE(SIGN(NVL(SUM(t.PROST_DAYS_CUM_YEAR),0) - NVL(ACC_HOURS_YEAR/24,0)),-1,0,(NVL(SUM(t.PROST_DAYS_CUM_YEAR),0) - NVL(ACC_HOURS_YEAR/24,0))) PROST_DAYS_CUM_YEAR " & Chr(10) & _
"      ,NVL((SUM(t.YTD_PROD_HOURS)/24),0) PROD_DAYS_CUM_YEAR " & Chr(10) & _
"      ,NVL((SUM(t.LTD_PROD_HOURS)/24),0) PROD_DAYS_CUM_FULL " & Chr(10) & _
"      ,ACC_HOURS_MTH, ACC_HOURS_YEAR, ACC_HOURS_FULL, MONTH_HOURS, NVL(SUM(t.PROST_HOURS),0) PROST_NAKOP " & Chr(10)
Sql = Sql & _
"    FROM " & Chr(10) & _
"    (SELECT MONTH_HIST_PL.ITEM_ID, MONTH_HIST_PL.FIELD_NAME, MONTH_HIST_PL.FIELD_ID, MONTH_HIST_PL.WELL_NAME, " & Chr(10) & _
"       (SELECT min(PROPERTY_STRING) FROM ITEM_PROPERTY WHERE ITEM_ID = MONTH_HIST_PL.ITEM_ID AND PROPERTY_TYPE='WELLPAD' AND last_day(MONTH_HIST.START_DATETIME) between START_DATETIME AND END_DATETIME) WELLPAD, " & Chr(10) & _
"       MONTH_HIST_PL.LAYER_NAME, MONTH_HIST_PL.LAYER_ID, MONTH_HIST_PL.OIL_FVF, MONTH_HIST_PL.WATER_DENSITY, " & Chr(10) & _
"       (SELECT min(p48.CODE_TEXT) AS INITIAL_TYPE_TEXT FROM ITEM i INNER JOIN ITEM_VERSION v ON v.ITEM_ID=i.ITEM_ID " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY p47 ON p47.ITEM_ID=i.ITEM_ID AND p47.START_DATETIME<=v.START_DATETIME AND p47.END_DATETIME>v.START_DATETIME AND p47.PROPERTY_TYPE='INITIAL_TYPE' " & Chr(10) & _
"          LEFT OUTER JOIN (SELECT CL48.CODE_TEXT, CL48.CODE FROM CODE_LIST CL48 WHERE CL48.LIST_TYPE='RSP_PROJECT_WELL_TYPE' AND CL48.CULTURE='ru-RU') p48 ON p47.PROPERTY_STRING=p48.CODE " & Chr(10) & _
"        WHERE i.ITEM_TYPE='COMPLETION' AND i.ITEM_ID=MONTH_HIST_PL.ITEM_ID AND last_day(MONTH_HIST.START_DATETIME) between v.START_DATETIME AND v.END_DATETIME) INITIAL_TYPE, " & Chr(10) & _
"       (SELECT to_char(IP_DATE_ON_PROD.PROPERTY_DATE,'DD/MM/YYYY') AS DATE_ON_PROD FROM ITEM I " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY IP_DATE_ON_PROD ON IP_DATE_ON_PROD.ITEM_ID=I.ITEM_ID AND IP_DATE_ON_PROD.START_DATETIME<=TRUNC(SYSDATE) " & Chr(10) & _
"            AND IP_DATE_ON_PROD.END_DATETIME>TRUNC(SYSDATE) AND IP_DATE_ON_PROD.PROPERTY_TYPE='DATE_ON_PROD' " & Chr(10) & _
"        WHERE I.ITEM_TYPE='COMPLETION' AND I.START_DATETIME<=TRUNC(SYSDATE) AND I.END_DATETIME>TRUNC(SYSDATE) AND I.ITEM_ID=MONTH_HIST_PL.ITEM_ID) DATE_ON_PROD, " & Chr(10) & _
"       MONTH_HIST.LIFT_TYPE_TEXT LIFT_TYPE, " & Chr(10) & _
"       DECODE(MONTH_HIST.LIFT_TYPE,STOCK_HIST.LIFT_TYPE,DECODE(STOCK_HIST.TYPE,'PRODUCTION',STOCK_HIST.STATUS_TEXT,''),'') WELL_STATUS, " & Chr(10) & _
"       NVL(( SELECT START_DATETIME  FROM VT_DOWNTIME_RU_RU WHERE item_id=MONTH_HIST_PL.ITEM_ID " & Chr(10) & _
"              AND ( SELECT TO_DATE(TO_CHAR(LAST_DAY(MAX(M.START_DATETIME)),'DD.MM.YYYY')||' 23:59:59','DD.MM.YYYY HH24:MI:SS') FROM VT_TOT_DET_MTH_RU_RU M " & Chr(10) & _
"                     WHERE M.ITEM_ID=MONTH_HIST.ITEM_ID AND M.LIFT_TYPE=MONTH_HIST.LIFT_TYPE AND M.START_DATETIME<=MONTH_HIST.START_DATETIME) BETWEEN START_DATETIME AND END_DATETIME), " & Chr(10) & _
"           ( SELECT TO_DATE(TO_CHAR(LAST_DAY(MAX(M.START_DATETIME)),'DD.MM.YYYY')||' 23:59:59','DD.MM.YYYY HH24:MI:SS') FROM VT_TOT_DET_MTH_RU_RU M " & Chr(10) & _
"             WHERE M.ITEM_ID = MONTH_HIST.ITEM_ID AND M.LIFT_TYPE= MONTH_HIST.LIFT_TYPE AND M.START_DATETIME<=MONTH_HIST.START_DATETIME)) PUMP_LST_DAY, " & Chr(10)
Sql = Sql & _
"       (SELECT max(choke_1) FROM vt_well_read_ru_ru WHERE item_id=MONTH_HIST_PL.ITEM_ID " & Chr(10) & _
"         AND last_day(MONTH_HIST.START_DATETIME) between START_DATETIME AND END_DATETIME) CHOKE_DIAM, " & Chr(10) & _
"       (SELECT max(CASING_PRESS4) FROM vt_well_read_ru_ru WHERE item_id=MONTH_HIST_PL.ITEM_ID " & Chr(10) & _
"         AND last_day(MONTH_HIST.START_DATETIME) between START_DATETIME AND END_DATETIME) PRESS_TUB, " & Chr(10) & _
"       (SELECT max(CASING_PRESS2) FROM vt_well_read_ru_ru WHERE item_id=MONTH_HIST_PL.ITEM_ID " & Chr(10) & _
"         AND last_day(MONTH_HIST.START_DATETIME) between START_DATETIME AND END_DATETIME) PRESS_BUF, " & Chr(10) & _
"       (SELECT min(downtime_type_text) FROM vt_downtime_ru_ru WHERE item_id=MONTH_HIST_PL.ITEM_ID AND MONTH_HIST.MTD_PROD_HOURS>0 " & Chr(10) & _
"         AND LAST_DAY_PROD between START_DATETIME AND END_DATETIME) PROST_TEXT, " & Chr(10) & _
"       FN_MTH_INACT_HOURS_OIL_RNI(MONTH_HIST.START_DATETIME,MONTH_HIST.ITEM_ID,MONTH_HIST.LIFT_TYPE) PROST_HOURS, " & Chr(10) & _
"       (SELECT SUM(FN_MTH_INACT_HOURS_OIL_RNI(MTH.START_DATETIME,MTH.ITEM_ID,MTH.LIFT_TYPE)) FROM vt_CUM_VOL_DET_ru_ru MTH WHERE MTH.ITEM_ID=MONTH_HIST_PL.ITEM_ID " & Chr(10) & _
"          AND TO_CHAR(MTH.START_DATETIME,'YYYY')=TO_CHAR(MONTH_HIST.START_DATETIME,'YYYY') AND TO_CHAR(MTH.START_DATETIME,'MM')<=TO_CHAR(MONTH_HIST.START_DATETIME,'MM') " & Chr(10) & _
"          AND MTH.LIFT_TYPE=MONTH_HIST.LIFT_TYPE)/24 PROST_DAYS_CUM_YEAR, " & Chr(10) & _
"       MONTH_HIST_PL.MTD_OIL_MASS        MTD_OIL_MASS, " & Chr(10) & _
"       MONTH_HIST_PL.YTD_OIL_MASS        YTD_OIL_MASS, " & Chr(10) & _
"       MONTH_HIST_PL.LTD_OIL_MASS        LTD_OIL_MASS, " & Chr(10) & _
"       MONTH_HIST_PL.MTD_OIL_VOL         MTD_OIL_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.YTD_OIL_VOL         YTD_OIL_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.LTD_OIL_VOL         LTD_OIL_VOL, " & Chr(10)
Sql = Sql & _
"       NVL((SELECT SUM(ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100) " & Chr(10) & _
"            FROM VT_TOT_DET_MTH_RU_RU t INNER JOIN VI_ZONE_ALL_RU_RU z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
"              ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK  WZ " & Chr(10) & _
"                  LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
"                WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
"            WHERE MONTH_HIST_PL.ZONE_ID=z.ITEM_ID AND t.ITEM_ID=MONTH_HIST_PL.ITEM_ID AND t.lift_type = MONTH_HIST.lift_type " & Chr(10) & _
"              AND t.START_DATETIME=MONTH_HIST.START_DATETIME ),0) MTD_OIL_VOL_PL, " & Chr(10) & _
"       NVL((SELECT SUM(ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100) " & Chr(10) & _
"            FROM VT_TOT_DET_MTH_RU_RU t INNER JOIN VI_ZONE_ALL_RU_RU z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
"              ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK  WZ " & Chr(10) & _
"                  LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
"                WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
"            WHERE MONTH_HIST_PL.ZONE_ID=z.ITEM_ID AND t.ITEM_ID=MONTH_HIST_PL.ITEM_ID AND t.lift_type = MONTH_HIST.lift_type " & Chr(10) & _
"              AND t.START_DATETIME between TRUNC(MONTH_HIST.START_DATETIME, 'YYYY') AND MONTH_HIST.START_DATETIME ),0) YTD_OIL_VOL_PL, " & Chr(10) & _
"       NVL((SELECT SUM(ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100) " & Chr(10) & _
"            FROM VT_TOT_DET_MTH_RU_RU t INNER JOIN VI_ZONE_ALL_RU_RU z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
"              ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK  WZ " & Chr(10) & _
"                  LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
"                WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
"            WHERE MONTH_HIST_PL.ZONE_ID=z.ITEM_ID AND t.ITEM_ID=MONTH_HIST_PL.ITEM_ID AND t.lift_type = MONTH_HIST.lift_type " & Chr(10) & _
"              AND t.START_DATETIME<=MONTH_HIST.START_DATETIME ),0) LTD_OIL_VOL_PL, " & Chr(10)
Sql = Sql & _
"       MONTH_HIST_PL.MTD_WATER_MASS      MTD_WATER_MASS, " & Chr(10) & _
"       MONTH_HIST_PL.YTD_WATER_MASS      YTD_WATER_MASS, " & Chr(10) & _
"       MONTH_HIST_PL.LTD_WATER_MASS      LTD_WATER_MASS, " & Chr(10) & _
"       MONTH_HIST_PL.MTD_WATER_VOL       MTD_WATER_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.YTD_WATER_VOL       YTD_WATER_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.LTD_WATER_VOL       LTD_WATER_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.MTD_GAS_VOL         MTD_GAS_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.YTD_GAS_VOL         YTD_GAS_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.LTD_GAS_VOL         LTD_GAS_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.MTD_GAS_RG_VOL         MTD_GAS_RG_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.YTD_GAS_RG_VOL         YTD_GAS_RG_VOL, " & Chr(10) & _
"       MONTH_HIST_PL.LTD_GAS_RG_VOL         LTD_GAS_RG_VOL, " & Chr(10) & _
"       MONTH_HIST.MTD_PROD_HOURS      MTD_PROD_HOURS, " & Chr(10) & _
"       MONTH_HIST.YTD_PROD_HOURS      YTD_PROD_HOURS, " & Chr(10) & _
"       MONTH_HIST.LTD_PROD_HOURS      LTD_PROD_HOURS " & Chr(10)
Sql = Sql & _
"       ,ROUND(NVL((select sum(tot.acc_time/60/60) from vt_tot_det_day_ru_ru totd inner join vt_totals_day_ru_ru tot on totd.item_id=tot.item_id and totd.start_datetime=tot.start_datetime " & Chr(10) & _
"                   where totd.item_id=MONTH_HIST_PL.ITEM_ID and totd.lift_type=MONTH_HIST.lift_type and trunc(totd.start_datetime,'mm')=MONTH_HIST.START_DATETIME " & Chr(10) & _
"        ),0),5) ACC_HOURS_MTH " & Chr(10) & _
"       ,ROUND(NVL((select sum(tot.acc_time/60/60) from vt_tot_det_day_ru_ru totd inner join vt_totals_day_ru_ru tot on totd.item_id=tot.item_id and totd.start_datetime=tot.start_datetime " & Chr(10) & _
"                   where totd.item_id=MONTH_HIST_PL.ITEM_ID and totd.lift_type=MONTH_HIST.lift_type and trunc(totd.start_datetime,'mm') between trunc(MONTH_HIST.START_DATETIME,'yyyy') and MONTH_HIST.START_DATETIME " & Chr(10) & _
"        ),0),5) ACC_HOURS_YEAR " & Chr(10) & _
"       ,ROUND(NVL((select sum(tot.acc_time/60/60) from vt_tot_det_day_ru_ru totd inner join vt_totals_day_ru_ru tot on totd.item_id=tot.item_id and totd.start_datetime=tot.start_datetime " & Chr(10) & _
"                   where totd.item_id=MONTH_HIST_PL.ITEM_ID and totd.lift_type=MONTH_HIST.lift_type and trunc(totd.start_datetime,'mm')<=MONTH_HIST.START_DATETIME " & Chr(10) & _
"        ),0),5) ACC_HOURS_FULL " & Chr(10) & _
"      ,NVL((SELECT (LEAST(LAST_DAY(to_date(':DATAOTCHETA','dd.mm.yyyy'))+1,p.end_DATETIME)-GREATEST(trunc(to_date(':DATAOTCHETA','dd.mm.yyyy'),'mm'), p.START_DATETIME))*24 " & Chr(10) & _
"             FROM ITEM_PROPERTY p  WHERE p.PROPERTY_TYPE='LIFT_TYPE' AND p.ITEM_ID=MONTH_HIST_PL.ITEM_ID and PROPERTY_STRING=MONTH_HIST.LIFT_TYPE_TEXT " & Chr(10) & _
"               and (trunc(p.START_DATETIME,'mm')=trunc(to_date(':DATAOTCHETA','dd.mm.yyyy'),'mm') or trunc(p.end_DATETIME,'mm')=trunc(to_date(':DATAOTCHETA','dd.mm.yyyy'),'mm'))), " & Chr(10) & _
"           TO_NUMBER(TO_CHAR(LAST_DAY(to_date(':DATAOTCHETA','dd.mm.yyyy')),'DD'))*24 ) MONTH_HOURS " & Chr(10)
Sql = Sql & _
"     FROM VT_RSP_STOCK_HIST_ru_RU STOCK_HIST, " & Chr(10) & SQL_MHP & " MONTH_HIST_PL " & Chr(10) & _
"       LEFT OUTER JOIN vt_CUM_VOL_DET_ru_ru MONTH_HIST ON MONTH_HIST.ITEM_ID=MONTH_HIST_PL.ITEM_ID AND MONTH_HIST.LIFT_TYPE=MONTH_HIST_PL.LIFT_TYPE AND MONTH_HIST.START_DATETIME=to_date(':DATAOTCHETA','DD.MM.YYYY') " & Chr(10) & _
"     WHERE MONTH_HIST.ITEM_ID=MONTH_HIST_PL.ITEM_ID " & Chr(10) & _
"       AND STOCK_HIST.START_DATETIME = last_day(to_date(':DATAOTCHETA','DD.MM.YYYY')) " & Chr(10) & _
"       AND MONTH_HIST_PL.ITEM_ID = STOCK_HIST.ITEM_ID " & Chr(10) & _
"       AND MONTH_HIST.LTD_OIL_MASS >0 AND MONTH_HIST.LTD_PROD_HOURS >0 " & Chr(10) & _
"    ) t " & Chr(10) & _
"    group by t.ITEM_ID, t.FIELD_NAME, t.WELL_NAME, t.WELLPAD, t.LIFT_TYPE, t.LAYER_NAME, t.INITIAL_TYPE, t.WELL_STATUS, t.PROST_TEXT, " & Chr(10) & _
"          t.CHOKE_DIAM , t.PUMP_LST_DAY, t.DATE_ON_PROD, t.PRESS_TUB, t.PRESS_BUF, t.WATER_DENSITY " & Chr(10) & _
"          ,ACC_HOURS_MTH, ACC_HOURS_YEAR, ACC_HOURS_FULL, MONTH_HOURS " & Chr(10) & _
"   ) Q order by str2num(WELL_NAME),LIFT_TYPE desc "
Sql = Replace(Sql, ":DATAOTCHETA", "01." & Month(StartDay) & "." & Year(StartDay))

'.Cells(1, 9) = = Sql

    Set RS_Layer = New ADODB.Recordset
    RS_Layer.ActiveConnection = cn ' Assign the Connection object.
    RS_Layer.CursorType = adOpenStatic
    RS_Layer.Open Sql ' Extract the required records.
     
CurrWELL_name = ""
well_num = 0
     
        ' Цикл перебора записей со скважинами по текущему набору месторождений-пластов
        If Not (RS_Layer.BOF And RS_Layer.EOF And RS_Layer.RecordCount = 0) Then
            For j = 0 To RS_Layer.RecordCount - 1 ' бежим по рекордсету
            Rows(CStr(RowCounter) & ":" & CStr(RowCounter)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            'Выставляем порядковые номера скважин только у первых записей
            If Not (RS_Layer.Fields("WELL_NAME") = CurrWELL_name) Then
            well_num = well_num + 1
            .Cells(RowCounter, 1) = well_num
            CurrWELL_name = RS_Layer.Fields("WELL_NAME")
            End If
        
            .Cells(RowCounter, 2) = RS_Layer.Fields("FIELD_NAME")     'Месторождение
            .Cells(RowCounter, 3) = RS_Layer.Fields("WELL_NAME")      'Номер скважины
            .Cells(RowCounter, 4) = RS_Layer.Fields("WELLPAD")        'Куст
            .Cells(RowCounter, 5) = RS_Layer.Fields("DATE_ON_PROD")   'Дата ввода в эксплуатацию
            .Cells(RowCounter, 6) = RS_Layer.Fields("LIFT_TYPE")      'Способ
            .Cells(RowCounter, 7) = RS_Layer.Fields("LAYER_NAME")     'Пласт
            .Cells(RowCounter, 8) = RS_Layer.Fields("INITIAL_TYPE")   'Категория проектная
            .Cells(RowCounter, 9) = RS_Layer.Fields("PUMP_TYPE")      'Насос
            .Cells(RowCounter, 10) = RS_Layer.Fields("PUMP_MFR")      'Завод изготовитель
            .Cells(RowCounter, 11) = RS_Layer.Fields("CHOKE_DIAM")    'Диаметр штуцера
            
            'Добыча нефти(в т.ч. конденсат), т
            .Cells(RowCounter, 12) = RS_Layer.Fields("OIL_MASS")            'За месяц
            .Cells(RowCounter, 13) = RS_Layer.Fields("OIL_MASS_CUM_YEAR")   'С начала года
            .Cells(RowCounter, 14) = RS_Layer.Fields("OIL_MASS_CUM_FULL")   'С начала разработки
            
            'Добыча воды, т
            .Cells(RowCounter, 15) = RS_Layer.Fields("WATER_MASS")          'За месяц
            .Cells(RowCounter, 16) = RS_Layer.Fields("WATER_MASS_CUM_YEAR") 'С начала года
            .Cells(RowCounter, 17) = RS_Layer.Fields("WATER_MASS_CUM_FULL") 'С начала разработки
            
            'Добыча воды, м3
            .Cells(RowCounter, 18) = RS_Layer.Fields("WATER_VOL")           'За месяц
            .Cells(RowCounter, 19) = RS_Layer.Fields("WATER_VOL_CUM_YEAR")  'С начала года
            .Cells(RowCounter, 20) = RS_Layer.Fields("WATER_VOL_CUM_FULL")  'С начала разработки
            
            'Добыча жидкости, т
            .Cells(RowCounter, 21) = RS_Layer.Fields("LIQ_MASS")            'За месяц
            .Cells(RowCounter, 22) = RS_Layer.Fields("LIQ_MASS_CUM_YEAR")   'С начала года
            .Cells(RowCounter, 23) = RS_Layer.Fields("LIQ_MASS_CUM_FULL")   'С начала разработки
            
            'Добыча жидкости в поверхн-ых условиях, м3
            .Cells(RowCounter, 24) = RS_Layer.Fields("LIQ_VOL")             'За месяц
            .Cells(RowCounter, 25) = RS_Layer.Fields("LIQ_VOL_CUM_YEAR")    'С начала года
            .Cells(RowCounter, 26) = RS_Layer.Fields("LIQ_VOL_CUM_FULL")    'С начала разработки
            
            'Добыча жидкости в пластовых условиях, м3
            .Cells(RowCounter, 27) = RS_Layer.Fields("LIQ_VOL_PL")          'За месяц
            .Cells(RowCounter, 28) = RS_Layer.Fields("LIQ_VOL_CUM_YEAR_PL") 'С начала года
            .Cells(RowCounter, 29) = RS_Layer.Fields("LIQ_VOL_CUM_FULL_PL") 'С начала разработки
            
            'Добыча растворенного газа, тыс.м3
            .Cells(RowCounter, 30) = RS_Layer.Fields("GAS_VOL")             'За месяц
            .Cells(RowCounter, 31) = RS_Layer.Fields("GAS_VOL_CUM_YEAR")    'С начала года
            .Cells(RowCounter, 32) = RS_Layer.Fields("GAS_VOL_CUM_FULL")    'С начала разработки
            
            .Cells(RowCounter, 33) = RS_Layer.Fields("GAS_FACTOR")           'Газовый фактор
            .Cells(RowCounter, 34) = RS_Layer.Fields("WC_MASS")              '% воды, весовой
            .Cells(RowCounter, 35) = RS_Layer.Fields("WC_VOL")               '% воды, объемный
            
            .Cells(RowCounter, 36) = RS_Layer.Fields("WATER_DENSITY")        'Удельный вес воды кг/м3
            .Cells(RowCounter, 37) = RS_Layer.Fields("PRESS_TUB")            'Давление затрубное, атм.
            .Cells(RowCounter, 38) = RS_Layer.Fields("PRESS_BUF")            'Давление буферное, атм.
           
            .Cells(RowCounter, 39) = RS_Layer.Fields("OIL_RATE_MASS")        'Уплотненный дебит нефти т/сут
            .Cells(RowCounter, 40) = RS_Layer.Fields("OIL_RATE_VOL")         'Уплотненный дебит нефти м3/сут
            .Cells(RowCounter, 41) = RS_Layer.Fields("LIQ_RATE_MASS")        'Уплотненный дебит жидк. т/сут
            .Cells(RowCounter, 42) = RS_Layer.Fields("LIQ_RATE_VOL")
            
            .Cells(RowCounter, 43) = RS_Layer.Fields("OIL_RATE_MASS_AVERAGE") 'Среднесуточный дебит нефти т/сут
            .Cells(RowCounter, 44) = RS_Layer.Fields("LIQ_RATE_MASS_AVERAGE") 'Среднесуточный дебит жидк. т/сут
            .Cells(RowCounter, 45) = RS_Layer.Fields("LIQ_RATE_VOL_AVERAGE")  'Среднесуточный дебит жидк. м3/сут
            
            .Cells(RowCounter, 46) = RS_Layer.Fields("PROD_DAYS_CUM_YEAR")  'Скв.сутки работы с начала года
            .Cells(RowCounter, 47) = RS_Layer.Fields("ACC_DAYS_YEAR")       'Скв.сутки накопления с начала года
            .Cells(RowCounter, 48) = RS_Layer.Fields("PROST_DAYS_CUM_YEAR") 'Скв. сутки простоя с начала года
            .Cells(RowCounter, 49) = RS_Layer.Fields("EKSPL_DAIS_CUM_FULL")  'Сутки экплуат. с начала разраб.

            
            'Уплотн.скв.часы за месяц
            '.Cells(RowCounter, 50) = RS_Layer.Fields("PROD_HOURS")    'Работы
            '.Cells(RowCounter, 51) = RS_Layer.Fields("ACC_HOURS_MTH") 'Накопления
            '.Cells(RowCounter, 52) = RS_Layer.Fields("PROST_HOURS")  'Простоя
            'HoursInMon = DateDiff("d", DateSerial(Year(StartDay), Month(StartDay), 1), DateAdd("m", 1, DateSerial(Year(StartDay), Month(StartDay), 1))) * 24 'Количество часов в отчетном месяце
            HoursInMon = RS_Layer.Fields("MONTH_HOURS") 'Количество часов в отчетном месяце
            HoursProd = RS_Layer.Fields("PROD_HOURS")  'Работы
            HoursNakop = RS_Layer.Fields("ACC_HOURS_MTH")        'Накопления
            HoursProst = RS_Layer.Fields("PROST_HOURS")          'Простоя
            .Cells(RowCounter, 51) = HoursNakop                  'Часы накопления
            If HoursProst <> 0 Then
              .Cells(RowCounter, 50) = HoursProd                 'Часы работы
              .Cells(RowCounter, 52) = IIf(HoursProd + HoursNakop + HoursProst > HoursInMon, _
                HoursInMon - HoursProd - HoursNakop, HoursProst) 'Часы простоя
            Else
              .Cells(RowCounter, 52) = HoursProst                'Часы простоя
              .Cells(RowCounter, 50) = IIf(HoursProd + HoursNakop + HoursProst > HoursInMon, _
                HoursInMon - HoursProst - HoursNakop, HoursProd) 'Часы работы
            End If
            
            'Уплотн.скв.сутки за месяц
            .Cells(RowCounter, 53) = RS_Layer.Fields("PROD_DAYS")     'Работы
            .Cells(RowCounter, 54) = RS_Layer.Fields("ACC_DAYS_MTH")  'Накопления
            .Cells(RowCounter, 55) = RS_Layer.Fields("PROST_DAYS")    'Простоя
            
            ' Если нет простоя, то и причину не выдаём
'            If (RS_Layer.Fields("PROST_HOURS") > 0) Then
            .Cells(RowCounter, 56) = RS_Layer.Fields("PROST_TEXT")    'Причина простоя
'            End If
            With Range("BD" & CStr(RowCounter) & ":BE" & CStr(RowCounter))
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlBottom
              .WrapText = False
              .Orientation = 0
              .AddIndent = False
              .IndentLevel = 0
              .ShrinkToFit = False
              .ReadingOrder = xlContext
              .MergeCells = False
              .Merge
            End With
            .Cells(RowCounter, 58) = RS_Layer.Fields("WELL_STATUS")   'Состояние на конец месяца
            
            Range("L" & CStr(RowCounter) & ":AF" & CStr(RowCounter)).NumberFormat = "0.000"
            Range("AG" & CStr(RowCounter) & ":BC" & CStr(RowCounter)).NumberFormat = "0.0"
        
            RowCounter = RowCounter + 1
        
            RS_Layer.MoveNext
            Next j
        End If
    
        RS_Layer.Close
        
        RowCounter = RowCounter + 1
    
    RS.MoveNext
    Next i
End If
RS.Close

'.Cells(30, 2).Font.Bold = 1
'.Cells(40, 8) = Sql

'Подпись
RowCounter = RowCounter + 2

Range(.Cells(RowCounter, 1), .Cells(RowCounter, 41)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge

.Cells(RowCounter, 1) = "Начальник отдела разработки      _________________  /                 /"
.Cells(RowCounter, 1).Font.Size = 14
'.Cells(RowCounter, 1).Font.FontStyle = "Bold"

cn.Close
Set RS = Nothing
Set RS_Layer = Nothing
Set cn = Nothing

'If Len(ParField_Item_Name) > 0 Then
'    OldFileName = Application.ThisWorkbook.Path & "\" & Application.ThisWorkbook.Name
'    Application.ThisWorkbook.SaveAs Application.ThisWorkbook.Path & _
'                          "\" & Choose(Month(StartDay), "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12") & "_" & ParField_Item_Name & "_" & Application.ThisWorkbook.Name
'    Kill OldFileName
'End If

Call MainExec2(rawParams)

Application.ThisWorkbook.Application.Visible = True
Application.ThisWorkbook.Application.ScreenUpdating = True
Application.Cursor = xlDefault

End With

Exit Sub

LabelErr:
    Application.Cursor = xlDefault
    Application.ThisWorkbook.Application.Visible = True
    Application.ThisWorkbook.Application.ScreenUpdating = True
    MsgBox "Ошибка выполнения отчета. Процедура MainExec. " + vbCrLf + vbCrLf + _
            Err.Description, vbCritical + vbOKOnly, "Ошибка: " + CStr(Err.Number)

End Sub




Sub MainExec2(rawParams() As Variant)
'On Error GoTo LabelErr
Application.Cursor = xlIBeam

Dim cn               As ADODB.Connection ' Connection object
Dim RS               As ADODB.Recordset  ' Recordset object
Dim RS_Layer         As ADODB.Recordset  ' Recordset object

Dim StartDay         As Date             ' Дата отчета
Dim RowCounter       As Integer          ' Счетчик для движения по строкам отчета
Dim RowCounter_Mest  As Integer          ' Счетчик для движения по строкам отчета для месторождения

Dim RowCounter_Spos       As Integer          ' Счетчик движения по строкам отчета для способа экспл.
Dim RowCounter_Mest_Spos  As Integer          ' Счетчик движения по строкам отчета для способа экспл. у месторождения

Dim RowCounter_temp  As Integer          ' Счетчик вспомогательный

Dim RowCounter_max   As Integer          ' Максимальный номер занимаемой строки

Dim i                As Integer          ' Счетчик для движения по строкам рекордсета
Dim j                As Integer          ' Счетчик для движения по строкам рекордсета
Dim well_num         As Integer          ' Счетчик порядкового номера скважины

Dim layer_count      As Integer          ' Число пластов на местрождении
Dim lift_type_count  As Integer          ' Число с/э на местрождении

Dim layer_count_prev     As Integer          ' Число пластов на местрождении (пред.значение)
Dim lift_type_count_prev As Integer          ' Число с/э на местрождении (пред.значение)
Dim is_new_mest          As Integer          ' Флаг появления записи с новым местрождением

Dim Sql_main         As String           ' Строка sql-запроса
Dim Sql              As String           ' Строка sql-запроса
Dim Sql_temp         As String           ' Строка sql-запроса

Dim CurrFIELD_name      As String           ' Текущее месторождение
Dim CurrGRP_name        As String
Dim CurrLAYER_name      As String           ' Текущий пласт
'Dim CurrWELL_name       As String           ' Текущая скважина
Dim CurrLIFT_TYPE_name  As String           ' Текущий способ эксплуатации

Dim CurrFIELD_code      As String           ' Текущий код месторождения
Dim CurrLAYER_code      As String           ' Текущий код пласта
Dim CurrLIFT_TYPE_code  As String           ' Текущий код способа эксплуатации
Dim ParField_Item_Id    As String        ' Id месторождения, выбранного пользователем
Dim ParField_Item_Name  As String        ' Name месторождения, выбранного пользователем
Dim OldFileName         As String
Dim TitleDate         As Date             ' Дата заголовка (+1 день)

Dim ThisMoment        As Date             ' Текущее время

'Application.ThisWorkbook.Application.Visible = False
'Application.ThisWorkbook.Application.ScreenUpdating = False

Dim params As New Collection
For i = 0 To UBound(rawParams) - 1
    params.Add rawParams(i + 1), rawParams(i)
    i = i + 1
Next i

Set cn = New ADODB.Connection
Dim strConn As String
strConn = "Provider=OraOLEDB.Oracle.1;" + params("CONNECTION")
cn.Open strConn

'Первая строчка после шапки
RowCounter = 37
RowCounter_Mest = 37
'17 - потому что к 14 добавляется 3 заголовочные строчки по месторождению

RowCounter_max = 37

StartDay = CDate(params("AS_OF_DATE"))
TitleDate = DateAdd("d", 1, StartDay)

If Len(params("ITEM")) > 0 Then
    'Id месторождения, выбранного пользователем
    ParField_Item_Id = Mid(params("ITEM"), InStr(params("ITEM"), ",") + 1, 32)
    'Name месторождения, выбранного пользователем
    ParField_Item_Name = Right(params("ITEM"), Len(params("ITEM")) - InStrRev(params("ITEM"), ","))
End If
With Sheets("Sheet1")

' Отобразим время начала формирования отчета
'ThisMoment = Now
'.Cells(1, 7) = ThisMoment

.Cells(2, 1) = "За :                   " & Choose(Month(StartDay), "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь") & " " & Year(StartDay) & " г."
.Cells(3, 1) = "Предприятие : " & GetUnitNameByFieldID(ParField_Item_Id)
.Cells(4, 1) = "Дата выдачи:  " & Day(Date) & " " & Choose(Month(Date), "Января", "Февраля", "Марта", "Апреля", "Мая", "Июня", "Июля", "Августа", "Сентября", "Октября", "Ноября", "Декабря") & " " & Year(Date) & " г."

' Зачистим хвосты, которые затем сами сгенерим
Range("A37:A84").ClearContents

Sql_main = _
"WITH D AS " & Chr(10) & _
"  ( select distinct WELL2UNITS.ORG_UNIT4_NAME FIELD_NAME, WELL2UNITS.ORG_UNIT4_LEGACY_ID FIELD_ID, " & Chr(10) & _
"      FORMATION.ITEM_NAME LAYER_NAME, FORMATION.ITEM_ID LAYER_ID, NVL(MONTH_HIST.lift_type,C.LIFT_TYPE) lift_type " & Chr(10) & _
"    FROM VCUSTOM_WELL2UNITS WELL2UNITS INNER JOIN " & Chr(10) & _
"      ( SELECT i.ITEM_ID, p16.PROPERTY_STRING LIFT_TYPE FROM ITEM i INNER JOIN ITEM_VERSION v ON v.ITEM_ID=i.ITEM_ID " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY p2 ON p2.PROPERTY_TYPE='TYPE' AND p2.ITEM_ID=i.ITEM_ID AND p2.START_DATETIME<=v.START_DATETIME AND p2.END_DATETIME>v.START_DATETIME " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY p6 ON p6.PROPERTY_TYPE='PRODUCT' AND p6.ITEM_ID=i.ITEM_ID AND p6.START_DATETIME<=v.START_DATETIME AND p6.END_DATETIME>v.START_DATETIME " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY p8 ON p8.PROPERTY_TYPE='STATUS' AND p8.ITEM_ID=i.ITEM_ID AND p8.START_DATETIME<=v.START_DATETIME AND p8.END_DATETIME>v.START_DATETIME " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY p16 ON p16.PROPERTY_TYPE='LIFT_TYPE' AND p16.ITEM_ID=i.ITEM_ID AND p16.START_DATETIME<=v.START_DATETIME AND p16.END_DATETIME>v.START_DATETIME " & Chr(10) & _
"          LEFT OUTER JOIN ITEM_PROPERTY p43 ON p43.PROPERTY_TYPE='BALANCE_STATUS' AND p43.ITEM_ID=i.ITEM_ID AND p43.START_DATETIME<=v.START_DATETIME AND p43.END_DATETIME>v.START_DATETIME " & Chr(10) & _
"        WHERE i.ITEM_TYPE='COMPLETION' AND p2.PROPERTY_STRING in ('PRODUCTION','EXPLORATION') AND p6.PROPERTY_STRING='OIL' AND p8.PROPERTY_STRING IN ('UNDER_DEVELOPMENT', 'PRODUCING','WELL_TESTING') AND p43.PROPERTY_STRING='BALANCE' " & Chr(10) & _
"          AND LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'))>=v.START_DATETIME AND LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'))<v.END_DATETIME ) C ON C.ITEM_ID=WELL2UNITS.ITEM_ID " & Chr(10) & _
"      LEFT OUTER JOIN vt_CUM_VOL_DET_RU_RU    MONTH_HIST          ON MONTH_HIST.ITEM_ID= WELL2UNITS.ITEM_ID " & Chr(10) & _
"      LEFT OUTER JOIN VL_WELL_ZONE_RU_RU      WELL_ZONE_LINK      ON WELL2UNITS.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND WELL_ZONE_LINK.TO_ITEM_TYPE='ZONE' AND WELL_ZONE_LINK.FROM_ITEM_TYPE='COMPLETION' " & Chr(10) & _
"      LEFT OUTER JOIN VL_FORMATION_ZONE_RU_RU FORMATION_ZONE_LINK ON WELL_ZONE_LINK.TO_ITEM_ID=FORMATION_ZONE_LINK.TO_ITEM_ID AND FORMATION_ZONE_LINK.TO_ITEM_TYPE='ZONE' AND FORMATION_ZONE_LINK.FROM_ITEM_TYPE='FORMATION' " & Chr(10) & _
"      LEFT OUTER JOIN VI_FORMATION_ALL_RU_RU  FORMATION           ON FORMATION.ITEM_ID=FORMATION_ZONE_LINK.FROM_ITEM_ID " & Chr(10) & _
"    WHERE MONTH_HIST.ITEM_ID=WELL2UNITS.ITEM_ID  and MONTH_HIST.LTD_OIL_MASS>0 and " & Chr(10) & _
"      FORMATION.ITEM_ID is not null " & Chr(10) & _
IIf(ParField_Item_Id <> "", " and WELL2UNITS.ORG_UNIT4_LEGACY_ID in (select LEGACY_ID from VI_ORG_UNIT4_RU_RU where ITEM_ID='" & ParField_Item_Id & "') ", "") & _
"  ) " & Chr(10)
Sql_main = Sql_main & _
"select N, FIELD_NAME, FIELD_ID, LAYER_NAME, LAYER_ID, lift_type, " & Chr(10) & _
"       (select CODE_TEXT from CODE_LIST where LIST_TYPE='WELL_LIFT_TYPE' AND CULTURE='ru-RU' AND CODE = LIFT_TYPE) LIFT_TYPE_TEXT " & Chr(10) & _
"       ,(count(distinct layer_id) over (partition by FIELD_ID /*ORDER BY FIELD_ID*/)) layer_count, " & Chr(10) & _
"       (count(distinct lift_type) over (partition by FIELD_ID, LAYER_ID /*ORDER BY FIELD_ID, LAYER_ID*/)) lift_type_count " & Chr(10) & _
"from " & Chr(10) & _
"  ( select '0PL_LT' N, FIELD_NAME, FIELD_ID, to_char(LAYER_NAME) LAYER_NAME, LAYER_ID, to_char(lift_type) lift_type FROM D UNION " & Chr(10) & _
"    select '1PL' N, FIELD_NAME, FIELD_ID, to_char(LAYER_NAME) LAYER_NAME, LAYER_ID, to_char('') lift_type FROM D UNION " & Chr(10) & _
"    select distinct '2MR_LT' N, FIELD_NAME, FIELD_ID, to_char('') LAYER_NAME, NULL LAYER_ID, to_char(lift_type) lift_type FROM D UNION " & Chr(10) & _
"    select distinct '3MR' N, FIELD_NAME, FIELD_ID, to_char('') LAYER_NAME, NULL LAYER_ID, to_char('') lift_type FROM D )" & Chr(10) & _
"order by FIELD_NAME, FIELD_ID,  LAYER_NAME, LAYER_ID, N, lift_type"

Sql_main = Replace(Sql_main, ":DATAOTCHETA", "01." & Month(StartDay) & "." & Year(StartDay))
'.Cells(1, 10) = Sql_main

Set RS = New ADODB.Recordset
RS.ActiveConnection = cn ' Assign the Connection object.
RS.CursorType = adOpenStatic
RS.Open Sql_main  ' Extract the required records.
'Sheets("Empty_Template").Range("D20").Value = Sql_main
'"       ,nvl(round(sum(PROST_DAYS),2),0) PROST_DAYS " & _

'Основной блок основного запроса >>>
Sql = Sql & _
"  SELECT count(distinct work_well_count)         WORK_WELL_COUNT " & Chr(10) & _
"       ,count(distinct expl_well_count)         ALL_WELL_COUNT " & Chr(10) & _
"       ,nvl(sum(OIL_MASS),0)                    OIL_MASS " & Chr(10) & _
"       ,nvl(sum(OIL_MASS_CUM_YEAR),0)           OIL_MASS_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(OIL_MASS_CUM_FULL),0)           OIL_MASS_CUM_FULL " & Chr(10) & _
"       ,nvl(sum(WATER_MASS),0)                  WATER_MASS " & Chr(10) & _
"       ,nvl(sum(WATER_MASS_CUM_YEAR),0)         WATER_MASS_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(WATER_MASS_CUM_FULL),0)         WATER_MASS_CUM_FULL " & Chr(10) & _
"       ,nvl(sum(WATER_VOL),0)                   WATER_VOL " & Chr(10) & _
"       ,nvl(sum(WATER_VOL_CUM_YEAR),0)          WATER_VOL_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(WATER_VOL_CUM_FULL),0) WATER_VOL_CUM_FULL " & Chr(10) & _
"       ,sum(nvl(OIL_MASS,0)+nvl(WATER_MASS,0))                    LIQ_MASS " & Chr(10) & _
"       ,sum(nvl(OIL_MASS_CUM_YEAR,0)+nvl(WATER_MASS_CUM_YEAR,0))  LIQ_MASS_CUM_YEAR " & Chr(10) & _
"       ,sum(nvl(OIL_MASS_CUM_FULL,0)+nvl(WATER_MASS_CUM_FULL,0))  LIQ_MASS_CUM_FULL " & Chr(10) & _
"       ,sum(nvl(OIL_VOL,0)+nvl(WATER_VOL,0))                      LIQ_VOL " & Chr(10) & _
"       ,sum(nvl(OIL_VOL_CUM_YEAR,0)+nvl(WATER_VOL_CUM_YEAR,0))    LIQ_VOL_CUM_YEAR " & Chr(10) & _
"       ,sum(nvl(OIL_VOL_CUM_FULL,0)+nvl(WATER_VOL_CUM_FULL,0))    LIQ_VOL_CUM_FULL " & Chr(10) & _
"       ,sum(nvl(nvl(OIL_VOL_PL,0)          + nvl(WATER_VOL,0),0))         LIQ_VOL_PL " & Chr(10) & _
"       ,sum(nvl(nvl(OIL_VOL_CUM_YEAR_PL,0) + nvl(WATER_VOL_CUM_YEAR,0),0)) LIQ_VOL_CUM_YEAR_PL " & Chr(10) & _
"       ,sum(nvl(nvl(OIL_VOL_CUM_FULL_PL,0) + nvl(WATER_VOL_CUM_FULL,0),0)) LIQ_VOL_CUM_FULL_PL " & Chr(10)
Sql = Sql & _
"       ,nvl(sum(GAS_VOL),0)          GAS_VOL " & Chr(10) & _
"       ,nvl(sum(GAS_VOL_CUM_YEAR),0) GAS_VOL_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(GAS_VOL_CUM_FULL),0) GAS_VOL_CUM_FULL " & Chr(10) & _
"       ,nvl(sum(GAS_RG_VOL),0)          GAS_RG_VOL " & Chr(10) & _
"       ,nvl(sum(GAS_RG_VOL_CUM_YEAR),0) GAS_RG_VOL_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(GAS_RG_VOL_CUM_FULL),0) GAS_RG_VOL_CUM_FULL " & Chr(10) & _
"       ,nvl(sum(GAS_GH_VOL),0)          GAS_GH_VOL " & Chr(10) & _
"       ,nvl(sum(GAS_GH_VOL_CUM_YEAR),0) GAS_GH_VOL_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(GAS_GH_VOL_CUM_FULL),0) GAS_GH_VOL_CUM_FULL " & Chr(10)
Sql = Sql & _
"       ,nvl(sum(GAS_VOL*1000/zvl(OIL_MASS)),0)                           GAS_FACTOR " & Chr(10) & _
"       ,nvl(sum(GAS_VOL_CUM_YEAR)*1000/zvl(sum(OIL_MASS_CUM_YEAR)),0)    GAS_FACTOR_YEAR " & Chr(10) & _
"       ,nvl(sum(WATER_MASS)*100/zvl(nvl(sum(OIL_MASS),0)+nvl(sum(WATER_MASS),0)),0)    WC_MASS " & Chr(10) & _
"       ,nvl(sum(WATER_VOL)*100/zvl(nvl(sum(OIL_VOL),0)+nvl(sum(WATER_VOL),0)),0)       WC_VOL " & Chr(10) & _
"       /*,nvl(max(WATER_DENSITY),0) WATER_DENSITY*/ " & Chr(10) & _
"       ,nvl(sum(WATER_MASS_CUM_YEAR)*100/zvl(nvl(sum(OIL_MASS_CUM_YEAR),0)+nvl(sum(WATER_MASS_CUM_YEAR),0)),0) WC_MASS_YEAR " & Chr(10) & _
"       ,nvl(sum(OIL_MASS)/zvl(sum(PROD_HOURS+ACC_HOURS_MTH)/24),0)                   OIL_RATE_MASS " & Chr(10) & _
"       ,nvl(sum(OIL_MASS_CUM_YEAR)/zvl(sum(PROD_DAYS_CUM_YEAR+ACC_HOURS_YEAR/24)),0) OIL_RATE_MASS_YEAR " & Chr(10) & _
"       ,nvl((nvl(sum(OIL_MASS),0)+nvl(sum(WATER_MASS),0))/zvl(sum(PROD_HOURS+ACC_HOURS_MTH)/24),0) LIQ_RATE_MASS " & Chr(10) & _
"       ,nvl((nvl(sum(OIL_VOL),0)+nvl(sum(WATER_VOL),0))/zvl(sum(PROD_HOURS+ACC_HOURS_MTH)/24),0)   LIQ_RATE_VOL " & Chr(10) & _
"       ,nvl((nvl(sum(OIL_MASS_CUM_YEAR),0)+nvl(sum(WATER_MASS_CUM_YEAR),0))/zvl(sum(PROD_DAYS_CUM_YEAR+ACC_HOURS_YEAR/24)),0) LIQ_RATE_MASS_YEAR " & Chr(10) & _
"       ,nvl(sum(GAS_VOL)/zvl(sum(PROD_HOURS+ACC_HOURS_MTH)/24),0)                   GAS_RATE_VOL " & Chr(10) & _
"       ,nvl(sum(GAS_VOL_CUM_YEAR)/zvl(sum(PROD_DAYS_CUM_YEAR+ACC_HOURS_YEAR/24)),0) GAS_RATE_VOL_YEAR " & Chr(10) & _
"       ,nvl(sum(OIL_MASS)/zvl(decode(sum(PROD_HOURS), 0, 0, to_number(to_char(last_day(to_date(':DATAOTCHETA','DD.MM.YYYY')),'DD'))),null),0) OIL_RATE_MASS_AVERAGE " & Chr(10) & _
"       ,nvl(sum(OIL_MASS)/zvl( nvl(sum(PROD_DAYS),0) + nvl(sum(PROST_HOURS/24),0) ),0) OIL_RATE_MASS_AVERAGE_2 " & Chr(10) & _
"       ,nvl((sum(nvl(OIL_VOL ,0))+sum(nvl(WATER_VOL ,0))) /zvl(decode(sum(PROD_HOURS), 0, 0, to_number(to_char(last_day(to_date(':DATAOTCHETA','DD.MM.YYYY')),'DD'))),null),0) LIQ_RATE_VOL_AVERAGE " & Chr(10) & _
"       ,nvl((sum(nvl(OIL_MASS,0))+sum(nvl(WATER_MASS,0))) /zvl(sum(decode(PROD_HOURS, 0, 0, to_number(to_char(last_day(START_DATETIME),'DD')))),null),0) LIQ_RATE_MASS_AVERAGE " & Chr(10)
Sql = Sql & _
"       ,nvl(sum(PROD_DAYS_CUM_YEAR),0) PROD_DAYS_CUM_YEAR " & Chr(10) & _
"       ,DECODE(SIGN(nvl(SUM(PROST_DAYS_CUM_YEAR),0) - nvl(SUM(ACC_HOURS_YEAR/24),0)),-1,0,(nvl(SUM(PROST_DAYS_CUM_YEAR),0) - nvl(SUM(ACC_HOURS_YEAR/24),0))) PROST_DAYS_CUM_YEAR " & Chr(10) & _
"       ,nvl(sum(PROD_HOURS),0) PROD_HOURS " & Chr(10) & _
"       ,DECODE(SIGN(nvl(SUM(PROST_HOURS),0) - nvl(SUM(ACC_HOURS_MTH),0)) ,-1,0,(nvl(SUM(PROST_HOURS),0) - nvl(SUM(ACC_HOURS_MTH),0)))  PROST_HOURS " & Chr(10) & _
"       ,nvl(sum(PROD_HOURS_CALENDAR),0) PROD_HOURS_CALENDAR " & Chr(10) & _
"       ,nvl(sum(PROD_DAYS_CUM_FULL),0) PROD_DAYS_CUM_FULL " & Chr(10) & _
"       ,nvl(sum(PROD_DAYS),0) PROD_DAYS " & Chr(10) & _
"       ,DECODE(SIGN(nvl(SUM(PROST_HOURS),0)-nvl(SUM(ACC_HOURS_MTH),0)),-1,0,(nvl(SUM(PROST_HOURS),0)-nvl(SUM(ACC_HOURS_MTH),0)))/24 PROST_DAYS " & Chr(10) & _
"       ,nvl(sum(PROD_HOURS+NVL(ACC_HOURS_MTH,0))/zvl(sum(PROD_HOURS+NVL(PROST_HOURS,0)),null),0) K_EKSPL " & Chr(10) & _
"       ,nvl(sum(PROD_DAYS_CUM_YEAR+NVL(ACC_HOURS_YEAR/24,0))/zvl(sum(PROD_DAYS_CUM_YEAR+NVL(ACC_HOURS_YEAR/24,0))+DECODE(SIGN(nvl(SUM(PROST_DAYS_CUM_YEAR),0) - nvl(SUM(ACC_HOURS_YEAR/24),0)),-1,0,(nvl(SUM(PROST_DAYS_CUM_YEAR),0) - nvl(SUM(ACC_HOURS_YEAR/24),0))),null),0) K_EKSPL_YEAR " & Chr(10) & _
"       ,nvl(sum(STATUS_FOR_PEREHOD),0) STATUS_FOR_PEREHOD " & Chr(10) & _
"       ,nvl(sum(PROD_HOURS+NVL(ACC_HOURS_MTH,0))/zvl(sum(CALENDAR_HOURS_FOR_EKSPL_SKVS),null),0) K_ISPOLZ " & Chr(10) & _
"       ,nvl(sum(ACC_HOURS_MTH    ),0) ACC_HOURS_MTH " & Chr(10) & _
"       ,nvl(sum(ACC_HOURS_MTH/24 ),0) ACC_DAYS_MTH " & Chr(10) & _
"       ,nvl(sum(ACC_HOURS_YEAR   ),0) ACC_HOURS_YEAR " & Chr(10) & _
"       ,nvl(sum(ACC_HOURS_YEAR/24),0) ACC_DAYS_YEAR " & Chr(10) & _
"       ,NVL(SUM(PROD_DAYS_CUM_YEAR+NVL(ACC_HOURS_YEAR/24,0)),0) EKSPL_DAIS_CUM_YEAR, NVL(SUM(PROD_DAYS_CUM_FULL+NVL(ACC_HOURS_FULL/24,0)),0) EKSPL_DAIS_CUM_FULL " & Chr(10)
Sql = Sql & _
"  FROM ( " & Chr(10) & _
"    SELECT /*distinct */ " & Chr(10) & _
"       MHP.work_well_count " & Chr(10) & _
"      ,CASE WHEN EKSPL_DAIS_CUM_FULL>0 THEN MHP.item_id END expl_well_count " & Chr(10) & _
"      ,MHP.item_id " & Chr(10) & _
"      ,MHP.OIL_FVF " & Chr(10) & _
"      ,MHP.START_DATETIME " & Chr(10) & _
"      /*,nvl(MHP.WATER_DENSITY,0) WATER_DENSITY*/ " & Chr(10) & _
"      ,nvl(MHP.MTD_OIL_MASS ,0) OIL_MASS " & Chr(10) & _
"      ,nvl(MHP.YTD_OIL_MASS ,0) OIL_MASS_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_OIL_MASS ,0) OIL_MASS_CUM_FULL " & Chr(10) & _
"      ,nvl(MHP.MTD_OIL_VOL  ,0) OIL_VOL " & Chr(10) & _
"      ,nvl(MHP.YTD_OIL_VOL  ,0) OIL_VOL_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_OIL_VOL  ,0) OIL_VOL_CUM_FULL " & Chr(10)
Sql = Sql & "      ,NVL(MHP_PL.OIL_VOL_PL,0) OIL_VOL_PL, NVL(MHP_PL.OIL_VOL_CUM_YEAR_PL,0) OIL_VOL_CUM_YEAR_PL, NVL(MHP_PL.OIL_VOL_CUM_FULL_PL,0) OIL_VOL_CUM_FULL_PL " & Chr(10)
'Sql = Sql & _
'"      ,NVL((SELECT SUM(ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100) " & Chr(10) & _
'"            FROM VT_TOT_DET_MTH_RU_RU t INNER JOIN VI_ZONE_ALL_RU_RU z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
'"              ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK  WZ " & Chr(10) & _
'"                  LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
'"                WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
'"            WHERE MHP.ZONE_ID=z.ITEM_ID AND t.ITEM_ID=MHP.ITEM_ID AND t.lift_type=MHP.lift_type " & Chr(10) & _
'"              AND t.START_DATETIME=MHP.START_DATETIME ),0) OIL_VOL_PL " & Chr(10) & _
'"      ,NVL((SELECT SUM(ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100) " & Chr(10) & _
'"            FROM VT_TOT_DET_MTH_RU_RU t INNER JOIN VI_ZONE_ALL_RU_RU z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
'"              ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK  WZ " & Chr(10) & _
'"                  LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
'"                WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
'"            WHERE MHP.ZONE_ID=z.ITEM_ID AND t.ITEM_ID=MHP.ITEM_ID AND t.lift_type=MHP.lift_type " & Chr(10) & _
'"              AND t.START_DATETIME between TRUNC(MHP.START_DATETIME, 'YYYY') AND MHP.START_DATETIME ),0) OIL_VOL_CUM_YEAR_PL " & Chr(10) & _
'"      ,NVL((SELECT SUM(ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100) " & Chr(10) & _
'"            FROM VT_TOT_DET_MTH_RU_RU t INNER JOIN VI_ZONE_ALL_RU_RU z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
'"              ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK  WZ " & Chr(10) & _
'"                  LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
'"                WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
'"            WHERE MHP.ZONE_ID=z.ITEM_ID AND t.ITEM_ID=MHP.ITEM_ID AND t.lift_type=MHP.lift_type " & Chr(10) & _
'"              AND t.START_DATETIME<=MHP.START_DATETIME ),0) OIL_VOL_CUM_FULL_PL " & Chr(10)
Sql = Sql & _
"      ,nvl(MHP.MTD_WATER_MASS ,0) WATER_MASS " & Chr(10) & _
"      ,nvl(MHP.YTD_WATER_MASS ,0) WATER_MASS_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_WATER_MASS ,0) WATER_MASS_CUM_FULL " & Chr(10) & _
"      ,nvl(MHP.MTD_WATER_VOL  ,0) WATER_VOL " & Chr(10) & _
"      ,nvl(MHP.YTD_WATER_VOL  ,0) WATER_VOL_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_WATER_VOL  ,0) WATER_VOL_CUM_FULL " & Chr(10) & _
"      ,nvl(MHP.MTD_GAS_VOL ,0) GAS_VOL " & Chr(10) & _
"      ,nvl(MHP.YTD_GAS_VOL ,0) GAS_VOL_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_GAS_VOL ,0) GAS_VOL_CUM_FULL " & Chr(10) & _
"      ,nvl(MHP.MTD_GAS_RG_VOL ,0) GAS_RG_VOL " & Chr(10) & _
"      ,nvl(MHP.YTD_GAS_RG_VOL ,0) GAS_RG_VOL_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_GAS_RG_VOL ,0) GAS_RG_VOL_CUM_FULL " & Chr(10) & _
"      ,nvl(MHP.MTD_GAS_VOL ,0)-nvl(MHP.MTD_GAS_RG_VOL ,0) GAS_GH_VOL " & Chr(10) & _
"      ,nvl(MHP.YTD_GAS_VOL ,0)-nvl(MHP.YTD_GAS_RG_VOL ,0) GAS_GH_VOL_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_GAS_VOL ,0)-nvl(MHP.LTD_GAS_RG_VOL ,0) GAS_GH_VOL_CUM_FULL " & Chr(10)
Sql = Sql & _
"      ,nvl(MHP.MTD_PROD_HOURS,0) PROD_HOURS " & Chr(10) & _
"      ,nvl(MHP.MTD_PROD_HOURS/24,0) PROD_DAYS " & Chr(10) & _
"      ,nvl(MHP.YTD_PROD_HOURS/24,0) PROD_DAYS_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.LTD_PROD_HOURS/24,0) PROD_DAYS_CUM_FULL " & Chr(10) & _
"      ,FN_MTH_INACT_HOURS_OIL_RNI(MHP.START_DATETIME,MHP.MH_ITEM_ID,MHP.LIFT_TYPE) PROST_HOURS " & Chr(10) & _
"      ,MHP.PROST_DAYS_CUM_YEAR " & Chr(10) & _
"      ,nvl(MHP.PROD_HOURS_CALENDAR,0) PROD_HOURS_CALENDAR " & Chr(10) & _
"      ,nvl(MHP.PROD_DAYS_CALENDAR_YEAR,0) PROD_DAYS_CALENDAR_YEAR " & Chr(10)
Sql = Sql & _
"      ,(SELECT max(p.property_date) DATE_ON_PROD FROM item i, item_property p " & Chr(10) & _
"         where i.ITEM_ID=p.ITEM_ID AND i.item_type='COMPLETION' AND p.property_type='DATE_ON_PROD' AND i.ITEM_ID=MHP.MH_ITEM_ID AND last_day(MHP.START_DATETIME) between p.START_DATETIME AND p.END_DATETIME " & Chr(10) & _
"       ) DATE_ON_PROD " & Chr(10) & _
"      ,(SELECT max(p.property_string) BORE_DIRECTION FROM item i, item_property p " & Chr(10) & _
"         where i.ITEM_ID=p.ITEM_ID AND i.item_type='COMPLETION' AND p.property_type='BORE_DIRECTION' AND i.ITEM_ID=MHP.MH_ITEM_ID AND last_day(MHP.START_DATETIME) between p.START_DATETIME AND p.END_DATETIME " & Chr(10) & _
"       ) BORE_DIRECTION " & Chr(10) & _
"      ,nvl((SELECT sum(decode(CHAR2/*status*/,'PRODUCING',0,'SHUT_IN',0,1)) FROM ITEM_EVENT WHERE EVENT_TYPE='RSP_STOCK_HIST' AND ITEM_ID=MHP.MH_ITEM_ID AND START_DATETIME=trunc(MHP.START_DATETIME,'YYYY')-1 " & Chr(10) & _
"           ),1) STATUS_FOR_PEREHOD " & Chr(10) & _
"      ,CASE WHEN /*MHP.LIFT_TYPE=(SELECT LIFT_TYPE FROM VT_RSP_STOCK_HIST_RU_RU WHERE ITEM_ID=MHP.ITEM_ID AND START_DATETIME=(SELECT MAX(START_DATETIME) FROM VT_RSP_STOCK_HIST_RU_RU WHERE ITEM_ID=MHP.ITEM_ID AND START_DATETIME<LAST_DAY(MHP.START_DATETIME)+1))*/ " & Chr(10) & _
"           EXISTS(SELECT LIFT_TYPE FROM VT_TOT_DET_DAY_RU_RU WHERE ITEM_ID=MHP.ITEM_ID AND LIFT_TYPE=MHP.LIFT_TYPE AND TRUNC(START_DATETIME,'MM')=TRUNC(MHP.START_DATETIME,'MM')) " & Chr(10) & _
"         THEN 24*(to_number(to_char(last_day(MHP.START_DATETIME),'DD'))-DECODE(TRUNC(MHP.COMPLETION_DATE_ON_PROD,'MM'),MHP.START_DATETIME,TO_NUMBER(TO_CHAR(MHP.COMPLETION_DATE_ON_PROD,'DD'))-1,0)) " & Chr(10) & _
"         ELSE NULL END CALENDAR_HOURS_FOR_EKSPL_SKVS " & Chr(10) & _
"      ,ROUND(NVL((select sum(tot.acc_time/60/60) from vt_tot_det_day_ru_ru totd inner join vt_totals_day_ru_ru tot on totd.item_id=tot.item_id and totd.start_datetime=tot.start_datetime " & Chr(10) & _
"                   where totd.item_id=MHP.ITEM_ID and totd.lift_type=MHP.lift_type and trunc(totd.start_datetime,'mm')=MHP.START_DATETIME " & Chr(10) & _
"        ),0),5) ACC_HOURS_MTH " & Chr(10) & _
"       ,ROUND(NVL((select sum(tot.acc_time/60/60) from vt_tot_det_day_ru_ru totd inner join vt_totals_day_ru_ru tot on totd.item_id=tot.item_id and totd.start_datetime=tot.start_datetime " & Chr(10) & _
"                   where totd.item_id=MHP.ITEM_ID and totd.lift_type=MHP.lift_type and trunc(totd.start_datetime,'mm') between trunc(MHP.START_DATETIME,'yyyy') and MHP.START_DATETIME " & Chr(10) & _
"        ),0),5) ACC_HOURS_YEAR " & Chr(10) & _
"       ,ROUND(NVL((select sum(tot.acc_time/60/60) from vt_tot_det_day_ru_ru totd inner join vt_totals_day_ru_ru tot on totd.item_id=tot.item_id and totd.start_datetime=tot.start_datetime " & Chr(10) & _
"                   where totd.item_id=MHP.ITEM_ID and totd.lift_type=MHP.lift_type and trunc(totd.start_datetime,'mm')<=MHP.START_DATETIME " & Chr(10) & _
"        ),0),5) ACC_HOURS_FULL " & Chr(10)
Sql = Sql & _
"    FROM " & Chr(10) & _
"   ( SELECT FORMATION_ZONE_LINK.TO_ITEM_ID AS ZONE_ID, " & Chr(10) & _
"       WELL2UNITS.ITEM_ID, WELL2UNITS.COMPLETION_NAME AS WELL_NAME, WELL2UNITS.ORG_UNIT4_LEGACY_ID AS FIELD_ID, WELL2UNITS.ORG_UNIT4_NAME AS FIELD_NAME, " & Chr(10) & _
"       FORMATION.ITEM_ID AS LAYER_ID, FORMATION.ITEM_NAME AS LAYER_NAME, MH.ITEM_ID AS MH_ITEM_ID, " & Chr(10) & _
"       MH.LIFT_TYPE/*NVL(MH.LIFT_TYPE,COMP.LIFT_TYPE)*/ LIFT_TYPE, MH.LIFT_TYPE_TEXT, " & Chr(10)
Sql = Sql & _
"       MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') AND MH.MTD_PROD_HOURS>0 THEN MH.ITEM_ID END) AS WORK_WELL_COUNT, " & Chr(10) & _
"       MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN FORMATION.OIL_FVF END) AS OIL_FVF, " & Chr(10) & _
"       /*MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(ZVL(FORMATION.WATER_DENSITY_1),1.005) END) AS WATER_DENSITY,*/ " & Chr(10) & _
"       TO_DATE(':DATAOTCHETA','DD.MM.YYYY') AS START_DATETIME, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS MTD_OIL_MASS, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS YTD_OIL_MASS, /*MH.YTD_OIL_MASS,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS LTD_OIL_MASS, /*MH.LTD_OIL_MASS,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_VOL*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS MTD_OIL_VOL, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_VOL*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS YTD_OIL_VOL, /*MH.YTD_OIL_VOL,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_OIL_VOL*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1) END) AS LTD_OIL_VOL, /*MH.LTD_OIL_VOL,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_MASS*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS MTD_WATER_MASS, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_MASS*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS YTD_WATER_MASS, /*MH.YTD_WATER_MASS,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_MASS*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS LTD_WATER_MASS, /*MH.LTD_WATER_MASS,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_VOL*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS MTD_WATER_VOL, " & Chr(10)
Sql = Sql & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_VOL*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS YTD_WATER_VOL, /*MH.YTD_WATER_VOL,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_WATER_VOL*NVL(WELL_ZONE_LINK.REL_WATER_VOL,1) END) AS LTD_WATER_VOL, /*MH.LTD_WATER_VOL,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1) END) AS MTD_GAS_VOL, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1) END) AS YTD_GAS_VOL, /*MH.YTD_GAS_VOL,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1) END) AS LTD_GAS_VOL, /*MH.LTD_GAS_VOL,*/ " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1)*WELL_ZONE_LINK.RNI_WGHOR,MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1)) END) AS MTD_GAS_RG_VOL, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1)*WELL_ZONE_LINK.RNI_WGHOR,MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1)) END) AS YTD_GAS_RG_VOL, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME<=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN NVL(MH.MTD_OIL_MASS*NVL(WELL_ZONE_LINK.REL_OIL_VOL,1)*WELL_ZONE_LINK.RNI_WGHOR,MH.MTD_GAS_VOL*NVL(WELL_ZONE_LINK.REL_GAS_VOL,1)) END) AS LTD_GAS_RG_VOL, " & Chr(10) & _
"       MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.MTD_PROD_HOURS END) AS MTD_PROD_HOURS, " & Chr(10) & _
"       MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.YTD_PROD_HOURS END) AS YTD_PROD_HOURS, " & Chr(10) & _
"       MAX(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN MH.LTD_PROD_HOURS END) AS LTD_PROD_HOURS, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN FN_MTH_INACT_HOURS_OIL_RNI(MH.START_DATETIME,MH.ITEM_ID,MH.LIFT_TYPE) END)/24 AS PROST_DAYS_CUM_YEAR, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME BETWEEN TRUNC(TO_DATE(':DATAOTCHETA','DD.MM.YYYY'),'YYYY') AND TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN decode(MH.MTD_PROD_HOURS, 0, 0, null, 0, to_char(last_day(MH.START_DATETIME),'DD')) END) AS PROD_DAYS_CALENDAR_YEAR, " & Chr(10) & _
"       SUM(CASE WHEN MH.START_DATETIME =TO_DATE(':DATAOTCHETA','DD.MM.YYYY') THEN decode(MH.MTD_PROD_HOURS, 0, 0, null, 0, to_char(last_day(MH.START_DATETIME),'DD')) END)*24 AS PROD_HOURS_CALENDAR  " & Chr(10)
Sql = Sql & _
"       ,WELL2UNITS.COMPLETION_DATE_ON_PROD " & Chr(10) & _
"       /*,SUM(CASE WHEN MH.START_DATETIME=TO_DATE(':DATAOTCHETA','DD.MM.YYYY') AND MH.MTD_PROD_HOURS>0 OR MH.LIFT_TYPE IS NULL " & Chr(10) & _
"                 THEN CASE WHEN ADD_MONTHS(MH.START_DATETIME,1)>WELL2UNITS.COMPLETION_DATE_ON_PROD " & Chr(10) & _
"                           THEN ADD_MONTHS(MH.START_DATETIME,1)-GREATEST(TRUNC(MH.START_DATETIME,'YYYY'),NVL((SELECT MIN(END_DATETIME) FROM VT_DOWNTIME_RU_RU WHERE ITEM_ID=WELL2UNITS.ITEM_ID AND DOWNTIME_TYPE='PP0000' AND TRUNC(START_DATETIME,'DD')=WELL2UNITS.COMPLETION_DATE_ON_PROD),WELL2UNITS.COMPLETION_DATE_ON_PROD)) " & Chr(10) & _
"                           ELSE 0 END + NVL(COMP.OSV_DAIS_CUM_YEAR,0) END) EKSPL_DAIS_CUM_YEAR*/ " & Chr(10) & _
"       ,SUM(CASE WHEN MH.START_DATETIME =TO_DATE(':DATAOTCHETA','DD.MM.YYYY') AND MH.MTD_PROD_HOURS>0 OR MH.LIFT_TYPE IS NULL " & Chr(10) & _
"                 THEN CASE WHEN ADD_MONTHS(MH.START_DATETIME,1)>WELL2UNITS.COMPLETION_DATE_ON_PROD " & Chr(10) & _
"                      THEN ADD_MONTHS(MH.START_DATETIME,1)-NVL((SELECT MIN(END_DATETIME) FROM VT_DOWNTIME_RU_RU WHERE ITEM_ID=WELL2UNITS.ITEM_ID AND DOWNTIME_TYPE='PP0000' AND TRUNC(START_DATETIME,'DD')=WELL2UNITS.COMPLETION_DATE_ON_PROD),WELL2UNITS.COMPLETION_DATE_ON_PROD) " & Chr(10) & _
"                 ELSE 0 END /*+ NVL(COMP.OSV_DAIS_CUM_FULL,0)*/ END) EKSPL_DAIS_CUM_FULL " & Chr(10)

Sql = Sql & _
"     FROM VCUSTOM_WELL2UNITS WELL2UNITS " & Chr(10) & _
"       LEFT OUTER JOIN vt_CUM_VOL_DET_ru_ru MH ON MH.ITEM_ID=WELL2UNITS.ITEM_ID " & Chr(10) & _
"       LEFT OUTER JOIN ( SELECT WZ.LINK_ID, WZ.TO_ITEM_ID, WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p0.PROPERTY_VALUE/100 AS REL_GAS_VOL, p1.PROPERTY_VALUE/100 AS REL_OIL_VOL, p2.PROPERTY_VALUE/100 AS REL_WATER_VOL, " & Chr(10) & _
"                           p9.PROPERTY_STRING AS RNI_ZONE_AGENT, DECODE(p9.PROPERTY_STRING,'OIL+WATER+GAS',p10.PROPERTY_VALUE,NULL)/1000 AS RNI_WGHOR FROM ITEM_LINK  WZ " & Chr(10) & _
"                           LEFT OUTER JOIN ITEM_LINK_PROPERTY p0 ON p0.LINK_TYPE='WELL_ZONE' AND p0.LINK_ID=WZ.LINK_ID AND p0.START_DATETIME<=WZ.START_DATETIME AND p0.END_DATETIME>WZ.START_DATETIME AND p0.PROPERTY_TYPE='REL_GAS_VOL' " & Chr(10) & _
"                           LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
"                           LEFT OUTER JOIN ITEM_LINK_PROPERTY p2 ON p2.LINK_TYPE='WELL_ZONE' AND p2.LINK_ID=WZ.LINK_ID AND p2.START_DATETIME<=WZ.START_DATETIME AND p2.END_DATETIME>WZ.START_DATETIME AND p2.PROPERTY_TYPE='REL_WATER_VOL' " & Chr(10) & _
"                           LEFT OUTER JOIN ITEM_LINK_PROPERTY p9 ON p9.LINK_TYPE='WELL_ZONE' AND p9.LINK_ID=WZ.LINK_ID AND p9.START_DATETIME<=WZ.START_DATETIME AND p9.END_DATETIME>WZ.START_DATETIME AND p9.PROPERTY_TYPE='RNI_ZONE_AGENT' " & Chr(10) & _
"                           LEFT OUTER JOIN ITEM_LINK_PROPERTY p10 ON p10.LINK_TYPE='WELL_ZONE' AND p10.LINK_ID=WZ.LINK_ID AND p10.START_DATETIME<=WZ.START_DATETIME AND p10.END_DATETIME>WZ.START_DATETIME AND p10.PROPERTY_TYPE='RNI_WGHOR' " & Chr(10) & _
"                         WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON WELL2UNITS.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND NVL(MH.START_DATETIME,LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY')))>=WELL_ZONE_LINK.START_DATETIME AND NVL(MH.START_DATETIME,LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY')))<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
"       LEFT OUTER JOIN VL_FORMATION_ZONE_RU_RU FORMATION_ZONE_LINK ON WELL_ZONE_LINK.TO_ITEM_ID=FORMATION_ZONE_LINK.TO_ITEM_ID  " & Chr(10)

'Sql = Sql & _
'"       LEFT OUTER JOIN VI_FORMATION_ALL_RU_RU FORMATION ON FORMATION.ITEM_ID=FORMATION_ZONE_LINK.FROM_ITEM_ID AND NVL(MH.START_DATETIME,LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY')))>=FORMATION.START_DATETIME AND NVL(MH.START_DATETIME,LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY')))<FORMATION.END_DATETIME " & Chr(10) & _
'"     WHERE NVL(MH.LTD_OIL_MASS,0)>0 " & Chr(10) & _
'"       /*AND NOT EXISTS(SELECT RNI_ZONE_AGENT FROM VL_WELL_ZONE_RU_RU WHERE WELL2UNITS.ITEM_ID=FROM_ITEM_ID AND FORMATION_ZONE_LINK.TO_ITEM_ID=TO_ITEM_ID AND RNI_ZONE_AGENT='OIL+WATER+GAS')*/ /*FOR VNZ*/ " & Chr(10) & _
'"       :USLOVIE1 " & Chr(10) & _
'"     GROUP BY WELL2UNITS.ITEM_ID, MH.ITEM_ID, WELL2UNITS.COMPLETION_NAME, WELL2UNITS.ORG_UNIT4_LEGACY_ID, WELL2UNITS.ORG_UNIT4_NAME, FORMATION_ZONE_LINK.TO_ITEM_ID, FORMATION.ITEM_ID, FORMATION.ITEM_NAME, MH.LIFT_TYPE/*NVL(MH.LIFT_TYPE,COMP.LIFT_TYPE)*/, MH.LIFT_TYPE_TEXT, WELL2UNITS.COMPLETION_DATE_ON_PROD ) " & Chr(10) & _
'" MHP ) " & Chr(10)'

Sql = Sql & _
"       LEFT OUTER JOIN ( " & Chr(10) & _
"         SELECT i.ITEM_ID, p0.PROPERTY_STRING AS ITEM_NAME, v.START_DATETIME, v.END_DATETIME, p24.PROPERTY_VALUE OIL_FVF/*, p35.PROPERTY_VALUE WATER_DENSITY_1*/ " & Chr(10) & _
"         FROM ITEM i INNER JOIN ITEM_VERSION v ON v.ITEM_ID=i.ITEM_ID AND i.ITEM_TYPE='FORMATION' " & Chr(10) & _
"           LEFT OUTER JOIN ITEM_PROPERTY p0 ON v.ITEM_ID=p0.ITEM_ID AND p0.START_DATETIME<=v.START_DATETIME AND p0.END_DATETIME>v.START_DATETIME AND p0.PROPERTY_TYPE='NAME' " & Chr(10) & _
"           LEFT OUTER JOIN ITEM_PROPERTY p24 ON p24.ITEM_ID=i.ITEM_ID AND p24.START_DATETIME<=v.START_DATETIME AND p24.END_DATETIME>v.START_DATETIME AND p24.PROPERTY_TYPE='OIL_FVF' " & Chr(10) & _
"           /*LEFT OUTER JOIN ITEM_PROPERTY p35 ON p35.ITEM_ID=i.ITEM_ID AND p35.START_DATETIME<=v.START_DATETIME AND p35.END_DATETIME>v.START_DATETIME AND p35.PROPERTY_TYPE='WATER_DENSITY_1'*/ " & Chr(10) & _
"        ) FORMATION ON FORMATION.ITEM_ID=FORMATION_ZONE_LINK.FROM_ITEM_ID AND NVL(MH.START_DATETIME,LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY')))>=FORMATION.START_DATETIME AND NVL(MH.START_DATETIME,LAST_DAY(TO_DATE(':DATAOTCHETA','DD.MM.YYYY')))<FORMATION.END_DATETIME " & Chr(10) & _
"     WHERE NVL(MH.LTD_OIL_MASS,0)>0 and WELL2UNITS.ORG_UNIT4_LEGACY_ID=':ORG' " & Chr(10) & "       :USLOVIE1 " & Chr(10) & _
"     GROUP BY WELL2UNITS.ITEM_ID, MH.ITEM_ID, WELL2UNITS.COMPLETION_NAME, WELL2UNITS.ORG_UNIT4_LEGACY_ID, WELL2UNITS.ORG_UNIT4_NAME, FORMATION_ZONE_LINK.TO_ITEM_ID, FORMATION.ITEM_ID, FORMATION.ITEM_NAME, MH.LIFT_TYPE, MH.LIFT_TYPE_TEXT, WELL2UNITS.COMPLETION_DATE_ON_PROD " & Chr(10) & _
"    ) MHP LEFT OUTER JOIN ( " & Chr(10)

Sql = Sql & _
"      SELECT t.ITEM_ID, z.ITEM_ID ZONE_ID, t.lift_type, " & Chr(10) & _
"        SUM( CASE WHEN t.START_DATETIME=to_date(':DATAOTCHETA','DD.MM.YYYY') THEN t.ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100 END) OIL_VOL_PL, " & Chr(10) & _
"        SUM( CASE WHEN t.START_DATETIME between TRUNC(to_date(':DATAOTCHETA','DD.MM.YYYY'), 'YYYY') AND to_date(':DATAOTCHETA','DD.MM.YYYY') THEN t.ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100 END) OIL_VOL_CUM_YEAR_PL, " & Chr(10) & _
"        SUM( CASE WHEN t.START_DATETIME<=to_date(':DATAOTCHETA','DD.MM.YYYY') THEN t.ACT_OIL_MASS*z.OIL_FVF*NVL(WELL_ZONE_LINK.REL_OIL_VOL,100)/100 END) OIL_VOL_CUM_FULL_PL " & Chr(10) & _
"      FROM VCUSTOM_WELL2UNITS WELL2UNITS INNER JOIN VT_TOT_DET_MTH_RU_RU t ON WELL2UNITS.ITEM_ID=T.ITEM_ID " & Chr(10) & _
"             AND WELL2UNITS.ORG_UNIT4_LEGACY_ID=':ORG' INNER JOIN " & Chr(10) & _
"        ( SELECT i.ITEM_ID, v.START_DATETIME, v.END_DATETIME, p15.PROPERTY_VALUE OIL_FVF FROM ITEM i INNER JOIN ITEM_VERSION v ON v.ITEM_ID=i.ITEM_ID AND i.ITEM_TYPE='ZONE' " & Chr(10) & _
"            LEFT OUTER JOIN ITEM_PROPERTY p15 ON p15.ITEM_ID=i.ITEM_ID AND p15.START_DATETIME<=v.START_DATETIME AND p15.END_DATETIME>v.START_DATETIME AND p15.PROPERTY_TYPE='OIL_FVF' " & Chr(10) & _
"        ) z ON t.START_DATETIME>=z.START_DATETIME AND t.START_DATETIME<z.END_DATETIME LEFT OUTER JOIN " & Chr(10) & _
"        ( SELECT WZ.START_DATETIME, WZ.END_DATETIME, WZ.FROM_ITEM_ID, p1.PROPERTY_VALUE AS REL_OIL_VOL FROM ITEM_LINK WZ " & Chr(10) & _
"            LEFT OUTER JOIN ITEM_LINK_PROPERTY p1 ON p1.LINK_TYPE='WELL_ZONE' AND p1.LINK_ID=WZ.LINK_ID AND p1.START_DATETIME<=WZ.START_DATETIME AND p1.END_DATETIME>WZ.START_DATETIME AND p1.PROPERTY_TYPE='REL_OIL_VOL' " & Chr(10) & _
"          WHERE WZ.LINK_TYPE='WELL_ZONE' ) WELL_ZONE_LINK ON T.ITEM_ID=WELL_ZONE_LINK.FROM_ITEM_ID AND T.START_DATETIME>=WELL_ZONE_LINK.START_DATETIME AND T.START_DATETIME<WELL_ZONE_LINK.END_DATETIME " & Chr(10) & _
"      GROUP BY t.ITEM_ID, z.ITEM_ID, t.lift_type " & Chr(10) & _
"    ) MHP_PL ON MHP.ZONE_ID=MHP_PL.ZONE_ID AND MHP_PL.ITEM_ID=MHP.ITEM_ID AND MHP_PL.lift_type=MHP.lift_type ) "

Sql = Replace(Sql, ":DATAOTCHETA", "01." & Month(StartDay) & "." & Year(StartDay))

'CurrWELL_name = ""
CurrFIELD_name = ""
CurrGRP_name = ""
CurrLAYER_name = ""
CurrLIFT_TYPE_name = ""

is_new_mest = 0

  RowCounter_temp = 35
  If Not (RS.BOF And RS.EOF And RS.RecordCount = 0) Then
    Sql = Replace(Sql, ":ORG", RS.Fields("FIELD_ID"))
    For i = 0 To RS.RecordCount - 1 ' Цикл перебора наборов месторождений-пластов-способов

      If CurrFIELD_name <> RS.Fields("FIELD_NAME") Then
        RowCounter_temp = RowCounter_temp + 2
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Месторождение : " & RS.Fields("FIELD_NAME")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Объект : "
      End If
      ' ПО ПЛАСТУ---------------------------------------------

      If RS.Fields("N") = "0PL_LT" And CurrLAYER_name <> RS.Fields("LAYER_NAME") Then
        RowCounter_temp = RowCounter_temp + 1
        ' Итого по пласту
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Итого по пласту : " & RS.Fields("LAYER_NAME")
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If
      End If

      If RS.Fields("N") = "0PL_LT" Then
        ' Итого по пласту - Итого по способу...
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'" & _
        "   and MH.LIFT_TYPE/*NVL(MH.LIFT_TYPE,COMP.LIFT_TYPE)*/ = '" & RS.Fields("LIFT_TYPE") & "'")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Итого по способу : " & RS.Fields("LIFT_TYPE_TEXT")
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If
      End If

      If RS.Fields("N") = "1PL" Then

        ' Итого по пласту - в том числе по новым :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'") & _
        " where date_on_prod between trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY')) "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по новым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - в том числе по старым :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по старым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - в том числе по переходящим :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and STATUS_FOR_PEREHOD = 0 "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по переходящим : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - в том числе по горизонтальным :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'") & _
        " where bore_direction = 'HORIZONTAL' "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по горизонтальным : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        '---------------------

        ' Итого по пласту - Итого по эксплуатационным скважинам
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Итого по эксплуатационным скважинам"
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - Итого по эксплуатационным скважинам - в том числе по новым :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where date_on_prod between trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY')) "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по новым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - Итого по эксплуатационным скважинам - в том числе по старым :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по старым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - Итого по эксплуатационным скважинам - в том числе по переходящим :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and STATUS_FOR_PEREHOD = 0 "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по переходящим : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по пласту - Итого по эксплуатационным скважинам - в том числе по горизонтальным :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and FORMATION.ITEM_ID = '" & RS.Fields("LAYER_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where bore_direction = 'HORIZONTAL' "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по горизонтальным : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If
      End If
      
      ' ПО МЕСТОРОЖДЕНИЮ---------------------------------------------
          
      If RS.Fields("N") = "2MR_LT" And CurrGRP_name = "1PL" Then
        RowCounter_temp = RowCounter_temp + 1
        ' Итого по месторождению
        Sql_temp = Replace(Sql, ":USLOVIE1", " and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Итого по месторождению : " & RS.Fields("FIELD_NAME")
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If
      End If

      If RS.Fields("N") = "2MR_LT" Then
        ' Итого по месторождению - Итого по способу...
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and MH.LIFT_TYPE/*NVL(MH.LIFT_TYPE,COMP.LIFT_TYPE)*/ = '" & RS.Fields("LIFT_TYPE") & "'")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Итого по способу : " & RS.Fields("LIFT_TYPE_TEXT")
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If
      End If

      If RS.Fields("N") = "3MR" Then
        
        ' Итого по месторождению - в том числе по новым :
        Sql_temp = Replace(Sql, ":USLOVIE1", "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'") & _
        " where date_on_prod between trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY')) "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по новым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - в том числе по старым :
        Sql_temp = Replace(Sql, ":USLOVIE1", "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по старым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - в том числе по переходящим :
        Sql_temp = Replace(Sql, ":USLOVIE1", "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and STATUS_FOR_PEREHOD = 0 "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по переходящим : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - в том числе по горизонтальным :
        Sql_temp = Replace(Sql, ":USLOVIE1", "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'") & _
        " where bore_direction = 'HORIZONTAL' "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по горизонтальным : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        '---------------------

        ' Итого по месторождению - Итого по эксплуатационным скважинам
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ")
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "Итого по эксплуатационным скважинам"
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - Итого по эксплуатационным скважинам - в том числе по новым :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where date_on_prod between trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY')) "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по новым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - Итого по эксплуатационным скважинам - в том числе по старым :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по старым : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - Итого по эксплуатационным скважинам - в том числе по переходящим :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where date_on_prod < trunc(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "', 'DD.MM.YYYY'),'YYYY') " & _
        "   and STATUS_FOR_PEREHOD = 0 "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по переходящим : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If

        ' Итого по месторождению - Итого по эксплуатационным скважинам - в том числе по горизонтальным :
        Sql_temp = Replace(Sql, ":USLOVIE1", _
        "   and WELL2UNITS.ORG_UNIT4_LEGACY_ID = '" & RS.Fields("FIELD_ID") & "'" & _
        "   and WELL2UNITS.ITEM_ID in (select ITEM_ID from VT_RSP_STOCK_HIST_RU_RU STOCK_HIST" & _
        "   where last_day(to_date('" & Day(StartDay) & "." & Month(StartDay) & "." & Year(StartDay) & "','DD.MM.YYYY')) = STOCK_HIST.START_DATETIME " & _
        "   and STATUS in ('PRODUCING','SHUT_IN','SHUT_IN_CURRENT_YEAR','SHUT_IN_PREVIOUS_YEARS','UNDER_DEVELOPMENT','UNDER_DEVELOPMENT_CURRENT_YEAR','UNDER_DEVELOPMENT_PREVIOUS_YEARS','AWAITING_DEVELOPMENT','AWAITING_DEVELOPMENT_CURRENT_YEAR','AWAITING_DEVELOPMENT_PREVIOUS_YEARS')) ") & _
        " where bore_direction = 'HORIZONTAL' "
        RowCounter_temp = RowCounter_temp + 1
        .Cells(RowCounter_temp, 1).Font.Bold = 1
        .Cells(RowCounter_temp, 1) = "в том числе по горизонтальным : "
        RowCounter_temp = RowCounter_temp + 1
        If Draw_by_Query(cn, Sql_temp, RowCounter_temp) = True Then
          .Rows(CStr(RowCounter_temp - 1) & ":" & CStr(RowCounter_temp)).ClearContents
          RowCounter_temp = RowCounter_temp - 2
        End If
      End If
      
      CurrGRP_name = RS.Fields("N")
      CurrLAYER_name = IIf(IsNull(RS.Fields("LAYER_NAME")), "", RS.Fields("LAYER_NAME"))
      CurrFIELD_name = IIf(IsNull(RS.Fields("FIELD_NAME")), "", RS.Fields("FIELD_NAME"))
      RS.MoveNext
    Next i  ' Цикл перебора наборов месторождений-пластов-способов
  End If
  RS.Close
  RowCounter_temp = RowCounter_temp + 6
  'Подпись в конце документа
  Range(.Cells(RowCounter_temp, 1), .Cells(RowCounter_temp, 42)).Select
  With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  Selection.Merge
  .Cells(RowCounter_temp, 1) = "Начальник отдела разработки      ___________________  /                   /"
  .Cells(RowCounter_temp, 1).Font.Size = 12

cn.Close
Set RS = Nothing
Set RS_Layer = Nothing
Set cn = Nothing


' Отбросим записи по способам эксплуатации с отсутствием данных
While RowCounter_max > 13
If InStr(1, .Cells(RowCounter_max, 1), "по способу", 1) > 0 Then
 If .Cells(RowCounter_max + 1, 1) = "0" Then
 .Rows(RowCounter_max + 1).Delete
 .Rows(RowCounter_max).Delete
 End If
End If
RowCounter_max = RowCounter_max - 1
Wend


If Len(ParField_Item_Name) > 0 Then
    OldFileName = Application.ThisWorkbook.Path & "\" & Application.ThisWorkbook.Name
    Application.ThisWorkbook.SaveAs Application.ThisWorkbook.Path & _
                          "\" & Choose(Month(StartDay), "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12") & "_" & ParField_Item_Name & "_" & Application.ThisWorkbook.Name
    Kill OldFileName
End If

Application.ThisWorkbook.Application.Visible = True
Application.ThisWorkbook.Application.ScreenUpdating = True
Application.Cursor = xlDefault

' Отобразим время окончания формирования отчета
'ThisMoment = Now
'.Cells(2, 7) = ThisMoment

End With

Exit Sub

'LabelErr:
'    Application.Cursor = xlDefault
'    Application.ThisWorkbook.Application.Visible = True
'    MsgBox "Ошибка выполнения отчета. Процедура MainExec. " + vbCrLf + vbCrLf + _
'            Err.Description, vbCritical + vbOKOnly, "Ошибка: " + CStr(Err.Number)
            
End Sub

'conn - соединение
'Sql_str - строка SQL-запроса
'CurrRow - номер строки для отображения в Excel-е
Function Draw_by_Query(conn As ADODB.Connection, Sql_str As String, CurrRow As Integer)
On Error GoTo LabelErr

Dim RS_Layer         As ADODB.Recordset  ' Recordset object
Dim j                As Integer          ' Счетчик для движения по строкам рекордсета
Dim RowIsEmpty As Boolean ' Является ли заполняемая строка пустой?

With Sheets("Sheet1")

'.Cells(nnn, 5) = "Тестик!!!"
    Set RS_Layer = Nothing
    Set RS_Layer = New ADODB.Recordset
    RS_Layer.ActiveConnection = conn ' Assign the Connection object.
    RS_Layer.CursorType = adOpenStatic
'    .Cells(CurrRow, 59) = Sql_str
    
    RS_Layer.Open Sql_str ' Extract the required records.
    
    'Sheets("Empty_Template").Range("D23").Value = Sql_str
     
        ' Цикл перебора записей со скважинами по текущему набору месторождений-пластов
        If Not (RS_Layer.BOF And RS_Layer.EOF And RS_Layer.RecordCount = 0) Then
            For j = 0 To RS_Layer.RecordCount - 1 ' бежим по рекордсету
        
            'Выставляем порядковые номера скважин только у первых записей
        '    If Not (RS_Layer.Fields("WELL_NAME") = CurrWELL_name) Then
        '    well_num = well_num + 1
        '    .Cells(RowCounter, 1) = well_num
        '    CurrWELL_name = RS_Layer.Fields("WELL_NAME")
        '    End If
        
            .Cells(CurrRow, 1) = RS_Layer.Fields("WORK_WELL_COUNT") 'Количество действ. скважин
            .Cells(CurrRow, 6) = RS_Layer.Fields("ALL_WELL_COUNT")  'Количество экспл. скважин
            With Range("A" & CStr(CurrRow) & ":E" & CStr(CurrRow))
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlBottom
              .WrapText = False
              .Orientation = 0
              .AddIndent = False
              .IndentLevel = 0
              .ShrinkToFit = False
              .ReadingOrder = xlContext
              .MergeCells = False
              .Merge
            End With
            With Range("F" & CStr(CurrRow) & ":K" & CStr(CurrRow))
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlBottom
              .WrapText = False
              .Orientation = 0
              .AddIndent = False
              .IndentLevel = 0
              .ShrinkToFit = False
              .ReadingOrder = xlContext
              .MergeCells = False
              .Merge
            End With

            'Добыча нефти (в т.ч. конденсат), т
            .Cells(CurrRow, 12) = RS_Layer.Fields("OIL_MASS")
            .Cells(CurrRow, 13) = RS_Layer.Fields("OIL_MASS_CUM_YEAR")
            .Cells(CurrRow, 14) = RS_Layer.Fields("OIL_MASS_CUM_FULL")
            
            'Добыча воды, т
            .Cells(CurrRow, 15) = RS_Layer.Fields("WATER_MASS")
            .Cells(CurrRow, 16) = RS_Layer.Fields("WATER_MASS_CUM_YEAR")
            .Cells(CurrRow, 17) = RS_Layer.Fields("WATER_MASS_CUM_FULL")
            
            'Добыча воды, м3
            .Cells(CurrRow, 18) = RS_Layer.Fields("WATER_VOL")
            .Cells(CurrRow, 19) = RS_Layer.Fields("WATER_VOL_CUM_YEAR")
            .Cells(CurrRow, 20) = RS_Layer.Fields("WATER_VOL_CUM_FULL")
            
            'Добыча жидкости, т
            .Cells(CurrRow, 21) = RS_Layer.Fields("LIQ_MASS")
            .Cells(CurrRow, 22) = RS_Layer.Fields("LIQ_MASS_CUM_YEAR")
            .Cells(CurrRow, 23) = RS_Layer.Fields("LIQ_MASS_CUM_FULL")
            
            'Добыча жидкости в поверхн-ых условиях, м3
            .Cells(CurrRow, 24) = RS_Layer.Fields("LIQ_VOL")
            .Cells(CurrRow, 25) = RS_Layer.Fields("LIQ_VOL_CUM_YEAR")
            .Cells(CurrRow, 26) = RS_Layer.Fields("LIQ_VOL_CUM_FULL")
            
            'Добыча жидкости в пластовых условиях, м3
            .Cells(CurrRow, 27) = RS_Layer.Fields("LIQ_VOL_PL")
            .Cells(CurrRow, 28) = RS_Layer.Fields("LIQ_VOL_CUM_YEAR_PL")
            .Cells(CurrRow, 29) = RS_Layer.Fields("LIQ_VOL_CUM_FULL_PL")

            
            ' new "Добыча газа", "Добыча растворенного газа", "Добыча газа газовой шапки"
            ' added 26-nov-2014
            'Добыча газа, тыс.м3
            .Cells(CurrRow, 30) = RS_Layer.Fields("GAS_VOL")
            .Cells(CurrRow, 31) = RS_Layer.Fields("GAS_VOL_CUM_YEAR")
            .Cells(CurrRow, 32) = RS_Layer.Fields("GAS_VOL_CUM_FULL")
            
            'Газовый фактор За месяц
            If RS_Layer.Fields("OIL_MASS") > 0 Then
             '.Cells(CurrRow, 33).NumberFormat = "0.00"
             .Cells(CurrRow, 33) = RS_Layer.Fields("GAS_VOL") * 1000 / RS_Layer.Fields("OIL_MASS")
            Else
             .Cells(CurrRow, 33) = RS_Layer.Fields("GAS_FACTOR")
            End If
            
            '.Cells(CurrRow, 34) = RS_Layer.Fields("GAS_FACTOR_YEAR")      'Газовый фактор С начала года
            .Cells(CurrRow, 34) = RS_Layer.Fields("WC_MASS")              '% воды весовой за месяц
            .Cells(CurrRow, 35) = RS_Layer.Fields("WC_VOL")               '% воды объемный за месяц
            '.Cells(CurrRow, 36) = RS_Layer.Fields("WATER_DENSITY") / 1000 'Удельный вес воды
            If RS_Layer.Fields("WATER_VOL") = 0 Then
              .Cells(CurrRow, 36) = 0
            Else
              .Cells(CurrRow, 36) = RS_Layer.Fields("WATER_MASS") / RS_Layer.Fields("WATER_VOL") * 1000
            End If
            '.Cells(CurrRow, 37) = RS_Layer.Fields("WC_MASS_YEAR")         '% воды весовой с н.г.
            
            .Cells(CurrRow, 39) = RS_Layer.Fields("OIL_RATE_MASS")      'Уплотненный дебит нефти т/сут
            .Cells(CurrRow, 40) = RS_Layer.Fields("OIL_RATE_MASS") / 0.84     'Уплотненный дебит нефти м3/сут
            '.Cells(CurrRow, 40) = RS_Layer.Fields("OIL_RATE_MASS_YEAR") 'Уплотненный дебит нефти за год, т/сут
            .Cells(CurrRow, 41) = RS_Layer.Fields("LIQ_RATE_MASS")      'Уплотненный дебит жидкости т/сут
            .Cells(CurrRow, 42) = RS_Layer.Fields("LIQ_RATE_VOL")       'Уплотненный дебит жидкости м3/сут
            '.Cells(CurrRow, 43) = RS_Layer.Fields("LIQ_RATE_MASS_YEAR") 'Уплотненный дебит жидк. за год, т/сут
            '.Cells(CurrRow, 44) = RS_Layer.Fields("GAS_RATE_VOL")       'Уплотненный дебит газа за месяц, тнм/сут
            '.Cells(CurrRow, 45) = RS_Layer.Fields("GAS_RATE_VOL_YEAR")  'Уплотненный дебит газа за год, тнм/сут
            
            '.Cells(CurrRow, 45) = RS_Layer.Fields("OIL_RATE_MASS_AVERAGE")   'Среднесуточная добыча нефти за месяц, т
            '.Cells(CurrRow, 46) = RS_Layer.Fields("LIQ_RATE_VOL_AVERAGE")    'Среднесуточная добыча жидкости за месяц, м3
            .Cells(CurrRow, 43) = RS_Layer.Fields("OIL_RATE_MASS_AVERAGE_2") 'Среднесуточный дебит нефти за месяц, т/сут
            .Cells(CurrRow, 44) = RS_Layer.Fields("LIQ_RATE_MASS_AVERAGE")   'Среднесуточный дебит жидк. за месяц, т/сут
            If RS_Layer.Fields("PROD_DAYS") = 0 Then
              .Cells(CurrRow, 45) = 0
            Else
              .Cells(CurrRow, 45) = RS_Layer.Fields("LIQ_VOL") / RS_Layer.Fields("PROD_DAYS") 'Среднесуточный дебит жидк. за месяц, м3/сут
            End If
            
            .Cells(CurrRow, 46) = RS_Layer.Fields("PROD_DAYS_CUM_YEAR")
            
            '.Cells(CurrRow, 49) = RS_Layer.Fields("EKSPL_DAIS_CUM_YEAR")  'Число суток экспл. с начала года
            .Cells(CurrRow, 47) = RS_Layer.Fields("ACC_DAYS_YEAR")       'Число суток накопл. с начала года
            .Cells(CurrRow, 48) = RS_Layer.Fields("PROST_DAYS_CUM_YEAR") 'Число суток прост. с начала года
            .Cells(CurrRow, 49) = RS_Layer.Fields("EKSPL_DAIS_CUM_FULL")  'Сутки эксплуатации с начала разраб., сут
            
            'Уплотненные скв.часы за месяц
            .Cells(CurrRow, 50) = RS_Layer.Fields("PROD_HOURS")    'Работы
            .Cells(CurrRow, 51) = RS_Layer.Fields("ACC_HOURS_MTH") 'Накопления
            .Cells(CurrRow, 52) = RS_Layer.Fields("PROST_HOURS")   'Простоя

            '.Cells(CurrRow, 53) = RS_Layer.Fields("PROD_HOURS") + RS_Layer.Fields("ACC_HOURS_MTH") + RS_Layer.Fields("PROST_HOURS") 'RS_Layer.Fields("PROD_HOURS_CALENDAR") 'Календарное время действ. фонда за месяц
            
            'Уплотненные скв.Сутки за месяц
            .Cells(CurrRow, 53) = RS_Layer.Fields("PROD_DAYS")    'Работы
            .Cells(CurrRow, 54) = RS_Layer.Fields("ACC_DAYS_MTH") 'Накопления
            .Cells(CurrRow, 55) = RS_Layer.Fields("PROST_DAYS")   'Простоя
            
            .Cells(CurrRow, 56) = RS_Layer.Fields("K_EKSPL")      'К эксплуатации За месяц
            .Cells(CurrRow, 57) = RS_Layer.Fields("K_EKSPL_YEAR") 'К эксплуатации С начала года
            .Cells(CurrRow, 58) = RS_Layer.Fields("K_ISPOLZ")     'К использования
            
            Range("L" & CStr(CurrRow) & ":AF" & CStr(CurrRow)).NumberFormat = "0.000"
            Range("AG" & CStr(CurrRow) & ":BF" & CStr(CurrRow)).NumberFormat = "0.0"
            RS_Layer.MoveNext
            Next j
        End If
    
        RS_Layer.Close
        
        RowIsEmpty = True
        For j = 1 To 54
          If .Cells(CurrRow, j) <> "" And .Cells(CurrRow, j) <> "0" And RowIsEmpty = True Then
            RowIsEmpty = False
          End If
        Next j
        Draw_by_Query = RowIsEmpty
End With

Exit Function

LabelErr:
    Application.Cursor = xlDefault
    Application.ThisWorkbook.Application.Visible = True
    MsgBox "Ошибка выполнения отчета. Процедура MainExec. " + vbCrLf + vbCrLf + _
            Err.Description, vbCritical + vbOKOnly, "Ошибка: " + CStr(Err.Number)
End Function
