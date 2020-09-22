Attribute VB_Name = "Calender_TMP"
Function DateBy_AddDayOnDay(DateNows As String, ToDay As Long, Optional MonthOut As Integer, Optional YearOut As Integer, Optional NextMonth As Boolean) As String
Dim s_nDatas_D() As String, s_nDatas_E() As String
Dim s_Day As Long, s_Month As Integer, s_Year As Integer
Dim s_iX As Integer, s_Int As Integer, s_Log As Long, s_y As Integer

GetString DateNows & "/", "/", s_nDatas_D()

s_Log = ToDay
s_Day = AddDayOnMonth(DateNows, ToDay, s_Month, MonthOut)
'MonthOut = MonthOut - 1
's_Month = s_Month - Val(s_nDatas_D(1))
''MonthOut = s_Month - Val(s_nDatas_D(1))
s_Year = s_Year + Int((s_Month - 0) / 12) + Val(s_nDatas_D(2))
''YearOut = Int(MonthOut / 12)
'If MonthOut Mod 12 <> 0 Then s_Year = s_Year + 1  '. . .
s_Month = s_Month Mod 12
If s_Month = 0 Then s_Month = 12: s_Year = s_Year - 1
DateBy_AddDayOnDay = s_Day & "/" & s_Month & "/" & s_Year
'MonthOut = MonthCount(DateNows, DateBy_AddDayOnDay)
YearOut = Int(MonthOut / 12)
End Function

Function DateBy_AddDayOnMonth(DateNows As String, ToMonth As Long, Optional DayOut As Long, Optional MonthOut As Integer) As String
Dim s_nDatas_D() As String
Dim s_Day As Long, s_Month As Integer, s_Year As Integer
Dim s_iX As Integer, s_Int As Integer, s_Log As Long

GetString DateNows & "/", "/", s_nDatas_D()

a = DateBy_AddMonthOnMonth(datenows, ToMonth,

s_Day = ToMonth

DateBy_AddDayOnMonth = sstop
End Function

Function DateBy_AddDayOnYear()

End Function

Function DateBy_AddYearOnYear(DateNows As String, ToYear As Long, Optional DayOut As Long, Optional MonthOut As Integer) As String
Dim s_nDatas_D() As String
Dim s_Day As Long, s_Month As Integer, s_Year As Integer
Dim s_iX As Integer, s_Int As Integer ', s_Int As Long

GetString DateNows & "/", "/", s_nDatas_D()

's_Int = ToYear * 12
MonthOut = ToYear * 12
DateBy_AddYearOnYear = DateBy_AddMonthOnMonth(DateNows, MonthOut, DayOut)
End Function

Function DateBy_AddMonthOnMonth(DateNows As String, ToMonth As Integer, Optional DayOut As Long, Optional YearOut) As String
Dim s_nDatas_D() As String
Dim s_Day As Long, s_Month As Integer, s_Year As Integer
Dim s_iX As Integer, s_Int As Integer, s_Log As Long

GetString DateNows & "/", "/", s_nDatas_D()

s_Log = ToMonth + Val(s_nDatas_D(1))
s_Year = Int((s_Log - 1) / 12) + Val(s_nDatas_D(2))
s_Month = s_Log Mod 12
If s_Month = 0 Then s_Month = 12
s_Day = DayCount("1/" & Val(s_nDatas_D(1)) + s_Int & "/" & Val(s_nDatas_D(2)), 1 & "/" & s_Month & "/" & s_Year)
''s_Day = DayCount(nJumlahHari(Val(s_nDatas_D(1)), Val(s_nDatas_D(2))) & "/" & Val(s_nDatas_D(1)) + s_Int & "/" & Val(s_nDatas_D(2)), nJumlahHari(s_Month, s_Year) & "/" & s_Month & "/" & s_Year)
's_Day = s_Day - Val(s_nDatas_D(0))

DayOut = s_Day + 0
YearOut = Int(ToMonth / 12)
DateBy_AddMonthOnMonth = DateBy_AddDayOnDay(DateNows, s_Day, 0)
End Function

