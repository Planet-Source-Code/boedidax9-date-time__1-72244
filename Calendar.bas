Attribute VB_Name = "Calendar"
Option Explicit

Function DayBy_Year(DateNows As String, ToYear As Long) As Integer
Dim s_nDatas_D() As String

GetString DateNows & "/", "/", s_nDatas_D()

DayBy_Month = DayCount(DateNows, DateBy_AddYearOnYear(DateNows, ToYear, 3))  '+ 1
Hapus
'30/10/1009 1500
End Function

Function DayBy_Month(DateNows As String, ToMonth As Long, Optional OutDate As String) As Long
'MsgBox DateNows & " " & DateBy_AddMonthOnMonth(DateNows, ToMonth - 0, 3)
'DateBy_AddMonthOnMonth
Dim s_nMonth As Integer, s_Day As Long, s_YearNow As Integer

OutDate = DateBy_AddMonthOnMonth(DateNows, ToMonth - 0, s_Day, s_Year)
DayBy_Month = DayCount(DateNows, OutDate) + 0
'DayBy_Month = DayCount(DateNows, DateBy_AddMonthOnMonth(DateNows, ToMonth, 3)) + 0

End Function

Function AddDayOnDay()

End Function

Function AddDayOnMonth(DateNows As String, ToDay As Long, Month As Integer, Optional ByRef TestError As Integer) As Integer
Dim s_iX As Integer, s_MonthToDay As Integer, s_TMPMonthToDay As Integer
Dim s_nDatas_D() As String, s_YearNow As Integer, s_DayCek As Integer
Dim s_nMonth As Integer, s_Day As Integer, s_Month As Integer, s_Year As Integer
Dim s_ToDay As Long, s_iMonth As Integer
'DateNows = "2/1/2009"
'ToDay = 364
GetString DateNows & "/", "/", s_nDatas_D()
s_Month = s_nDatas_D(1)
s_Year = s_nDatas_D(2)

's_ToDay = AddDayOnYear(DateNows, ToDay, s_YearNow) ' + 1
s_ToDay = ToDay ' - 365
'MsgBox ToDay - s_ToDay
ToDay = ToDay + Val(s_nDatas_D(0))
If ToDay >= MonthToDay(12, Val(s_nDatas_D(2))) Then
    'If TestError > 0 Then ToDay = ToDay - 1
    'TestError = (s_ToDay \ 365) * 12
    ToDay = AddDayOnYear(DateNows, ToDay - Val(s_nDatas_D(0)), s_YearNow) - 0
    's_ToDay = 1
    Month = 12 - Val(s_nDatas_D(1)) '<---------- HAPUS
    s_DayCek = nJumlahHari(Val(s_nDatas_D(1)), Val(s_nDatas_D(2))) '<---------- HAPUS
    If s_DayCek = (s_DayCek - Val(s_nDatas_D(0))) + 1 Then Month = Month + 1 '<---------- HAPUS
    Month = (s_YearNow - Val(s_nDatas_D(2))) * 12
    s_nDatas_D(0) = "1"
    s_nDatas_D(1) = "1"
    s_nDatas_D(2) = s_YearNow
End If

s_iX = Val(s_nDatas_D(1)) - 1
s_nMonth = s_nMonth + s_iX
Do
    s_iX = s_iX + 1
    s_nMonth = s_nMonth + 1
    If s_iX > 12 Then
        s_iX = 1
        s_nDatas_D(2) = Val(s_nDatas_D(2)) + 1
    End If
    s_TMPMonthToDay = s_MonthToDay
    s_MonthToDay = s_MonthToDay + nJumlahHari(s_iX, Val(s_nDatas_D(2)))
        
    If s_MonthToDay >= ToDay Then
        Dim sPusing As Integer
        Month = Month + s_nMonth '- 1 's_DayCek
        AddDayOnMonth = ToDay - s_TMPMonthToDay
        
        
        'Form1.Command2.Caption = Month - s_Month & " " & s_TMPMonthToDay 'Month - 12 'Month - s_Month 's_TMPMonthToDay & " " & s_MonthToDay & " " & s_iX
        'TestError = Month - s_Month
        
        
        'Form1.Command2.Caption = (Month - s_Month) Mod 12
        If (Month - s_Month) Mod 12 <> 0 Then TestError = YearToYearInDay(s_Year, s_Year + ((Month - s_Month) \ 12)) _
        Else TestError = 0
        'TestError = YearToYearInDay(s_Year, Val(s_nDatas_D(2)) - 1)
        'TestError = YearToYearInDay(s_Year, s_Year + ((Month) \ 12))
        TestError = s_ToDay - TestError '+ 5
        'If TestError < 0 Then TestError = s_ToDay
        If TestError >= 0 Then
            'If s_Year <> Val(s_nDatas_D(2)) Then s_Year = Val(s_nDatas_D(2)) - 1
            's_Year = 2008
            s_iX = s_Month
            Do
                sPusing = sPusing + nJumlahHari(s_iX, s_Year)
                If sPusing > TestError Then Exit Do
                'If TestError < MonthToDay(s_iX, Val(s_nDatas_D(2))) Then Exit Do
                s_iX = s_iX + 1
                If s_iX > 12 Then s_iX = 1: s_Year = s_Year + 1
                s_iMonth = s_iMonth + 1
'                Stop
            Loop
            'If TestError = MonthToDay(s_iX - 1, Val(s_nDatas_D(2))) Then Stop 's_iMonth = s_iMonth + 1
            'MsgBox MonthToDay(s_iX, Val(s_nDatas_D(2)))
            'Form1.Caption = TestError & " " & ((((Month - s_Month) \ 12)) * 12) & " " & s_iMonth
            If (Month - s_Month) Mod 12 <> 0 Then TestError = ((((Month - s_Month) \ 12)) * 12) + s_iMonth _
            Else TestError = s_iMonth  '(((Month - s_Month) \ 12) * 12) + (Month - s_Month) \ 12 '- Val(ghg) ' & " " & ghg
        Else
            TestError = Month - s_Month ' - 1
        End If
        Exit Do
    Else
    End If
Loop

'Form1.Command2.Caption = s_iMonth & " " & Month
End Function

Function AddDayOnYear(DateNows As String, ToDay As Long, YearNow As Integer) As Long
Dim s_nDatas_D() As String, s_DayCount As Long, s_DayCountTo As Long
Dim s_YearToYearInDay As Long

GetString DateNows & "/", "/", s_nDatas_D()

''If ToDay < MonthToDay(12, Val(s_nDatas_D(2))) Then Stop

s_DayCount = DayCount(DateNows, "31/12/" & s_nDatas_D(2)) + 1
s_DayCount = ToDay - s_DayCount

YearNow = s_DayCount \ 365
YearNow = YearNow + Val(s_nDatas_D(2))

Do
    s_YearToYearInDay = YearToYearInDay(Val(s_nDatas_D(2)), YearNow)
    AddDayOnYear = s_DayCount - s_YearToYearInDay
    If AddDayOnYear < 0 Then YearNow = YearNow - 1 Else Exit Do
Loop
AddDayOnYear = AddDayOnYear + 1
YearNow = YearNow + 1
End Function

Function MonthBy_Day(DateNows As String, ToDay As Long)
Dim s_nDatas_D() As String, s_nDatas_E() As String
Dim s_Day As Long, s_Month As Integer, s_Year As Integer
Dim s_iX As Integer, s_Int As Integer, s_Log As Long, s_y As Integer

GetString DateNows & "/", "/", s_nDatas_D()
GetString DateBy_AddDayOnDay & "/", "/", s_nDatas_E()

s_Day = s_Log
If s_Day >= MonthToDay(12, Val(s_nDatas_D(2))) Then
    s_Log = DayCount(DateNows, "1/1/" & Val(s_nDatas_D(2)) + 1)
    s_Day = s_Day - s_Log
    MonthOut = MonthCount(DateNows, "1/1/" & Val(s_nDatas_D(2)) + 1)
    If s_Day >= MonthToDay(12, Val(s_nDatas_D(2)) + 1) Then
        s_Log = DayCount("1/1/" & Val(s_nDatas_D(2)) + 1, "1/1/" & Val(s_nDatas_E(2)))
        s_Day = s_Day - s_Log
        MonthOut = MonthOut + MonthCount("1/1/" & Val(s_nDatas_D(2)) + 1, "1/1/" & Val(s_nDatas_E(2)))
        s_nDatas_D(1) = "1"
        s_nDatas_D(2) = s_nDatas_E(2)
    Else
        s_nDatas_D(1) = "1"
        s_nDatas_D(2) = Val(s_nDatas_D(2)) + 1
    End If
End If

s_Month = Val(s_nDatas_D(1))
s_Year = Val(s_nDatas_D(2))
Do
    s_Log = nJumlahHari(s_Month, s_Year)
    If s_Day >= s_Log Then
        MonthOut = MonthOut + 1
        s_Day = s_Day - s_Log
    Else
        Exit Do
    End If
    s_Month = s_Month + 1
    If s_Month > 12 Then
        s_Month = 1
        s_Year = s_Year + 1
    End If
Loop
End Function

Function aXXXXAddDayOnMonth(DateNows As String, ToDay As Long, Month As Integer, Optional TestError As Integer) As Integer
Dim s_iX As Integer, s_MonthToDay As Integer, s_TMPMonthToDay As Integer
Dim s_nDatas_D() As String, s_YearNow As Integer, s_DayCek As Integer
Dim s_nMonth As Integer


DateNows = "2/1/2009"
ToDay = 363

GetString DateNows & "/", "/", s_nDatas_D()

ToDay = ToDay + Val(s_nDatas_D(0))
If ToDay >= MonthToDay(12, Val(s_nDatas_D(2))) Then
    'If TestError > 0 Then ToDay = ToDay - 1
    ToDay = AddDayOnYear(DateNows, ToDay - TestError, s_YearNow) - 0
    Month = 12 - Val(s_nDatas_D(1))
    s_DayCek = nJumlahHari(Val(s_nDatas_D(1)), Val(s_nDatas_D(2)))
    If s_DayCek = (s_DayCek - Val(s_nDatas_D(0))) + 1 Then Month = Month + 1
    Month = (s_YearNow - Val(s_nDatas_D(2))) * 12
    s_nDatas_D(0) = "1"
    s_nDatas_D(1) = "1"
    s_nDatas_D(2) = s_YearNow
End If

s_iX = Val(s_nDatas_D(1)) - 1
s_nMonth = s_nMonth + s_iX
Do
    s_iX = s_iX + 1
    s_nMonth = s_nMonth + 1
    If s_iX > 12 Then s_iX = 1
    s_TMPMonthToDay = s_MonthToDay
    s_MonthToDay = s_MonthToDay + nJumlahHari(s_iX, Val(s_nDatas_D(2)))
    
    If s_MonthToDay >= ToDay Then
        Month = Month + s_nMonth '- 1 's_DayCek
        AddDayOnMonth = ToDay - s_TMPMonthToDay
'        DateNext = Format(ToDay - s_TMPMonthToDay, "0#") & "/" & Format(s_iX, "0#")
        Exit Do
    End If
Loop
End Function

Function DateTimeOnDay(DateTimeNows As String, OnDay As Long) As String
Dim s_nDatas_DTNow() As String, s_nDatas_D() As String
Dim s_YearNow As Integer, s_MonthNow As Integer, s_DayNow As Long

s_DayNow = OnDay

GetString DateTimeNows & " ", " ", s_nDatas_DTNow()
GetString s_nDatas_DTNow(0) & "/", "/", s_nDatas_D()

s_DayNow = AddDayOnYear(s_nDatas_DTNow(0), s_DayNow, s_YearNow)
s_DayNow = AddDayOnMonth("0/1/" & s_YearNow, s_DayNow, s_MonthNow)

's_DayNow = s_DayNow - 1
DateTimeOnDay = Format(s_DayNow, "0#") & "/" & Format(s_MonthNow, "0#") & "/" & s_YearNow
End Function

Function XXXMonthModDay(DateNows As String, ToDay As Long, Month As Integer) As Integer
Dim s_iX As Integer, s_MonthToDay As Integer, s_TMPMonthToDay As Integer
Dim s_nDatas_D() As String, s_YearNow As Integer, s_DayCek As Integer
Dim s_nMonth As Integer

GetString DateNows & "/", "/", s_nDatas_D()

If ToDay > MonthToDay(12, Val(s_nDatas_D(2))) Then
    ToDay = YearModDay(DateNows, ToDay, s_YearNow) - 1
    Month = 12 - Val(s_nDatas_D(1))
    s_DayCek = nJumlahHari(Val(s_nDatas_D(1)), Val(s_nDatas_D(2)))
    If s_DayCek = (s_DayCek - Val(s_nDatas_D(0))) + 1 Then Month = Month + 1
    Month = (s_YearNow - Val(s_nDatas_D(2))) * 12
    s_nDatas_D(0) = "1"
    s_nDatas_D(1) = "1"
    s_nDatas_D(2) = s_YearNow
Else
    Month = Val(s_nDatas_D(1)) - 1
    s_DayCek = nJumlahHari(Val(s_nDatas_D(1)), Val(s_nDatas_D(2)))
    If s_DayCek <> (s_DayCek - Val(s_nDatas_D(0))) + 1 Then
        s_DayCek = ToDay - (s_DayCek - Val(s_nDatas_D(0)))
        If s_DayCek > 0 Then
            If ToDay > nJumlahHari(Val(s_nDatas_D(1)) + 1, Val(s_nDatas_D(2))) Then
            ToDay = s_DayCek
                s_iX = Val(s_nDatas_D(1)) '+ 1
                s_DayCek = 0
                s_nMonth = 0
            End If
        Else
            s_DayCek = 0
            s_nMonth = s_iX
        End If
        'Stop
    Else
        s_DayCek = 0
        s_nMonth = s_iX
    End If
End If


's_iX = Val(s_nDatas_D(1)) - 1
Do
    s_iX = s_iX + 1
    s_nMonth = s_nMonth + 1
    If s_iX > 12 Then s_iX = 1
    s_TMPMonthToDay = s_MonthToDay
    s_MonthToDay = s_MonthToDay + nJumlahHari(s_iX, Val(s_nDatas_D(2)))
    
    If s_MonthToDay >= ToDay Then
        Month = Month + s_nMonth - 1 's_DayCek
        MonthModDay = ToDay - s_TMPMonthToDay
'        DateNext = Format(ToDay - s_TMPMonthToDay, "0#") & "/" & Format(s_iX, "0#")
        Exit Do
    End If
Loop
End Function

Function XXXYearModDay(DateNows As String, ToDay As Long, YearNow As Integer) As Long
Dim s_nDatas_D() As String, s_DayCount As Long, s_DayCountTo As Long
Dim s_YearToYearInDay As Long

GetString DateNows & "/", "/", s_nDatas_D()

''If ToDay < MonthToDay(12, Val(s_nDatas_D(2))) Then Stop

s_DayCount = DayCount(DateNows, "31/12/" & s_nDatas_D(2)) + 1
s_DayCount = ToDay - s_DayCount

YearNow = s_DayCount \ 365
YearNow = YearNow + Val(s_nDatas_D(2))

Do
    s_YearToYearInDay = YearToYearInDay(Val(s_nDatas_D(2)), YearNow)
    YearModDay = s_DayCount - s_YearToYearInDay
    If YearModDay < 0 Then YearNow = YearNow - 1 Else Exit Do
Loop
YearModDay = YearModDay + 1
YearNow = YearNow + 1
End Function

Function DateTimeSelisih(DateTimeBefores As String, DateTimeAfters As String) As String
Dim s_nDatas_Bef() As String, s_nDatas_Aft() As String
Dim s_Day As Long, s_Time As Long

GetString DateTimeBefores & " ", " ", s_nDatas_Bef()
GetString DateTimeAfters & " ", " ", s_nDatas_Aft()

s_Day = DayCount(s_nDatas_Bef(0), s_nDatas_Aft(0))
s_Time = FomatSecond(s_nDatas_Aft(1)) - FomatSecond(s_nDatas_Bef(1))

DateTimeSelisih = s_Day & " " & s_Time
End Function

Function DateModDay(DateNows As String, ToDay As Long, Optional YearNow As Integer) As Long
Dim s_nDatas_D() As String, s_DayCount As Long, s_DayCountTo As Long
Dim a As Long, A2 As Integer, B As Long, D1 As Long, D2 As Long, F As Long
Dim XXX As Integer
'DateNows = "31/12/2008"
GetString DateNows & "/", "/", s_nDatas_D()

If Val(s_nDatas_D(2)) Mod 4 = 0 Then A2 = 366 Else A2 = 365
GetString DateNows & "/", "/", s_nDatas_D()
a = DayCount(DateNows, "31/12/" & s_nDatas_D(2)) + 1
a = ToDay - a

's = A - (A Mod 366)

's = A - (365 - (A Mod 366))
F = a \ 365
'k = A Mod 365
F = F + Val(s_nDatas_D(2))
Do
    D1 = Val(s_nDatas_D(2)) + 0
    'F = 3019
    D2 = F '- r1
    'D2 = 15
    B = YearToYearInDay(D1, D2)
    'B = B + ((D2 - D1) \ 4)
    c = a - B
    If c < 0 Then
    '    MsgBox "Eroor"
        F = F - 1
    Else
        Exit Do
    End If
    XXX = XXX + 1
Loop
DateModDay = c + 1
'MsgBox XXX
YearNow = D2 + 1
'If YearNow Mod 4 <> 0 Then DateModDay = DateModDay + 1
'MsgBox YearToYearInDay(Val(s_nDatas_D(2)) + 1, (Val(s_nDatas_D(2)) + 1) + A)
''A = 2008 - 2011 '+ 1
''A = A \ 4
''A = YearToYearInDay(2010, 2012)
''Stop
''End
End Function


Function Asli_DateTimeToDateTime(DateTimeNows As String, ToTimeDate As String) As String
'("1/10/2009", "1/10/2019")

'Exit Function
Dim iX As Integer, bbb As Integer, Ccc As Integer

Dim XXX As String

DateTimeNows = "29/2/2008 0:0:0" '"01/10/2009 0:0:0"
XXX = "29/2/2012" '"29/12/2016"

'DateTimeNows = _
"25/7/2001 0:0:0"
'XXX = "29/7/2001"

Dim s_nDatas_DTNow() As String, s_nDatas_D() As String, s_nDatas_T() As String
Dim s_CountDays As Integer, s_OutPutProses As Long, s_YearNow As Integer

GetString DateTimeNows & " ", " ", s_nDatas_DTNow()
GetString s_nDatas_DTNow(0) & "/", "/", s_nDatas_D()
GetString s_nDatas_DTNow(1) & ".", ".", s_nDatas_T()

MsgBox DateTo(s_nDatas_DTNow(0), DayCount(s_nDatas_DTNow(0), XXX)) & " test"
'MsgBox (DayCount("29/2/2008", "1/3/2008"))
'Exit Function

s_OutPutProses = Val(s_nDatas_D(2))
s_CountDays = DateModDay(s_nDatas_DTNow(0), DayCount(s_nDatas_DTNow(0), XXX), s_YearNow)
's_CountDays = s_CountDays - 1
s_OutPutProses = MonthModDay(s_CountDays, s_nDatas_D(0) & "/" & s_nDatas_D(1) & "/" & s_nDatas_D(2), DateTimeToDateTime)

DateTimeToDateTime = DateTimeToDateTime & "/" & s_YearNow 's_nDatas_D(0) & "/" & s_nDatas_D(1) & "/" & s_nDatas_D(2)
End Function

Function DateTimeToDay(DateTimeNows As String, Days As String)
Dim s_nDatas_DTNow() As String, s_nDatas() As String

GetString DateTimeBefores & " ", " ", s_nDatas_Bef()

End Function

Function DateTimeNext(DateTimeNows As String, TimeDateAfters As String)
Dim s_nDatas_DTNow() As String, s_nDatas() As String

GetString DateTimeBefores & " ", " ", s_nDatas_Bef()

End Function

Function YearCount(DateBefores As String, DateAfters As String) As Long
    
End Function

Function MonthCount(DateBefores As String, DateAfters As String) As Long
Dim s_iX As Integer, s_MonthToDay As Integer, s_TMPMonthToDay As Integer
Dim s_nDatas_D1() As String, s_nDatas_D2() As String, s_YearNow As Integer, s_DayCek As Integer
'Dim s_nMonth As Integer,

GetString DateBefores & "/", "/", s_nDatas_D1()
GetString DateAfters & "/", "/", s_nDatas_D2()

MonthCount = DayCount(DateBefores, DateAfters)  '& " " & nJumlahHari(2, 2009)

Form1.Caption = MonthCount

s_Year = Val(s_nDatas_D2(2)) - Val(s_nDatas_D1(2))
s_nMonth = 12 * s_Year + Val(s_nDatas_D2(1))
MonthCount = s_nMonth - Val(s_nDatas_D1(1))
'If MonthCount < 12 Then MonthCount = MonthCount + 1
End Function

Function YearCountBy_Day(DateBefores As String, ToDay As Long) As Long
Dim s_nDatas_D() As String

GetString DateNows & "/", "/", s_nDatas_D()

End Function

Function MonthCountBy_Day(DateBefores As String, ToDay As Long) As Long
Dim s_nDatas_D() As String, s_YearNow As Integer
Dim s_DayCount As Long

'ToDay = 427
'DateBefores = "30/5/2008"




GetString DateBefores & "/", "/", s_nDatas_D()

s_DayCount = DayCount(DateBefores, "31/12/" & s_nDatas_D(2))
MonthCountBy_Day = AddDayOnYear(DateBefores, ToDay, s_YearNow)
MsgBox ToDay - MonthCountBy_Day
Stop
End Function

Function xxxxMonthCount(DateBefores As String, DateAfters As String) As Long
Dim s_iX As Integer, s_MonthToDay As Integer, s_TMPMonthToDay As Integer
Dim s_nDatas_D1() As String, s_nDatas_D2() As String, s_YearNow As Integer, s_DayCek As Integer
Dim s_nMonth As Integer

GetString DateBefores & "/", "/", s_nDatas_D1()
GetString DateAfters & "/", "/", s_nDatas_D2()

s_Year = Val(s_nDatas_D2(2)) - Val(s_nDatas_D1(2))
s_nMonth = 12 * s_Year + Val(s_nDatas_D2(1))
MonthCount = s_nMonth - Val(s_nDatas_D1(1))
'If MonthCount < 12 Then MonthCount = MonthCount + 1
End Function

Function DayCount(DateBefores As String, DateAfters As String) As Long
Dim nDatasBefore() As String, nDatasAfter() As String
Dim Years As Integer
Dim Day1 As Long, Day2 As Long, Day3 As Long
Dim Year1 As Integer, Year2 As Integer

If DateBefores = "" Or DateAfters = "" Then
DayCount = -100
Exit Function
End If

GetString DateBefores & "/", "/", nDatasBefore()
GetString DateAfters & "/", "/", nDatasAfter()

Years = Val(nDatasAfter(2)) - Val(nDatasBefore(2))
If Years = 0 Then
    Years = Val(nDatasBefore(2))
    
    Day1 = MonthToDay(Val(nDatasBefore(1)) - 1, Val(nDatasBefore(2))) + _
    Val(nDatasBefore(0))

    Day2 = MonthToDay(Val(nDatasAfter(1)) - 1, Val(nDatasAfter(2))) + _
    Val(nDatasAfter(0))
    
    DayCount = Day2 - Day1
Else
    Year1 = Val(nDatasBefore(2) + 1)
    Year2 = Val(nDatasAfter(2) - 1)
    
    If (Year2 - Year1) + 1 >= 1 Then Day1 = YearToYearInDay(Year1 - 1, Year2)
    
    Day2 = MonthToDay(Val(nDatasBefore(1)) - 1, Val(nDatasBefore(2))) + _
    Val(nDatasBefore(0))
    Day2 = MonthToDay(12, Val(nDatasBefore(2))) - Day2
    Day1 = Day1 + Day2
    
    Day2 = MonthToDay(Val(nDatasAfter(1)) - 1, Val(nDatasAfter(2))) + _
    Val(nDatasAfter(0))
    DayCount = Day1 + Day2 ' + 1
End If
End Function

