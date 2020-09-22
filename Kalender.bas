Attribute VB_Name = "Kalender"
'Type PointAPI
'    X As Long
'    Y As Long
'End Type
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'cek kembali masalah vocernya oiiiiiiiiiii
'Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

'Public Const WM_SETHOTKEY = &H32
'Public Const WM_SHOWWINDOW = &H18

' contoh tombol
'Public Const HK_ALTZ = &H45A
'Public Const HK_SHIFTA = &H141 'Shift + A
'Public Const HK_SHIFTB = &H142 'Shift + B
'Public Const HK_CONTROLA = &H241 'Control + A
'Public Const HK_CONTROLB = &H242 'Control + B
'Public Const HK_CONTROLC = &H243 'Control + B
'Public Const HK_CONTROLD = &H244 'Control + B

'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Public TestTimeVirtual As Long

'Public Const DayToSecond As Long = 86400

'Private Type Bonuss
'    Names As String
'    Hits  As Integer
'    Outs  As String
'End Type
'Dim SetBonus() As Bonuss

'Type CountBill
'    iRuns As Integer
'    iTimeOffs As Integer
'    iLogOffs As Integer
'    iEmptys As Integer
'    iWait As Integer
    
'    iPendReal As Single
'End Type
'Public CountBillFor As CountBill'

'Type KeperluanAdmin
'    MyData As String
    
'    Nama As String
'    NamaLicensi As String
'    Password As String
'    code As String
    
'    NamaWarnet As String
'    AlamatWarnet As String
'    InfoWarnet As String
'End Type
'Public MasterAdmin As KeperluanAdmin

'Type KeperluanKasir
'    Tgl_Login As String
'
'    IDs  As String
'    Nama As String
'    NickName As String
'    Password As String
        
'    WaktuFirst As String
'    WaktuEnd   As String
'    Shifs    As String
'    LogOn   As String
'End Type
'Public Kasir As KeperluanKasir
'Public Kasir_Ganti As KeperluanKasir
'Public TmpKasir As KeperluanKasir
'Public KasirLogin As Boolean

'Type FileLogin
'    RandomID_Login As Long
'    RandonNick_Login As String
'
'    Tgl_Login As String
'
'    IDs_Login As String
'    Nama_Login As String
'    NickName_Login As String
'    Password_Login As String
'
'    WaktuFirst_Login As String
'    WaktuEnd_Login   As String
'    Shift_Login As String
'    LogOn_Login As String
'End Type
'Public Kasir_FileLogin As FileLogin'

'Public DateNow As String, TimeNow As String
'Public Billing As String
'Public LuckyAscii As String

'Type SkinClients
'    Namas As String
'    Directorys As String
'End Type

'Public nSkinClient() As SkinClients
'Public DataMarketWaktu As String
'Public CountMarketWaktu As Integer

'Public AutoWaktuSetClient As Boolean

'Public DataBahasa As String

'Public FrmNew() As New frmMarket
'Public CntFrmNew As Integer

'Public IndexInfoBil As Integer
'Public IndexInfoBil_Tab As Integer, IndexInfoBil_TabAll As Boolean

'Public iCurrency As String'

'Public TmpIcoIndex() As Integer
Public Const DayToSecond As Long = 86400

Public Function nJumlahHari(nBulan As Integer, nTahun As Integer) As Integer
Select Case nBulan
Case 1
    nJumlahHari = 31
Case 2
    If nTahun Mod 4 = 0 Then nJumlahHari = 29 Else nJumlahHari = 28
Case 3
    nJumlahHari = 31
Case 4
    nJumlahHari = 30
Case 5
    nJumlahHari = 31
Case 6
    nJumlahHari = 30
Case 7
    nJumlahHari = 31
Case 8
    nJumlahHari = 31
Case 9
    nJumlahHari = 30
Case 10
    nJumlahHari = 31
Case 11
    nJumlahHari = 30
Case 12
    nJumlahHari = 31
End Select
End Function

Function CekTglPenghabisan(DateIn As String, DateOut As String, DateNow As String) As Boolean
'    If DateIn <> 0 Then '====> Tambahkan
        CekTglPenghabisan = DateToLow(DateTo(DateIn, DateOut), DateNow)
'    End If
End Function

Function DateTo(DateIn As String, ByVal DateOut As String) As String
Dim nDatas() As String
Dim nData1() As String, CountDay As Integer
Dim nDatasT() As String
Dim PlusT As Long

    'DateOut = "0#51"
    
    GetString DateOut & "#", "#", nDatasT()
    DateOut = Val(nDatasT(0))
    
    If UBound(nDatasT()) > 0 Then _
       nDatasT(1) = FormatTime(DayToSecond / 100 * Val(nDatasT(1)), "0#")
    
    GetString DateIn & " ", " ", nDatas()
    GetString nDatas(0) & "/", "/", nData1()
    DateOut = DateOut + Val(nData1(0))
        
    If UBound(nData1()) <> 2 Then
        DateTo = "#Error"
        Exit Function
    Else
        For X = 0 To UBound(nData1())
            If Not IsNumeric(nData1(X)) = True Then
                DateTo = "#Error"
                Exit Function
            End If
        Next X
    End If
    
    If UBound(nDatasT()) > 0 Then
        nDatas(1) = FomatSecond(nDatas(1)) + FomatSecond(nDatasT(1))
        PlusT = (nDatas(1) - FomatSecond("23:59:59"))  ', "0#")
        nDatas(1) = FormatTime((nDatas(1) Mod FomatSecond("23:59:59")), "0#") 'FormatTime(Val(nDatas(1)), "0#")
        If PlusT > 0 Then DateOut = DateOut + 1
    End If
        
    Do
        CountDay = nJumlahHari(Val(nData1(1)), Val(nData1(2)))
        If DateOut > CountDay Then
            nData1(1) = Val(nData1(1)) + 1
            If Val(nData1(1)) > 12 Then
                nData1(1) = 1
                nData1(2) = Val(nData1(2)) + 1
            End If
        Else
            Exit Do
        End If
        DateOut = DateOut - CountDay
    Loop
    
    If UBound(nDatas()) = 0 Then
        DateTo = Format(DateOut, "0#") & "/" & Format(nData1(1), "0#") & "/" & nData1(2)
    ElseIf UBound(nDatas()) = 1 Then
        DateTo = Format(DateOut, "0#") & "/" & Format(nData1(1), "0#") & "/" & nData1(2) & " " & nDatas(1)
    Else
        MsgBox "Error"
    End If
    'MsgBox DateTo
    'DateTo = DateOut & "/" & Format(nData1(1), "0#") & "/" & nData1(2)
'    a0 = a0
End Function

Function DateToLow(DateCek As String, ByVal DateNow As String) As Boolean
Dim nData1() As String, nData2() As String, nData3(2) As Boolean
Dim nDatasT1() As String, nDatasT2() As String
Dim Count1 As Long, Count2 As Long
    
    If DateCek = "#Error" Then
        DateToLow = True
        Exit Function 'Cek Kembali
    End If
    
    
    GetString DateCek & " ", " ", nDatasT1()
    GetString nDatasT1(0) & "/", "/", nData1()
    'GetString DateCek & "/", "/", nData1()
'    nData1(0) =
    '    GetString nData1(0) & ".", ".", nData4()
    GetString DateNow & " ", " ", nDatasT2()
    GetString nDatasT2(0) & "/", "/", nData2()
              
    'If UBound(nData4()) = 1 Then
        
    'Else
        
    'End If
    
    'TimeNow = "11:05:05"
'    MsgBox TimeNow
'MsgBox FomatSecond(nDatasT1(1)) & " " & FomatSecond(nDatasT2(1))
    Count1 = ((Val(nData1(0)) * 1) + (Val(nData1(1)) * 31) + (Val(nData1(2)) * 372)) 'DateCek
    Count1 = Count1 '+ FomatSecond(nDatasT1(1))

'    Count1 = Count1 + FomatSecond(nDatasT1(1))
    Count2 = ((Val(nData2(0)) * 1) + (Val(nData2(1)) * 31) + (Val(nData2(2)) * 372)) 'DateNow
    Count2 = Count2 '+ FomatSecond(nDatasT2(1))
    If Count1 <= Count2 Then DateToLow = True
    If Count1 = Count2 Then
        If FomatSecond(nDatasT1(1)) > FomatSecond(nDatasT2(1)) Then DateToLow = False
    End If
    
End Function

Function FormatTime(TimeSecond As Long, Optional DigitFormatHour As String = "00#") As String
    FormatTime = Format(TimeSecond \ 3600, DigitFormatHour) & ":" & Format((TimeSecond Mod 3600) \ 60, "0#") & ":" & Format(TimeSecond Mod 60, "0#")
End Function

Function FomatSecond(FTime As String) As Long
Dim nDatas() As String

If FTime = "" Then
    FomatSecond = 0
Else
    On Error GoTo 10
    GetString FTime & ":", ":", nDatas()
    FomatSecond = ((Val(nDatas(0)) * 3600) + (Val(nDatas(1)) * 60) + (Val(nDatas(2)) * 1))
End If

Exit Function
10
    FomatSecond = 0
End Function

Function DateTimeToOver(DateNows As String, TimeNows As String, TimeSecondOvers As Long, Optional ShowDates As Boolean) As String
Dim TimeSecond As Long, TimeNowsSecond As Long, DateOvers As String, TimesOvers As String

TimeNowsSecond = FomatSecond(TimeNows)
TimeSecond = ((TimeNowsSecond + TimeSecondOvers))

TimesOvers = FormatTime(Val(TimeSecond Mod DayToSecond), "0#")
DateOvers = DateTo(DateNows, TimeSecond \ DayToSecond)
If ShowDates = False Then DateTimeToOver = DateOvers & " " & TimesOvers Else DateTimeToOver = DateOvers
End Function

Function MonthToDay(Months As Integer, Years As Integer) As Integer
Select Case Months
Case 0
    MonthToDay = 0
Case 1
    MonthToDay = 31
Case 2
    MonthToDay = 60
Case 3
    MonthToDay = 91
Case 4
    MonthToDay = 121
Case 5
    MonthToDay = 152
Case 6
    MonthToDay = 182
Case 7
    MonthToDay = 213
Case 8
    MonthToDay = 244
Case 9
    MonthToDay = 274
Case 10
    MonthToDay = 305
Case 11
    MonthToDay = 335
Case 12
    MonthToDay = 366
End Select

    If Years Mod 4 <> 0 And Months > 1 Then
        MonthToDay = MonthToDay - 1
    End If
End Function

Function xxxDayCount(DateBefores As String, DateAfters As String) As Long
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

'    Day2 = MonthToDay(Val(nDatasAfter(1)), Years) - MonthToDay(Val(nDatasBefore(1)), Years) - nJumlahHari(Val(nDatasAfter(1)), Years)
'    DayCount = Day2 + Day1 + Val(nDatasAfter(0))
   
Exit Function
    Years = Val(nDatasBefore(2))
    
    Day1 = nJumlahHari(Val(nDatasBefore(1)), Years) - Val(nDatasBefore(0))

    Day2 = MonthToDay(Val(nDatasAfter(1)), Years) - MonthToDay(Val(nDatasBefore(1)), Years) - nJumlahHari(Val(nDatasAfter(1)), Years)
    DayCount = Day2 + Day1 + Val(nDatasAfter(0))
Else
    'nDatasBefore(2) = 2008
    'nDatasAfter(2) = 2008
    
    Year1 = Val(nDatasBefore(2) + 1)
    Year2 = Val(nDatasAfter(2) - 1)
    
'    If Year1 = Year2 Then Day1 = MonthToDay(12, Year1) Else Day1 = YearToYearInDay(Year1, Year2)
    If (Year2 - Year1) + 1 >= 1 Then Day1 = YearToYearInDay(Year1 - 1, Year2)

    
    Day2 = MonthToDay(Val(nDatasBefore(1)) - 1, Val(nDatasBefore(2))) + _
    Val(nDatasBefore(0))
    Day2 = MonthToDay(12, Val(nDatasBefore(2))) - Day2
    Day1 = Day1 + Day2
    
    Day2 = MonthToDay(Val(nDatasAfter(1)) - 1, Val(nDatasAfter(2))) + _
    Val(nDatasAfter(0))
    DayCount = Day1 + Day2 ' + 1
Exit Function
    Day1 = YearToYearInDay(2009, 2010)
    Day2 = MonthToDay(Val(nDatasBefore(1)) - 1, Val(nDatasBefore(2))) + _
    Val(nDatasBefore(0))
    Day2 = MonthToDay(12, Val(nDatasBefore(2))) - Day2
    Day1 = Day1 + Day2
    Day2 = MonthToDay(Val(nDatasAfter(1)) - 1, Val(nDatasAfter(2))) + _
    Val(nDatasAfter(0))
    DayCount = Day1 + Day2 ' + 1

Exit Function
    Day1 = YearToYearInDay(Val(nDatasBefore(2)), Val(nDatasAfter(2)))

    Day2 = MonthToDay(Val(nDatasBefore(1)) - 1, Val(nDatasBefore(2))) + _
    Val(nDatasBefore(0))

    DayCount = Day1 - Day2
    
    Day1 = Val(nDatasAfter(0)) + MonthToDay(Val(nDatasAfter(1)) - 1, Val(nDatasAfter(2)))

    DayCount = DayCount + Day1 '- 0
'    Stop
End If
End Function

Function SecondOnDay(DayTimeBefore As String, DayTimeAfter As String) As Long
Dim nDatas1() As String, nDatas2() As String
Dim Days As Long, Times As Long

If DayTimeBefore = "" Or DayTimeAfter = "" Then
SecondOnDay = -100
Exit Function
End If

GetString DayTimeBefore & " ", " ", nDatas1()
GetString DayTimeAfter & " ", " ", nDatas2()

Days = DayCount(nDatas1(0), nDatas2(0))
Times = Val(FomatSecond(nDatas2(1))) - Val(FomatSecond(nDatas1(1)))

'MsgBox FomatSecond(nDatas2(1)) & " " & FomatSecond(nDatas1(1))

SecondOnDay = (Days * DayToSecond) + Times
End Function

Function YearToYearInDay(ByVal YearBefore As Long, Optional ByVal YearAfter As Long, Optional YearBeforOn As Boolean) As Long
'YearBefore = YearBefore - 1
If YearBeforOn = True Then
    YearBefore = YearBefore - 1
    YearAfter = YearAfter - 1
End If

YearBefore = (YearBefore * 365) + (YearBefore \ 4)
If YearAfter > 0 Then
    YearAfter = (YearAfter * 365) + (YearAfter \ 4)
    YearToYearInDay = YearAfter - YearBefore  '+ 365
Else
    YearToYearInDay = YearBefore
End If
End Function

Function XXXXXYearToYearInDay(ByVal YearBefore As Long, Optional ByVal YearAfter As Long) As Long
'YearAfter = YearAfter - 1
'YearBefore = 4
'MsgBox (YearBefore * 365) + (YearBefore \ 4)

YearBefore = (YearBefore * 365) + (YearBefore \ 4)
If YearAfter > 0 Then
    YearAfter = (YearAfter * 365) + (YearAfter \ 4)
    YearToYearInDay = YearAfter - YearBefore '+ 365
Else
    YearToYearInDay = YearBefore
End If
End Function

Function ThisNameDay(MyDate As String) As String
Dim nDatas() As String

GetString MyDate & "/", "/", nDatas()

nDatas(1) = DayCount("1/1/" & nDatas(2), MyDate) + 1
nDatas(2) = YearToYearInDay(Val(nDatas(2)) - 1)

ThisNameDay = MyNameDays(((nDatas(1) + nDatas(2)) Mod 7))
End Function

Function MyNameDays(iDayNames As Integer) As String

Select Case iDayNames
Case 0
    MyNameDays = "Sabtu"
Case 1
    MyNameDays = "Minggu"
Case 2
    MyNameDays = "Senin"
Case 3
    MyNameDays = "Selasa"
Case 4
    MyNameDays = "Rabu"
Case 5
    MyNameDays = "Kamis"
Case 6
    MyNameDays = "Jum'at"
End Select
End Function

Function MyNameIndexDays(iDayNames As String) As Integer

Select Case iDayNames
Case "Sabtu"
    MyNameIndexDays = 0
Case "Minggu"
    MyNameIndexDays = 1
Case "Senin"
    MyNameIndexDays = 2
Case "Selasa"
    MyNameIndexDays = 3
Case "Rabu"
    MyNameIndexDays = 4
Case "Kamis"
    MyNameIndexDays = 5
Case "Jum'at"
    MyNameIndexDays = 6
End Select
End Function



