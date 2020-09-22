VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalender 
   AutoRedraw      =   -1  'True
   Caption         =   "Planet-Source-Code.com"
   ClientHeight    =   6780
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDateBy_AddDayOnDay 
      Caption         =   "DateBy_AddDayOnDay"
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3615
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   66
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   64
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   6
         Left            =   960
         TabIndex        =   52
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   51
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   17
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   15
         Text            =   "427"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   12
         Text            =   "2008"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   10
         Text            =   "12"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdAddDayOnDay 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtAddDayOnDay 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Text            =   "30"
         Top             =   360
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDAddDayOnDay 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   31
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddDayOnDay 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   12
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddDayOnDay 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2008
         Max             =   3000
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddDayOnDay 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   18
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   99999
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "None"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "MonthMod"
         Height          =   255
         Index           =   65
         Left            =   960
         TabIndex        =   65
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "DayMod"
         Height          =   255
         Index           =   64
         Left            =   960
         TabIndex        =   63
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "YearOut"
         Height          =   255
         Index           =   52
         Left            =   120
         TabIndex        =   54
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "MonthOut"
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   53
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DateOut"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Day"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Date In"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frmDateBy_AddMonthOnMonth 
      Caption         =   "DateBy_AddMonthOnMonth"
      Height          =   4335
      Left            =   4080
      TabIndex        =   35
      Top             =   240
      Width           =   3615
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   68
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   6
         Left            =   960
         TabIndex        =   56
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   55
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   50
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   48
         Text            =   "14"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   46
         Text            =   "2008"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   44
         Text            =   "12"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdAddMonthOnMonth 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   37
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtAddMonthOnMonth 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   36
         Text            =   "30"
         Top             =   360
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDAddMonthOnMonth 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   38
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   31
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddMonthOnMonth 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   45
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   31
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddMonthOnMonth 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   47
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2008
         Max             =   3000
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddMonthOnMonth 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   49
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   99999
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "None"
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "MonthMod"
         Height          =   255
         Index           =   25
         Left            =   960
         TabIndex        =   69
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "YearOut"
         Height          =   255
         Index           =   54
         Left            =   120
         TabIndex        =   58
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DayOut"
         Height          =   255
         Index           =   53
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DateOut"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   255
         Index           =   12
         Left            =   1680
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Date In"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   39
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame frmDateBy_AddYearOnYear 
      Caption         =   "DateBy_AddYearOnYear"
      Height          =   4335
      Left            =   7920
      TabIndex        =   19
      Top             =   240
      Width           =   3615
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   6
         Left            =   960
         TabIndex        =   60
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   59
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   32
         Text            =   "47"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   30
         Text            =   "2008"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   28
         Text            =   "12"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtAddYearOnYear 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Text            =   "30"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdAddYearOnYear 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   3840
         Width           =   975
      End
      Begin MSComCtl2.UpDown UpDAddYearOnYear 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   22
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   31
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddYearOnYear 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   29
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   31
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddYearOnYear 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   31
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2008
         Max             =   3000
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDAddYearOnYear 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   33
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   99999
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "None"
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   71
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "MonthOut"
         Height          =   255
         Index           =   56
         Left            =   120
         TabIndex        =   62
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DayOut"
         Height          =   255
         Index           =   55
         Left            =   120
         TabIndex        =   61
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Date In"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "/"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DateOut"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicStripDurasi 
      Height          =   495
      Left            =   7680
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   11760
      Width           =   2295
   End
   Begin VB.PictureBox PicStripRP 
      Height          =   495
      Left            =   8280
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   11040
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "BoedidaX9 - Nurfaststar@Yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   74
      Top             =   6000
      Width           =   11535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Upgrade Date My Function"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   73
      Top             =   5400
      Width           =   11535
   End
   Begin VB.Label Label2 
      Caption         =   "input only abs(integer)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   72
      Top             =   4920
      Width           =   11535
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   528
      X2              =   768
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   272
      X2              =   512
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line Line1 
      X1              =   16
      X2              =   256
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Shape Shape1 
      Height          =   4695
      Index           =   2
      Left            =   7800
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   4695
      Index           =   1
      Left            =   3960
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   4695
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblMember 
      Caption         =   "Label1"
      Height          =   495
      Index           =   3
      Left            =   9840
      TabIndex        =   0
      Top             =   11040
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sl_Day As Long, sl_Month As Integer, sl_Year As Integer
Dim sl_Date As String

Private Sub cmdAddDayOnDay_Click()
sl_Date = txtAddDayOnDay(0).Text & "/" & txtAddDayOnDay(1).Text & "/" & txtAddDayOnDay(2).Text
sl_Month = 0
txtAddDayOnDay(4).Text = DateBy_AddDayOnDay(sl_Date, txtAddDayOnDay(3), sl_Month, sl_Year, True)
txtAddDayOnDay(5).Text = sl_Month
txtAddDayOnDay(6).Text = sl_Year

DateBy_AddMonthOnMonth sl_Date, sl_Month, sl_Day
txtAddDayOnDay(7).Text = Val(txtAddDayOnDay(3).Text) - sl_Day
txtAddDayOnDay(8).Text = sl_Month - (sl_Year * 12)

Label1(0).Caption = ThisNameDay(sl_Date) & " " & sl_Date & " + " & txtAddDayOnDay(3) & " Day = " & vbCrLf & ThisNameDay(txtAddDayOnDay(4).Text) & " " & txtAddDayOnDay(4).Text & vbCrLf & txtAddDayOnDay(6) & " Year, " & txtAddDayOnDay(8) & " Month, " & txtAddDayOnDay(7) & " Day."
End Sub

Private Sub cmdAddMonthOnMonth_Click()
sl_Date = txtAddMonthOnMonth(0).Text & "/" & txtAddMonthOnMonth(1).Text & "/" & txtAddMonthOnMonth(2).Text
txtAddMonthOnMonth(4).Text = DateBy_AddMonthOnMonth(sl_Date, txtAddMonthOnMonth(3), sl_Day, sl_Year)
txtAddMonthOnMonth(5).Text = sl_Day
txtAddMonthOnMonth(6).Text = sl_Year

txtAddMonthOnMonth(7).Text = txtAddMonthOnMonth(3) - (sl_Year * 12)

Label1(1).Caption = ThisNameDay(sl_Date) & " " & sl_Date & " + " & txtAddMonthOnMonth(3) & " Month = " & vbCrLf & ThisNameDay(txtAddMonthOnMonth(4)) & " " & txtAddMonthOnMonth(4) & vbCrLf & txtAddMonthOnMonth(6) & " Year, " & txtAddMonthOnMonth(7) & " Month, 0 Day."
End Sub

Private Sub cmdAddYearOnYear_Click()
sl_Date = txtAddYearOnYear(0).Text & "/" & txtAddYearOnYear(1).Text & "/" & txtAddYearOnYear(2).Text
txtAddYearOnYear(4).Text = DateBy_AddYearOnYear(sl_Date, txtAddYearOnYear(3), sl_Day, sl_Month)
txtAddYearOnYear(5).Text = sl_Day
txtAddYearOnYear(6).Text = sl_Month

Label1(2).Caption = ThisNameDay(sl_Date) & " " & sl_Date & " + " & txtAddYearOnYear(3) & " Year = " & vbCrLf & ThisNameDay(txtAddYearOnYear(4)) & " " & txtAddYearOnYear(4).Text & vbCrLf & txtAddYearOnYear(3) & " Year, 0 Month, 0 Day."
End Sub

Private Sub txtAddDayOnDay_Change(Index As Integer)
UpDAddDayOnDay(3).Value = txtAddDayOnDay(3).Text
End Sub

Private Sub UpDAddDayOnDay_Change(Index As Integer)
txtAddDayOnDay(Index).Text = UpDAddDayOnDay(Index).Value
cmdAddDayOnDay_Click
End Sub

Private Sub UpDAddMonthOnMonth_Change(Index As Integer)
txtAddMonthOnMonth(Index).Text = UpDAddMonthOnMonth(Index).Value
cmdAddMonthOnMonth_Click
End Sub

Private Sub UpDAddYearOnYear_Change(Index As Integer)
txtAddYearOnYear(Index).Text = UpDAddYearOnYear(Index).Value
cmdAddYearOnYear_Click
End Sub

Private Sub UpDDayBy_Month_Change(Index As Integer)
Dim s_Date As String
txtDayBy_Month(Index).Text = UpDDayBy_Month(Index).Value

txtDayBy_Month(4).Text = DayBy_Month(txtDayBy_Month(0).Text & "/" & txtDayBy_Month(1).Text & "/" & txtDayBy_Month(2).Text, Val(txtDayBy_Month(3)), s_Date)
txtDayBy_Month(5).Text = s_Date
End Sub
