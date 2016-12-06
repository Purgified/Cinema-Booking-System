VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MonthlyDash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Report"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Start date of report"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   3015
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Arabic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110166017
         CurrentDate     =   41693
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate report"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   375
   End
End
Attribute VB_Name = "MonthlyDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
    Dashboard.Show
End Sub

Private Sub cmdGenerate_Click()
    'Adds one month to the selected DTPicker value
    Dim X As String
    X = DateAdd("m", 1, dtStart.Value)
    'Sets the SQL statement of the report to the date selected by the user
    DataEnvironment1.Commands("Monthly_report").CommandType = adCmdText
    DataEnvironment1.Commands("Monthly_report").CommandText = "SELECT * FROM Transactions WHERE TransDate>= #" & dtStart.Value & "# AND TransDate<= #" & X & "#"
    drMonthly.Show
End Sub

Private Sub dtStart_LostFocus()
    'Checks if user asks for a report that doesn't even exist (by determining future value)
    If DateValue(dtStart) > DateValue(Now) Then
        MsgBox "You cannot generate a report on date that hasn't already passed."
        dtStart.Value = DateValue(Now)
    End If
End Sub
