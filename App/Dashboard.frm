VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Dashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   5880
   ClientLeft      =   7530
   ClientTop       =   3615
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   3870
   Begin VB.CommandButton cmdMonthly 
      Caption         =   "Monthly report"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   7
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdDaily 
      Caption         =   "Daily report"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   6240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdEmp 
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMovies 
      Caption         =   "Movies"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
      Begin VB.CommandButton cmdShows 
         Caption         =   "Shows"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   0
      Picture         =   "Dashboard.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conf As Integer

Private Sub cmdEmp_Click()
    Unload Me
    EmpDash.Show
End Sub

Private Sub cmdLogout_Click()
    conf = MsgBox("Are you sure you want to log out?", vbYesNo + vbQuestion + vbDefaultButton2, App.Title)
    If conf = vbYes Then
        Unload Me
        Login.Show
    End If
End Sub

Private Sub cmdMonthly_Click()
    Unload Me
    MonthlyDash.Show
End Sub

Private Sub cmdMovies_Click()
    Unload Me
    MoviesDash.Show
End Sub

Private Sub cmdDaily_Click()
    Unload Me
    DailyDash.Show
End Sub

Private Sub cmdSettings_Click()
    'Checks for administrative rights
    Adodc1.Recordset.Filter = "Username='Administrator'"
    If Adodc1.Recordset.Fields("LogAdmin") = True Then
        Unload Me
        CinemaSettings.Show
    Else
        MsgBox "You must be logged in as an administrator to access this form."
    End If
End Sub

Private Sub cmdShows_Click()
    Unload Me
    ShowsDash.Show
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Login
    Adodc1.RecordSource = "Login"
    Adodc1.Refresh
End Sub
