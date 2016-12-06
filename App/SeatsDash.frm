VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SeatsSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seats settings"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
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
      Left            =   4320
      TabIndex        =   10
      Top             =   2640
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Screen 2 booked seats"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   4695
      Begin VB.CommandButton cmdUpdate2 
         Caption         =   "Unbook seat"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdGet2 
         Caption         =   "Get seats"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtBooked2 
         Height          =   405
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblBooked2 
         Caption         =   "Seats booked"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   4440
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
   Begin VB.Frame Frame1 
      Caption         =   "Screen 1 booked seats"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdUpdate1 
         Caption         =   "Unbook seat"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtBooked1 
         Height          =   405
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdGet1 
         Caption         =   "Get seats"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblBooked1 
         Caption         =   "Seats booked"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "SeatsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim screen1 As Integer
Dim screen2 As Integer

Private Sub cmdBack_Click()
    Unload Me
    CinemaSettings.Show
End Sub

Private Sub cmdGet1_Click()
    'Loop to find how many seats are booked in screen 1
    While Adodc1.Recordset.EOF = False
        If Adodc1.Recordset.Fields("ScreenID") = 1 And Adodc1.Recordset.Fields("Booked") = True Then
            screen1 = screen1 + 1
        End If
        Adodc1.Recordset.MoveNext
    Wend
    Adodc1.Recordset.MoveFirst
    txtBooked1.Text = screen1
End Sub

Private Sub cmdGet2_Click()
    'Loop to find how many seats are booked in screen 2
    While Adodc1.Recordset.EOF = False
        If Adodc1.Recordset.Fields("ScreenID") = 2 And Adodc1.Recordset.Fields("Booked") = True Then
            screen2 = screen2 + 1
        End If
        Adodc1.Recordset.MoveNext
    Wend
    Adodc1.Recordset.MoveFirst
    txtBooked2.Text = screen2
End Sub

Private Sub cmdUpdate1_Click()
    Unload Me
    UnbookSeat.Show
    'Seats next form's screen to 1
    UnbookSeat.cmboScreen.Text = "One"
End Sub

Private Sub cmdUpdate2_Click()
    Unload Me
    UnbookSeat.Show
    'Seats next form's screen to 2
    UnbookSeat.cmboScreen.Text = "Two"
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Seats
    Adodc1.RecordSource = "Seats"
    Adodc1.Refresh
    Adodc1.Recordset.MoveFirst
    
    screen1 = 0
    screen2 = 0
End Sub

Private Sub txtBooked1_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub

Private Sub txtBooked2_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub
