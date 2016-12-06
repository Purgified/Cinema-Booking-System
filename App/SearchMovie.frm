VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SearchMovie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Movie"
   ClientHeight    =   4980
   ClientLeft      =   5835
   ClientTop       =   2640
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7575
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
      Left            =   7200
      TabIndex        =   8
      Top             =   4800
      Width           =   375
   End
   Begin VB.ComboBox cmboStyle 
      Height          =   315
      ItemData        =   "SearchMovie.frx":0000
      Left            =   3840
      List            =   "SearchMovie.frx":000A
      TabIndex        =   7
      Text            =   "2D"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cmboRating 
      Height          =   315
      ItemData        =   "SearchMovie.frx":0016
      Left            =   3840
      List            =   "SearchMovie.frx":0029
      TabIndex        =   6
      Text            =   "Everyone"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtRelease 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   124846081
      CurrentDate     =   41685
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   6360
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
      Caption         =   "Search catagory"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cmboCatagory 
         Height          =   315
         ItemData        =   "SearchMovie.frx":004D
         Left            =   240
         List            =   "SearchMovie.frx":005D
         TabIndex        =   1
         Text            =   "Name"
         Top             =   360
         Width           =   2655
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SearchMovie.frx":0084
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SearchMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmboCatagory_Click()
    If cmboCatagory.Text = "Name" Then
        txtName.Visible = True
        cmboRating.Visible = False
        cmboStyle.Visible = False
        dtRelease.Visible = False
    ElseIf cmboCatagory.Text = "Release date" Then
        txtName.Visible = False
        cmboRating.Visible = False
        cmboStyle.Visible = False
        dtRelease.Visible = True
    ElseIf cmboCatagory.Text = "Rating" Then
        txtName.Visible = False
        cmboRating.Visible = True
        cmboStyle.Visible = False
        dtRelease.Visible = False
    ElseIf cmboCatagory.Text = "2d/3d" Then
        txtName.Visible = False
        cmboRating.Visible = False
        cmboStyle.Visible = True
        dtRelease.Visible = False
    End If
End Sub

Private Sub cmboCatagory_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Please select one of the values from the list."
End Sub

Private Sub cmboRating_Keypress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Please select one of the values from the list."
End Sub

Private Sub cmboStyle_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Please select one of the values from the list."
End Sub

Private Sub cmdBack_Click()
    Unload Me
    MoviesDash.Show
End Sub

Private Sub cmdSearch_Click()
    'Checks which combobox value was selected and then filters to the appropriate field(s)
    If cmboCatagory.Text = "Name" Then
        Adodc1.Recordset.Filter = "MovieName='" & txtName.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Movie name not found"
            Adodc1.Refresh
        End If
    ElseIf cmboCatagory.Text = "Release date" Then
        Adodc1.Recordset.Filter = "ReleaseDate=" & dtRelease.Value
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Release date not found"
            Adodc1.Refresh
        End If
    ElseIf cmboCatagory.Text = "Rating" Then
        Adodc1.Recordset.Filter = "Rating='" & cmboRating.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Rating not found"
            Adodc1.Refresh
        End If
    ElseIf cmboCatagory.Text = "2d/3d" Then
        Adodc1.Recordset.Filter = "Dimension='" & cmboStyle.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Style not found"
            Adodc1.Refresh
        End If
    Else
        MsgBox "Make sure a value is selected."
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Movies
    Adodc1.RecordSource = "Movies"
    Adodc1.Refresh
    
    'Initializes all combobox list values
    cmboStyle.List(0) = "2D"
    cmboStyle.List(1) = "3D"
    
    cmboRating.List(0) = "Everyone"
    cmboRating.List(1) = "7+"
    cmboRating.List(2) = "10+"
    cmboRating.List(3) = "PG 13"
    cmboRating.List(4) = "17+"
    cmboRating.List(5) = "Mature"
    
    cmboCatagory.List(0) = "Name"
    cmboCatagory.List(1) = "Release date"
    cmboCatagory.List(2) = "Rating"
    cmboCatagory.List(3) = "2d/3d"
End Sub

