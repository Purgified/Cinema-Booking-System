VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SearchVIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search VIP members"
   ClientHeight    =   5220
   ClientLeft      =   4425
   ClientTop       =   3195
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11415
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
      Left            =   11040
      TabIndex        =   14
      Top             =   5040
      Width           =   375
   End
   Begin VB.OptionButton opDOR 
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      ToolTipText     =   "ID"
      Top             =   480
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Date of registration"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      TabIndex        =   11
      ToolTipText     =   "Identification Number"
      Top             =   120
      Width           =   2055
      Begin MSComCtl2.DTPicker dtDOR 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   129761281
         CurrentDate     =   41687
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   6480
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.OptionButton opNumber 
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "ID"
      Top             =   480
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Number"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Identification Number"
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtNumber 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.OptionButton opName 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "Last Name"
      Top             =   480
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.OptionButton opID 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "First name"
      Top             =   480
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtID 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SearchVIP.frx":0000
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
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
Attribute VB_Name = "SearchVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
    'ID textbox will be used for the search
    If opID.Value = True Then
        Adodc1.Recordset.Filter = "MemberID=" & Val(txtID.Text)
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "ID not found"
            Adodc1.Refresh
        End If
    'Name textbox will be used for the search
    ElseIf opName.Value = True Then
        Adodc1.Recordset.Filter = "MemName='" & txtName.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Name not found"
            Adodc1.Refresh
        End If
    'Number textbox will be used for the search
    ElseIf opNumber.Value = True Then
        Adodc1.Recordset.Filter = "MemNumber='" & txtNumber.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Number not found"
            Adodc1.Refresh
        End If
    'DOR textbox will be used for the search
    ElseIf opDOR.Value = True Then
        Adodc1.Recordset.Filter = "DOR='" & dtDOR.Value & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Date Of Registration not found"
            Adodc1.Refresh
        End If
    Else
        MsgBox "Please make sure you selected or entered values in one of the fields."
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    VIPDash.Show
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be the VIP table
    Adodc1.RecordSource = "VIP"
    Adodc1.Refresh
End Sub

Private Sub txtID_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter digits only."
    End If
End Sub

Private Sub txtName_KeyPress(Keyascii As Integer)
    'Checks for letters (capital and small), backspace, space, shift and shift keypress
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 20 Or Keyascii = 32) Then
        Keyascii = 0
        MsgBox "Please enter letters only."
    End If
End Sub

Private Sub txtName_LostFocus()
    txtName = StrConv(txtName.Text, vbProperCase)
End Sub

Private Sub txtNumber_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter numbers only"
    End If
    'Checks for limit of field and that backspace is not pressed
    If Len(txtNumber.Text) >= 11 Then
        If Not (Keyascii = 8) Then
            Keyascii = 0
            MsgBox "Cannot enter more numbers in this field"
        End If
    End If
End Sub

Private Sub txtNumber_LostFocus()
    'Checks wheter total 11 digits were entered or not
    If Len(txtNumber.Text) <> 11 Then
        MsgBox "Please make sure 11 digits were entered."
    End If
End Sub
