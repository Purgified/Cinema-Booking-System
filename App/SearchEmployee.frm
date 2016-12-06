VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SearchEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Employee"
   ClientHeight    =   6780
   ClientLeft      =   5280
   ClientTop       =   1800
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9075
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
      Left            =   8640
      TabIndex        =   25
      Top             =   6600
      Width           =   375
   End
   Begin VB.OptionButton opFirst 
      Height          =   495
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "First name"
      Top             =   480
      Width           =   255
   End
   Begin VB.OptionButton opLast 
      Height          =   495
      Left            =   2280
      TabIndex        =   23
      ToolTipText     =   "Last Name"
      Top             =   480
      Width           =   255
   End
   Begin VB.OptionButton opContact 
      Height          =   495
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Contact"
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton opPosition 
      Height          =   735
      Left            =   6960
      TabIndex        =   21
      ToolTipText     =   "Position"
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton opID 
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      ToolTipText     =   "ID"
      Top             =   480
      Width           =   255
   End
   Begin VB.OptionButton opDOA 
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      ToolTipText     =   "DOA"
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton opAddress 
      Height          =   495
      Left            =   5280
      TabIndex        =   18
      ToolTipText     =   "Address"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame Frame7 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "DOA"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Date Of Admission"
      Top             =   1200
      Width           =   1695
      Begin MSComCtl2.DTPicker dtDOA 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   122159105
         CurrentDate     =   41678
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Position"
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
      Left            =   7320
      TabIndex        =   12
      Top             =   120
      Width           =   1695
      Begin VB.ComboBox cmboPosition 
         Height          =   315
         ItemData        =   "SearchEmployee.frx":0000
         Left            =   240
         List            =   "SearchEmployee.frx":0002
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
      Begin VB.TextBox txtContact2 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtContact1 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Identification Number"
      Top             =   120
      Width           =   1455
      Begin VB.TextBox txtID 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last Name"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtLast 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11280
      Top             =   3480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SearchEmployee.frx":0004
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7858
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
      Height          =   855
      Left            =   7440
      TabIndex        =   2
      ToolTipText     =   "Search through records with the selected combo box"
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "First Name"
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
      Begin VB.TextBox txtFirst 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "SearchEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
    EmpDash.Show
End Sub

Private Sub cmdSearch_Click()
    'First Name textbox will be used for the search
    If opFirst.Value = True Then
        Adodc1.Recordset.Filter = "FirstName='" & txtFirst.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "First name not found"
            Adodc1.Refresh
        End If
    'Last Name textbox will be used for the search
    ElseIf opLast.Value = True Then
        Adodc1.Recordset.Filter = "LastName='" & txtLast.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Last name not found"
            Adodc1.Refresh
        End If
    'ID textbox will be used for the search
    ElseIf opID.Value = True Then
        Adodc1.Recordset.Filter = "ID='" & txtID.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "ID not found"
            Adodc1.Refresh
        End If
    'Contact No textbox will be used for the search
    ElseIf opContact.Value = True Then
        Adodc1.Recordset.Filter = "Contact='" & (txtContact1.Text & txtContact2.Text) & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Contact number not found"
            Adodc1.Refresh
        End If
    'Position textbox will be used for the search
    ElseIf opPosition.Value = True Then
        Adodc1.Recordset.Filter = "Position='" & cmboPosition.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Position not found"
            Adodc1.Refresh
        End If
    'Doa textbox will be used for the search
    ElseIf opDOA.Value = True Then
        Adodc1.Recordset.Filter = "DOA='" & dtDOA.Value & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Date Of Admission not found"
            Adodc1.Refresh
        End If
    'Address textbox will be used for the search
    ElseIf opAddress.Value = True Then
        Adodc1.Recordset.Filter = "Address='" & txtAddress.Text & "'"
        If Adodc1.Recordset.RecordCount = 0 Then
            MsgBox "Address not found"
            Adodc1.Refresh
        End If
    Else
        MsgBox "Please select one of the options from the option boxes given."
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be the Employee Info table
    Adodc1.RecordSource = "EmpInfo"
    Adodc1.Refresh
    
    'Initializes combobox list values
    cmboPosition.List(0) = "Manager"
    cmboPosition.List(1) = "Ticketer"
    cmboPosition.List(2) = "Receptionist"
    cmboPosition.List(3) = "Regular"
    cmboPosition.List(4) = "Other"
End Sub

Private Sub cmboPosition_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Please select one of the values from the list."
End Sub

Private Sub txtAddress_KeyPress(Keyascii As Integer)
    'Checks for letters (capital and small), backspace, space, shift and shift keypress
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 32 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only."
    End If
End Sub

Private Sub txtContact1_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter numbers only"
    End If
    'Checks for limit of field and that backspace is not pressed
    If Len(txtContact1.Text) >= 4 Then
        If Not (Keyascii = 8) Then
            Keyascii = 0
            MsgBox "Cannot enter more numbers in this field"
        End If
    End If
End Sub

Private Sub txtContact1_LostFocus()
    'Checks wheter total 4 digits were entered or not
    If Len(txtContact1.Text) <> 4 Then
        MsgBox "Please make sure 4 digits were entered"
    End If
End Sub

Private Sub txtContact2_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter numbers only"
    End If
    'Checks for limit of field and that backspace is not pressed
    If Len(txtContact2.Text) >= 7 Then
        If Not (Keyascii = 8) Then
            Keyascii = 0
            MsgBox "Cannot enter more numbers in this field"
        End If
    End If
End Sub

Private Sub txtContact2_LostFocus()
    'Checks wheter total 7 digits were entered or not
    If Len(txtContact2.Text) <> 7 Then
        MsgBox "Please make sure 7 digits were entered."
    End If
End Sub

Private Sub txtID_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter digits only."
    End If
End Sub

Private Sub txtFirst_KeyPress(Keyascii As Integer)
    'Checks for letters (capital and small), backspace, space, shift and shift keypress
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only"
    End If
End Sub

Private Sub txtLast_KeyPress(Keyascii As Integer)
    'Checks for letters (capital and small), backspace, space, shift and shift keypress
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only."
    End If
End Sub

Private Sub txtLast_LostFocus()
    'Converts first letter of every word to its capital letter
    txtLast = StrConv(txtLast.Text, vbProperCase)
End Sub

Private Sub txtFirst_LostFocus()
    'Converts first letter of every word to its capital letter
    txtFirst = StrConv(txtFirst.Text, vbProperCase)
End Sub


