VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form AddEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Employee"
   ClientHeight    =   5775
   ClientLeft      =   7815
   ClientTop       =   2355
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4695
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
      Left            =   4200
      TabIndex        =   17
      Top             =   5520
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   5280
      Visible         =   0   'False
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
   Begin VB.CommandButton cmdAddEmp 
      Caption         =   "Add Employee"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   16
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Frame Frame7 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   4215
      Begin VB.TextBox txtNote 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
      Begin VB.TextBox txtAddress 
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
      Begin VB.ComboBox cmboPos 
         Height          =   315
         ItemData        =   "AddEmployee.frx":0000
         Left            =   120
         List            =   "AddEmployee.frx":0002
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Date of Admission"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      Begin MSComCtl2.DTPicker dtDOA 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   134610945
         CurrentDate     =   41677
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
      Begin VB.TextBox txtContact2 
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtContact1 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Last name"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
      Begin VB.TextBox txtLast 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "First name"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtFirst 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "AddEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conf As Integer
Dim cont As String

Private Sub cmboPos_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "You cannot alter this field. Select a value from the list."
End Sub

Private Sub cmdBack_Click()
    Unload Me
    EmpDash.Show
End Sub

Private Sub cmdAddEmp_Click()
    'Checks for all main fields to be filled
    If Not (txtFirst.Text = "" Or txtLast.Text = "" Or cmboPos.Text = "" _
    Or txtContact1.Text = "" Or txtContact2.Text = "" Or txtAddress.Text = "" _
    Or Len(txtContact1.Text) <> 4 Or Len(txtContact2.Text) <> 7) Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to add this record?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            'Concatenates the two fields
            cont = txtContact1.Text & txtContact2.Text
            Dim exists As Boolean
            exists = False
            'Checks if contact, the unique field, exists in the database or not
            While Adodc1.Recordset.EOF = False
                If Adodc1.Recordset.Fields("Contact") = cont Then
                    exists = True
                End If
                Adodc1.Recordset.MoveNext
            Wend
            If Adodc1.Recordset.RecordCount <> 0 Then
                Adodc1.Recordset.MoveFirst
            End If
            If exists = False Then
                'Adds a new row to the table
                Adodc1.Recordset.AddNew
                
                Adodc1.Recordset.Fields("FirstName") = txtFirst.Text
                Adodc1.Recordset.Fields("LastName") = txtLast.Text
                Adodc1.Recordset.Fields("Contact") = txtContact1.Text & txtContact2.Text
                Adodc1.Recordset.Fields("Position") = cmboPos.Text
                Adodc1.Recordset.Fields("Address") = txtAddress.Text
                Adodc1.Recordset.Fields("Note") = txtNote.Text
                Adodc1.Recordset.Fields("DOA") = dtDOA.Value
                
                'Updates the records in the table to show the new field values
                Adodc1.Recordset.Update
                
                MsgBox "Employee has been added to the database."
            Else
                MsgBox "The contact number you entered already exists in the database."
            End If
        End If
    Else
        MsgBox "Please make sure all mandatory fields are filled in correctly."
    End If
End Sub

Private Sub dtDOA_LostFocus()
    'Checks to make sure date selected is not in the future
    If DateValue(dtDOA.Value) > DateValue(Now) Then
        MsgBox "You cannot add an employee to a date that hasn't already passed."
        dtDOA.Value = DateValue(Now)
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be EmpInfo
    Adodc1.RecordSource = "EmpInfo"
    Adodc1.Refresh
    
    'Initializes combobox list values
    cmboPos.List(0) = "Manager"
    cmboPos.List(1) = "Ticketer"
    cmboPos.List(2) = "Receptionist"
    cmboPos.List(3) = "Regular"
    cmboPos.List(4) = "Other"
    
    dtDOA.Value = DateValue(Now)
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
    'Focuses back to the same field if amount of numbers is not four
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
    If Len(txtContact2.Text) <> 7 Then
        MsgBox "Please make sure 7 digits were entered"
    End If
End Sub

Private Sub txtFirst_KeyPress(Keyascii As Integer)
    'Checks for letters, backspace, caps lock and shift
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only"
    End If
End Sub

Private Sub txtLast_KeyPress(Keyascii As Integer)
    'Checks for letters, backspace, caps lock and shift
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only"
    End If
End Sub

Private Sub txtLast_LostFocus()
    'Capitalizes the first letter of the word
    txtLast = StrConv(txtLast.Text, vbProperCase)
End Sub

Private Sub txtFirst_LostFocus()
    txtFirst = StrConv(txtFirst.Text, vbProperCase)
End Sub
