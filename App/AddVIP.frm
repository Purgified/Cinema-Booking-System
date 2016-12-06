VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form AddVIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add VIP"
   ClientHeight    =   2880
   ClientLeft      =   7530
   ClientTop       =   4740
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4425
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
      Left            =   3960
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   4080
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add VIP Member"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date Of Registration"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1935
      Begin MSComCtl2.DTPicker dtDOR 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   113770497
         CurrentDate     =   41687
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtNumber 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "AddVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conf As Integer

Private Sub cmdAdd_Click()
    If Not (txtName.Text = "" Or txtNumber.Text = "" Or Len(txtNumber.Text) <> 11) Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to add this record?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            'Checks if the member's contact number (unique field) exists in db already
            Dim exists As Boolean
            While Adodc1.Recordset.EOF = False
                If Adodc1.Recordset.Fields("MemNumber") = Val(txtNumber.Text) Then
                    exists = True
                End If
                Adodc1.Recordset.MoveNext
            Wend
            Adodc1.Recordset.MoveFirst
            If exists = False Then
                'Create new row for new records
                Adodc1.Recordset.AddNew
                
                Adodc1.Recordset.Fields("MemName") = txtName.Text
                Adodc1.Recordset.Fields("MemNumber") = Val(txtNumber.Text)
                Adodc1.Recordset.Fields("DOR") = dtDOR.Value
                
                Adodc1.Recordset.Update
                Adodc1.Refresh
                
                MsgBox "The has been given VIP membership and has been added to the database."
            Else
                MsgBox "This member already exists in the database."
            End If
        End If
    Else
        MsgBox "Please make sure all fields were filled in."
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    VIPDash.Show
End Sub

Private Sub dtDOR_LostFocus()
    'Checks to make sure date selected is not in the future
    If DateValue(dtDOR.Value) > DateValue(Now) Then
        MsgBox "You cannot add a member to a date that hasn't already passed."
        dtDOR.Value = DateValue(Now)
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be VIP
    Adodc1.RecordSource = "VIP"
    Adodc1.Refresh
End Sub

Private Sub txtName_KeyPress(Keyascii As Integer)
    'Checks for letters (capital and small), backspace, space, shift and shift keypress
    If Not ((Keyascii >= 97 And Keyascii <= 122) Or _
    (Keyascii >= 65 And Keyascii <= 90) Or Keyascii = 8 Or _
    Keyascii = 16 Or Keyascii = 32 Or Keyascii = 20) Then
        Keyascii = 0
        MsgBox "Please enter letters only"
    End If
End Sub

Private Sub txtName_LostFocus()
    'Converts first letter of every word to its capital letter
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
        MsgBox "Please make sure 11 digits were entered"
    End If
End Sub
