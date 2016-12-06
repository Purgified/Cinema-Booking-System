VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2520
   ClientLeft      =   7755
   ClientTop       =   3870
   ClientWidth     =   4620
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4620
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.ComboBox cmboUser 
         Height          =   315
         ItemData        =   "Login.frx":19B65
         Left            =   1800
         List            =   "Login.frx":19B6F
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
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
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4560
      Top             =   2880
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
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmboUser_KeyPress(Keyascii As Integer)
    Keyascii = 0
End Sub

Private Sub cmdCancel_Click()
    'When called, the form will close completely
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    'Make the program check the username/password combination with the database to see if they exist
    Adodc1.Recordset.Filter = "Username='" & cmboUser.Text & "' And Password='" & txtPass.Text & "'"
    
    'Calls the sub routine to validate the login credentials
    Call validateDetails
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Login
    Adodc1.RecordSource = "Login"
    Adodc1.Refresh
    
    'Tooltip indicating the purpose of the label bar.
    'Green means successful, red means failure and orange means in progress
    lblStatus.ToolTipText = "Progress bar showing whether login was successful or not."
End Sub

Private Sub validateDetails()
    'If no records found, then login details will be incorrect
    If Adodc1.Recordset.RecordCount = 0 Then
        lblStatus.BackColor = &HFF&
        MsgBox "Invalid username/password combination. Please try again."
        
        'Clear any value present in password textbox
        txtPass.Text = vbNullString
        txtPass.SetFocus
    Else
        'Record found which also means login details are correct
        lblStatus.BackColor = &HFF00&
        MsgBox "Successfully logged in as " & cmboUser.Text
        If Adodc1.Recordset.Fields("Username") = "Administrator" Then
            Adodc1.Recordset.Filter = "Username='Administrator'"
            Adodc1.Recordset.Fields("LogAdmin") = True
            Adodc1.Recordset.Update
        ElseIf Adodc1.Recordset.Fields("Username") = "Employee" Then
            Adodc1.Recordset.Filter = "Username='Administrator'"
            Adodc1.Recordset.Fields("LogAdmin") = False
            Adodc1.Recordset.Update
        End If
        Unload Me
        Dashboard.Show
    End If
End Sub
