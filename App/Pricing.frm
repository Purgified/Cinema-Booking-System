VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Pricing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pricing"
   ClientHeight    =   4140
   ClientLeft      =   6825
   ClientTop       =   2640
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4860
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
      Left            =   4440
      TabIndex        =   12
      Top             =   3960
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7200
      Top             =   2640
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
   Begin VB.Frame Frame2 
      Caption         =   "Update prices"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
      Begin VB.TextBox txtNew2D 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNew3D 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Prices"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblNew2D 
         Caption         =   "2D tickets"
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
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblNew3D 
         Caption         =   "3D tickets"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current prices"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdGet 
         Caption         =   "Get prices"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtCurrent3D 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtCurrent2D 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl3DCurrent 
         Caption         =   "3D tickets"
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
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl2DCurrent 
         Caption         =   "2D tickets"
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
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "Pricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conf As Integer

Private Sub cmdBack_Click()
    Unload Me
    CinemaSettings.Show
End Sub

Private Sub cmdGet_Click()
    Adodc1.Recordset.MoveFirst
    txtCurrent2D.Text = "Rs " & Adodc1.Recordset.Fields("Price")
    Adodc1.Recordset.MoveNext
    txtCurrent3D.Text = "Rs " & Adodc1.Recordset.Fields("Price")
End Sub

Private Sub cmdUpdate_Click()
    If Not (txtNew2D.Text = "" Or txtNew3D.Text = "" Or Val(txtNew2D.Text) < 150 Or _
    Val(txtNew2D.Text) > 800 Or Val(txtNew3D.Text) < 200 Or Val(txtNew3D.Text) > 900) Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to update the price?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            Adodc1.Recordset.MoveFirst
            Adodc1.Recordset.Fields("Price") = txtNew2D.Text
            Adodc1.Recordset.MoveNext
            Adodc1.Recordset.Fields("Price") = txtNew3D.Text
            Adodc1.Recordset.Update
            MsgBox "Prices have been updated in the database."
            Adodc1.Refresh
        End If
    Else
        MsgBox "Please make sure all fields are filled in correctly."
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Pricing
    Adodc1.RecordSource = "Pricing"
    Adodc1.Refresh
End Sub

Private Sub txtCurrent2D_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub

Private Sub txtCurrent3D_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "This field cannot be altered."
End Sub

Private Sub txtNew3D_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter numbers only"
    End If
    'Checks for limit of field and whether backspace was pressed
    If Len(txtNew3D.Text) >= 3 Then
        If Not (Keyascii = 8) Then
            Keyascii = 0
            MsgBox "More digits are not allowed in this field."
        End If
    End If
End Sub

Private Sub txtNew3D_LostFocus()
    If Not (txtNew3D.Text = "") Then
        'Sets minimum price value of 3D which comply's with Cinepax's pricing system
        If Val(txtNew3D.Text) < 200 Then
            MsgBox "Cannot have 3D price less than Rs 200."
        'Sets maximum price value of 3D which comply's with Cinepax's pricing system
        ElseIf Val(txtNew3D.Text) > 900 Then
            MsgBox "Cannot have 3D price more than Rs 900."
        End If
    End If
End Sub

Private Sub txtNew2D_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter numbers only"
    End If
    'Checks for limit of field and whether backspace was pressed
    If Len(txtNew2D.Text) >= 3 Then
        If Not (Keyascii = 8) Then
            Keyascii = 0
            MsgBox "More digits are not allowed in this field."
        End If
    End If
End Sub

Private Sub txtNew2D_LostFocus()
    If Not (txtNew2D.Text = "") Then
        'Sets minimum price value of 2D which comply's with Cinepax's pricing system
        If Val(txtNew2D.Text) < 150 Then
            MsgBox "Cannot have 2D price less than Rs 150."
        'Sets maximum price value of 2D which comply's with Cinepax's pricing system
        ElseIf Val(txtNew2D.Text) > 800 Then
            MsgBox "Cannot have 2D price more than Rs 800."
        End If
    End If
End Sub
