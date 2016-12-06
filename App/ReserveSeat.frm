VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReserveSeat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reserve Seat"
   ClientHeight    =   4905
   ClientLeft      =   6825
   ClientTop       =   2640
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4320
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
      Left            =   3840
      TabIndex        =   17
      Top             =   4680
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3240
      Top             =   6000
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2160
      Top             =   6000
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1200
      Top             =   6000
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   6000
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
   Begin VB.CommandButton cmdReserve 
      Caption         =   "Reserve"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   12
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "VIP details"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin MSComCtl2.DTPicker dtDOR 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106102785
         CurrentDate     =   41687
      End
      Begin VB.TextBox txtNumber 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblDOR 
         Caption         =   "Member DOR"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblNumber 
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Member Name"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblID 
         Caption         =   "Member ID"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seat details"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   4095
      Begin VB.ComboBox cmboSeat 
         Height          =   315
         Left            =   2760
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmboRow 
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Caption         =   "Screen"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1215
         Begin VB.ComboBox cmboScreen 
            Height          =   315
            ItemData        =   "ReserveSeat.frx":0000
            Left            =   120
            List            =   "ReserveSeat.frx":000A
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Row"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Seat number"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "ReserveSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conf As Integer

Private Sub cmdBack_Click()
    Unload Me
    BookSeat.Show
End Sub

Private Sub cmdReserve_Click()
    If Not (txtID.Text = "" Or txtName.Text = "" Or txtNumber.Text = "" Or Len(txtNumber.Text) <> 11) Then
        'Prompts confirmation
        conf = MsgBox("Are you sure you want to reserve this seat?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
        If conf = vbYes Then
            Adodc1.Recordset.Filter = "MemberID=" & Val(txtID.Text)
            
            'Checks if ID given is in database
            If Adodc1.Recordset.RecordCount = 0 Then
                MsgBox "Member ID not found"
            Else
                'Checks each field individually
                If Adodc1.Recordset.Fields("MemName") <> txtName.Text Then
                    MsgBox "Incorrect name of that member ID"
                ElseIf Adodc1.Recordset.Fields("MemNumber") <> txtNumber.Text Then
                    MsgBox "Incorrect number of that member ID"
                ElseIf Adodc1.Recordset.Fields("DOR") <> dtDOR.Value Then
                    MsgBox "Incorrect Date Of Registration of that member ID"
                End If
                
                'Checks if all fields were correct
                If Adodc1.Recordset.Fields("MemberID") = Val(txtID.Text) And _
                Adodc1.Recordset.Fields("MemName") = txtName.Text And _
                Adodc1.Recordset.Fields("MemNumber") = txtNumber.Text And _
                Adodc1.Recordset.Fields("DOR") = dtDOR.Value Then
                    If cmboScreen.Text = "One" Then
                        Adodc2.Recordset.Filter = "SeatID=" & getSeatID(cmboRow.Text, Val(cmboSeat.Text), 1)
                    ElseIf cmboScreen.Text = "Two" Then
                        Adodc2.Recordset.Filter = "SeatID=" & getSeatID(cmboRow.Text, Val(cmboSeat.Text), 2)
                    Else
                        MsgBox "No screen selected"
                    End If
                    
                    If Adodc2.Recordset.RecordCount <> 0 Then
                        'Checks if seat is not already booked, to prevent double booking.
                        If Adodc2.Recordset.Fields("Reserved") = False And _
                        Adodc2.Recordset.Fields("Booked") = False Then
                            'Reserves the seat
                            Adodc2.Recordset.Fields("Reserved") = True
                            Adodc2.Recordset.Fields("Booked") = True
                            Adodc2.Recordset.Update
                            
                            'Compiles all linked tables values into a single table
                            Adodc4.Recordset.AddNew
                            Adodc4.Recordset.Fields("ShowID") = Adodc3.Recordset.Fields("ShowID")
                            Adodc4.Recordset.Fields("MovieID") = Adodc3.Recordset.Fields("MovieID")
                            Adodc4.Recordset.Fields("SeatID") = Adodc2.Recordset.Fields("SeatID")
                            Adodc4.Recordset.Fields("MemberID") = Adodc1.Recordset.Fields("MemberID")
                            If cmboScreen.Text = "One" Then
                                Adodc4.Recordset.Fields("ScreenID") = 1
                            ElseIf cmboScreen.Text = "Two" Then
                                Adodc4.Recordset.Fields("ScreenID") = 2
                            End If
                            Adodc4.Recordset.Update
                            Adodc4.Refresh
                            
                            MsgBox "Seat reserved"
                            drTicket.Show
                        Else
                            MsgBox "This seat is already reserved/booked."
                        End If
                    Else
                        MsgBox "Cannot find record"
                    End If
                    Adodc2.Refresh
                End If
            End If
        End If
    Else
        MsgBox "Please make sure all fields are filled in correctly."
    End If
End Sub

Private Function getSeatID(ByVal row As String, ByVal seat As Integer, ByVal screen As Integer) As Integer
    If screen = 1 Then
        If row = "A" Then
            getSeatID = (seat - 15) + 15
        ElseIf row = "B" Then
            getSeatID = (seat - 15) + 30
        ElseIf row = "C" Then
            getSeatID = (seat - 15) + 45
        ElseIf row = "D" Then
            getSeatID = (seat - 15) + 60
        ElseIf row = "E" Then
            getSeatID = (seat - 15) + 75
        ElseIf row = "F" Then
            getSeatID = (seat - 15) + 90
        ElseIf row = "G" Then
            getSeatID = (seat - 15) + 105
        ElseIf row = "H" Then
            getSeatID = (seat - 15) + 120
        ElseIf row = "I" Then
            getSeatID = (seat - 15) + 135
        ElseIf row = "J" Then
            getSeatID = (seat - 15) + 150
        End If
    ElseIf screen = 2 Then
        If row = "A" Then
            getSeatID = (seat - 15) + 165
        ElseIf row = "B" Then
            getSeatID = (seat - 15) + 180
        ElseIf row = "C" Then
            getSeatID = (seat - 15) + 195
        ElseIf row = "D" Then
            getSeatID = (seat - 15) + 210
        ElseIf row = "E" Then
            getSeatID = (seat - 15) + 225
        ElseIf row = "F" Then
            getSeatID = (seat - 15) + 240
        ElseIf row = "G" Then
            getSeatID = (seat - 15) + 255
        ElseIf row = "H" Then
            getSeatID = (seat - 15) + 270
        ElseIf row = "I" Then
            getSeatID = (seat - 15) + 285
        ElseIf row = "J" Then
            getSeatID = (seat - 15) + 300
        End If
    End If
End Function

Private Sub dtDOR_LostFocus()
    'Checks if future value was selected
    If DateValue(dtDOR) > DateValue(Now) Then
        MsgBox "A member cannot have a DOR to a date that hasn't already passed."
        dtDOR.Value = DateValue(Now)
    End If
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be VIP
    Adodc1.RecordSource = "VIP"
    Adodc1.Refresh
    
    'Establish a connection with the database
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Login
    Adodc2.RecordSource = "Seats"
    Adodc2.Refresh
    
    'Establish a connection with the database
    Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Shows
    Adodc3.RecordSource = "Shows"
    Adodc3.Refresh
     
    'Establish a connection with the database
    Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Transactions
    Adodc4.RecordSource = "Transactions"
    Adodc4.Refresh

    'Initializes combobox list values
    cmboScreen.List(0) = "One"
    cmboScreen.List(1) = "Two"
    
    cmboRow.List(0) = "A"
    cmboRow.List(1) = "B"
    cmboRow.List(2) = "C"
    cmboRow.List(3) = "D"
    cmboRow.List(4) = "E"
    cmboRow.List(5) = "F"
    cmboRow.List(6) = "G"
    cmboRow.List(7) = "H"
    cmboRow.List(8) = "I"
    cmboRow.List(9) = "J"
    cmboRow.Text = "A"
    
    cmboSeat.List(0) = 1
    cmboSeat.List(1) = 2
    cmboSeat.List(2) = 3
    cmboSeat.List(3) = 4
    cmboSeat.List(4) = 5
    cmboSeat.List(5) = 6
    cmboSeat.List(6) = 7
    cmboSeat.List(7) = 8
    cmboSeat.List(8) = 9
    cmboSeat.List(9) = 10
    cmboSeat.List(10) = 11
    cmboSeat.List(11) = 12
    cmboSeat.List(12) = 13
    cmboSeat.List(13) = 14
    cmboSeat.List(14) = 15
    cmboSeat.Text = 1

End Sub

Private Sub txtID_KeyPress(Keyascii As Integer)
    'Checks for only numerical and backspace key
    If Not ((Keyascii >= 48 And Keyascii <= 57) Or Keyascii = 8) Then
        Keyascii = 0
        MsgBox "Please enter numbers only"
    End If
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
