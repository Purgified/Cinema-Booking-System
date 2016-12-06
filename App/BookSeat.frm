VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BookSeat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Seat"
   ClientHeight    =   3465
   ClientLeft      =   7680
   ClientTop       =   4605
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3960
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
      Left            =   3480
      TabIndex        =   11
      Top             =   3240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2640
      Top             =   6720
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
      Left            =   1440
      Top             =   6720
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
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
      Begin VB.CommandButton cmdBook 
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBooked 
         Caption         =   "Booked seats"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton cmdReserve 
         Caption         =   "Reserve"
         BeginProperty Font 
            Name            =   "Myriad Arabic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6720
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
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox cmboSeat 
         Height          =   315
         Left            =   2760
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmboRow 
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Caption         =   "Screen"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         Begin VB.ComboBox cmboScreen 
            Height          =   315
            ItemData        =   "BookSeat.frx":0000
            Left            =   120
            List            =   "BookSeat.frx":0002
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
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
   End
End
Attribute VB_Name = "BookSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conf As Integer

Private Sub cmboSCreen_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "You cannot alter this field"
End Sub

Private Sub cmdBook_Click()
    'Prompts confirmation
    conf = MsgBox("Are you sure you want to book this seat?", vbYesNo + vbQuestion + vbDefaultButton2, App.title)
    If conf = vbYes Then
        If cmboScreen.Text = "One" Then
            'Locks onto the exact Seat ID with the SeatID field
            Adodc1.Recordset.Filter = "SeatID=" & getSeatID(cmboRow.Text, Val(cmboSeat.Text), 1)
            
            If Adodc1.Recordset.RecordCount <> 0 Then
                'Checks if seat is not already booked, to prevent double booking.
                If Adodc1.Recordset.Fields("Booked") = False Then
                    'Books the seat and update the database with the new value
                    Adodc1.Recordset.Fields("Booked") = True
                    Adodc1.Recordset.Update
                        
                    Adodc3.Recordset.AddNew
                    Adodc3.Recordset.Fields("ShowID") = Adodc2.Recordset.Fields("ShowID")
                    Adodc3.Recordset.Fields("MovieID") = Adodc2.Recordset.Fields("MovieID")
                    Adodc3.Recordset.Fields("SeatID") = Adodc1.Recordset.Fields("SeatID")
                    Adodc3.Recordset.Fields("ScreenID") = 1
                    Adodc3.Recordset.Fields("TransDate") = DateValue(Now)
                    Adodc3.Recordset.Update
                    Adodc3.Refresh
                        
                    MsgBox "Seat Booked!"
                    drTicket.Show
                Else
                    MsgBox "This seat is already booked."
                End If
            Else
                MsgBox "Cannot find record"
            End If
            
            Adodc1.Refresh
        ElseIf cmboScreen.Text = "Two" Then
            'Locks onto the exact Seat ID with the SeatID field
            Adodc1.Recordset.Filter = "SeatID=" & getSeatID(cmboRow.Text, Val(cmboSeat.Text), 2)
            
            If Adodc1.Recordset.RecordCount <> 0 Then
                'Checks if seat is not already booked, to prevent double booking.
                If Adodc1.Recordset.Fields("Booked") = False Then
                    'Books the seat
                    Adodc1.Recordset.Fields("Booked") = True
                    Adodc1.Recordset.Update
                    
                    Adodc3.Recordset.AddNew
                    Adodc3.Recordset.Fields("ShowID") = Adodc2.Recordset.Fields("ShowID")
                    Adodc3.Recordset.Fields("MovieID") = Adodc2.Recordset.Fields("MovieID")
                    Adodc3.Recordset.Fields("SeatID") = Adodc1.Recordset.Fields("SeatID")
                    Adodc3.Recordset.Fields("ScreenID") = 2
                    Adodc3.Recordset.Fields("TransDate") = DateValue(Now)
                    Adodc3.Recordset.Update
                    Adodc3.Refresh
                    
                    MsgBox "Seat Booked!"
                Else
                    MsgBox "This seat is already booked."
                End If
            Else
                MsgBox "Cannot find record"
            End If
            
            Adodc1.Refresh
        End If
    End If
End Sub

'Function that calculates the Seat ID as the database classifies every seat by its ID
'This narrows down the Row and Seat No to the ID of that Row and Seat No
'It takes 3 arguments: row name, seat number & screen number
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

Private Sub cmdBooked_Click()
    'Filters records shown in other form to screen 1 only
    If cmboScreen.Text = "One" Then
        SeatsBooked.Show
        SeatsBooked.Adodc1.Recordset.Filter = "ScreenID='" & "1" & "' And Booked=" & True
    'Filters records shown in other form to screen 2 only
    ElseIf cmboScreen.Text = "Two" Then
        SeatsBooked.Show
        SeatsBooked.Adodc1.Recordset.Filter = "ScreenID='" & "2" & "' And Booked=" & True
    Else
        MsgBox "No screen selected"
    End If
End Sub

Private Sub cmdReserve_Click()
    Me.Hide
    ReserveSeat.Show
    Dim X As Integer
    X = Adodc2.Recordset.Fields("ShowID")
    'Passes previous form values to next form
    ReserveSeat.Adodc3.Recordset.Filter = "ShowID=" & X
    If cmboScreen.Text = "One" Then
        ReserveSeat.cmboScreen.Text = "One"
    ElseIf cmboScreen.Text = "Two" Then
        ReserveSeat.cmboScreen.Text = "Two"
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    BookShow.Show
End Sub

Private Sub Form_Load()
    'Establish a connection with the database
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Seats
    Adodc1.RecordSource = "Seats"
    Adodc1.Refresh
    
    'Establish a connection with the database
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Shows
    Adodc2.RecordSource = "Shows"
    Adodc2.Refresh
     
    'Establish a connection with the database
    Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Cinema.mdb;Persist Security Info=False"
    'Set the source table to be Transactions
    Adodc3.RecordSource = "Transactions"
    Adodc3.Refresh
    
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

Private Sub cmboRow_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Select one of the values given from the list of this box."
End Sub

Private Sub cmboSeat_KeyPress(Keyascii As Integer)
    Keyascii = 0
    MsgBox "Select one of the values given from the list of this box."
End Sub
