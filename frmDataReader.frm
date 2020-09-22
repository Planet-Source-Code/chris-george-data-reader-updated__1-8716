VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDataReader 
   Caption         =   "DataReader"
   ClientHeight    =   4020
   ClientLeft      =   6780
   ClientTop       =   5865
   ClientWidth     =   2400
   Icon            =   "frmDataReader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   2400
   Begin VB.ComboBox cmbValue 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmDataReader.frx":0442
      Left            =   1080
      List            =   "frmDataReader.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         Picture         =   "frmDataReader.frx":045D
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         Picture         =   "frmDataReader.frx":070F
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         Picture         =   "frmDataReader.frx":09C1
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         Picture         =   "frmDataReader.frx":0C73
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSQL 
         Caption         =   "SQL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         ToolTipText     =   "Query the Database"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   315
         Left            =   360
         Picture         =   "frmDataReader.frx":0F25
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete the Current Record"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAddNew 
         Height          =   315
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDataReader.frx":1027
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add a New Record"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.Data Data1 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   480
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdOpenRecordset 
         Height          =   345
         Left            =   1800
         Picture         =   "frmDataReader.frx":1111
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Open a Database"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtFieldInfo 
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   315
      BorderStyle     =   0
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3765
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ComboBox cmbTableName 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblFields 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fields"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblTable 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Table"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Data Reader

'Created By:  Chris George

'Purpose:  To display fields, field types, and field values in a small window that will always stay on top


Public LastClicked As Long

Private Sub cmbTableName_Click()
Dim I As Integer
'hide the textbox and combo box
Me.txtValue.Visible = False
Me.cmbValue.Visible = False
'query the database for the table selected
Me.Data1.RecordSource = "Select * from [" + Me.cmbTableName.Text + "]"
Me.Data1.Refresh
'refresh the grid with the fields for that table
RefreshGrid
'get a record count
If Not Me.Data1.Recordset.EOF Then Data1.Recordset.MoveLast
Me.StatusBar1.Panels(1).Text = "Records: " & Me.Data1.Recordset.RecordCount
If Not Me.Data1.Recordset.EOF Then Data1.Recordset.MoveFirst


End Sub

Private Sub cmbValue_Click()
'the combo box becomes visible when the user clicks on a boolean field.
'this allows the user to simply click true or false instead of having to
'type it in, and therefor reduces the risk of errors.

If Not Me.Data1.Recordset.EOF Then
    'edit the database field and insert the value that was clicked
    Me.Data1.Recordset.Edit
    Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Value = cmbValue.Text
    Me.Data1.Recordset.Update
    'update the grid with the new value
    Me.MSFlexGrid1.Text = Me.cmbValue.Text
Else
    'if there are no records then tell the user
    MsgBox "No Current Record.", vbInformation + vbOKOnly
End If

Exit Sub

'error trap
Record_Error:
    MsgBox "Error: " + Error, vbExclamation + vbOKOnly
End Sub

Private Sub cmdAddNew_Click()
Dim I As Integer
'add a new record to the table
Me.Data1.Recordset.AddNew
'fill the grid with the fields in the table
For I = 0 To Me.Data1.Recordset.Fields.Count - 1
    If Me.Data1.Recordset.Fields(I).DefaultValue <> "" Then
        Me.Data1.Recordset.Fields(I).Value = Me.Data1.Recordset.Fields(I).DefaultValue
    End If
Next
'update the database
Me.Data1.Recordset.Update
Me.Data1.Recordset.MoveLast
'refresh the grid
RefreshGrid
End Sub

Private Sub cmdDelete_Click()
'delete the current record
If Not Me.Data1.Recordset.EOF Then
    Me.Data1.Recordset.Delete
    RefreshGrid
End If

End Sub

Private Sub cmdFirst_Click()
Me.cmbValue.Visible = False
Me.txtValue.Visible = False

Data1.Recordset.MoveFirst
RefreshGrid
End Sub

Private Sub cmdLast_Click()
Me.cmbValue.Visible = False
Me.txtValue.Visible = False

Data1.Recordset.MoveLast
RefreshGrid
End Sub

Private Sub cmdNext_Click()
Me.cmbValue.Visible = False
Me.txtValue.Visible = False

Data1.Recordset.MoveNext
If Not Data1.Recordset.EOF Then
    RefreshGrid
Else
    Data1.Recordset.MoveLast
    RefreshGrid
End If
End Sub

Private Sub cmdOpenRecordset_Click()
'opens the database and gets the tables and fields for the first table
On Error GoTo Cancel
Dim I As Integer
'default the open dialog to open mdb's
CommonDialog1.FileName = "*.mdb"
'show the open common dialog window
CommonDialog1.Action = 1
'set the data object to the database selected
Data1.DatabaseName = CommonDialog1.FileName
Data1.RecordSource = ""
Data1.Refresh

'get all of the table names and load them into the combo box
Me.cmbTableName.Clear
For I = 0 To Me.Data1.Database.TableDefs.Count - 1
    If InStr(1, Me.Data1.Database.TableDefs(I).Name, "MSys") = 0 Then
        Me.cmbTableName.AddItem Me.Data1.Database.TableDefs(I).Name
    End If
Next

'set the combo box to the first table
cmbTableName.ListIndex = 0

Cancel:

End Sub

Private Sub cmdPrevious_Click()
Me.cmbValue.Visible = False
Me.txtValue.Visible = False
Data1.Recordset.MovePrevious

If Not Data1.Recordset.BOF Then
    RefreshGrid
Else
    Data1.Recordset.MoveFirst
    RefreshGrid
End If

End Sub

Private Sub cmdSQL_Click()
'allows the user to query for records in the table
Dim SQL As String
On Error GoTo Cancel
'show the input box and get the string to use for the query.
'default the string to select all records from the current table.
SQL = InputBox("Enter the SQL Statement to use for the query.", "SQL", "Select * from [" & Me.cmbTableName.Text & "]")
'if the user doesn't enter anything then exit
If SQL = "Select * from" Or SQL = "" Then Exit Sub
On Error GoTo SQLError
'query the database with the string
Me.Data1.RecordSource = SQL
Me.Data1.Refresh
'if there are records then show a record count in the status bar
If Not Me.Data1.Recordset.EOF Then
    Me.Data1.Recordset.MoveLast
    Me.StatusBar1.Panels(1).Text = "Records:  " & Me.Data1.Recordset.RecordCount
    Me.Data1.Recordset.MoveFirst
End If

'refresh the grid
RefreshGrid
Exit Sub

'error trapping
SQLError:
MsgBox Error, vbExclamation + vbOKOnly

Cancel:

End Sub

Private Sub Data1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.cmbValue.Visible = False
Me.txtValue.Visible = False
'Data1.Refresh
RefreshGrid
End Sub


Private Sub Form_Load()
Dim lRetVal As Long
'set the window to stay on top.  Comment this out if you don't want it to stay on top
lRetVal = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub Form_Resize()
'resize all of the objects
Me.txtValue.Visible = False
Me.cmbValue.Visible = False

If Me.WindowState <> 1 Then
    Me.cmbTableName.Width = Me.Width - 350
    Me.MSFlexGrid1.Width = Me.Width - 350
    Me.MSFlexGrid1.ColWidth(0) = Me.MSFlexGrid1.Width / 2 - 125
    Me.MSFlexGrid1.ColWidth(1) = Me.MSFlexGrid1.Width / 2 - 125
    
    Me.Frame1.Width = Me.Width - 350
    Me.lblTable.Width = Me.Width - 350
    Me.lblFields.Width = Me.Width - 350
    Me.txtFieldInfo.Width = Me.Width - 350
    Me.cmdOpenRecordset.Left = Me.Width - 350 - Me.cmdOpenRecordset.Width
    Me.Data1.Width = Me.Width - Me.cmdOpenRecordset.Width - 350 - 60
    Me.Frame1.Top = Me.Height - Me.Frame1.Height - 700
    Me.MSFlexGrid1.Height = Me.Height - Frame1.Height - 1800
End If
End Sub


Private Sub MSFlexGrid1_Click()
On Error GoTo ClickError
cmbValue.Visible = False
txtValue.Visible = False

Dim FType As String     'stores the type of field selected

'get the type of field the user clicked on
Select Case Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Type
    Case 1
        FType = "True/False"
    Case 4
        FType = "Number"
    Case 5
        FType = "Currency"
    Case 8
        FType = "Date/Time"
    Case 11
        FType = "OLE Object"
    Case 10
        FType = "Text"
    Case 12
        FType = "Memo"
    Case Else
        FType = "Unknown"
End Select

'display the field name and type in the textbox below the grid
Me.txtFieldInfo.Text = Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Name + "(" + FType + "):"
'get the properties for that field
For I = 1 To Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Properties.Count - 1
    On Error Resume Next
    Me.txtFieldInfo.Text = Me.txtFieldInfo.Text + vbCrLf + Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Properties(I).Name + ": " & Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Properties(I).Value
Next

'if it is a true/false field then show the combo box
If Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Type = 1 Then
    Me.MSFlexGrid1.Col = 1
    Me.cmbValue.Left = Me.MSFlexGrid1.CellLeft + Me.MSFlexGrid1.Left
    Me.cmbValue.Width = Me.MSFlexGrid1.CellWidth
    Me.cmbValue.Top = Me.MSFlexGrid1.CellTop + Me.MSFlexGrid1.Top
    Me.cmbValue.Text = Me.MSFlexGrid1.Text
    Me.cmbValue.Visible = True
    Me.txtValue.Visible = False
Else
    'otherwise show the textbox
    Me.MSFlexGrid1.Col = 1
    Me.txtValue.Left = Me.MSFlexGrid1.CellLeft + Me.MSFlexGrid1.Left
    Me.txtValue.Width = Me.MSFlexGrid1.CellWidth
    Me.txtValue.Top = Me.MSFlexGrid1.CellTop + Me.MSFlexGrid1.Top
    Me.txtValue.Text = Me.MSFlexGrid1.Text
    Me.txtValue.Height = Me.MSFlexGrid1.CellHeight
    Me.cmbValue.Visible = False
    Me.txtValue.Visible = True
End If

Exit Sub

ClickError:
    MsgBox "Database Not Selected.", vbInformation + vbOKOnly
    
End Sub

Private Sub MSFlexGrid1_Scroll()
'hide the combo box and textbox
Me.cmbValue.Visible = False
Me.txtValue.Visible = False
End Sub

Private Sub txtValue_GotFocus()
'if the user clicks on the textbox then set lastclicked to the row so you know
'what field to update when the user is done
LastClicked = Me.MSFlexGrid1.Row
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
'if the user hits return then update the field with the new value
If KeyCode = vbKeyReturn Then
    On Error GoTo Record_Error
    Me.Data1.Recordset.Edit
    Me.Data1.Recordset.Fields(Me.MSFlexGrid1.Row - 1).Value = txtValue.Text
    Me.Data1.Recordset.Update
    Me.MSFlexGrid1.Text = Me.txtValue.Text
End If

Exit Sub

'error trap
Record_Error:
MsgBox "Error: " + Error, vbExclamation + vbOKOnly
End Sub

Private Sub txtValue_LostFocus()
'if the user clicks off of the field then update the field with the value currently in the
'textbox
On Error GoTo Record_Error
Me.Data1.Recordset.Edit
Me.Data1.Recordset.Fields(LastClicked - 1).Value = txtValue.Text
Me.Data1.Recordset.Update
Me.MSFlexGrid1.TextMatrix(LastClicked, 1) = Me.txtValue.Text

Exit Sub

Record_Error:
MsgBox "Error: " + Error, vbExclamation + vbOKOnly

End Sub
Public Sub RefreshGrid()
'fills the grid with the fields and values in the fields
Dim I As Integer
On Error Resume Next

'clear the grid
Me.MSFlexGrid1.Clear
'set up the grid to the number of fields in the table
Me.MSFlexGrid1.Rows = Me.Data1.Recordset.Fields.Count + 1
'set the first row to show column headers
Me.MSFlexGrid1.TextMatrix(0, 0) = "Name"
Me.MSFlexGrid1.TextMatrix(0, 1) = "Value"
'loop through the fields and add them to the grid
For I = 0 To Me.Data1.Recordset.Fields.Count - 1
    Me.MSFlexGrid1.TextMatrix(I + 1, 0) = Me.Data1.Recordset.Fields(I).Name
    If Not Me.Data1.Recordset.EOF And Me.Data1.Recordset.RecordCount > 0 Then
        'depending on the field type fill the textmatrix with the data
        Select Case Me.Data1.Recordset.Fields(I).Type
            Case 1
                FType = "True/False"
                     Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
            Case 4
                FType = "Number"
                     Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
            Case 5
                FType = "Currency"
                Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
            Case 8
                FType = "Date/Time"
                Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
            Case 11
                FType = "OLE Object"
            Case 10
                FType = "Text"
                If Me.Data1.Recordset.Fields(I).Value <> "" Then
                    Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
                Else
                    Me.MSFlexGrid1.TextMatrix(I + 1, 1) = ""
                End If
            
            Case 12
                FType = "Memo"
                If Me.Data1.Recordset.Fields(I).Value <> "" Then
                    Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
                Else
                    Me.MSFlexGrid1.TextMatrix(I + 1, 1) = ""
                End If
            Case Else
                FType = "Unknown"
                If Me.Data1.Recordset.Fields(I).Value <> "" Then
                    Me.MSFlexGrid1.TextMatrix(I + 1, 1) = Me.Data1.Recordset.Fields(I).Value
                Else
                    Me.MSFlexGrid1.TextMatrix(I + 1, 1) = ""
                End If
        End Select
    End If
        
Next

'set the alignment of the grid to left center
Me.MSFlexGrid1.ColAlignment(1) = flexAlignLeftCenter
'show the number of records in the table
Me.StatusBar1.Panels(1).Text = "Records: " & Me.Data1.Recordset.RecordCount
End Sub
