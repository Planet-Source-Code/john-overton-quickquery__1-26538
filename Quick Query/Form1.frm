VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Quick Query"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add SQL"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame frDB 
      Caption         =   "SQL Server Database"
      Height          =   615
      Left            =   8640
      TabIndex        =   13
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox cboDB 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data"
      Height          =   1935
      Left            =   240
      TabIndex        =   9
      Top             =   6000
      Width           =   11415
      Begin MSDataGridLib.DataGrid dg 
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2778
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
   Begin VB.Frame Frame4 
      Caption         =   "Database Type"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
      Begin VB.ComboBox cboDBType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   11160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tables/Fields"
      Height          =   4695
      Left            =   8640
      TabIndex        =   4
      Top             =   960
      Width           =   3015
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7646
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.TextBox text1 
      Height          =   2415
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1320
      Width           =   5655
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H80000012&
      Height          =   4350
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "SQL"
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "SQL Statment"
      Height          =   3255
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdexecute 
         Caption         =   "Execute"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'John Overton
'overtonjohn@yahoo.com


Dim rs As ADODB.Recordset
Dim rss As ADODB.Recordset
Dim rsExecute As ADODB.Recordset
Dim nodx As Node
Dim nody As Node
Dim nodz As Node
Dim Selected As String
Dim LastSQL
Dim cnt
Dim cnt1
Sub Read_Files()
Dim LinesFromFile, NextLine As String

Open App.Path & "\TEST.txt" For Input As #1

'filenum = App.Path & "\TEST.txt"
Do Until EOF(1)
   Line Input #1, NextLine
    List1.AddItem Trim(NextLine)
Loop
Close #1

End Sub



Private Sub cboDB_Click()
         cmdexecute.Enabled = True
    cmdClear.Enabled = True
    cmdSave.Enabled = True
Set nody = TreeView1.Nodes.Add(, 1, "Database", cboDB.Text)
   nody.EnsureVisible
    nody.Expanded = True
   DB = cboDB.Text
End Sub

Private Sub cboDBType_Click()

On Error GoTo errhandler
   Dim cnt
   Dim cnt1
    
    If cboDBType.ListIndex = 0 Then
        DBType = 1
        Dim strCheckForDatabase As String

    
    
     dlgCommon.DialogTitle = "Pick A Database"
    'Give the file selection window a title.
    
    dlgCommon.InitDir = App.Path
    'The file selection window will start in the
    'applications directory.
    
    'Allow the user to view only Access files.
    dlgCommon.Filter = "Access Databases (*.mdb)|*.mdb|" & _
                       "All Files (*.*)|*.*"
    dlgCommon.ShowOpen
    'Open the file selection window.
    
    strCheckForDatabase = Right(dlgCommon.FileName, 4)
    'Select the last four letters of the file selected.
    
    Select Case strCheckForDatabase
       Case vbNullString
            'Do not allow empty strings.
            Exit Sub
            
        Case ".mdb"
             'Assign the chosen file to the path string.
            DatabasePath = dlgCommon.FileName
             'Do not allow the user to select another DB until
            'clear is clicked.             cmdOpenDB.Enabled = False
             
    End Select
        
        Open_cn
        LoadAccess
    Else
    
        DBType = 2
        Form2.Show (vbModal)
        If ServerName = "" Or IsNull(ServerName) = True Then
            Exit Sub
        End If
        
        LoadSQL
    End If

    cboDBType.Enabled = False
  Read_Files
errhandler:
    Exit Sub
    
End Sub

Private Sub cmdAdd_Click()
Form3.Show (vbModal)

     Open App.Path & "\test.txt" For Append As #1  ' Open/Create file
                                         '      for output.
    Print #1, AddSQL
    
    Close #1
    List1.Clear
    Read_Files
End Sub

Private Sub cmdClear_Click()
    Text1.Text = ""
    Set dg.DataSource = Nothing
    
End Sub

Private Sub cmdExecute_Click()

    On Error GoTo errhandler
    Open_cn
    Set rsExecute = New ADODB.Recordset
    rsExecute.Open Text1.Text, cn, adOpenStatic, adLockOptimistic, _
        adCmdText
    Set dg.DataSource = rsExecute
    Clipboard.Clear
    Clipboard.SetText (Text1.Text)
    Exit Sub
errhandler:
   MsgBox (Err.Description)
    
End Sub

Private Sub cmdSave_Click()
 Create_File
    Call MsgBox("Query is saved to " & App.Path & "\query.txt", vbOKOnly, "Query")
End Sub
Sub Create_File()
    Open App.Path & "\query.txt" For Append As #1  ' Open/Create file
    Print #1, ""                                         '      for output.
    Print #1, Text1.Text
    
    Close #1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()


   
    cboDB.Visible = False
    frDB.Visible = False
    cboDBType.AddItem ("MS Access")
    cboDBType.AddItem ("SQL Server")
    Clipboard.Clear
    cmdexecute.Enabled = False
    cmdClear.Enabled = False
    cmdSave.Enabled = False
    
    
End Sub
Sub LoadAccess()
Dim cnt
Dim cnt1
         cmdexecute.Enabled = True
    cmdClear.Enabled = True
    cmdSave.Enabled = True
    Set rs = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
         Set nodx = TreeView1.Nodes.Add(, , DatabasePath, "MS Access")
 
        cnt1 = 1
        
        Do While Not rs.EOF
        cnt = "c" & cnt1
        Set nodx = TreeView1.Nodes.Add(DatabasePath, tvwChild, "Table" & cnt, rs.Fields("table_name"))
     
        cnt1 = cnt1 + 1
        
Dim str1
Dim i
Dim jnt
Dim jnt1
Set rss = cn.OpenSchema(adSchemaForeignKeys)
       ' setKeysAC (rs.RecordCount)
        ' str1 = "select * from [" & Node.Text & "]"
        Set rss = cn.Execute("Select * from [" & rs.Fields("table_name") & "]")
        For i = 0 To rss.Fields.Count - 1
            jnt = "c" & jnt1
            jnt1 = jnt1 + 1
              Set nody = TreeView1.Nodes.Add(nodx, 4, "Column" & jnt, rss.Fields(i).Name)
       Next i
       rs.MoveNext
        Loop
         nodx.EnsureVisible
       
        
    


End Sub

Private Sub List1_Click()



    If List1.ListIndex > 1 And List1.ListIndex <= 11 Then
        Text1.Text = Trim(Text1.Text & vbCrLf & " " & List1.Text)
    Else
        Text1.Text = Text1.Text & " " & List1.Text
    End If
     LastSQL = Trim(List1.Text)
End Sub

Private Sub TreeView1_DblClick()



If Left(LastSQL, 6) = "SELECT" Then
    If Left(Selected, 5) = "Table" Then
       Call MsgBox("Must be a field", vbOKOnly, "Field Needed")
        Exit Sub
    End If
End If
If Left(LastSQL, 4) = "FROM" Then
    If Left(Selected, 1) = "C" Then
       Call MsgBox("must be a TABLE", vbOKOnly, "Table Needed")
        Exit Sub
    End If
End If
If Left(LastSQL, 5) = "WHERE" Then
    If Left(Selected, 1) = "T" Then
       Call MsgBox("must be a Field", vbOKOnly, "Field Needed")
        Exit Sub
    End If
End If

    If Left(Selected, 1) = "D" Then
        Exit Sub
    End If
    
    If Right(Text1.Text, 1) = "]" Then
        Text1.Text = Text1.Text & ", [" & TreeView1.SelectedItem.Parent & "].[" & TreeView1.SelectedItem & "]"
    Else
        If Left(Selected, 1) = "T" Then
            Text1.Text = Text1.Text & " [" & TreeView1.SelectedItem & "]"
        Else
         Text1.Text = Text1.Text & " [" & TreeView1.SelectedItem.Parent & "].[" & TreeView1.SelectedItem & "]"
        End If
    End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

On Error Resume Next
 Dim jnt
 Dim jnt1
Selected = Node.Key
If DBType = 1 Then
    Open_cn
   
   
End If

If DBType = 2 Then

    If Left(Node.Key, 8) = "Database" Then
    DB = Node.Text
    Open_cn
       Set rss = cn.OpenSchema(adSchemaTables)
            Do Until rss.EOF
                  cnt1 = cnt1 + 1
            cnt = "c" & cnt1
                If UCase(Left(rss!table_name, 4)) <> "MSYS" Then
                    Set nody = TreeView1.Nodes.Add(Node.Index, 4, "Table" & cnt & Node.Index, rss.Fields!table_name)
                End If
                rss.MoveNext
            Loop
           
        nodx.EnsureVisible
        nodx.Expanded = True
         
    End If
    
    If Left(Node.Key, 5) = "Table" Then
     Open_cn
        
       ' txtSQL.Text = "Select * from [" & Node.Text & "]"
        Set rss = New ADODB.Recordset
        Set rss = cn.Execute("sp_columns [" & TreeView1.SelectedItem & "]")
         Do While Not rss.EOF
         cnt1 = cnt1 + 1
         cnt = "cs" & cnt1
            Set nodz = TreeView1.Nodes.Add(Node.Index, 4, "Column" & cnt & Node.Index, rss!column_name)
            rss.MoveNext
       Loop
       
       
    End If
    nodx.EnsureVisible
        nodx.Expanded = True
        
End If
End Sub


Sub setKeysAC(recCnt As Long)

        ReDim PKTab(recCnt) As String
        ReDim PKCol(recCnt) As String
        
        ReDim FKTab(recCnt) As String
        ReDim FKCol(recCnt) As String
        
        incr = 0
        If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            PKTab(incr) = rs.Fields(2).Value
            PKCol(incr) = rs.Fields(3).Value
            
            FKTab(incr) = rs.Fields(8).Value
            FKCol(incr) = rs.Fields(9).Value
        incr = incr + 1
        rs.MoveNext
                   
        Loop
        End If

End Sub
Sub LoadSQL()
    cboDB.Visible = True
    frDB.Visible = True
    Open_cn
        Set rs = New ADODB.Recordset
        Set rs = cn.Execute("sp_databases")


Set rs = cn.OpenSchema(adSchemaCatalogs, Array(CATALOG_NAME))
                        
    
                        Do While Not rs.EOF
                  
                            cboDB.AddItem (rs.Fields(0))
                      
                            rs.MoveNext
                        
                        Loop
                   Close_cn
    
End Sub


