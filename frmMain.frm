VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Web Site End-2-End Times"
   ClientHeight    =   7755
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5535
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":0000
      TabIndex        =   1
      ToolTipText     =   "chart"
      Top             =   2160
      Width           =   11055
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph It !"
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtToDate 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Text            =   "12/31/02"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtFromDate 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Text            =   "12/01/02"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox cboCategories 
      Height          =   315
      ItemData        =   "frmMain.frx":24B8
      Left            =   7680
      List            =   "frmMain.frx":24BA
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMonData 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   12648447
      FixedCols       =   0
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      TextStyle       =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "    To Date:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblDataValue 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Point"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Web Site :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveGraph 
         Caption         =   "Save Graph"
      End
      Begin VB.Menu mnuPrintGraph 
         Caption         =   "Print Graph"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopyGraph 
         Caption         =   "Copy Data"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'//
'                           MS-Chart example
'
'  I use this to graph the web site performace for end-2-end times.  The data is loaded
'  into an Access database via FTP data transfer.  The database holds the Project ID (web site)
'  end-to-end seconds (time it takes to load a page) and availability %
'
'  * This example uses a 2 dimensional array to load the data for the chart
'  * Provides a date range for the data to graph (typical is one month)
'  * Uses resizing of controls to get a better picture for the graph (bigger screens)
'  * Shows how to get a specific datapoint on the graph
'  * How to load a combo box with data
'  * How to attach a recordset to a grid
'\\
Private Sub cboCategories_Click()

Dim chrtArray()

'//
'  Close the record set if open
'\\
If rst.State <> adStateClosed Then
    rst.Close
End If

'//
'  Given the "project id" from the drop down box and "From" and "TO" dates
'  Find all the monitoring data (date recorded, end to end times, availability)
'\\
Set cmd.ActiveConnection = cnn
cmd.CommandText = "SELECT MonitorData.TodayDate, MonitorData.EndToEnd, MonitorData.AvailPerc" _
    & " FROM MonitorData" _
    & " Where MonitorData.ProjID ='" & cboCategories.Text & "'" _
    & " And MonitorData.TodayDate >= #" & txtFromDate.Text & "#" _
    & " And MonitorData.TodayDate <= #" & txtToDate.Text & "#" _
    & " ORDER BY MonitorData.TodayDate"

rst.CursorLocation = adUseClient
rst.Open cmd, , adOpenStatic, adLockBatchOptimistic

If rst.RecordCount = 0 Then
    MsgBox "site does not have any availability data"
    Exit Sub
End If
'//
'  Give the record set to the gird for display
'\\
Set grdMonData.Recordset = rst
grdMonData.Refresh
'//
'                    C H A R T
'
' Dynamic 2-dimensional array to store series
' The first index (x) is the total number of series
' The second index will store the Date(1) and its end-to-end time(2).
' x-axis value in the 1st slot (i.e. chrtArray(x,1)
' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
'
'\\
ReDim chrtArray(1 To rst.RecordCount, 1 To 2)

MSChart1.ShowLegend = True
MSChart1.chartType = VtChChartType2dLine
'//
'  Chart Title centered on top
'\\
MSChart1.Title.Text = cboCategories.Text & " web site end-to-end time"
'//
'  Chart X and Y axis titles
'\\
MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "Date"
MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Seconds"
'//
'  Chart Foot note
'\\
MSChart1.FootnoteText = "some note"
'//
'  Load the array with data
'\\
For X = 1 To rst.RecordCount
    chrtArray(X, 1) = rst!TodayDate
    chrtArray(X, 2) = rst!EndToEnd
    
    rst.MoveNext
Next X
'//
'  Attach the array of data to MS-CHART
'  setup the column lable (the project id)
'\\
With MSChart1
    .ChartData = chrtArray
    .ColumnCount = 1
    .ColumnLabelCount = 1
    .Column = 1
    .ColumnLabel = cboCategories.Text
    .Refresh
End With

End Sub

Private Sub cboCategories_KeyPress(KeyAscii As Integer)
    '//
    '  don't let the user type anything into drop down box
    '\\
    KeyAscii = 0
End Sub

Private Sub cmdGraph_Click()
    '//
    '  after the user chooses the project id, go and graph it
    '\\
    Call cboCategories_Click
End Sub

Private Sub Form_Load()

    Dim cmdCategories As New ADODB.Command
    Dim rstCategories As New ADODB.Recordset

    Call Open_Database


'//
'  Select all the project id's (WEB site ids) and put in combo box
'\\
    Set cmdCategories.ActiveConnection = cnn
    cmdCategories.CommandText = "SELECT Projects.ProjID From Projects ORDER BY Projects.ProjID"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic

'//
'  Add project id names to drop down cbo
'\\
    Do While Not rstCategories.EOF
        cboCategories.AddItem rstCategories!ProjID
        rstCategories.MoveNext
    Loop

    Set cmdCategories = Nothing
    rstCategories.Close

'//
'  Set the first projec id as default id
'\\
    cboCategories.ListIndex = 0
'//
'  Setup the grd behavior
'\\
    grdMonData.Rows = 2
    grdMonData.Cols = 3
    grdMonData.FixedRows = 1
    grdMonData.TextMatrix(0, 0) = "Date"
    grdMonData.TextMatrix(0, 1) = "End-2-End Seconds"
    grdMonData.TextMatrix(0, 2) = "Avail %"

'grdMonData.FixedCols = 0
    grdMonData.ColWidth(1) = 2000
    grdMonData.ColWidth(2) = 2000

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '//
    '  Clean up by removing objects from memory
    '\\
    rst.Close
    Call Close_Database
    Set cmd = Nothing

End Sub
Private Sub Form_Resize()
    MSChart1.Width = Me.ScaleWidth - 200

    'grdMonData.ColWidth(0) = 960
    'grdMonData.ColWidth(1) = Me.ScaleWidth - 960 - 2025
    'grdMonData.ColWidth(2) = 2025
End Sub
Private Sub mnuCopyGraph_Click()
    '//
    '  This gives the data points only, you can paste that into
    '  an Excel sheet for manupolation of data points
    '\\
    MSChart1.EditCopy
End Sub

Private Sub mnuExit_Click()

    Unload Me
End Sub

Private Sub mnuPrintGraph_Click()
    '//
    '  Send the graph to your default printer
    '\\
    MSChart1.EditCopy
    Printer.Print " "
    Printer.PaintPicture Clipboard.GetData(), 0, 0
    Printer.EndDoc
End Sub

Private Sub mnuSaveGraph_Click()
    '//
    '  Common Dialog area for saving picture of the graph
    '\\
    On Error GoTo saverr
    Dim sFileName As String
  
    With CommonDialog1
        .Filter = "Pictures (*.bmp)|*.bmp"
        .DefaultExt = "bmp"
        .CancelError = True
        .ShowSave
        sFileName = .FileName
    End With
  
    If sFileName = "" Then Exit Sub
  
    MSChart1.EditCopy
    SavePicture Clipboard.GetData, sFileName
    Exit Sub
  
saverr:
  MsgBox Err.Description

End Sub

Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    '//
    '  Get the chart's selected point's coordinates
    '  Get the tool tip text to show the point
    '  Also, display it in the label
    '\\

    MSChart1.Row = DataPoint
    MSChart1.ToolTipText = "Point " & DataPoint & " " & MSChart1.Data
    'MsgBox MSChart1.Data
    lblDataValue.Caption = "Point " & DataPoint & " " & MSChart1.Data
    
End Sub


