Attribute VB_Name = "OpenDatabase"
Global cnn As New ADODB.Connection
Global cmd As New ADODB.Command
Global rst As New ADODB.Recordset

Public Sub Open_Database()
'# Opes a connection to the NorthWind Database                          #
'# The DataSource section MUST match the path to the database on        #
'# your own computer (i.e. Data Source= [Correct Path] & "\NWIND.MDB    #
Dim sFileName As String

sFileName = App.Path & "/" & "Monitor_data.mdb"

cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & " Data Source= " & sFileName & ";Persist Security Info=False"
cnn.Open

End Sub

Public Sub Close_Database()
    cnn.Close
End Sub
Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    '//
    '  Function waits for seconds given
    '\\
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds

    Do Until GetTickCount > EndTime
        DoEvents
        Loop
    End Function


