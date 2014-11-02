VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShivsPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT question, name, question_group, start_date, " & _
            "end_date, status " & _
            "FROM wp_shivs_polls WHERE question_group IS NOT NULL"
    
    Record.Open query, Connect
    
    Dim ctr As Integer
    
    For ctr = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = ctr
            
            .Col = 1: .Text = Record!question
            .Col = 2: .Text = Record!question_group
            .Col = 3: .Text = Record!start_date
            .Col = 4: .Text = Record!end_date
            .Col = 5: .Text = Record!Status
            .Col = 6: .Text = Record!Name
        End With
        
        Record.MoveNext
    Next ctr
End Function
