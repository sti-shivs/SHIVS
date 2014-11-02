VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShivsVoters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT wp_shivs_poll_voters.id, display_name, question, answer " & _
            "FROM wp_shivs_poll_voters " & _
            "INNER JOIN wp_shivs_polls " & _
            "ON wp_shivs_poll_voters.poll_id = wp_shivs_polls.id " & _
            "INNER JOIN wp_users " & _
            "ON wp_shivs_poll_voters.user_id = wp_users.ID " & _
            "INNER JOIN wp_shivs_poll_answers " & _
            "ON wp_shivs_poll_voters.poll_id = wp_shivs_poll_answers.poll_id"
    
    Record.Open query, Connect
    
    Dim ctr As Integer
    
    For ctr = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = ctr
            
            .Col = 1: .Text = Record!display_name
            .Col = 2: .Text = Record!question
            .Col = 3: .Text = Record!answer
        End With
        
        Record.MoveNext
    Next ctr
End Function
