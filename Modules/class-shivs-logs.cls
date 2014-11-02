VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShivsLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
            
    query = "SELECT wp_shivs_poll_logs.id, wp_shivs_polls.name, " & _
            "vote_id, wp_shivs_poll_answers.answer, ip, " & _
            "wp_users.display_name, user_type, http_referer, " & _
            "tr_id, other_answer_value, host, vote_date " & _
            "FROM wp_shivs_poll_logs " & _
            "INNER JOIN wp_shivs_polls ON wp_shivs_poll_logs.poll_id = wp_shivs_polls.id " & _
            "INNER JOIN wp_shivs_poll_answers ON wp_shivs_poll_logs.answer_id = wp_shivs_poll_answers.id " & _
            "INNER JOIN wp_users ON wp_shivs_poll_logs.user_id = wp_users.id"
    
    Record.Open query, Connect
    
    Dim ctr As Integer
    
    For ctr = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = ctr
            
            .Col = 1: .Text = Record!Name
            .Col = 2: .Text = Record!vote_id
            .Col = 3: .Text = Record!answer
            .Col = 4: .Text = Record!ip
            .Col = 5: .Text = Record!display_name
            .Col = 6: .Text = Record!user_type
            .Col = 7: .Text = Record!http_referer
            .Col = 8: .Text = Record!tr_id
            .Col = 9: .Text = Record!other_answer_value
            .Col = 10: .Text = Record!host
            .Col = 11: .Text = Record!vote_date
        End With
        
        Record.MoveNext
    Next ctr
End Function
