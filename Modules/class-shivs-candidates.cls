VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShivsCandidates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayCandidates(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT answer, question, question_group, wp_bp_groups.name " & _
            "FROM wp_shivs_poll_answers " & _
            "INNER JOIN wp_shivs_polls ON wp_shivs_poll_answers.poll_id = wp_shivs_polls.id " & _
            "INNER JOIN wp_bp_groups ON wp_shivs_poll_answers.group_id = wp_bp_groups.id " & _
            "WHERE question_group IS NOT NULL"
    
    Record.Open query, Connect
    
    Dim ctr As Integer
    
    For ctr = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = ctr
            
            .Col = 1: .Text = Record!answer
            .Col = 2: .Text = Record!question
            .Col = 3: .Text = Record!question_group
            .Col = 4: .Text = Record!Name
        End With
        Record.MoveNext
    Next ctr
End Function

Public Function DisplayElectoralParty(Control As ComboBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT name FROM wp_bp_groups WHERE parent_id = 19"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!Name
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function

Public Function DisplayElectoralPartyMembers(GroupID As Integer, Control As ComboBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT display_name " & _
          "FROM wp_bp_groups_members " & _
          "INNER JOIN wp_users ON wp_bp_groups_members.user_id = wp_users.ID " & _
          "WHERE group_id = '" & GroupID & "' " & _
          "ORDER BY display_name ASC"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!display_name
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function

Public Function DisplayElectoralPartyMembersIDNumber(UserID As Integer, Control As Label)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT value FROM wp_bp_xprofile_data WHERE field_id = 2 AND user_id = '" & UserID & "'"
    
    Record.Open query, Connect
    
    If Not Record.EOF Then
        Control.ForeColor = vbBlack
        Control.Caption = Record!value
    Else
        Control.ForeColor = vbRed
        Control.Caption = "Missing"
    End If
End Function

Public Function DisplayElectoralPartyMembersCourse(UserID As Integer, Control As Label)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT value FROM wp_bp_xprofile_data WHERE field_id = 8 AND user_id = '" & UserID & "'"
    
    Record.Open query, Connect
    
    If Not Record.EOF Then
        Control.ForeColor = vbBlack
        Control.Caption = Record!value
    Else
        Control.ForeColor = vbRed
        Control.Caption = "Missing"
    End If
End Function

Public Function DisplayElectoralPartyMembersYearLvl(UserID As Integer, Control As Label)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT value FROM wp_bp_xprofile_data WHERE field_id = 16 AND user_id = '" & UserID & "'"
    
    Record.Open query, Connect
    
    If Not Record.EOF Then
        Control.ForeColor = vbBlack
        Control.Caption = Record!value
    Else
        Control.ForeColor = vbRed
        Control.Caption = "Missing"
    End If
End Function

Public Function DisplayElectoralPartyMembersGender(UserID As Integer, Control As Label)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT value FROM wp_bp_xprofile_data WHERE field_id = 28 AND user_id = '" & UserID & "'"
    
    Record.Open query, Connect
    
    If Not Record.EOF Then
        Control.ForeColor = vbBlack
        Control.Caption = Record!value
    Else
        Control.ForeColor = vbRed
        Control.Caption = "Missing"
    End If
End Function

Public Function DisplayElectoralPartyMembersBirthday(UserID As Integer, Control As Label)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT value FROM wp_bp_xprofile_data WHERE field_id = 57 AND user_id = '" & UserID & "'"
    
    Record.Open query, Connect
    
    If Not Record.EOF Then
        Control.ForeColor = vbBlack
        Control.Caption = Format(Record!value, "mm" & "-" & "dd" & "-" & "yyyy")
    Else
        Control.ForeColor = vbRed
        Control.Caption = "Missing"
    End If
End Function

'*** This function is deprecated ***'
Public Function DisplayMembersAvatar(Control As Image)
    Control.Picture = LoadPicture("C:\xampp\htdocs\shivs\wp-content\uploads\avatars\19\2698ddb1eabb84fa6ebf12ac0b409b72-bpfull.jpg")
End Function
