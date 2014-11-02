Attribute VB_Name = "HeaderFor"
'*******************************************************
'* SHIVS Grid Headers
'* Author: Ephramar A. Telog
'* Created: February 24, 2014
'* Email: ephramar@outlook.com
'*
'* Copyright 2014
'*******************************************************

Public Function PositionsGrid()
    With Position.grdPositions
        .Cols = 7: .Rows = 1
        
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 5000
        .ColWidth(6) = 5000
        
        .Col = 1: .Text = "Position"
        .Col = 2: .Text = "Intended for"
        .Col = 3: .Text = "Start Date"
        .Col = 4: .Text = "End Date"
        .Col = 5: .Text = "Status"
        .Col = 6: .Text = "Description"
    End With
End Function

Public Function CandidatesGrid()
    With Candidate.grdCandidates
        .Cols = 5: .Rows = 1
        
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 2500
        .ColWidth(3) = 5000
        .ColWidth(4) = 4500
        
        .Col = 1: .Text = "Candidate"
        .Col = 2: .Text = "Position"
        .Col = 3: .Text = "Type"
        .Col = 4: .Text = "Partylist"
    End With
End Function

Public Function VoterGrid()
    With Voter.grdVoters
        .Cols = 4: .Rows = 1
        
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 2500
        .ColWidth(3) = 2500
        
        .Col = 1: .Text = "Student Name"
        .Col = 2: .Text = "Position"
        .Col = 3: .Text = "Candidate"
    End With
End Function

Public Function LogGrid()
    With Log.grdLogs
        .Cols = 12: .Rows = 1
        
        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 3000
        .ColWidth(3) = 2000
        .ColWidth(5) = 2000
        .ColWidth(7) = 4500
        .ColWidth(10) = 2000
        .ColWidth(11) = 3000
        
        .Col = 1: .Text = "Poll Name"
        .Col = 2: .Text = "Vote ID"
        .Col = 3: .Text = "Candidate"
        .Col = 4: .Text = "IP"
        .Col = 5: .Text = "Voter"
        .Col = 6: .Text = "User Type"
        .Col = 7: .Text = "HTTP Referer"
        .Col = 8: .Text = "Tracking ID"
        .Col = 9: .Text = "Other Answer"
        .Col = 10: .Text = "Host"
        .Col = 11: .Text = "Vote Date"
    End With
End Function
