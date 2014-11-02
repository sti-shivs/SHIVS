Attribute VB_Name = "Global"
'*******************************************************
'* SHIVS Core Version 3.0
'* Author: Hot Swap LLC
'* Created: February 6, 2014
'* Email: ephramar@outlook.com
'*
'* Copyright 2014
'*******************************************************

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'ADODB Variables
Public Connect As ADODB.Connection
Public Record As ADODB.Recordset

'MySQL prefix
Public wpdb As String

'SQL String
Public query As String

'Poll is just a prefix, not a part of the program
Public PollCandidate As ShivsCandidates
Public PollResult As ShivsLogs
Public PollVoter As ShivsVoters
Public PollLanding As ShivsLanding
Public PollPosition As ShivsPositions
Public MD5 As TsunaMD5

Sub Main()
    Dim ConnectionString As String
    wpdb = wp_
    
    Set Connect = New ADODB.Connection
    Set Record = New ADODB.Recordset
    
    Set PollCandidate = New ShivsCandidates
    Set PollResult = New ShivsLogs
    Set PollVoter = New ShivsVoters
    Set PollLanding = New ShivsLanding
    Set PollPosition = New ShivsPositions
    Set MD5 = New TsunaMD5
    
    With Record
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    
    ConnectionString = "Driver={Mysql ODBC 3.51 Driver};" & _
                       "Server=localhost;" & _
                       "Port=3306;" & _
                       "Database=shivs;" & _
                       "User=root;" & _
                       "Password=2n329fdx;" & _
                       "Option=3;"
                       
    Connect.Open ConnectionString
    
    Shivs_Init
End Sub

Public Function Shivs_Init() As Boolean
    Dim Username As String, Password As String
    Dim Fail As Boolean, Successful As Boolean
    
    Randomize
    
    Username = GetSetting(App.EXEName, "Settings", "LastUser", "")
    
    Fail = Security.GetUserInfo(Username, Password, Index)
    
    Do While Fail
        If Record.State = 1 Then Record.Close
        
        query = "SELECT ID, user_login, user_pass FROM wp_users WHERE user_login = '" & Replace(Username, "'", "''") & "'"
        
        Record.Open query, Connect
        
        If Record.RecordCount = 0 Then GoTo Bye
        
        If LCase(MD5.DigestStrToHexStr(Password)) = Record!user_pass Then
            CurrentUserID = Record!ID
            CurrentUserName = Record!user_login
            
            SaveSetting App.EXEName, "Setting", "LastUser", CurrentUserName
            
            Successful = True
            
            Landing.Show
            
            Exit Do
        End If
        
Bye:
        If Not Successful Then
            Fail = False
            
            If MsgBox("Invalid username or password" & vbCrLf & "Do you want to try again?", vbQuestion + vbYesNo, Security.Caption) = vbYes Then
                Sleep 200 + 300 * Rnd
                Fail = Security.GetUserInfo(Username, Password, Index)
            End If
        End If
    Loop
    
    Shivs_Init = Successful
End Function



