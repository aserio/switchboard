Attribute VB_Name = "Switchboard"
'Copyright (c) 2020-2021 Adrian S. Lemoine
'
'Distributed under the Boost Software License, Version 1.0.
'(See accompanying file LICENSE_1_0.txt or copy at
'http://www.boost.org/LICENSE_1_0.txt)

Option Explicit

Sub Switchboard()
'' This subroutine is able to read in tasks from a file and update the Project tasks
''
'' Asumptions:
''    Board status have been added which match Python Output (Not Started, In progress, Under Review, Done)
''    Configuration file named config.txt is present in working directory
''    User will provide a match pattern to retreive the sprint ID number from a GitHub ID name
''      - Can use "default" to just find the number automatically
''      - Can pass and empty string ("") to skip the sprint assignment step
''      - Can pass a regular expression which will return the sprint ID number
    Dim ProjectName As String
    Dim ProjectPfx As String
    Dim PathName As String
    Dim CsvFilePath As String
    Dim ConfigFile As String
    Dim PythonPath As String
    Dim GitHubCordPath As String
    Dim GitHubRepo As String
    Dim SprintLength As String
    Dim SprintPattern As String
    Dim ProjectFieldDur As Long
    Dim ProjectFieldBS As Long
    Dim i As Integer
    Dim LineFromFile As String
    Dim LineItems() As String
    Dim gid As Integer
    Dim NewTask As Task
    
    ProjectName = GetUNCPath(Application.ActiveProject.FullName)
    ProjectPfx = Replace(StripFileName(ProjectName), ".mpp", "")
    PathName = GetUNCPath(Application.ActiveProject.path)
    CsvFilePath = PathName + "\" + ProjectPfx + ".csv"
        
    ConfigFile = "config.txt"
    ''' Fetch configuration information
    Open PathName + "\" + ConfigFile For Input As #1
    Line Input #1, LineFromFile ' Ignore header
    If LineFromFile = "[System Information]" Then
      ' Extract Python Path
      Line Input #1, LineFromFile
      PythonPath = extract_path(LineFromFile)
      ' Extract path to Switchboard Python file
      Line Input #1, LineFromFile
      GitHubCordPath = extract_path(LineFromFile, "github_cord.py")
      
    Else
      Err.Raise vbObjectError + 513, "Switchboard Module", _
        "Configuration file is not set up correctly."
    End If
    
    Line Input #1, LineFromFile ' Ignore paragraph break
    
    Line Input #1, LineFromFile
      If LineFromFile = "[Project Information]" Then
        Line Input #1, LineFromFile 'Ignore header
      Else
        Err.Raise vbObjectError + 513, "Switchboard Module", _
          "Configuration file is not set up correctly."
      End If
    
    ' Fetch project configuration
    Do Until EOF(1)
      Line Input #1, LineFromFile
      LineItems = parse_line(LineFromFile)
      If LineItems(0) = ProjectPfx Then
        GitHubRepo = LineItems(1)
        SprintLength = LineItems(2)
        SprintPattern = LineItems(3)
      End If
    Loop
    
    ' Clean up
    Close #1
    LineFromFile = ""
   
   ''' Call Python script to fetch GitHub issues
    Dim wshell As Object
    Dim Args As String
    Dim error_code As Double
    
    Set wshell = CreateObject("WScript.Shell")
    
    Args = "--github_repo " & Chr(34) & GitHubRepo & Chr(34) & _
      " --csv_file " & Chr(34) & CsvFilePath & Chr(34) & " --sprint_length " & SprintLength
    'MsgBox (PythonPath & " " & GitHubCordPath & " " & Args)

    error_code = wshell.Run(PythonPath & " " & GitHubCordPath & " " & Args, 1, True)
    
    Open CsvFilePath For Input As #2
      
    ''' Get field IDs
    ProjectFieldDur = FieldNameToFieldConstant("Duration", pjProject)
    ProjectFieldBS = FieldNameToFieldConstant("Board Status", pjProject)
    
    ''' Create dictionary relating task ID to GitHub ID
    Dim unique_id As Long
    Dim git_id As Long
    Dim dict
    Dim ii As Integer

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To ActiveProject.Tasks.count
      If Application.ActiveProject.Tasks(i).Text2 = vbNullString Then
      Else
        unique_id = Application.ActiveProject.Tasks(i).UniqueID
        git_id = Application.ActiveProject.Tasks(i).Text2
        dict.Add git_id, unique_id
      End If
    Next i
    
    ' Skip header row
    Line Input #2, LineFromFile
    
    ''' Update and add tasks with information from CSV file
    Do Until EOF(2)
      Line Input #2, LineFromFile
      LineItems = parse_line(LineFromFile)
      If dict.exists(CInt(LineItems(7))) Then
        gid = dict(CInt(LineItems(7)))
        Application.ActiveProject.Tasks.UniqueID(gid).Name = LineItems(0)
        Application.ActiveProject.Tasks.UniqueID(gid).SetField FieldID:=ProjectFieldDur, Value:=LineItems(2)
        'Only set start date. Rely on duration to calculate finsih date
        Application.ActiveProject.Tasks.UniqueID(gid).Start = LineItems(3)
        ' If a pattern has been supplied set the sprint
        If SprintPattern = "" Then
        Else
          Application.ActiveProject.Tasks.UniqueID(gid).Sprint = set_sprint(LineItems(5), SprintPattern)
        End If
        Application.ActiveProject.Tasks.UniqueID(gid).SetField FieldID:=ProjectFieldBS, Value:=LineItems(6)
        'Add Labels
        Application.ActiveProject.Tasks.UniqueID(gid).Text6 = LineItems(8)
        Application.ActiveProject.Tasks.UniqueID(gid).Text8 = LineItems(9)
        Application.ActiveProject.Tasks.UniqueID(gid).Text9 = LineItems(5)
        ' Set Percent Complete last to stop Project from overwriting
        If LineItems(1) = vbNullString Then
          Application.ActiveProject.Tasks.UniqueID(gid).PercentComplete = 1
        Else
          Application.ActiveProject.Tasks.UniqueID(gid).PercentComplete = CInt(LineItems(1))
        End If
      Else
        Set NewTask = Application.ActiveProject.Tasks.Add(LineItems(0))
        NewTask.SetField FieldID:=ProjectFieldDur, Value:=LineItems(2)
        'Only set start date. Rely on duration to calculate finsih date
        NewTask.Start = LineItems(3)
        ' If a pattern has been supplied set the sprint
        If SprintPattern = "" Then
        Else
          NewTask.Sprint = set_sprint(LineItems(5), SprintPattern)
        End If
        NewTask.SetField FieldID:=ProjectFieldBS, Value:=LineItems(6)
        'Add GitHub Issue Number
        NewTask.Text2 = LineItems(7)
        'Add Labels
        NewTask.Text6 = LineItems(8)
        NewTask.Text8 = LineItems(9)
        NewTask.Text9 = LineItems(5)
        ' Set Percent Complete last to stop Project from overwriting
        If LineItems(1) = vbNullString Then
           NewTask.PercentComplete = 1
        Else
           NewTask.PercentComplete = CInt(LineItems(1))
        End If
        
      End If

    Loop
        
    Close #2
    
    '' Recalculate Project
    Application.CalculateProject
    
    '' Remove CSV file
    Kill (CsvFilePath)

End Sub

Function parse_line(str As String) As String()
Dim RegEx As Object
Dim pattern
Dim str_array() As String
Dim i As Integer

' Find only the commas outside of the quotes
pattern = ",(?=([^" & Chr(34) & "]*" & Chr(34) & "[^" & Chr(34) & "]*" & Chr(34) & ")*(?![^" & Chr(34) & "]*" & Chr(34) & "))"
Set RegEx = CreateObject("vbscript.regexp")
RegEx.Global = True
RegEx.pattern = pattern
str_array = Split(RegEx.Replace(str, ";"), ";")

' Remove leading whitespace
pattern = "^\s+"
RegEx.pattern = pattern
For i = LBound(str_array) To UBound(str_array)
  str_array(i) = RegEx.Replace(str_array(i), "")
Next i

' Remove outer quotes
pattern = "^" & Chr(34) & "|" & Chr(34) & "$"
RegEx.pattern = pattern
For i = LBound(str_array) To UBound(str_array)
  str_array(i) = RegEx.Replace(str_array(i), "")
Next i

parse_line = str_array

End Function

Function extract_path(str As String, Optional fname As String = "") As String
Dim RegEx As Object
Dim pattern

' Find "Python Path:" or "Switchboard Path:" and remove from string
pattern = ".*Path:\s*"
Set RegEx = CreateObject("vbscript.regexp")
RegEx.Global = True
RegEx.pattern = pattern
str = RegEx.Replace(str, "")

' Set up quotes and optionally add string
'' Remove current quotes
pattern = """"
RegEx.pattern = pattern
str = RegEx.Replace(str, "")

'' Add string
If fname = "" Then
Else
  str = str + "\" + fname
End If

'' Remove double \
pattern = "\\\\"
RegEx.pattern = pattern
If RegEx.Test(str) Then
  str = RegEx.Replace(str, "\")
End If

'' Add quotes
''' Chr(34) is the double quotes character
str = Chr(34) & str & Chr(34)
extract_path = str

End Function

Function StripFileName(path As String) As String
  Dim PathElems As Variant
  
  PathElems = Split(path, "\")
  StripFileName = PathElems(UBound(PathElems))

End Function

Function GetUNCPath(path As String) As String
  Dim CurrentDrive As String
  Dim network As Object
  Dim drives As Object
  Dim el As Variant
  Dim RegEx As Object
  Dim RegEx2 As Object
  Dim pattern As String
  Dim str As String
  Dim NewPath As String
  Dim PathDirs As Variant
  Dim NewPathDirs As Variant
  Dim i As Integer
  Dim ii As Integer
  
  ''' Skip function if the string is not a path
  pattern = "[:\\/]"
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Global = True
  RegEx.pattern = pattern
  If Not RegEx.Test(path) Then
    GetUNCPath = path
    Exit Function
  End If
  
  CurrentDrive = Left(path, 2)
  ''' If the file in not on an external drive exit function
  If CurrentDrive = "C:" Or CurrentDrive = "\\" Then
    GetUNCPath = path
    Exit Function
  End If
  
  ''' Otherwise find the UNC Path (Universal Nameing Convention)
  Set network = CreateObject("WScript.Network")
  Set drives = network.enumnetworkdrives
  
  ''' Find the inital part (directory) of the path
  pattern = ".*:[\\/]*"
  RegEx.pattern = pattern
  str = RegEx.Replace(path, "")
  pattern = "[\\/].*"
  RegEx.pattern = pattern
  str = RegEx.Replace(str, "")
  
  ''' Match the path to a drive
  For Each el In drives
    RegEx.pattern = str
    If RegEx.Test(el) Then
      NewPath = el
      Exit For
    End If
  Next
  
  pattern = "/"
  RegEx.pattern = pattern
  Set RegEx2 = CreateObject("vbscript.regexp")
  pattern = "\\"
  RegEx2.pattern = pattern
  
  ''' Add remaining path
  ' Check the slashes used
  If RegEx.Test(path) Then
    PathDirs = Split(path, "/")
  ElseIf RegEx2.Test(path) Then
    PathDirs = Split(path, "\")
  Else
    MsgBox ("Error: Unable to parse path!")
  End If
  NewPathDirs = Split(NewPath, "\")
  ' Compare the paths to find the missing portion
  For i = UBound(PathDirs) To 1 Step -1
    If PathDirs(i) = NewPathDirs(UBound(NewPathDirs)) Then
      ' Skip if the file is in the top level directory
      If i < UBound(PathDirs) Then
        For ii = i + 1 To UBound(PathDirs)
          NewPath = NewPath + "\" + PathDirs(ii)
        Next ii
      End If
      Exit For
    End If
  Next i
  
  GetUNCPath = NewPath
End Function

Function set_sprint(str As String, pattern As String) As String
  Dim RegEx As Object
  Dim Matches As Object
  Dim match
  Dim count As Integer
  Dim sprintID As String
  Dim returnstr As String
  
  ' If the user has used the default setting just find the number
  If pattern = "default" Then
    pattern = "(\d+)"
  End If
  
  Set RegEx = CreateObject("vbscript.regexp")
  RegEx.Global = True
  RegEx.pattern = pattern
  If RegEx.Test(str) Then
    Set Matches = RegEx.Execute(str)
    
    count = 0
    For Each match In Matches
      count = count + 1
      sprintID = match.subMatches.Item(0)
    Next match
    
    ' Raise an error if the pattern matches more than one section of the string
    If count > 1 Then
      Err.Raise VBA.vbObjectError + 514, "Function: set_sprint", _
        "The match pattern has matched more than one section of the milestone string." _
         + " Please provide a regular expression which will only match the ID number of the sprint."
    End If
    
    returnstr = "Sprint " & sprintID
  Else
    returnstr = ""
  End If
  set_sprint = returnstr
End Function
