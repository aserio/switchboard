Attribute VB_Name = "Module1"
Option Explicit

Sub switchboard()
'' This subroutine is able to read in tasks from a file and update the Project tasks
''
'' Asumptions:
''    Board status have been added which match Python Output (Not Started, In progress, Under Review, Done)
''    Configuration file named config.txt is present in working directory
Dim FileName As String
    Dim ProjectName As String
    Dim PathName As String
    Dim FilePath As String
    Dim ConfigFile As String
    Dim GitHubRepo As String
    Dim SprintLength As String
    Dim ProjectFieldDur As Long
    Dim ProjectFieldBS As Long
    Dim i As Integer
    Dim LineFromFile As String
    Dim LineItems() As String
    Dim gid As Integer
    Dim NewTask As Task
    Dim DateValue As Date
    
    ProjectName = Replace(Application.ActiveProject.Name, ".mpp", "")
    PathName = Application.ActiveProject.Path
    FileName = ProjectName + ".csv"
    FilePath = PathName + "\" + FileName
    
    ConfigFile = "config.txt"
    ''' Fetch configuration information for the project
    Open PathName + "\" + ConfigFile For Input As #1
    Line Input #1, LineFromFile ' Ignore header
    Do Until EOF(1)
      Line Input #1, LineFromFile
      LineItems = parse_line(LineFromFile)
      If LineItems(0) = ProjectName Then
        GitHubRepo = LineItems(1)
        SprintLength = LineItems(2)
      End If
    Loop
    
    ' Clean up
    Close #1
    LineFromFile = ""
   
   ''' Call Python script to fetch GitHub issues
    Dim wshell As Object
    Dim process As Object
    Dim PythonExe As String
    Dim PythonScript As String
    Dim Args As String
    Dim error_code As Double
    
    Set wshell = CreateObject("WScript.Shell")
    
    'You have to excape the inner quotes with another set of quotes (ie. " "" text ...)
    PythonExe = """C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python37_64\python.exe"""
    PythonScript = """C:\Users\alemoine\OneDrive - Advanced Micro Devices Inc\AMD-AUS-LX-ALEMOINE\Experimental\vba_examples\production_test\pygithub_prod_test.py"""
    ' Chr(34) is the double quotes character
    Args = "--github_repo " & Chr(34) & GitHubRepo & Chr(34) & " --csv_file " & Chr(34) & FilePath & Chr(34) & " --sprint_length " & SprintLength

    error_code = wshell.Run(PythonExe & " " & PythonScript & " " & Args, 1, True)
    'MsgBox (PythonExe & " " & PythonScript & " " & Args)
    
    Open FilePath For Input As #2
      
    ''' Get field IDs
    ProjectFieldDur = FieldNameToFieldConstant("Duration", pjProject)
    ProjectFieldBS = FieldNameToFieldConstant("Board Status", pjProject)
    
    ''' Create dictionary relating task ID to GitHub ID
    Dim unique_id As Long
    Dim git_id As Long
    Dim dict
    Dim ii As Integer

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To ActiveProject.Tasks.Count
      If Application.ActiveProject.Tasks(i).Text1 = vbNullString Then
      Else
        unique_id = Application.ActiveProject.Tasks(i).UniqueID
        git_id = Application.ActiveProject.Tasks(i).Text1
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
        Application.ActiveProject.Tasks.UniqueID(gid).Sprint = LineItems(5)
        Application.ActiveProject.Tasks.UniqueID(gid).SetField FieldID:=ProjectFieldBS, Value:=LineItems(6)
        'Add Labels
        Application.ActiveProject.Tasks.UniqueID(gid).Text6 = LineItems(8)
        Application.ActiveProject.Tasks.UniqueID(gid).Text8 = LineItems(9)
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
        NewTask.Sprint = LineItems(5)
        NewTask.SetField FieldID:=ProjectFieldBS, Value:=LineItems(6)
        'Add GitHub Issue Number
        NewTask.Text1 = LineItems(7)
        'Add Labels
        NewTask.Text6 = LineItems(8)
        NewTask.Text8 = LineItems(9)
        ' Set Percent Complete last to stop Project from overwriting
        If LineItems(1) = vbNullString Then
           NewTask.PercentComplete = 1
        Else
           NewTask.PercentComplete = CInt(LineItems(1))
        End If
        
      End If

    Loop
        
    Close #2

End Sub

Function parse_line(str As String) As String()
Dim RegEx As Object
Dim pattern
Dim str_array() As String
Dim el As Variant
Dim i As Integer

' Find only the commas outside of the quotes
pattern = ",(?=([^" & Chr(34) & "]*" & Chr(34) & "[^" & Chr(34) & "]*" & Chr(34) & ")*(?![^" & Chr(34) & "]*" & Chr(34) & "))"
Set RegEx = CreateObject("vbscript.regexp")
RegEx.Global = True
RegEx.pattern = pattern
'parse_line = Split(RegEx.Replace(str, ";"), ";")
str_array = Split(RegEx.Replace(str, ";"), ";")

' Remove leading whitespace
pattern = "^\s+"
RegEx.pattern = pattern
For i = LBound(str_array) To UBound(str_array)
  str_array(i) = RegEx.Replace(str_array(i), "")
Next i

parse_line = str_array

End Function
