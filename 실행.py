Sub RunPythonScript()
    ' 방법 1: Shell 명령어 사용
    Dim pythonPath As String
    Dim scriptPath As String
    Dim cmd As String
    
    ' Python 실행 파일 경로 (실제 경로로 수정 필요)
    pythonPath = "C:\Python39\python.exe"
    ' Python 스크립트 경로 (실제 경로로 수정 필요)
    scriptPath = "E:\7 자료정보수집\제주도교육청채용정보.py"
    
    ' 명령어 구성
    cmd = pythonPath & " " & scriptPath
    
    ' 실행
    Shell cmd, vbNormalFocus
    
    ' 방법 2: WScript.Shell 사용
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' 실행
    wsh.Run cmd, 1, True
End Sub
