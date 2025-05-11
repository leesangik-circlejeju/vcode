Option Explicit
' FindWindow : 윈도우의 핸들을 찾는 함수
' FindWindowEx : 부모 윈도우 내에서 자식 윈도우의 핸들을 찾는 함수
' SendMessage : 윈도우에 메시지를 보내는 함수 (예: 텍스트 입력 등)
' PostMessage : 윈도우에 메시지를 비동기적으로 보내는 함수 (예: 키 입력 등)
' keybd_event : 가상 키보드 이벤트를 발생시키는 함수 (키보드 입력 시뮬레이션)
' GetKeyState : 특정 키의 현재 상태(눌림/안눌림)를 확인하는 함수
' GetWindow : 윈도우의 관계(부모/자식/다음 윈도우 등)를 가져오는 함수
' GetClassName : 윈도우의 클래스 이름을 가져오는 함수
' GetWindowText : 윈도우의 제목(캡션)을 가져오는 함수
' GetWindowHandle 함수: FindWindow API를 래핑

Function GetWindowHandle(lpClassName As String, lpWindowName As String) As LongPtr
    GetWindowHandle = FindWindow(lpClassName, lpWindowName)
End Function

' FindWindowEx 래핑 함수
Function GetChildWindowHandle(hwndParent As LongPtr, hwndChildAfter As LongPtr, lpszClass As String, lpszWindow As String) As LongPtr
    GetChildWindowHandle = FindWindowEx(hwndParent, hwndChildAfter, lpszClass, lpszWindow)
End Function

Sub main()
    Dim data(1 To 20, 1 To 4) As String ' 1:순번, 2:카톡이름, 3:실명, 4:발송여부
    Dim i As Integer
    Dim 카톡이름 As String
    Dim 실명 As String

    ' 데이터 입력 (발송순번, 카톡이름, 실명)
    For i = 1 To 20
        data(i, 1) = CStr(i)                ' 발송순번
        data(i, 2) = "카톡이름" & i         ' 카톡이름
        data(i, 3) = "실명" & i             ' 실명
        data(i, 4) = ""                    ' 발송여부 (초기값)
    Next i

    ' 처리 예시
    For i = 1 To 20
        카톡이름 = data(i, 2)
        실명 = data(i, 3)

        ' (여기에 실제 발송 함수 호출, 예: SendKakao(카톡이름, ...))
        ' 예시로 발송여부를 "성공"으로 처리
        data(i, 4) = "성공"

        ' 1초 대기
        Application.Wait Now + TimeValue("0:00:01")
    Next i

    ' 결과 Immediate 창에 출력
    For i = 1 To 20
        Debug.Print "순번:" & data(i, 1) & ", 카톡이름:" & data(i, 2) & ", 실명:" & data(i, 3) & ", 발송여부:" & data(i, 4)
    Next i

    MsgBox "전송이 끝났습니다."
End Sub

'######################################################################
' 공통 상수 선언
'######################################################################

' 가상키 ASCII 코드
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11
Public Const VK_ESC = &H1B
Public Const KEYEVENTF_KEYUP As Long = 2
Public Const VK_A = &H41
Public Const VK_V = &H56
Public Const VK_C = &H43
Public Const VK_DOWN = &H28
Public Const VK_UP = &H26

'######################################################################
' 친구에게 카카오톡 메세지 전송합니다.
'######################################################################
Function SendKakao(Target As String, Message As String, Optional ChatRoomSearch As Boolean = False, Optional iDelay As Long = 1) As Boolean

    ' 변수 선언
    Dim hwnd_RichEdit As LongPtr       ' 채팅입력창 hwnd

    ' 친구 채팅입력창 hwnd 찾기
    hwnd_RichEdit = FindReciepientHwnd(Target, ChatRoomSearch, iDelay)

    '보낼 대상의 채팅창을 찾을 경우 안내메세지 출력 후 종료


If hwnd_RichEdit = 0 Then
        MsgBox "보낼 대상의 카카오톡 대화창을 찾을 수 없습니다.", vbExclamation
        Exit Function
    End If

    ' 메세지 보내기
    Send_TextMsg Message, hwnd_RichEdit

    SendKakao = True

End Function

'######################################################################
' 카카오톡 메세지를 전송합니다.
'######################################################################
Sub Send_TextMsg(Message As String, hwnd_RichEdit As LongPtr)

    ' 대화상대 채팅 입력창 hWnd 에 메세지 입력
    Call SendMessage(hwnd_RichEdit, WM_SETTEXT, 0, ByVal Message)

    ' 사용자 Ctrl 키 입력여부 확인
    If IsCtrlKeyDown = False Then
        ' Ctrl 키 미입력 시, 메세지 전송
        Call PostMessage(hwnd_RichEdit, WM_KEYDOWN, VK_RETURN, 0)
    Else
        ' Ctrl 키 입력중일 경우, 강제로 Ctrl 키 떼었다 떼는 동작 -> 메세지 전송 -> Ctrl 키 재입력
        keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
        Call PostMessage(hwnd_RichEdit, WM_KEYDOWN, VK_RETURN, 0)
        keybd_event VK_CONTROL, 0, 0, 0
    End If

End Sub

'######################################################################
' 보낼 대상의 hWnd 값을 찾습니다.
' 보낼 대상의 채팅창이 없을 경우 False 를 반환합니다.
'######################################################################
Function FindReciepientHwnd(Target As String, ChatRoomSearch As Boolean, iDelay As Long) As LongPtr

    ' 변수 선언
    Dim dStart As Double
    Dim hwnd_KakaoTalk As LongPtr       ' 친구 채팅창 hwnd
    Dim hwnd_RichEdit As LongPtr       ' 채팅입력창 hwnd

    ' 보낼 대상의 카톡창 실행
    ActiveChat Target, iDelay, ChatRoomSearch

    '보낼 대상의 카톡창 Hwnd 찾기
    hwnd_KakaoTalk = GetWindowHandle(vbNullString, Target)


If hwnd_KakaoTalk = 0 Then
        dStart = Now
        While hwnd_KakaoTalk = 0
            hwnd_KakaoTalk = GetWindowHandle(vbNullString, Target)
            ' 1초 지났는데 창을 못찾으면 함수 강제 종료
            If DateDiff("s", dStart, Now) > 1 Then
                FindReciepientHwnd = 0
                Exit Function
            End If
            DoEvents
        Wend
    End If

    ' 친구의 채팅입력창 hWnd 찾기
    hwnd_RichEdit = FindWindowEx(hwnd_KakaoTalk, 0, "RichEdit50W", vbNullString) ' 카카오톡 버전에 따른 RichEdit ClassName 차이
    If hwnd_RichEdit = 0 Then hwnd_RichEdit = FindWindowEx(hwnd_KakaoTalk, 0, "RichEdit20W", vbNullString)

    FindReciepientHwnd = hwnd_RichEdit

End Function

'######################################################################
' 친구 채팅창이 열려있지 않을 경우 채팅창을 활성화 합니다.
'######################################################################
Function ActiveChat(Target As String, iDelay As Long, ChatRoomSearch As Boolean)

    ' 변수 설정
    Dim hwndMain As LongPtr: Dim hwndChild1 As LongPtr: Dim hwndChild2 As LongPtr: Dim hwndEdit As LongPtr
    Dim i As Long

    ' 대화상대 창 이미 열려있을 시 명령문 종료
    If GetWindowHandle(vbNullString, Target) > 0 Then Exit Function

    ' 카카오톡 메인 hWnd 검색 [▶▶▶▶▶ F5 키]
    hwndMain = GetWindowHandle("KakaoTalk.exe", vbNullString)
    If hwndMain = 0 Then hwndMain = GetWindowHandle("Eva_Window", vbNullString) '최신 버전 대응

    ' 카카오톡 메인 창 없으면 카톡 실행 안됨 -> 함수 False 반환 후 명령문 종료
    If hwndMain = 0 Then
        ActiveChat = False
        Exit Function
    End If

    ' 메인 [▶▶▶▶▶ F5 키]
    hwndChild1 = FindWindowEx(hwndMain, 0, "EVA_ChildWindow", vbNullString)

    ' 연락처 목록 검색 -> EVA_Window 첫번째 항목
    hwndChild2 = FindWindowEx(hwndChild1, 0, "EVA_Window", vbNullString)

    ' 채팅창 목록 검색 -> EVA_Window 두번째 항목
    If ChatRoomSearch = True Then hwndEdit = FindWindowEx(hwndChild2, hwndChild2, "EVA_Window", vbNullString)


' EVA_Window 의 대화상대 검색창
    hwndEdit = FindWindowEx(hwndChild2, 0, "Edit", vbNullString)

    ' 검색창에 대화상대 복사/붙여넣기
    Call SendMessage(hwndEdit, WM_SETTEXT, 0, ByVal Target): MyDelay iDelay
    ' 채팅창 목록 맨위로 땡겨와서 검색창 활성화
    If ChatRoomSearch = True Then Call PostMessage(hwndEdit, WM_KEYDOWN, VK_UP, 0): MyDelay iDelay
    ' 엔터키로 채팅창 열기
    Call PostMessage(hwndEdit, WM_KEYDOWN, VK_RETURN, 0): MyDelay iDelay

End Function

'######################################################################
' 카카오톡 실행여부 확인 후, 실행중일 시 메인창의 hWnd 를 반환합니다.
'######################################################################
Private Function FindHwndEVA() As LongPtr

    ' 변수 설정
    Dim hwnd As LongPtr: Dim lngT As Long: Dim strT As String

    ' 현재 Desktop 에서 실행중인 첫번째 프로그램 hWnd
    hwnd = GetWindowHandle(vbNullString, vbNullString)

    ' 모든 hWnd 를 돌아가며 검색
    While hwnd <> 0
        ' hwnd 를 돌아가며 ClassName 확인
        strT = String(100, Chr(0))
        lngT = GetClassName(hwnd, strT, 100)
        strT = Left$(strT, lngT)

        ' hwnd 의 ClassName 에 "EVA_Window_Dblclk" 이 포함될 경우
        If InStr(1, Left$(strT, lngT), "EVA_Window_Dblclk") > 0 Then
            ' 해당 hWnd 의 채팅창 이름을 받아옴.
            strT = String(100, Chr(0))
            lngT = GetWindowText(hwnd, strT, 100)
            strT = Left$(strT, lngT)

            ' 채팅창 이름이 "카카오톡" 또는 "KakaoTalk" (영문 OS 사용시) 일 경우, hWnd 를 함수 결과로 반환 후 종료
            If InStr(1, Left$(strT, lngT), "카카오톡") > 0 Or InStr(1, Left$(strT, lngT), "KakaoTalk") > 0 Then
                FindHwndEVA = hwnd: Exit Function
            End If
        End If

        hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
    Wend

End Function

'######################################################################
' 오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
' MyDelay 함수

' 윈도우 지연 코드
' MyDelay : 지연 속도 조절 (클수록 지연 많음, 1 = 4ms)
'######################################################################
Sub MyDelay(Optional iDelay As Long = 1)

    Dim i As Long

    For i = 1 To iDelay * 1000000: i = i + 1: Next

End Sub

'######################################################################
' 오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
' IsCtrlKeyDown 함수
' 키보드 Ctrl 키 누름여부를 확인합니다.
' 인수 설명
' __Optional LeftRightKey : 왼쪽/오른쪽 Ctrl 키 누름여부를 정합니다.
' 1 : 왼쪽 Ctrl 키 입력 시 TRUE
' 2 : 오른쪽 Ctrl 키 입력 시 TRUE
' 3 : 양쪽 Ctrl 키 동시 입력 시 TRUE
' 0 : 둘 중 하나라도 입력 시 TRUE
'######################################################################
Private Function IsCtrlKeyDown(Optional LeftRightKey As Long = 0) As Boolean

    Const VK_LCTRL = &HA2: Const VK_RCTRL = &HA3: Const KEY_MASK As Integer = &H80

    Dim Result As Long

    Select Case LeftRightKey
        ' 왼쪽 Ctrl 키 입력여부 확인
        Case 1: Result = GetKeyState(VK_LCTRL) And KEY_MASK
        ' 오른쪽 Ctrl 키 입력여부 확인
        Case 2: Result = GetKeyState(VK_RCTRL) And KEY_MASK
        ' 양쪽 Ctrl 키 동시 입력여부 확인
        Case 3: Result = GetKeyState(VK_LCTRL) And GetKeyState(VK_RCTRL) And KEY_MASK
        ' Ctrl 키 둘중 하나의 입력여부 확인
        Case Else: Result = GetKeyState(vbKeyControl) And KEY_MASK
    End Select

    IsCtrlKeyDown = CBool(Result)

End Function




' =============================
' 64비트용 Windows API 선언부

' FindWindow : 윈도우의 핸들을 찾는 함수
' FindWindowEx : 부모 윈도우 내에서 자식 윈도우의 핸들을 찾는 함수
' SendMessage : 윈도우에 메시지를 보내는 함수 (예: 텍스트 입력 등)
' PostMessage : 윈도우에 메시지를 비동기적으로 보내는 함수 (예: 키 입력 등)
' keybd_event : 가상 키보드 이벤트를 발생시키는 함수 (키보드 입력 시뮬레이션)
' GetKeyState : 특정 키의 현재 상태(눌림/안눌림)를 확인하는 함수
' GetWindow : 윈도우의 관계(부모/자식/다음 윈도우 등)를 가져오는 함수
' GetClassName : 윈도우의 클래스 이름을 가져오는 함수
' GetWindowText : 윈도우의 제목(캡션)을 가져오는 함수
' =============================



' FindWindow : 윈도우의 핸들을 찾는 함수
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As LongPtr

' FindWindowEx : 부모 윈도우 내에서 자식 윈도우의 핸들을 찾는 함수
Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
    ByVal hwndParent As LongPtr, _
    ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String _
) As LongPtr

' SendMessage : 윈도우에 메시지를 보내는 함수 (예: 텍스트 입력 등)
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As LongPtr, _
    ByRef lParam As Any _
) As LongPtr

' PostMessage : 윈도우에 메시지를 비동기적으로 보내는 함수 (예: 키 입력 등)
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As LongPtr, _
    ByRef lParam As Any _
) As LongPtr

' keybd_event : 가상 키보드 이벤트를 발생시키는 함수 (키보드 입력 시뮬레이션)
Private Declare PtrSafe Sub keybd_event Lib "user32.dll" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As LongPtr _
)

' GetKeyState : 특정 키의 현재 상태(눌림/안눌림)를 확인하는 함수
Private Declare PtrSafe Function GetKeyState Lib "user32" ( _
    ByVal nVirtKey As Long _
) As Long

' GetWindow : 윈도우의 관계(부모/자식/다음 윈도우 등)를 가져오는 함수
Public Declare PtrSafe Function GetWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal wCmd As Long _
) As LongPtr

' GetClassName : 윈도우의 클래스 이름을 가져오는 함수
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long

' GetWindowText : 윈도우의 제목(캡션)을 가져오는 함수
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpString As String, _
    ByVal cch As Long _
) As Long
