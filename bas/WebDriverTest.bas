Attribute VB_Name = "WebDriverTest"
Option Explicit

Sub EdgeDriverExecute()
  Shell EdgeDriverPath, vbMinimizedNoFocus
End Sub

'
' �e�X�g���W���[��
'
Sub SendRequestTest()
  
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count = 1 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  ' �G���[�p�^�[��
  Set ResultParam = SendRequest("POST", "http://localhost:9515/sessions", params)("value")
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
    MsgBox ResultParam("error")
  End If
  
  Set ResultParam = Nothing
  Set params = Nothing
  
  ' ���N�G�X�g�Ɏ��s�����ꍇ
End Sub

Sub CloseBrowserTest()
  ' �u���E�U�����e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  ' �Z�b�V����ID���w�肳��Ă��Ȃ��ꍇ
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub NavigateBrowserTest()
  ' URL���J���e�X�g
  
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser("https://google.co.jp")
  
  ' URL���ԈႦ�Ă���ꍇ
  Call NavigateBrowser("https://google.co.j")
  
  ' --- �����
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub GetTitleTest()
  ' �^�C�g�����擾����e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  
  ' URL���w�肳��Ă��Ȃ��ꍇ
  Call NavigateBrowser("https://google.co.jp")
  
  Debug.Print GetTitle
  
  ' URL���ԈႦ�Ă���ꍇ
  Call NavigateBrowser("https://google.co.j")
  
  Debug.Print GetTitle
  
  ' --- �����
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub
Sub GetElementByCssSelectorTest()
  ' CSS Selector �ɂ��G�������g�擾�̃e�X�g
End Sub


Sub SendKeyValueTest()
  
  ' �e�L�X�g����͂���e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser("https://google.co.jp")
  Debug.Print GetTitle
  
  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("name", "q")
  Call SendKeyValue(QueryElementId, "test")
  
  ' --- �����
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub ClickElementTest()
  ' �N���b�N�e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser("https://google.co.jp")
  Debug.Print GetTitle

  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("name", "q")
  Call SendKeyValue(QueryElementId, "test")

  Dim InputElementId As String
  InputElementId = GetElementByCssSelector("name", "btnK")
  Call ClickElement(InputElementId)
  
  ' --- �����
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub GetElementPropertyTest()
  ' �G�������g�̃v���p�e�B�̎擾�e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "inner_html.html")
  Debug.Print GetTitle

  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("name", "q")
  
  Debug.Print GetElementProperty(QueryElementId, "innerHTML")
  
  ' --- �����
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub FindInputElementsTest()
  ' �����input �v�f��T��
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "find_input_elements.html")
  Debug.Print GetTitle

  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("type", "button")
  
  Debug.Print QueryElementId
  
  Debug.Print GetElementProperty(QueryElementId, "value")
  
  ' --- �����
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub FindAnchorElementsTest()
  ' ����̃A���J�[�^�O��T���e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "find_anchor_element.html")
  Debug.Print GetTitle

  Dim AnchorElementId As String
  AnchorElementId = FindAnchorElements("Google")
  
  Debug.Print GetElementProperty(AnchorElementId, "innerHTML")
  
  Call ClickElement(AnchorElementId)
  
  ' --- �����
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub FindFrameElementsTest()
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "frames.html")
  
  Dim FrameElementId As String
  FrameElementId = FindFrameElements("test")
  
  Debug.Print GetElementProperty(FrameElementId, "src")
  ' --- �����
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub ExcuteScriptSyncTest()
  ' �X�N���v�g���s�e�X�g
End Sub

Sub GetCookieTest()
  ' Cookie�擾�e�X�g
End Sub

Sub SwitchToWinodowTest()
  ' �E�B���h�E�ɃX�C�b�`����e�X�g

  
End Sub

Sub SwitchToFrameTest()
  ' �t���[���ɃX�C�b�`����e�X�g
  Call EdgeDriverExecute
  
  ' SendRequest ���e�X�g
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' �u���E�U�N���p�����[�^�̍쐬
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' �u���E�U�N��
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' ���������ꍇ��Count �v���p�e�B��1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ���������肽������������
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "frames.html")
  Debug.Print SwitchToFrame("frame", "test")

  ' --- �����
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub
