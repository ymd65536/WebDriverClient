Attribute VB_Name = "WebDriver"
Option Explicit

' �u���E�U�̃Z�b�V����ID
Public SessionId            As String
' �G�������g���擾����ׂ�ID
Public Const ElementConst   As String = "element-6066-11e4-a52e-4f735466cecf"
' �t���[���̒萔
Public Const FrameConst     As String = "frame-075b-4da1-b6ba-e579c2d3230a"""
' �G�������g����肷��ID
Public ElementId            As String
' �G���h�|�C���g
Public Const EndPointUrl    As String = "http://localhost:9515/session"
' WebDriber�̃p�X
Public Const EdgeDriverPath As String = "D:\edgedriver_win32\msedgedriver.exe"
Private fs                  As FileSystemObject
Public params               As Dictionary

Function DriverStatus() As Boolean
  Dim StatusDic As Dictionary
  
  Set StatusDic = SendRequest("GET", "http://localhost:9515/status", New Dictionary)
  DriverStatus = StatusDic("value")("ready")
End Function

Sub SetSessionId(testSessionId As String)
  SessionId = testSessionId
End Sub

Function GetSessionId() As String
  GetSessionId = SessionId
End Function
'
' WebDriver�ɑ΂��ă��N�G�X�g�𑗂�ۂɗ��p
'
Public Function SendRequest(method As String, url As String, Optional Data As Dictionary = Nothing) As Dictionary
  Dim Json As Object
  ' �N���C�A���g�̋N��
  Dim client As Object
  Set client = CreateObject("MSXML2.ServerXMLHTTP")

  ' ���\�b�h�ɉ����ă��N�G�X�g���M
  client.Open method, url
  If method = "POST" Or method = "PUT" Then
    client.setRequestHeader "Content-Type", "application/json"
    client.send JsonConverter.ConvertToJson(Data)
  Else
    client.send
  End If

  ' ���M�����҂�
  Do While client.readyState < 4
    DoEvents
  Loop
  ' ���X�|���X��Dictionary�ɕϊ����ă��^�[��
  ' Debug.Print client.responseText
  Set Json = JsonConverter.ParseJson(client.responseText)
  If IsNull(Json("value")) Then
    Set Data = New Dictionary
    Data.Add "value", "null"
    Set SendRequest = Data
  Else
    Set SendRequest = Json
  End If
End Function
'
' �u���E�U���J��
'
Public Function OpenBrowser() As Boolean
  
  Dim ResultParam  As Dictionary
  Set fs = New FileSystemObject
  ' WebDriver�̋N���B�f�t�H���g��9515�ԃ|�[�g���Ď�
  If fs.FileExists(EdgeDriverPath) Then
    Shell EdgeDriverPath, vbMinimizedNoFocus
    Set params = New Dictionary
    ' �u���E�U�N���p�����[�^�̍쐬
    params.Add "capabilities", New Dictionary
    params.Add "desiredCapabilities", Nothing
    ' �u���E�U�N��
    Set ResultParam = Nothing
    Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
    
    If ResultParam.Count < 3 Then
      SessionId = ResultParam("sessionId")
      OpenBrowser = True
    Else
      MsgBox ResultParam("error")
      OpenBrowser = False
    End If
    Set params = Nothing
  
  Else
    MsgBox EdgeDriverPath + "������܂���I�I"
    OpenBrowser = False
  End If
  Set fs = Nothing
End Function
'
' �����URL��Navigate
'
Public Function NavigateBrowser(url As String) As Boolean
  Dim ResultParam As Dictionary
  If url = "" Then
    NavigateBrowser = False
  ElseIf SessionId = "" Then
    MsgBox "�Z�b�V����ID������܂���"
    NavigateBrowser = False
  Else
    Set params = Nothing
    Set params = New Dictionary
    params.Add "url", url
    Set ResultParam = SendRequest("POST", EndPointUrl + "/" + SessionId + "/url", params)
    If TypeName(ResultParam("value")) = "String" Then
      NavigateBrowser = True
    Else
      MsgBox ResultParam("value")("error")
      NavigateBrowser = False
    End If
  End If
End Function
'
' �Z�b�V����ID�����ɃE�B���h�E�^�C�g����Ԃ�
'
Public Function GetTitle() As String
  If SessionId = "" Then
    GetTitle = ""
  Else
    Set params = Nothing
    Set params = New Dictionary
    Set params = SendRequest("GET", EndPointUrl + "/" + SessionId + "/title", params)
    
    If params Is Nothing Then
      GetTitle = ""
    Else
      GetTitle = params("value")
    End If
  End If
End Function
'
' �u���E�U �����
'
Public Sub CloseBrowser()
  Dim CloseObj As Object
  If SessionId = "" Then
  Else
    Set params = Nothing
    Set params = New Dictionary
    ' �E�B���h�E�����
    Set CloseObj = SendRequest("DELETE", EndPointUrl + "/" + SessionId, params)
  End If
End Sub
'
' �f�[�^���͗p���\�b�h
'
Public Sub SendKeyValue(ElementId As String, text As String)
  ' �l���͗p�̃p�����[�^������
  Dim chars()   As String
  Dim CharCnt   As Integer
  ReDim chars(Len(text) - 1)
  ' 1�������ɋ�؂�
  For CharCnt = 0 To UBound(chars)
    chars(CharCnt) = Mid(text, CharCnt + 1, 1)
  Next
  Set params = Nothing
  Set params = New Dictionary
  params.Add "text", text
  params.Add "value", chars
  ' ���ɓ��͂���Ă��镶��������
  SendRequest "POST", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/clear", New Dictionary
  ' �l���͂̎w��
  SendRequest "POST", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/value", params
End Sub
'
' CSS Selector ��ElementId��Ԃ�
'
Public Function GetElementByCssSelector(AttrName As String, AttrValue As String) As String
  
  On Error GoTo ErrExit
  Dim SetValue As String
  Set params = New Dictionary
  params.Add "using", "css selector"
  SetValue = "[" + AttrName + "=" + Chr(34) + AttrValue + Chr(34) + "]"
  params.Add "value", SetValue
  ElementId = SendRequest("POST", EndPointUrl + "/" + SessionId + "/elements", params)("value").Item(1)(ElementConst)
  GetElementByCssSelector = ElementId
  Exit Function
ErrExit:
  MsgBox Err.Description
  End
End Function
'
' ����̑������Ƒ����l���w�肵�čŏ��Ɉ�v�����G�������g���N���b�N����
'
Public Function ClickElement(SetElementId As String) As Boolean

  ElementId = SetElementId

  Dim ClickResult As New Dictionary
  Set params = Nothing
  Set params = New Dictionary
  params.Add "handle", """"""
  
  Set ClickResult = SendRequest("POST", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/click", params)
  
  ' Debug.print ClickResult("value")
  
  ClickElement = True

End Function
'
' �G�������g�̃v���p�e�B���擾����
' �v���p�e�B�̗�FinnerHTML,children
'
Public Function GetElementProperty(SetElementId As String, Property As String) As String

  If SetElementId = "null" Then
    GetElementProperty = "null"
  Else
    ElementId = SetElementId
    If SessionId = "" Then
      GetElementProperty = ""
    Else
      Set params = Nothing
      Set params = New Dictionary
      Set params = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/property/" + Property, params)
      
      If params Is Nothing Then
        GetElementProperty = ""
      Else
        If TypeName(params("value")) = "String" Then
          If params("value") = "" Then
            GetElementProperty = "none"
          Else
            GetElementProperty = params("value")
          End If
        Else
          If params("value").Count = 0 Then
            GetElementProperty = "zero"
            Exit Function
          Else
            If TypeName(params("value")) = "Collection" Then
              GetElementProperty = TypeName(params("value"))
            End If
          End If
        End If
      End If
    End If
  End If
End Function
'
' �����input�v�f��T��
'
Function FindInputElements(TypeAttrValue As String, ValueAttrValue As String) As String

  On Error GoTo ErrExit
  Dim SetValue    As String
  Dim param       As Collection
  Dim DicTagName  As Dictionary
  Dim TypeAttr    As Variant
  Dim ValueAttr   As Variant
  
  ElementId = ""
  Set params = New Dictionary
  params.Add "using", "tag name"
  params.Add "value", "input"
  Set params = SendRequest("POST", EndPointUrl + "/" + SessionId + "/elements", params)
  
  Dim param_tmp As Variant
  Set param = params("value")

  If param.Count = 0 Then
    ElementId = ""
  Else
    For Each param_tmp In param
      ElementId = param_tmp(ElementConst)
      Set DicTagName = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/name", New Dictionary)
      TypeAttr = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/attribute/type", New Dictionary)("value")
      ValueAttr = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementId + "/attribute/value", New Dictionary)("value")
      If ValueAttrValue = "null" Then
        If TypeAttr = TypeAttrValue And ValueAttr = "null" Then
          FindInputElements = ElementId
          Exit For
        End If
      ElseIf TypeAttr = TypeAttrValue And ValueAttr = ValueAttrValue Then
        FindInputElements = ElementId
        Exit For
      End If
    Next
  End If
Exit Function

ErrExit:
  MsgBox Err.Description
  ElementId = ""
  FindInputElements = ElementId
End Function

'
' �����a�v�f��T��
'
Function FindAnchorElements(innerHTML As String) As String
  
  On Error GoTo ErrExit
  
  Dim SetValue    As String
  Dim param       As Collection
  Dim DicTagName  As Dictionary
  Dim ElementItem  As String
  
  Dim InnerHTMLProperty    As Variant
  
  ElementId = ""
  Set params = New Dictionary
  params.Add "using", "tag name"
  params.Add "value", "a"
  
  Set params = SendRequest("POST", EndPointUrl + "/" + SessionId + "/elements", params)
  
  Dim param_tmp As Variant
  Set param = params("value")
  
  If param.Count = 0 Then
    ElementId = ""
  Else
    For Each param_tmp In param
      ElementItem = param_tmp(ElementConst)
      Set DicTagName = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementItem + "/name", New Dictionary)
      InnerHTMLProperty = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementItem + "/property/innerHTML", New Dictionary)("value")
      If InnerHTMLProperty = innerHTML Then
        ElementId = ElementItem
        FindAnchorElements = ElementId
        Exit For
      End If
    Next
  End If
Exit Function

ErrExit:
  MsgBox Err.Description
  ElementId = ""
  FindAnchorElements = ElementId
End Function
'
' �����Frame�v�f��T��
'
Function FindFrameElements(NameValue As String, Optional TagName As String = "frame") As String
  
  On Error GoTo ErrExit
  
  Dim param       As Collection
  Dim DicTagName  As Dictionary
  Dim FrameElement As Variant
  Dim ElementItem  As String
  
  ElementId = ""
  Set params = New Dictionary
  params.Add "using", "tag name"
  params.Add "value", TagName
  
  ' �����̃G�������g��T��
  Set params = SendRequest("POST", EndPointUrl + "/" + SessionId + "/elements", params)
  
  Dim param_tmp As Variant
  Set param = params("value")
  
  If param.Count = 0 Then
    ElementId = ""
  Else
    ' For �� �G�������g�����o��
    For Each FrameElement In param
      ElementItem = FrameElement(ElementConst)
      Set DicTagName = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementItem + "/attribute/name", New Dictionary)
        ' name ���������邩
      If TypeName(DicTagName("value")) = "String" Then
        ' name ������Value���`�F�b�N
        If DicTagName("value") = NameValue Then
          ElementId = ElementItem
          FindFrameElements = ElementItem
          Exit For
        End If
      End If
    Next
  End If
Exit Function

ErrExit:
  MsgBox Err.Description
  ElementId = ""
  FindFrameElements = ElementId
End Function
'
' JavaScript �����s
'
Public Sub ExcuteScriptSync(Script As String)
  Dim SetValue    As String
  Set params = New Dictionary
  params.Add "script", Script
  params.Add "args", New Collection
  Set params = SendRequest("POST", EndPointUrl + "/" + SessionId + "/execute/sync", params)
End Sub
'
' Cookie ���擾
'
Function GetCookie() As String
  Set params = New Dictionary
  Set params = SendRequest("GET", EndPointUrl + "/" + SessionId + "/cookie", params)
  GetCookie = params("value").Item(1)("value")
End Function
'
' Window �ɃX�C�b�`
'
Function SwitchToWinodow(Optional handle As String = "") As String
  Set params = Nothing
  Set params = New Dictionary
  params.Add "handle", handle
  SendRequest "POST", EndPointUrl + "/" + SessionId + "/window", params
  If TypeName(params("value")) = "String" Then
    SwitchToWinodow = "null"
  Else
    SwitchToWinodow = "error"
  End If
End Function
'
' �C�ӂ�Frame�ɃX�C�b�`
'
Function SwitchToFrame(TagName As String, FrameName As String, Optional DefaultFrameNo As Integer = 0) As String
  
  Dim handle As String
  
  ' �X�C�b�`����t���[�������
  Dim param_tmp As Variant
  Dim param        As Collection
  Dim DicTagName   As Dictionary
  Dim FrameElement As Variant
  Dim ElementItem  As String
  
  Dim FrameIdx        As Integer: FrameIdx = 0
  Dim FrameIdxTmp     As Integer: FrameIdxTmp = 0
  
  ' �R���e�L�X�g�̃X�C�b�`�󋵂��C��
  handle = ""
  Set params = Nothing
  Set params = New Dictionary
  params.Add "handle", handle
  SendRequest "POST", EndPointUrl + "/" + SessionId + "/window", params
  
  ElementId = ""
  Set params = New Dictionary
  params.Add "using", "tag name"
  params.Add "value", TagName
  
  ' �����̃G�������g��T��
  Set params = SendRequest("POST", EndPointUrl + "/" + SessionId + "/elements", params)
  Set param = params("value")
  
  If param.Count = 0 Then
    ElementId = ""
  Else
    ' For �� �G�������g�����o��
    For Each FrameElement In param
      ElementItem = FrameElement(ElementConst)
      Set DicTagName = SendRequest("GET", EndPointUrl + "/" + SessionId + "/element/" + ElementItem + "/attribute/name", New Dictionary)
        ' name ���������邩
      If TypeName(DicTagName("value")) = "String" Then
        ' name ������Value���`�F�b�N
        If DicTagName("value") = FrameName Then
          FrameIdx = FrameIdxTmp + 1
          Exit For
        End If
        FrameIdxTmp = FrameIdxTmp + 1
      End If
    Next
  End If
  
  ' �J�E���g�l-1 �� �t���[��ID�ɂȂ�
  FrameIdx = FrameIdx - 1
  
  ' �f�t�H���g�̃t���[���i���o�[�͌����Ƀq�b�g����菬�����Ȃ���΂Ȃ�Ȃ��B
  If FrameIdx < 0 Then
    If DefaultFrameNo >= FrameIdx Then
      MsgBox "�t���[��������܂���"
      End
    Else
      FrameIdx = DefaultFrameNo
    End If
  End If
  
  ' ����̃t���[���ɃX�C�b�`
  Set params = Nothing
  Set params = New Dictionary
  
  params.Add "id", FrameIdx
  SendRequest "POST", EndPointUrl + "/" + SessionId + "/frame", params
  If TypeName(params("value")) = "String" Then
    SwitchToFrame = "null"
  Else
    SwitchToFrame = "error"
  End If
End Function


