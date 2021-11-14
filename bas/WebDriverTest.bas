Attribute VB_Name = "WebDriverTest"
Option Explicit

Sub EdgeDriverExecute()
  Shell EdgeDriverPath, vbMinimizedNoFocus
End Sub

'
' テストモジュール
'
Sub SendRequestTest()
  
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count = 1 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  ' エラーパターン
  Set ResultParam = SendRequest("POST", "http://localhost:9515/sessions", params)("value")
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
    MsgBox ResultParam("error")
  End If
  
  Set ResultParam = Nothing
  Set params = Nothing
  
  ' リクエストに失敗した場合
End Sub

Sub CloseBrowserTest()
  ' ブラウザを閉じるテスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  ' セッションIDが指定されていない場合
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub NavigateBrowserTest()
  ' URLを開くテスト
  
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser("https://google.co.jp")
  
  ' URLを間違えている場合
  Call NavigateBrowser("https://google.co.j")
  
  ' --- おわり
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub GetTitleTest()
  ' タイトルを取得するテスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  
  ' URLが指定されていない場合
  Call NavigateBrowser("https://google.co.jp")
  
  Debug.Print GetTitle
  
  ' URLを間違えている場合
  Call NavigateBrowser("https://google.co.j")
  
  Debug.Print GetTitle
  
  ' --- おわり
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub
Sub GetElementByCssSelectorTest()
  ' CSS Selector によるエレメント取得のテスト
End Sub


Sub SendKeyValueTest()
  
  ' テキストを入力するテスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser("https://google.co.jp")
  Debug.Print GetTitle
  
  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("name", "q")
  Call SendKeyValue(QueryElementId, "test")
  
  ' --- おわり
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub ClickElementTest()
  ' クリックテスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser("https://google.co.jp")
  Debug.Print GetTitle

  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("name", "q")
  Call SendKeyValue(QueryElementId, "test")

  Dim InputElementId As String
  InputElementId = GetElementByCssSelector("name", "btnK")
  Call ClickElement(InputElementId)
  
  ' --- おわり
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub GetElementPropertyTest()
  ' エレメントのプロパティの取得テスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "inner_html.html")
  Debug.Print GetTitle

  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("name", "q")
  
  Debug.Print GetElementProperty(QueryElementId, "innerHTML")
  
  ' --- おわり
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub FindInputElementsTest()
  ' 特定のinput 要素を探す
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "find_input_elements.html")
  Debug.Print GetTitle

  Dim QueryElementId As String
  QueryElementId = GetElementByCssSelector("type", "button")
  
  Debug.Print QueryElementId
  
  Debug.Print GetElementProperty(QueryElementId, "value")
  
  ' --- おわり
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub FindAnchorElementsTest()
  ' 特定のアンカータグを探すテスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "find_anchor_element.html")
  Debug.Print GetTitle

  Dim AnchorElementId As String
  AnchorElementId = FindAnchorElements("Google")
  
  Debug.Print GetElementProperty(AnchorElementId, "innerHTML")
  
  Call ClickElement(AnchorElementId)
  
  ' --- おわり
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub FindFrameElementsTest()
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "frames.html")
  
  Dim FrameElementId As String
  FrameElementId = FindFrameElements("test")
  
  Debug.Print GetElementProperty(FrameElementId, "src")
  ' --- おわり
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub

Sub ExcuteScriptSyncTest()
  ' スクリプト実行テスト
End Sub

Sub GetCookieTest()
  ' Cookie取得テスト
End Sub

Sub SwitchToWinodowTest()
  ' ウィンドウにスイッチするテスト

  
End Sub

Sub SwitchToFrameTest()
  ' フレームにスイッチするテスト
  Call EdgeDriverExecute
  
  ' SendRequest をテスト
  Dim SessionId    As String
  Dim ResultParam  As Dictionary
  
  Set params = New Dictionary
  
  ' ブラウザ起動パラメータの作成
  params.Add "capabilities", New Dictionary
  params.Add "desiredCapabilities", Nothing

  ' ブラウザ起動
  Set ResultParam = SendRequest("POST", EndPointUrl, params)("value")
  
  ' 成功した場合はCount プロパティは1
  If ResultParam.Count < 3 Then
    SessionId = ResultParam("sessionId")
  Else
    SessionId = ""
  End If
  
  WebDriver.SetSessionId (SessionId)
  
  ' --- ここからやりたい処理を書く
  Call NavigateBrowser(ThisWorkbook.Path + "\" + "frames.html")
  Debug.Print SwitchToFrame("frame", "test")

  ' --- おわり
  
  Call WebDriver.CloseBrowser
  
  Set ResultParam = Nothing
  Set params = Nothing
End Sub
