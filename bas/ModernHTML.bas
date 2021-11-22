Attribute VB_Name = "ModernHTML"
Option Explicit

'
' モダンブラウザにおけるHTMLをイイカンジに操作できるようにする
'

' 画面遷移しない画面更新対策
' innerHTMLを取得する場合は画面更新が終わったことを確認して
' 取得しなければならない。

Function DocumentWait(innerHTML As String, HtmlDoc As IHTMLDocument) As HTMLDocument
  If Not innerHTML = "" Then
    Set HtmlDoc = New HTMLDocument
    HtmlDoc.Write innerHTML
    Set DocumentWait = HtmlDoc
  Else
    Set DocumentWait = Nothing
  End If
End Function
