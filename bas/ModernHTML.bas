Attribute VB_Name = "ModernHTML"
Option Explicit

'
' ���_���u���E�U�ɂ�����HTML���C�C�J���W�ɑ���ł���悤�ɂ���
'

' ��ʑJ�ڂ��Ȃ���ʍX�V�΍�
' innerHTML���擾����ꍇ�͉�ʍX�V���I��������Ƃ��m�F����
' �擾���Ȃ���΂Ȃ�Ȃ��B

Function DocumentWait(innerHTML As String, HtmlDoc As IHTMLDocument) As HTMLDocument
  If Not innerHTML = "" Then
    Set HtmlDoc = New HTMLDocument
    HtmlDoc.Write innerHTML
    Set DocumentWait = HtmlDoc
  Else
    Set DocumentWait = Nothing
  End If
End Function
