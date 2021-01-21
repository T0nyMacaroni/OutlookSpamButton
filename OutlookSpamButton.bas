Attribute VB_Name = "OutlookSpamButton"
Sub ForwardAndDeleteSpam()
'
' Takes currently highlighted e-mail, sends it as an attachment to
' verdacht@safeonweb.be and then deletes the message from the inbox.
'

    Set objItem = GetCurrentItem()
    Set objMsg = Application.CreateItem(olMailItem)

    With objMsg
        .Attachments.Add objItem, olEmbeddeditem
        .Subject = objItem.Subject()
        .To = "verdacht@safeonweb.be"
        'NL: verdacht@safeonweb.be
        'FR: suspect@safeonweb.be
        'EN: suspicious@safeonweb.be
        .Send
    End With
    objItem.Delete

    Set objItem = Nothing
    Set objMsg = Nothing
End Sub

Function GetCurrentItem() As Object
    On Error Resume Next
    Select Case TypeName(Application.ActiveWindow)
    Case "Explorer"
        Set GetCurrentItem = Application.ActiveExplorer.Selection.Item(1)
    Case "Inspector"
        Set GetCurrentItem = Application.ActiveInspector.CurrentItem
    Case Else
        ' anything else will result in an error, which is
        ' why we have the error handler above
    End Select

    Set objApp = Nothing
End Function

