Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Cancel = True
        If Item.Attachments.Count > 0 Then
            answer = MsgBox("Сообщение содержит вложенные файлы. Неархивированные вложения будут удалены. Отправить?", vbYesNo)
            If answer = vbNo Then
                Cancel = True
            Else
                ' Удаление всех вложений кроме 7z, png
                For i = Item.Attachments.Count To 1 Step -1
                    box = MsgBox("file:" + Item.Attachments.Item(i).FileName, vbOKOnly)
                    If InStr(".7z png zip rar", LCase(Right(Item.Attachments.Item(i).FileName, 3))) > 0 Then
                        Item.Attachments.Item(i).Delete
                    End If
                Next i
            End If
        End If
End Sub
