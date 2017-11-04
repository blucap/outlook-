Private Function GetNextWorkday(ByVal MinDaysOffset As Long) As Date
  Dim d As Date
  d = DateAdd("d", MinDaysOffset, Date)
  Select Case Weekday(d, vbMonday)
  Case 6: d = DateAdd("d", 2, d)
  Case 7: d = DateAdd("d", 1, d)
  End Select
  GetNextWorkday = d
End Function

Private Function IsWorkDay()
  IsWorkDay = True
  Select Case Weekday(Date, vbMonday)
    Case 6: IsWorkDay = False
    Case 7: IsWorkDay = False
  End Select
  ' feel free to adjust your working hours 
  If Hour(Time) > 18 Or Hour(Time) < 8 Then IsWorkDay = False
End Function

' This sub may not work, but it does not do harm either
Private Sub ShowOptionDialog(oMail As Outlook.MailItem)
  Dim oBars As Office.CommandBars
  Dim oBtn As Office.CommandBarButton
  Set oBars = oMail.GetInspector.CommandBars
  Set oBtn = oBars.FindControl(, 5598)
  oBtn.Execute
End Sub

'Main routine
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error Resume Next
    lbadFound = False
    lijst = ""
    ' to prevent missing out when an error is raised:
    splits = 1
    Set Recipients = Item.Recipients
    receivers = Recipients.Count
    For i = receivers To 1 Step -1
        Set recip = Recipients.Item(i)
        ' Add your list here
        If InStr(1, LCase(recip), "lists.my_list_here.com") >= 1 Then
            lbadFound = True
        End If
        If i = receivers Then
            lijst = recip
        Else
            lijst = lijst & ", " & recip
        End If
    Next i
    splits = UBound(Split(Item.To, ";"))
    If receivers > 1 Or lbadFound Or splits > 0 Then
        If MsgBox("Do you want to continue sending the mail to: " & lijst & "?", vbOKCancel) <> vbOK Then
            Cancel = True
        End If
    End If

    If TypeOf Item Is Outlook.MailItem Then
    ShowOptionDialog Item
    Set Mail = Item
    If Not (IsWorkDay()) And MsgBox("Do you want to expedite sending this mail ?", vbYesNoCancel) <> vbYes Then
        Item.Categories = "delayed"
    End If
    If LCase(Mail.Categories) = "delayed" Then
      Mail.DeferredDeliveryTime = GetNextWorkday(1) & " 08:30 AM"
      MsgBox ("Delayed until " & GetNextWorkday(1) & " 08:30 AM")
    End If
  End If 
End Sub

