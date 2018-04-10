Private Function GetNextWorkday() As Date
  Dim d As Date
  d = Date
  Select Case Weekday(d, vbMonday)
    Case 6: d = DateAdd("d", 2, d)
    Case 7: d = DateAdd("d", 1, d)
    Case Else: d = DateAdd("d", Abs((Time > TimeValue("6:00 PM")) * 1), d)
  End Select
  GetNextWorkday = d
End Function
Private Function IsWorkDay()
  IsWorkDay = True
  Select Case Weekday(Date, vbMonday)
    Case 6: IsWorkDay = False
    Case 7: IsWorkDay = False
  End Select
End Function
Private Function IsWorkHour()
  IsWorkHour = True
  If (Time > TimeValue("6:00 PM")) Or (Time < TimeValue("8:30 AM")) Then IsWorkHour = False
End Function
Private Sub ShowOptionDialog(oMail As Outlook.MailItem)
  Dim oBars As Office.CommandBars
  Dim oBtn As Office.CommandBarButton

  Set oBars = oMail.GetInspector.CommandBars
  Set oBtn = oBars.FindControl(, 5598)
  oBtn.Execute
End Sub
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error Resume Next
    Early = False
    DelayDay = 0
    lbadFound = False
    lijst = ""
    ' to prevent missing out when an error is raised:
    splits = 1
    Set Recipients = Item.Recipients
    receivers = Recipients.Count
    For i = receivers To 1 Step -1
        Set recip = Recipients.Item(i)
        If InStr(1, LCase(recip), "lists.vuw.ac.nz") >= 1 Then
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

    If Not (IsWorkDay()) Or Not (IsWorkHour()) Then
        Antwoord = MsgBox("Do you want to delay sending this mail?", vbYesNoCancel)
        Select Case Antwoord
            Case vbCancel
                Cancel = True
            Case vbYes
                Item.Categories = "delayed"
                Mail.DeferredDeliveryTime = GetNextWorkday() & " 08:30 AM"
                MsgBox ("Delayed until " & GetNextWorkday() & " 08:30 AM")
        End Select
    End If
  End If
End Sub
