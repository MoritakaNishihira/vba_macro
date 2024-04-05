Const DIC_KEY_KIKANINFO_KIKAN_FROM = "from"
Const DIC_KEY_KIKANINFO_KIKAN_TO = "to"
Const DIC_KEY_KIKANINFO_KIKAN_MONTHS = "months"

Sub Gmailに転送()
    Dim dicObj As Object
    Set dicObj = CreateObject("Scripting.Dictionary")

    Dim kikan_Info As Object
    Set kikan_Info = getKikanInfo()

    If kikan_Info Is Nothing Then
        Exit Sub
    End If

    Call mainProcess(kikan_Info)

End Sub

Private Function getKikanInfo() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim kikan As String
    kikan = Format(Date, "yyyy/MM/dd")

    Dim targetMonths As String
    targetMonths = "6"

    Dim kikan_From As Date
    Dim kikan_To As Date
    kikan_To = Format(Date, "yyyy/MM/dd")
    kikan_From = DateAdd("m", targetMonths * -1, kikan_To)

    dic.Add DIC_KEY_KIKANINFO_KIKAN_FROM, kikan_From
    dic.Add DIC_KEY_KIKANINFO_KIKAN_TO, kikan_To
    dic.Add DIC_KEY_KIKANINFO_KIKAN_MONTHS, targetMonths

    Set getKikanInfo = dic
End Function

Private Sub mainProcess(ByVal kikan_Info As Object)
On Error GoTo ErrorHandler

    Dim forwardCount As Integer
    forwardCount = 1

    ' Outlook の Application オブジェクトを取得
    Dim objOutlook As Outlook.Application
    Dim nameSpase As Outlook.NameSpace
    Set objOutlook = New Outlook.Application

    ' プライベートメールボックスを取得
    Dim inboxMailItems As Outlook.Folder
    Set nameSpase = objOutlook.GetNamespace("MAPI")
    Set inboxMailItems = nameSpase.GetDefaultFolder(olFolderInbox)

    ' フォルダー内のアイテムをすべて処理
    Dim mailItem As Variant
    For Each mailItem In inboxMailItems.Items

        'MailItem以外は対象外
        If TypeName(mailItem) <> "MailItem" Then
            GoTo ContinueFor
        End If

        '期間外はスキップ
        Dim targetDate As Date
        targetDate = CDate(Format(CDate(mailItem.SentOn), "yyyy/MM/dd"))
        If isKikangai(targetDate, kikan_Info) Then
            GoTo ContinueFor
        End If

        '対象のみ
        If isTarget(mailItem.Sender.Address) Then
            'メール転送
            Call forwardMail(mailItem)
            forwardCount = forwardCount + 1
        End If

ContinueFor:

    Next

    Call MsgBox("処理完了 転送件数： " & forwardCount & " 件")

    GoTo Finally

ErrorHandler:
    MsgBox "[No:" & Err.Number & "]" & Err.Description, vbCritical & vbOKOnly, "エラー"
    Resume Finally

Finally:

End Sub

Private Function isKikangai(ByVal targetDate As Date, ByVal kikan_Info As Object)
    isKikangai = False

    Dim kikan_From As Date
    Dim kikan_To As Date

    kikan_From = kikan_Info(DIC_KEY_KIKANINFO_KIKAN_FROM)
    kikan_To = kikan_Info(DIC_KEY_KIKANINFO_KIKAN_TO)

    If targetDate < kikan_From Or targetDate > kikan_To Then
        isKikangai = True
    End If
End Function

Private Function isTarget(ByVal targetStr As String)
    Dim forwardTagetAccounts()
    ReDim forwardTagetAccounts(26)
    forwardTagetAccounts(0) = "mail@contact.vpass.ne.jp"
    forwardTagetAccounts(1) = "statement@vpass.ne.jp"
    forwardTagetAccounts(2) = "magazine@member.startheaters.jp"
    forwardTagetAccounts(3) = "info@okinawa-basketball.jp"
    forwardTagetAccounts(4) = "helpdesk@j-com.co.jp"
    forwardTagetAccounts(5) = "no-reply-gig@demae-can.co.jp"
    forwardTagetAccounts(6) = "noreply@email.apple.com"
    forwardTagetAccounts(7) = "post_master@netbk.co.jp"
    forwardTagetAccounts(8) = "emagazine@daiichilife.com"
    forwardTagetAccounts(9) = "moritaka_nishihira@pvcjp.com"
    forwardTagetAccounts(10) = "donotreply@psrv.jp"
    forwardTagetAccounts(11) = "microsoft-noreply@microsoft.com"
    forwardTagetAccounts(12) = "info-api@ryugin.co.jp"
    forwardTagetAccounts(13) = "media@ryugin.co.jp"
    forwardTagetAccounts(14) = "autoreply@shop.bleague-info.jp"
    forwardTagetAccounts(15) = "direct@ryugin.co.jp"
    forwardTagetAccounts(16) = "mobileidticket@psrv.jp"
    forwardTagetAccounts(17) = "info@prepaid.smbc-card.com"
    forwardTagetAccounts(18) = "info-api@ryugin.co.jp"
    forwardTagetAccounts(19) = "info@ma.axa-direct.co.jp"
    forwardTagetAccounts(20) = "ticket@ml.smart-theater.com"
    forwardTagetAccounts(21) = "okigin@ib.finemax.net"
    forwardTagetAccounts(22) = "okigin@ib.finemax.net"
    forwardTagetAccounts(23) = "order_acknowledgment@orders.apple.com"
    forwardTagetAccounts(24) = "no-reply@accounts.google.com"

    isTarget = False

    Dim t As Variant
    For Each t In forwardTagetAccounts
        If targetStr = t Then
            isTarget = True
            Exit For
        End If
    Next
End Function

Sub forwardMail(ByVal mailItem As Object)

    Dim fwMail As Object
    Set fwMail = mailItem.Forward

    fwMail.To = "moritaka.nishihira@gmail.com"
    fwMail.Subject = "【転送】" & mailItem.Subject

    fwMail.Send

End Sub