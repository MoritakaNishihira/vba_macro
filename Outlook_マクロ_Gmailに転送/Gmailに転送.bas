Attribute VB_Name = "Module1"
Const DIC_KEY_KIKANINFO_KIKAN_FROM = "from"
Const DIC_KEY_KIKANINFO_KIKAN_TO = "to"
Const DIC_KEY_KIKANINFO_KIKAN_MONTHS = "months"

Const DIC_KEY_MAILBOXTYPE_PRIVATE = "private"
Const DIC_KEY_MAILBOXTYPE_SHARED = "shared"

Const SHARED_MAILBOX_NAME = "RICOH EDW設定代行サービス"

Const PROCESS_SKIP_STRING_MAIL_SUBJECT_1 = "【DocuWare保守 問合せ】1"
Const PROCESS_SKIP_STRING_MAIL_SUBJECT_2 = "【DocuWare保守 問合せ】2"
Const PROCESS_SKIP_STRING_MAIL_SUBJECT_3 = "【DocuWare保守 問合せ】3"

Const EXCEL_HEADER_NAME_NO = "No"
Const EXCEL_HEADER_NAME_SENTDATE = "送信日時"
Const EXCEL_HEADER_NAME_SENDER_NAME = "差出人"
Const EXCEL_HEADER_NAME_SENDER_MAILADDRESS = "差出人アドレス"
Const EXCEL_HEADER_NAME_TO = "宛先"
Const EXCEL_HEADER_NAME_CC = "CC"
Const EXCEL_HEADER_NAME_SUBJECT = "件名"

Const Excel_FILE_NAME_TEMPLATE = "DW_{0}_送信先リスト.xlsx"

Sub 送信済みアイテムのリスト化()

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

    Dim kikan_def As String
    kikan_def = Format(Date, "yyyy/MM")

    Dim kikan As String
    kikan = InputBox("出力年月（開始月）をyyyy/MM形式で入力してください", "出力年月（開始月）を入力してください", kikan_def)

    If kikan = "" Then
        Set getKikanInfo = Nothing
        Exit Function
    End If

    Dim targetMonths As String
    targetMonths = InputBox("何か月分を取得しますか？" & vbCrLf & "入力例：1→1カ月　0→出力年月から全期間", , "1")

    If targetMonths = "" Then
        Set getKikanInfo = Nothing
        Exit Function
    End If

    Dim kikan_From As Date
    Dim kikan_To As Date
    kikan_From = CDate(Format(kikan & "/01", "yyyy/MM/dd"))

    If targetMonths = 0 Then
        kikan_To = CDate("2100/12/31")
    Else
        kikan_To = DateAdd("d", -1, DateAdd("m", CInt(targetMonths), CDate(kikan_From)))
    End If

    dic.Add DIC_KEY_KIKANINFO_KIKAN_FROM, kikan_From
    dic.Add DIC_KEY_KIKANINFO_KIKAN_TO, kikan_To
    dic.Add DIC_KEY_KIKANINFO_KIKAN_MONTHS, targetMonths

    Set getKikanInfo = dic
End Function

Private Sub mainProcess(ByVal kikan_Info As Object)
On Error GoTo ErrorHandler

    Dim skipStrings() As String
    ReDim skipStrings(2)
    skipStrings(0) = PROCESS_SKIP_STRING_MAIL_SUBJECT_1
    skipStrings(1) = PROCESS_SKIP_STRING_MAIL_SUBJECT_2
    skipStrings(2) = PROCESS_SKIP_STRING_MAIL_SUBJECT_3

    Dim objOutlook As Outlook.Application
    Dim nameSpase As Outlook.NameSpace

    Dim privateSentMailItems As Outlook.Folder

    Dim sharedMailBox As Outlook.Store
    Dim sharedSentMailItems As Outlook.Folder

    ' Outlook の Application オブジェクトを取得
    Set objOutlook = New Outlook.Application

    ' プライベートメールボックスを取得
    Set nameSpase = objOutlook.GetNamespace("MAPI")
    Set privateSentMailItems = nameSpase.GetDefaultFolder(olFolderSentMail)

    ' 共有メールボックスを取得
    Set sharedMailBox = objOutlook.Session.Stores.item(SHARED_MAILBOX_NAME)
    Set sharedSentMailItems = sharedMailBox.GetDefaultFolder(olFolderSentMail)

    Dim sentMailItems As Object
    Set sentMailItems = CreateObject("Scripting.Dictionary")
    sentMailItems.Add DIC_KEY_MAILBOXTYPE_PRIVATE, privateSentMailItems
    sentMailItems.Add DIC_KEY_MAILBOXTYPE_SHARED, sharedSentMailItems

    Dim excelData As Object
    Set excelData = CreateObject("Scripting.Dictionary")

    ' プライベートの送信済みアイテムと共有メールボックスの送信済みアイテムを処理
    Dim mailBoxType As Variant
    For Each mailBoxType In sentMailItems.Keys
        Dim sentMailItem As Object
        Set sentMailItem = sentMailItems(mailBoxType)

        Dim data As String
        'ヘッダー
        data = initExcelHeader()

        Dim no As String
        'No列
        no = 1

        ' フォルダー内のアイテムをすべて処理
        Dim mailItem As Variant
        For Each mailItem In sentMailItem.Items

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

            '件名に特定の文字が含まれていたらスキップ
            If isSome(mailItem.Subject, skipStrings) Then
                GoTo ContinueFor
            End If

            data = data & no & vbTab
            data = data & mailItem.SentOn & vbTab
            data = data & mailItem.SenderName & vbTab
            data = data & mailItem.SenderEmailAddress & vbTab

            Dim dics() As Object

            'to
            Dim recipient As Variant
            Dim idx As Integer

            idx = 0
            For Each recipient In mailItem.Recipients
                If recipient.Type = OlMailRecipientType.olTo Or recipient.Type = OlMailRecipientType.olCC Then
                    ReDim Preserve dics(idx)
                    Set dics(idx) = GetRecipientEmailAddress(recipient.AddressEntry, recipient.Type)
                    idx = idx + 1
                End If
            Next

            Dim dic As Variant
            Dim concat_to As String
            Dim concat_cc As String

            For Each dic In dics
                Dim key As Variant
                For Each key In dic.Keys
                    If key = "To" Then
                        concat_to = concat_to & dic("To") & " , "
                    ElseIf key = "CC" Then
                        concat_cc = concat_cc & dic("CC") & " , "
                    End If
                Next
            Next

            data = data & concat_to & vbTab
            data = data & concat_cc & vbTab

            data = data & mailItem.Subject & vbCrLf

            no = no + 1
ContinueFor:

        Next
        excelData.Add mailBoxType, data
    Next

    Call excelOut(excelData, kikan_Info)

    Call MsgBox("処理完了")

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

Private Function isSome(ByVal targetStr As String, ByRef skipStrings() As String)
    isSome = False

    Dim skipString As Variant
    For Each skipString In skipStrings
        If targetStr Like "*" & skipString & "*" Then
            isSome = True
            Exit For
        End If
    Next
End Function

Private Sub excelOut(ByVal excelData As Object, ByVal kikan_Info As Object)

    'Excelに出力
    Dim xlsObj As Object
    Dim xlsBook As Object
    Dim xlsSheet As Object
    Set xlsObj = CreateObject("Excel.Application")
    Set xlsBook = xlsObj.Workbooks.Add

    Dim key As Variant
    For Each key In excelData

        Dim data As String
        data = excelData(key)

        If key = DIC_KEY_MAILBOXTYPE_PRIVATE Then
            Set xlsSheet = xlsBook.Worksheets(1)
            xlsSheet.Name = "個人アドレス"
        ElseIf key = DIC_KEY_MAILBOXTYPE_SHARED Then
            xlsBook.Worksheets.Add after:=xlsBook.Worksheets(1)
            Set xlsSheet = xlsBook.Worksheets(2)
            xlsSheet.Name = "共有メールボックス"
        Else
            GoTo ExitFor
        End If
        xlsSheet.cells(1, 1).Select

        'クリップボードに一次的に保持
        Dim cb As Object
        Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        With cb
            .setText data
            .PutInClipboard
        End With

        xlsSheet.Paste
        xlsSheet.cells(1, 1).Select
    Next

ExitFor:

    xlsBook.Worksheets(1).Select

    Dim nengetsu As Date
    nengetsu = kikan_Info(DIC_KEY_KIKANINFO_KIKAN_FROM)

    '保存（自身のダウンロードフォルダ）
    Dim xlsFileName As String
    xlsFileName = Replace(Excel_FILE_NAME_TEMPLATE, "{0}", Format(nengetsu, "yyyy年MM月度"))

    Dim myDownLoadFolderPath As String
    myDownLoadFolderPath = Environ("UserProfile") & "\" & "Downloads"

    Dim saveFileFullPath As String
    saveFileFullPath = myDownLoadFolderPath & "\" & xlsFileName

    Call xlsBook.SaveAs(saveFileFullPath)

    xlsBook.Close
    Set xlsSheet = Nothing
    Set xlsBook = Nothing
    Set xlsObj = Nothing

End Sub

Private Function GetRecipientEmailAddress(ByVal oAddressEntry As Object, ByVal mailRecipientType As Integer) As Object

    Dim oSender As Object
    Dim oExUser As Object

    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    If mailRecipientType = 1 Then
        If oAddressEntry.AddressEntryUserType = olExchangeUserAddressEntry _
            Or oAddressEntry.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            Set oExUser = oAddressEntry.GetExchangeUser
            dic.Add "To", oExUser.PrimarySmtpAddress

        ElseIf oAddressEntry.AddressEntryUserType = olSmtpAddressEntry Then
            dic.Add "To", oAddressEntry.Address

        End If
    ElseIf mailRecipientType = 2 Then
        If oAddressEntry.AddressEntryUserType = olSmtpAddressEntry Then
            dic.Add "CC", oAddressEntry.Address

        End If
    End If

    Set GetRecipientEmailAddress = dic
End Function

Private Function initExcelHeader()
    Dim result As String
    result = ""

    result = result & EXCEL_HEADER_NAME_NO & vbTab
    result = result & EXCEL_HEADER_NAME_SENTDATE & vbTab
    result = result & EXCEL_HEADER_NAME_SENDER_NAME & vbTab
    result = result & EXCEL_HEADER_NAME_SENDER_MAILADDRESS & vbTab
    result = result & EXCEL_HEADER_NAME_TO & vbTab
    result = result & EXCEL_HEADER_NAME_CC & vbTab
    result = result & EXCEL_HEADER_NAME_SUBJECT & vbCrLf

    initExcelHeader = result
End Function
