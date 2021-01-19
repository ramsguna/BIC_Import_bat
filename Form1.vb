Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim SqlCmd1 As SqlClient.SqlCommand
    Dim DaList1 = New SqlClient.SqlDataAdapter
    Dim DsList1, DsList2, DsList3 As New DataSet
    Dim WK_DsList1 As New DataSet
    Dim DtView1, DtView2 As DataView
    Dim WK_DtView1, WK_DtView2, WK_DtView3 As DataView
    Dim DtView_A, DtView_C As DataView
    Dim dttable As DataTable
    Dim dtRow As DataRow

    Dim strSQL, strDATA(40), Err_F As String
    Dim folder, filename, strTEL As String
    Dim i, j, pos, len, cnt, kensuu, Skip As Integer
    Dim WK_date As Date

    Dim DownLordFld As String = "C:\ftproot"            '取込み元
    Dim DownLordFld2 As String = "C:\ftproot\log"       '取込み済み移動先
    'Dim DownLordFld As String = "D:\ftproot"            '取込み元
    'Dim DownLordFld2 As String = "D:\ftproot\log"       '取込み済み移動先
    Dim strFile, strFile2 As String


#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更してください。  
    ' コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.ClientSize = New System.Drawing.Size(194, 39)
        Me.ControlBox = False
        Me.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data import"

    End Sub

#End Region

    '起動時
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'On Error GoTo err
        Call DB_INIT()
        DB_OPEN()

        'イベントログ書き込み
        System.Diagnostics.EventLog.WriteEntry("BIC Import START", "データの取り込みを開始しました。", System.Diagnostics.EventLogEntryType.Information)
        cnt = 0

        If System.IO.Directory.Exists(DownLordFld) = True Then
            For Each strFile In System.IO.Directory.GetFiles(DownLordFld, "*.*")
                filename = strFile.Substring(strFile.LastIndexOf("\") + 1)
                'MsgBox(filename)

                WK_date = "20" & Mid(filename, 7, 2) & "/" & Mid(filename, 9, 2) & "/" & Mid(filename, 11, 2)
                'WK_date = "20" & Mid(filename, 6, 2) & "/" & Mid(filename, 8, 2) & "/" & Mid(filename, 10, 2)

                '重複取込みCheck
                DsList1.Clear()
                strSQL = "SELECT Inport_File FROM Inport_log WHERE (Inport_File = '" & filename & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                DaList1.SelectCommand = SqlCmd1
                SqlCmd1.CommandTimeout = 600
                DaList1.Fill(DsList1, "Inport_log")
                DtView1 = New DataView(DsList1.Tables("Inport_log"), "", "", DataViewRowState.CurrentRows)
                If DtView1.Count = 0 Then
                    Call read()             'テキスト取込み開始
                    System.Diagnostics.EventLog.WriteEntry("BIC Import END " & filename & "  " & DtView2.Count & " 件", "データの取り込みを完了しました。", System.Diagnostics.EventLogEntryType.Information)
                    If DtView2.Count <> kensuu Then Err_output("2", DtView2.Count, kensuu) 'エラー出力
                    If Skip <> 0 Then Err_output("4", Skip, DtView2.Count) ' : Err_List() 'エラー出力
                Else
                    System.Diagnostics.EventLog.WriteEntry("BIC Import Warning", filename & "は既に取込み済みです。", System.Diagnostics.EventLogEntryType.Warning)
                    Err_output("5", 0, 0)  'エラー出力
                End If

                'ﾌｧｲﾙ移動
                strFile2 = strFile.Replace(".", "") '.を取る
                System.IO.File.Move(strFile, DownLordFld2 & Mid(strFile2, DownLordFld.Length + 1, 100) & "_" & Format(Now, "yyyyMMddhhmmss"))
            Next
        End If

        DB_CLOSE()
        System.Diagnostics.EventLog.WriteEntry("BIC Import END", "", System.Diagnostics.EventLogEntryType.Information)
        Application.Exit()
        Exit Sub
err:
        'イベントログ書き込み
        If Err.Number = 57 Then    '57はコピー中
            System.Diagnostics.EventLog.WriteEntry("BIC Import Warning", "データの取り込みに失敗しました。：" & Err.Number & ":" & Err.Description, System.Diagnostics.EventLogEntryType.Warning)
        Else
            Err_output("1", 0, 0)  'エラー出力
            System.Diagnostics.EventLog.WriteEntry("BIC Import Error", "データの取り込みに失敗しました。：" & Err.Number & ":" & Err.Description, System.Diagnostics.EventLogEntryType.Error)
        End If
        Application.Exit()
    End Sub

    Sub read()          'テキスト取込み開始
        Call txt_all_clr()                  'txt_data_all_tempをクリア
        kensuu = 0
        Skip = 0


        Dim srFile As New System.IO.StreamReader(strFile, System.Text.Encoding.Default)
        Dim strLine As String = srFile.ReadLine()
        While Not strLine Is Nothing
            If Mid(strLine, 9, 20) <> "00000000000000000000" Then

                strLine = strLine.Replace("'", "`") 'ｼﾝｸﾞﾙｺｰﾃｰｼｮﾝを`に置き換え

                pos = 1 : len = 8 : strDATA(1) = MidB(strLine, pos, len)
                pos = pos + len : len = 14 : strDATA(2) = MidB(strLine, pos, len)
                pos = pos + len : len = 6 : strDATA(3) = MidB(strLine, pos, len)
                pos = pos + len : len = 4 : strDATA(4) = MidB(strLine, pos, len)
                pos = pos + len : len = 13 : strDATA(5) = MidB(strLine, pos, len)
                pos = pos + len : len = 48 : strDATA(6) = MidB(strLine, pos, len)
                pos = pos + len : len = 6 : strDATA(7) = MidB(strLine, pos, len)
                pos = pos + len : len = 50 : strDATA(8) = MidB(strLine, pos, len)
                pos = pos + len : len = 4 : strDATA(9) = MidB(strLine, pos, len)
                pos = pos + len : len = 50 : strDATA(10) = MidB(strLine, pos, len)
                pos = pos + len : len = 9 : strDATA(11) = MidB(strLine, pos, len)
                pos = pos + len : len = 9 : strDATA(12) = MidB(strLine, pos, len)
                pos = pos + len : len = 2 : strDATA(13) = MidB(strLine, pos, len)
                pos = pos + len : len = 2 : strDATA(14) = MidB(strLine, pos, len)
                pos = pos + len : len = 8 : strDATA(15) = MidB(strLine, pos, len)
                pos = pos + len : len = 6 : strDATA(16) = MidB(strLine, pos, len)
                pos = pos + len : len = 13 : strDATA(17) = MidB(strLine, pos, len)
                pos = pos + len : len = 30 : strDATA(18) = MidB(strLine, pos, len)
                pos = pos + len : len = 3 : strDATA(19) = MidB(strLine, pos, len)
                pos = pos + len : len = 4 : strDATA(20) = MidB(strLine, pos, len)
                pos = pos + len : len = 60 : strDATA(21) = MidB(strLine, pos, len)
                pos = pos + len : len = 60 : strDATA(22) = MidB(strLine, pos, len)
                pos = pos + len : len = 1 : strDATA(23) = MidB(strLine, pos, len)
                pos = pos + len : len = 8 : strDATA(24) = MidB(strLine, pos, len)
                pos = pos + len : len = 25 : strDATA(25) = MidB(strLine, pos, len)
                pos = pos + len : len = 25 : strDATA(26) = MidB(strLine, pos, len)

                Call F_Check()
                If Err_F = "1" Then
                    Skip = Skip + 1
                    Call inport_txt_err_log()   '項目エラーでLOG出力
                End If

                Call inport_txt_all_temp()  'テキストデータをそのまま取込み
                Call Master_ADD()           'マスタデータ追加
            Else
                kensuu = kensuu + CInt(Mid(strLine, 33, 9))
            End If
            strLine = srFile.ReadLine()
        End While
        srFile.Close()

        Call Inport_Log()                   'インポートLOG
        Call LAST_IMPORT_FILE()             '最終取込みファイル名保存
        Call inport()                       '取込み

    End Sub

    Sub F_Check()
        Err_F = "0"

        '保証加入日
        If IsDate(Mid(strDATA(1), 1, 4), Mid(strDATA(1), 5, 2), Mid(strDATA(1), 7, 2)) = False Then
            Err_F = "1" : Exit Sub
        End If

        '実売価格
        If numeric_check(strDATA(11)) = "NG" Then
            Err_F = "1" : Exit Sub
        End If

        '保証料
        If numeric_check(strDATA(12)) = "NG" Then
            Err_F = "1" : Exit Sub
        End If

        '作成日
        If strDATA(15) <> "00000000" And strDATA(15) <> "        " Then
            If IsDate(Mid(strDATA(15), 1, 4), Mid(strDATA(15), 5, 2), Mid(strDATA(15), 7, 2)) = False Then
                Err_F = "1" : Exit Sub
            End If
        End If

        '締め月
        If strDATA(16) <> "000000" And strDATA(16) <> "      " Then
            If IsDate(Mid(strDATA(16), 1, 4), Mid(strDATA(16), 5, 6), "01") = False Then
                Err_F = "1" : Exit Sub
            End If
        End If

        ''誕生日
        'If strDATA(24) <> "00000000" Then
        '    If IsDate(Mid(strDATA(24), 1, 4), Mid(strDATA(24), 5, 2), Mid(strDATA(24), 7, 2)) = False Then
        '        Err_F = "1" : Exit Sub
        '    End If
        'End If

    End Sub

    Sub inport()        '取込み

        DsList2.Clear()
        strSQL = "SELECT * FROM txt_data_all_temp"
        'strSQL = strSQL & " ORDER BY WRN_DATE, SALE_STS"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 600
        DaList1.Fill(DsList2, "txt_data_all_temp")
        DtView2 = New DataView(DsList2.Tables("txt_data_all_temp"), "", "", DataViewRowState.CurrentRows)
        If DtView2.Count <> 0 Then

            For i = 0 To DtView2.Count - 1

                DsList1.Clear()
                strSQL = "SELECT * FROM txt_data_all WHERE (WRN_NO = '" & DtView2(i)("WRN_NO") & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                DaList1.SelectCommand = SqlCmd1
                SqlCmd1.CommandTimeout = 600
                DaList1.Fill(DsList1, "txt_data_all")
                WK_DtView1 = New DataView(DsList1.Tables("txt_data_all"), "", "", DataViewRowState.CurrentRows)

                If DtView2(i)("SALE_STS") = "00" Then  'Ａデータ
                    If WK_DtView1.Count = 0 Then
                        Call add_AC()       'データ取込み
                    Else
                        DtView_A = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '00'", "", DataViewRowState.CurrentRows)
                        DtView_C = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '09'", "", DataViewRowState.CurrentRows)
                        If DtView_A.Count >= DtView_C.Count Then
                            Call upd_A()    '余ったＡデータで更新
                        Else
                            WK_DsList1.Clear()
                            strSQL = "SELECT WRN_NO FROM WRN_DATA WHERE (WRN_NO = '" & DtView2(i)("WRN_NO") & "')"
                            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                            DaList1.SelectCommand = SqlCmd1
                            SqlCmd1.CommandTimeout = 600
                            DaList1.Fill(WK_DsList1, "WRN_DATA")
                            WK_DtView1 = New DataView(WK_DsList1.Tables("WRN_DATA"), "", "", DataViewRowState.CurrentRows)
                            If WK_DtView1.Count = 0 Then
                                add_C()         'Ｃデータで取込み
                            End If
                        End If
                    End If
                Else        'Ｃデータ
                    If WK_DtView1.Count = 0 Then
                        Call add_AC()       'データ取込み
                    Else
                        DtView_A = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '00'", "", DataViewRowState.CurrentRows)
                        DtView_C = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '09'", "", DataViewRowState.CurrentRows)
                        If DtView_A.Count - 1 <= DtView_C.Count Then
                            Call upd_C()    'Ｃデータで更新
                        Else
                            Call upd_A()    '余ったＡデータで更新
                        End If
                    End If
                End If
                Call txt_all_copy()         'txt_data_all_tempからコピー
            Next
        End If

    End Sub

    Sub add_AC()                    'データ取込み
        DsList3.Clear()
        strSQL = "SELECT WRN_NO FROM WRN_DATA"
        strSQL = strSQL & " WHERE (WRN_NO = '" & DtView2(i)("WRN_NO") & "')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 600
        DaList1.Fill(DsList3, "WRN_DATA")
        WK_DtView3 = New DataView(DsList3.Tables("WRN_DATA"), "", "", DataViewRowState.CurrentRows)
        If WK_DtView3.Count = 0 Then
            strSQL = "INSERT INTO WRN_DATA"
            strSQL = strSQL & " (WRN_DATE, WRN_NO, SHOP_CODE, ITEM_CODE, MODEL, CAT_CODE"
            strSQL = strSQL & ", CAT_NAME, MKR_CODE, MKR_NAME, PRICE, WRN_PRICE, WRN_PRD"
            strSQL = strSQL & ", SALE_STS, CRT_DATE, CLS_MNTH, PNT_NO, CUST_NAME, ZIP1"
            strSQL = strSQL & ", ZIP2, ADRS1, ADRS2, SEX, BRTH_DATE, TEL_NO, CNT_NO, CXL_DATE)"
            If IsDate(Mid(DtView2(i)("WRN_DATE"), 1, 4), Mid(DtView2(i)("WRN_DATE"), 5, 2), Mid(DtView2(i)("WRN_DATE"), 7, 2)) = True Then
                strSQL = strSQL & " VALUES ('" & Mid(DtView2(i)("WRN_DATE"), 1, 4) & "/" & Mid(DtView2(i)("WRN_DATE"), 5, 2) & "/" & Mid(DtView2(i)("WRN_DATE"), 7, 2) & "'"
            Else
                strSQL = strSQL & " VALUES (NULL"
            End If
            strSQL = strSQL & ", '" & DtView2(i)("WRN_NO") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("SHOP_CODE") & "', '" & DtView2(i)("ITEM_CODE") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("MODEL") & "', '" & DtView2(i)("CAT_CODE") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("CAT_NAME") & "', '" & DtView2(i)("MKR_CODE") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("MKR_NAME") & "'"
            strSQL = strSQL & ", " & CInt(DtView2(i)("PRICE")) & ", " & CInt(DtView2(i)("WRN_PRICE"))
            strSQL = strSQL & ", '" & DtView2(i)("WRN_PRD") & "', '" & DtView2(i)("SALE_STS") & "'"
            If IsDate(Mid(DtView2(i)("CRT_DATE"), 1, 4), Mid(DtView2(i)("CRT_DATE"), 5, 2), Mid(DtView2(i)("CRT_DATE"), 7, 2)) = True Then
                strSQL = strSQL & ", '" & Mid(DtView2(i)("CRT_DATE"), 1, 4) & "/" & Mid(DtView2(i)("CRT_DATE"), 5, 2) & "/" & Mid(DtView2(i)("CRT_DATE"), 7, 2) & "'"
            Else
                strSQL = strSQL & ", NULL"
            End If
            If IsDate(Mid(DtView2(i)("CLS_MNTH"), 1, 4), Mid(DtView2(i)("CLS_MNTH"), 5, 2), "01") = True Then
                strSQL = strSQL & ", '" & Mid(DtView2(i)("CLS_MNTH"), 1, 4) & "/" & Mid(DtView2(i)("CLS_MNTH"), 5, 2) & "/01'"
            Else
                strSQL = strSQL & ", NULL"
            End If
            strSQL = strSQL & ", '" & DtView2(i)("PNT_NO") & "', '" & DtView2(i)("CUST_NAME") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("ZIP1") & "', '" & DtView2(i)("ZIP2") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("ADRS1") & "', '" & DtView2(i)("ADRS2") & "'"
            strSQL = strSQL & ", '" & DtView2(i)("SEX") & "'"
            If IsDate(Mid(DtView2(i)("BRTH_DATE"), 1, 4), Mid(DtView2(i)("BRTH_DATE"), 5, 2), Mid(DtView2(i)("BRTH_DATE"), 7, 2)) = True Then
                strSQL = strSQL & ", '" & Mid(DtView2(i)("BRTH_DATE"), 1, 4) & "/" & Mid(DtView2(i)("BRTH_DATE"), 5, 2) & "/" & Mid(DtView2(i)("BRTH_DATE"), 7, 2) & "'"
            Else
                strSQL = strSQL & ", NULL"
            End If
            strTEL = Trim(DtView2(i)("TEL_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
            strSQL = strSQL & ", '" & strTEL & "'"
            strTEL = Trim(DtView2(i)("CNT_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
            strSQL = strSQL & ", '" & strTEL & "'"
            If DtView2(i)("SALE_STS") = "00" Then
                strSQL = strSQL & ", NULL)"
            Else
                If IsDate(Mid(DtView2(i)("WRN_DATE"), 1, 4), Mid(DtView2(i)("WRN_DATE"), 5, 2), Mid(DtView2(i)("WRN_DATE"), 7, 2)) = True Then
                    strSQL = strSQL & ", '" & Mid(DtView2(i)("WRN_DATE"), 1, 4) & "/" & Mid(DtView2(i)("WRN_DATE"), 5, 2) & "/" & Mid(DtView2(i)("WRN_DATE"), 7, 2) & "')"
                Else
                    strSQL = strSQL & ", NULL)"
                End If
            End If
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.CommandTimeout = 600
            SqlCmd1.ExecuteNonQuery()
        Else
            strSQL = "UPDATE WRN_DATA"
            If IsDate(Mid(DtView2(i)("WRN_DATE"), 1, 4), Mid(DtView2(i)("WRN_DATE"), 5, 2), Mid(DtView2(i)("WRN_DATE"), 7, 2)) = True Then
                strSQL = strSQL & " SET WRN_DATE = '" & Mid(DtView2(i)("WRN_DATE"), 1, 4) & "/" & Mid(DtView2(i)("WRN_DATE"), 5, 2) & "/" & Mid(DtView2(i)("WRN_DATE"), 7, 2) & "'"
            Else
                strSQL = strSQL & " SET WRN_DATE = NULL"
            End If
            strSQL = strSQL & ", SHOP_CODE = '" & DtView2(i)("SHOP_CODE") & "'"
            strSQL = strSQL & ", ITEM_CODE = '" & DtView2(i)("ITEM_CODE") & "'"
            strSQL = strSQL & ", MODEL = '" & DtView2(i)("MODEL") & "'"
            strSQL = strSQL & ", CAT_CODE = '" & DtView2(i)("CAT_CODE") & "'"
            strSQL = strSQL & ", CAT_NAME = '" & DtView2(i)("CAT_NAME") & "'"
            strSQL = strSQL & ", MKR_CODE = '" & DtView2(i)("MKR_CODE") & "'"
            strSQL = strSQL & ", MKR_NAME = '" & DtView2(i)("MKR_NAME") & "'"
            strSQL = strSQL & ", PRICE = " & CInt(DtView2(i)("PRICE")) & ""
            strSQL = strSQL & ", WRN_PRICE = " & CInt(DtView2(i)("WRN_PRICE")) & ""
            strSQL = strSQL & ", WRN_PRD = '" & DtView2(i)("WRN_PRD") & "'"
            strSQL = strSQL & ", SALE_STS = '" & DtView2(i)("SALE_STS") & "'"
            If IsDate(Mid(DtView2(i)("CRT_DATE"), 1, 4), Mid(DtView2(i)("CRT_DATE"), 5, 2), Mid(DtView2(i)("CRT_DATE"), 7, 2)) = True Then
                strSQL = strSQL & ", CRT_DATE = '" & Mid(DtView2(i)("CRT_DATE"), 1, 4) & "/" & Mid(DtView2(i)("CRT_DATE"), 5, 2) & "/" & Mid(DtView2(i)("CRT_DATE"), 7, 2) & "'"
            Else
                strSQL = strSQL & ", CRT_DATE = NULL"
            End If
            If IsDate(Mid(DtView2(i)("CLS_MNTH"), 1, 4), Mid(DtView2(i)("CLS_MNTH"), 5, 2), Mid(DtView2(i)("CLS_MNTH"), 7, 2)) = True Then
                strSQL = strSQL & ", CLS_MNTH = '" & Mid(DtView2(i)("CLS_MNTH"), 1, 4) & "/" & Mid(DtView2(i)("CLS_MNTH"), 5, 2) & "/01'"
            Else
                strSQL = strSQL & ", CLS_MNTH = NULL"
            End If
            strSQL = strSQL & ", PNT_NO = '" & DtView2(i)("PNT_NO") & "'"
            strSQL = strSQL & ", CUST_NAME = '" & DtView2(i)("CUST_NAME") & "'"
            strSQL = strSQL & ", ZIP1 = '" & DtView2(i)("ZIP1") & "'"
            strSQL = strSQL & ", ZIP2 = '" & DtView2(i)("ZIP2") & "'"
            strSQL = strSQL & ", ADRS1 = '" & DtView2(i)("ADRS1") & "'"
            strSQL = strSQL & ", ADRS2 = '" & DtView2(i)("ADRS2") & "'"
            strSQL = strSQL & ", SEX = '" & DtView2(i)("SEX") & "'"
            If IsDate(Mid(DtView2(i)("BRTH_DATE"), 1, 4), Mid(DtView2(i)("BRTH_DATE"), 5, 2), Mid(DtView2(i)("BRTH_DATE"), 7, 2)) = True Then
                strSQL = strSQL & ", BRTH_DATE = '" & Mid(DtView2(i)("BRTH_DATE"), 1, 4) & "/" & Mid(DtView2(i)("BRTH_DATE"), 5, 2) & "/" & Mid(DtView2(i)("BRTH_DATE"), 7, 2) & "'"
            Else
                strSQL = strSQL & ", BRTH_DATE = NULL"
            End If
            strTEL = Trim(DtView2(i)("TEL_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
            strSQL = strSQL & ", TEL_NO = '" & strTEL & "'"
            strTEL = Trim(DtView2(i)("CNT_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
            strSQL = strSQL & ", CNT_NO = '" & strTEL & "'"
            If DtView2(i)("SALE_STS") = "00" Then
                strSQL = strSQL & ", CXL_DATE = NULL"
            Else
                If IsDate(Mid(DtView2(i)("WRN_DATE"), 1, 4), Mid(DtView2(i)("WRN_DATE"), 5, 2), Mid(DtView2(i)("WRN_DATE"), 7, 2)) = True Then
                    strSQL = strSQL & ", CXL_DATE = '" & Mid(DtView2(i)("WRN_DATE"), 1, 4) & "/" & Mid(DtView2(i)("WRN_DATE"), 5, 2) & "/" & Mid(DtView2(i)("WRN_DATE"), 7, 2) & "'"
                Else
                    strSQL = strSQL & ", CXL_DATE = NULL"
                End If
            End If
            strSQL = strSQL & " WHERE (WRN_NO = '" & DtView2(i)("WRN_NO") & "')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.CommandTimeout = 600
            SqlCmd1.ExecuteNonQuery()
        End If
    End Sub

    Sub upd_A()                 '余ったＡデータで更新
        'dttable = DsList1.Tables("txt_data_all")
        'dtRow = dttable.NewRow
        'dtRow("WRN_DATE") = DtView2(i)("WRN_DATE")
        'dtRow("WRN_NO") = DtView2(i)("WRN_NO")
        'dtRow("SHOP_CODE") = DtView2(i)("SHOP_CODE")
        'dtRow("ITEM_CODE") = DtView2(i)("ITEM_CODE")
        'dtRow("MODEL") = DtView2(i)("MODEL")
        'dtRow("CAT_CODE") = DtView2(i)("CAT_CODE")
        'dtRow("CAT_NAME") = DtView2(i)("CAT_NAME")
        'dtRow("MKR_CODE") = DtView2(i)("MKR_CODE")
        'dtRow("MKR_NAME") = DtView2(i)("MKR_NAME")
        'dtRow("PRICE") = DtView2(i)("PRICE")
        'dtRow("WRN_PRICE") = DtView2(i)("WRN_PRICE")
        'dtRow("WRN_PRD") = DtView2(i)("WRN_PRD")
        'dtRow("SALE_STS") = DtView2(i)("SALE_STS")
        'dtRow("CRT_DATE") = DtView2(i)("CRT_DATE")
        'dtRow("CLS_MNTH") = DtView2(i)("CLS_MNTH")
        'dtRow("PNT_NO") = DtView2(i)("PNT_NO")
        'dtRow("CUST_NAME") = DtView2(i)("CUST_NAME")
        'dtRow("ZIP1") = DtView2(i)("ZIP1")
        'dtRow("ZIP2") = DtView2(i)("ZIP2")
        'dtRow("ADRS1") = DtView2(i)("ADRS1")
        'dtRow("ADRS2") = DtView2(i)("ADRS2")
        'dtRow("SEX") = DtView2(i)("SEX")
        'dtRow("BRTH_DATE") = DtView2(i)("BRTH_DATE")
        'dtRow("TEL_NO") = DtView2(i)("TEL_NO")
        'dtRow("CNT_NO") = DtView2(i)("CNT_NO")
        'dtRow("IMPT_FILE") = DtView2(i)("IMPT_FILE")
        'dttable.Rows.Add(dtRow)
        'DtView_C = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '09'", "", DataViewRowState.CurrentRows)
        'If DtView_C.Count <> 0 Then
        '    For j = 0 To DtView_C.Count - 1
        '        DtView_A = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '00' AND PRICE = " & DtView_C(j)("PRICE"), "IMPT_FILE, WRN_DATE", DataViewRowState.CurrentRows)
        '        If DtView_A.Count <> 0 Then
        '            DtView_A(0)("PRICE") = -999
        '        End If
        '    Next
        'End If
        'DtView_A = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '00' AND PRICE <> -999", "IMPT_FILE DESC, WRN_DATE DESC", DataViewRowState.CurrentRows)
        'If DtView_A.Count <> 0 Then
        '    strSQL = "UPDATE WRN_DATA"
        '    If IsDate(Mid(DtView_A(0)("WRN_DATE"), 1, 4), Mid(DtView_A(0)("WRN_DATE"), 5, 2), Mid(DtView_A(0)("WRN_DATE"), 7, 2)) = True Then
        '        strSQL = strSQL & " SET WRN_DATE = '" & Mid(DtView_A(0)("WRN_DATE"), 1, 4) & "/" & Mid(DtView_A(0)("WRN_DATE"), 5, 2) & "/" & Mid(DtView_A(0)("WRN_DATE"), 7, 2) & "'"
        '    Else
        '        strSQL = strSQL & " SET WRN_DATE = NULL"
        '    End If
        '    strSQL = strSQL & ", SHOP_CODE = '" & DtView_A(0)("SHOP_CODE") & "'"
        '    strSQL = strSQL & ", ITEM_CODE = '" & DtView_A(0)("ITEM_CODE") & "'"
        '    strSQL = strSQL & ", MODEL = '" & DtView_A(0)("MODEL") & "'"
        '    strSQL = strSQL & ", CAT_CODE = '" & DtView_A(0)("CAT_CODE") & "'"
        '    strSQL = strSQL & ", CAT_NAME = '" & DtView_A(0)("CAT_NAME") & "'"
        '    strSQL = strSQL & ", MKR_CODE = '" & DtView_A(0)("MKR_CODE") & "'"
        '    strSQL = strSQL & ", MKR_NAME = '" & DtView_A(0)("MKR_NAME") & "'"
        '    strSQL = strSQL & ", PRICE = " & CInt(DtView_A(0)("PRICE")) & ""
        '    strSQL = strSQL & ", WRN_PRICE = " & CInt(DtView_A(0)("WRN_PRICE")) & ""
        '    strSQL = strSQL & ", WRN_PRD = '" & DtView_A(0)("WRN_PRD") & "'"
        '    strSQL = strSQL & ", SALE_STS = '" & DtView_A(0)("SALE_STS") & "'"
        '    If IsDate(Mid(DtView_A(0)("CRT_DATE"), 1, 4), Mid(DtView_A(0)("CRT_DATE"), 5, 2), Mid(DtView_A(0)("CRT_DATE"), 7, 2)) = True Then
        '        strSQL = strSQL & ", CRT_DATE = '" & Mid(DtView_A(0)("CRT_DATE"), 1, 4) & "/" & Mid(DtView_A(0)("CRT_DATE"), 5, 2) & "/" & Mid(DtView_A(0)("CRT_DATE"), 7, 2) & "'"
        '    Else
        '        strSQL = strSQL & ", CRT_DATE = NULL"
        '    End If
        '    If IsDate(Mid(DtView_A(0)("CLS_MNTH"), 1, 4), Mid(DtView_A(0)("CLS_MNTH"), 5, 2), Mid(DtView_A(0)("CLS_MNTH"), 7, 2)) = True Then
        '        strSQL = strSQL & ", CLS_MNTH = '" & Mid(DtView_A(0)("CLS_MNTH"), 1, 4) & "/" & Mid(DtView_A(0)("CLS_MNTH"), 5, 2) & "/01'"
        '    Else
        '        strSQL = strSQL & ", CLS_MNTH = NULL"
        '    End If
        '    strSQL = strSQL & ", PNT_NO = '" & DtView_A(0)("PNT_NO") & "'"
        '    'strSQL = strSQL & ", CUST_NAME_KANA = NULL"
        '    strSQL = strSQL & ", CUST_NAME = '" & DtView_A(0)("CUST_NAME") & "'"
        '    strSQL = strSQL & ", ZIP1 = '" & DtView_A(0)("ZIP1") & "'"
        '    strSQL = strSQL & ", ZIP2 = '" & DtView_A(0)("ZIP2") & "'"
        '    strSQL = strSQL & ", ADRS1 = '" & DtView_A(0)("ADRS1") & "'"
        '    strSQL = strSQL & ", ADRS2 = '" & DtView_A(0)("ADRS2") & "'"
        '    strSQL = strSQL & ", SEX = '" & DtView_A(0)("SEX") & "'"
        '    If IsDate(Mid(DtView_A(0)("BRTH_DATE"), 1, 4), Mid(DtView_A(0)("BRTH_DATE"), 5, 2), Mid(DtView_A(0)("BRTH_DATE"), 7, 2)) = True Then
        '        strSQL = strSQL & ", BRTH_DATE = '" & Mid(DtView_A(0)("BRTH_DATE"), 1, 4) & "/" & Mid(DtView_A(0)("BRTH_DATE"), 5, 2) & "/" & Mid(DtView_A(0)("BRTH_DATE"), 7, 2) & "'"
        '    Else
        '        strSQL = strSQL & ", BRTH_DATE = NULL"
        '    End If
        '    strTEL = Trim(DtView_A(0)("TEL_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
        '    strSQL = strSQL & ", TEL_NO = '" & strTEL & "'"
        '    strTEL = Trim(DtView_A(0)("CNT_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
        '    strSQL = strSQL & ", CNT_NO = '" & strTEL & "'"
        '    'strSQL = strSQL & ", MODEL_UPD_RSN = NULL"
        '    strSQL = strSQL & ", CXL_DATE = NULL"
        '    strSQL = strSQL & " WHERE (WRN_NO = '" & DtView_A(0)("WRN_NO") & "')"
        '    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        '    SqlCmd1.CommandTimeout = 600
        '    SqlCmd1.ExecuteNonQuery()
        'End If

        strSQL = "UPDATE WRN_DATA"
        If IsDate(Mid(DtView2(i)("WRN_DATE"), 1, 4), Mid(DtView2(i)("WRN_DATE"), 5, 2), Mid(DtView2(i)("WRN_DATE"), 7, 2)) = True Then
            strSQL = strSQL & " SET WRN_DATE = '" & Mid(DtView2(i)("WRN_DATE"), 1, 4) & "/" & Mid(DtView2(i)("WRN_DATE"), 5, 2) & "/" & Mid(DtView2(i)("WRN_DATE"), 7, 2) & "'"
        Else
            strSQL = strSQL & " SET WRN_DATE = NULL"
        End If
        strSQL = strSQL & ", SHOP_CODE = '" & DtView2(i)("SHOP_CODE") & "'"
        strSQL = strSQL & ", ITEM_CODE = '" & DtView2(i)("ITEM_CODE") & "'"
        strSQL = strSQL & ", MODEL = '" & DtView2(i)("MODEL") & "'"
        strSQL = strSQL & ", CAT_CODE = '" & DtView2(i)("CAT_CODE") & "'"
        strSQL = strSQL & ", CAT_NAME = '" & DtView2(i)("CAT_NAME") & "'"
        strSQL = strSQL & ", MKR_CODE = '" & DtView2(i)("MKR_CODE") & "'"
        strSQL = strSQL & ", MKR_NAME = '" & DtView2(i)("MKR_NAME") & "'"
        strSQL = strSQL & ", PRICE = " & CInt(DtView2(i)("PRICE")) & ""
        strSQL = strSQL & ", WRN_PRICE = " & CInt(DtView2(i)("WRN_PRICE")) & ""
        strSQL = strSQL & ", WRN_PRD = '" & DtView2(i)("WRN_PRD") & "'"
        strSQL = strSQL & ", SALE_STS = '" & DtView2(i)("SALE_STS") & "'"
        If IsDate(Mid(DtView2(i)("CRT_DATE"), 1, 4), Mid(DtView2(i)("CRT_DATE"), 5, 2), Mid(DtView2(i)("CRT_DATE"), 7, 2)) = True Then
            strSQL = strSQL & ", CRT_DATE = '" & Mid(DtView2(i)("CRT_DATE"), 1, 4) & "/" & Mid(DtView2(i)("CRT_DATE"), 5, 2) & "/" & Mid(DtView2(i)("CRT_DATE"), 7, 2) & "'"
        Else
            strSQL = strSQL & ", CRT_DATE = NULL"
        End If
        If IsDate(Mid(DtView2(i)("CLS_MNTH"), 1, 4), Mid(DtView2(i)("CLS_MNTH"), 5, 2), Mid(DtView2(i)("CLS_MNTH"), 7, 2)) = True Then
            strSQL = strSQL & ", CLS_MNTH = '" & Mid(DtView2(i)("CLS_MNTH"), 1, 4) & "/" & Mid(DtView2(i)("CLS_MNTH"), 5, 2) & "/01'"
        Else
            strSQL = strSQL & ", CLS_MNTH = NULL"
        End If
        strSQL = strSQL & ", PNT_NO = '" & DtView2(i)("PNT_NO") & "'"
        strSQL = strSQL & ", CUST_NAME = '" & DtView2(i)("CUST_NAME") & "'"
        strSQL = strSQL & ", ZIP1 = '" & DtView2(i)("ZIP1") & "'"
        strSQL = strSQL & ", ZIP2 = '" & DtView2(i)("ZIP2") & "'"
        strSQL = strSQL & ", ADRS1 = '" & DtView2(i)("ADRS1") & "'"
        strSQL = strSQL & ", ADRS2 = '" & DtView2(i)("ADRS2") & "'"
        strSQL = strSQL & ", SEX = '" & DtView2(i)("SEX") & "'"
        If IsDate(Mid(DtView2(i)("BRTH_DATE"), 1, 4), Mid(DtView2(i)("BRTH_DATE"), 5, 2), Mid(DtView2(i)("BRTH_DATE"), 7, 2)) = True Then
            strSQL = strSQL & ", BRTH_DATE = '" & Mid(DtView2(i)("BRTH_DATE"), 1, 4) & "/" & Mid(DtView2(i)("BRTH_DATE"), 5, 2) & "/" & Mid(DtView2(i)("BRTH_DATE"), 7, 2) & "'"
        Else
            strSQL = strSQL & ", BRTH_DATE = NULL"
        End If
        strTEL = Trim(DtView2(i)("TEL_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
        strSQL = strSQL & ", TEL_NO = '" & strTEL & "'"
        strTEL = Trim(DtView2(i)("CNT_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
        strSQL = strSQL & ", CNT_NO = '" & strTEL & "'"
        strSQL = strSQL & ", CXL_DATE = NULL"
        strSQL = strSQL & " WHERE (WRN_NO = '" & DtView2(i)("WRN_NO") & "')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()

    End Sub

    Sub upd_C()                 'Ｃデータで更新
        strSQL = "Update WRN_DATA"
        strSQL = strSQL & " SET SALE_STS = '" & DtView2(i)("SALE_STS") & "'"
        If IsDate(Mid(DtView2(i)("WRN_DATE"), 1, 4), Mid(DtView2(i)("WRN_DATE"), 5, 2), Mid(DtView2(i)("WRN_DATE"), 7, 2)) = True Then
            strSQL = strSQL & ", CXL_DATE = '" & Mid(DtView2(i)("WRN_DATE"), 1, 4) & "/" & Mid(DtView2(i)("WRN_DATE"), 5, 2) & "/" & Mid(DtView2(i)("WRN_DATE"), 7, 2) & "'"
        End If
        strSQL = strSQL & " WHERE (WRN_NO = '" & DtView2(i)("WRN_NO") & "')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub add_C()                 'Ｃデータで取込み
        DtView_C = New DataView(DsList1.Tables("txt_data_all"), "SALE_STS = '09'", "WRN_DATE DESC", DataViewRowState.CurrentRows)

        strSQL = "INSERT INTO WRN_DATA"
        strSQL = strSQL & " (WRN_DATE, WRN_NO, SHOP_CODE, ITEM_CODE, MODEL, CAT_CODE"
        strSQL = strSQL & ", CAT_NAME, MKR_CODE, MKR_NAME, PRICE, WRN_PRICE, WRN_PRD"
        strSQL = strSQL & ", SALE_STS, CRT_DATE, CLS_MNTH, PNT_NO, CUST_NAME, ZIP1"
        strSQL = strSQL & ", ZIP2, ADRS1, ADRS2, SEX, BRTH_DATE, TEL_NO, CNT_NO, CXL_DATE)"
        If IsDate(Mid(DtView_C(0)("WRN_DATE"), 1, 4), Mid(DtView_C(0)("WRN_DATE"), 5, 2), Mid(DtView_C(0)("WRN_DATE"), 7, 2)) = True Then
            strSQL = strSQL & " VALUES ('" & Mid(DtView_C(0)("WRN_DATE"), 1, 4) & "/" & Mid(DtView_C(0)("WRN_DATE"), 5, 2) & "/" & Mid(DtView_C(0)("WRN_DATE"), 7, 2) & "'"
        Else
            strSQL = strSQL & " VALUES (NULL"
        End If
        strSQL = strSQL & ", '" & DtView_C(0)("WRN_NO") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("SHOP_CODE") & "', '" & DtView_C(0)("ITEM_CODE") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("MODEL") & "', '" & DtView_C(0)("CAT_CODE") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("CAT_NAME") & "', '" & DtView_C(0)("MKR_CODE") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("MKR_NAME") & "'"
        strSQL = strSQL & ", " & CInt(DtView_C(0)("PRICE")) & ", " & CInt(DtView_C(0)("WRN_PRICE"))
        strSQL = strSQL & ", '" & DtView_C(0)("WRN_PRD") & "', '" & DtView_C(0)("SALE_STS") & "'"
        If IsDate(Mid(DtView_C(0)("CRT_DATE"), 1, 4), Mid(DtView_C(0)("CRT_DATE"), 5, 2), Mid(DtView_C(0)("CRT_DATE"), 7, 2)) = True Then
            strSQL = strSQL & ", '" & Mid(DtView_C(0)("CRT_DATE"), 1, 4) & "/" & Mid(DtView_C(0)("CRT_DATE"), 5, 2) & "/" & Mid(DtView_C(0)("CRT_DATE"), 7, 2) & "'"
        Else
            strSQL = strSQL & ", NULL"
        End If
        If IsDate(Mid(DtView_C(0)("CLS_MNTH"), 1, 4), Mid(DtView_C(0)("CLS_MNTH"), 5, 2), "01") = True Then
            strSQL = strSQL & ", '" & Mid(DtView_C(0)("CLS_MNTH"), 1, 4) & "/" & Mid(DtView_C(0)("CLS_MNTH"), 5, 2) & "/01'"
        Else
            strSQL = strSQL & ", NULL"
        End If
        strSQL = strSQL & ", '" & DtView_C(0)("PNT_NO") & "', '" & DtView_C(0)("CUST_NAME") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("ZIP1") & "', '" & DtView_C(0)("ZIP2") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("ADRS1") & "', '" & DtView_C(0)("ADRS2") & "'"
        strSQL = strSQL & ", '" & DtView_C(0)("SEX") & "'"
        If IsDate(Mid(DtView_C(0)("BRTH_DATE"), 1, 4), Mid(DtView_C(0)("BRTH_DATE"), 5, 2), Mid(DtView_C(0)("BRTH_DATE"), 7, 2)) = True Then
            strSQL = strSQL & ", '" & Mid(DtView_C(0)("BRTH_DATE"), 1, 4) & "/" & Mid(DtView_C(0)("BRTH_DATE"), 5, 2) & "/" & Mid(DtView_C(0)("BRTH_DATE"), 7, 2) & "'"
        Else
            strSQL = strSQL & ", NULL"
        End If
        strTEL = Trim(DtView_C(0)("TEL_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
        strSQL = strSQL & ", '" & strTEL & "'"
        strTEL = Trim(DtView_C(0)("CNT_NO")).Replace("-", "") : strTEL = strTEL.Replace(" ", "")
        strSQL = strSQL & ", '" & strTEL & "'"
        If IsDate(Mid(DtView_C(0)("WRN_DATE"), 1, 4), Mid(DtView_C(0)("WRN_DATE"), 5, 2), Mid(DtView_C(0)("WRN_DATE"), 7, 2)) = True Then
            strSQL = strSQL & ", '" & Mid(DtView_C(0)("WRN_DATE"), 1, 4) & "/" & Mid(DtView_C(0)("WRN_DATE"), 5, 2) & "/" & Mid(DtView_C(0)("WRN_DATE"), 7, 2) & "')"
        Else
            strSQL = strSQL & ", NULL)"
        End If
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub txt_all_clr()           'txt_data_all_tempをクリア
        strSQL = "DELETE FROM txt_data_all_temp"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()

        strSQL = "DBCC CHECKIDENT (txt_data_all_temp ,RESEED ,0)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()

    End Sub

    Sub inport_txt_err_log()   '項目エラーがあった時エラーログ出力
        strSQL = "INSERT INTO Inport_Err_Log (WRN_DATE, WRN_NO, SHOP_CODE, ITEM_CODE, MODEL, CAT_CODE, CAT_NAME"
        strSQL = strSQL & ", MKR_CODE, MKR_NAME, PRICE, WRN_PRICE, WRN_PRD, SALE_STS, CRT_DATE, CLS_MNTH, PNT_NO"
        strSQL = strSQL & ", CUST_NAME, ZIP1, ZIP2, ADRS1, ADRS2, SEX, BRTH_DATE, TEL_NO, CNT_NO, IMPT_FILE, mnt_flg)"
        strSQL = strSQL & " VALUES ('" & RTrim(strDATA(1)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(2)) & RTrim(strDATA(3)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(4)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(5)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(6)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(7)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(8)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(9)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(10)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(11)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(12)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(13)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(14)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(15)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(16)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(17)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(18)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(19)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(20)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(21)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(22)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(23)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(24)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(25)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(26)) & "'"
        strSQL = strSQL & ", '" & filename & "'"
        strSQL = strSQL & ", 0)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub inport_txt_all_temp()   'テキストデータをそのまま取込み
        strSQL = "INSERT INTO txt_data_all_temp"
        strSQL = strSQL & " (WRN_DATE, WRN_NO, SHOP_CODE, ITEM_CODE, MODEL, CAT_CODE, CAT_NAME, MKR_CODE"
        strSQL = strSQL & ", MKR_NAME, PRICE, WRN_PRICE, WRN_PRD, SALE_STS, CRT_DATE, CLS_MNTH, PNT_NO"
        strSQL = strSQL & ", CUST_NAME, ZIP1, ZIP2, ADRS1, ADRS2, SEX, BRTH_DATE, TEL_NO, CNT_NO, IMPT_FILE)"
        strSQL = strSQL & " VALUES ('" & RTrim(strDATA(1)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(2)) & RTrim(strDATA(3)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(4)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(5)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(6)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(7)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(8)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(9)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(10)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(11)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(12)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(13)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(14)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(15)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(16)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(17)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(18)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(19)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(20)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(21)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(22)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(23)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(24)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(25)) & "'"
        strSQL = strSQL & ", '" & RTrim(strDATA(26)) & "'"
        strSQL = strSQL & ", '" & filename & "')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub txt_all_copy()   'txt_data_all_tempからコピー
        strSQL = "INSERT INTO txt_data_all"
        strSQL = strSQL & " (WRN_DATE, WRN_NO, SHOP_CODE, ITEM_CODE, MODEL, CAT_CODE, CAT_NAME, MKR_CODE"
        strSQL = strSQL & ", MKR_NAME, PRICE, WRN_PRICE, WRN_PRD, SALE_STS, CRT_DATE, CLS_MNTH, PNT_NO"
        strSQL = strSQL & ", CUST_NAME, ZIP1, ZIP2, ADRS1, ADRS2, SEX, BRTH_DATE, TEL_NO, CNT_NO, IMPT_FILE)"
        strSQL = strSQL & " VALUES ('" & DtView2(i)("WRN_DATE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("WRN_NO") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("SHOP_CODE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("ITEM_CODE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("MODEL") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("CAT_CODE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("CAT_NAME") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("MKR_CODE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("MKR_NAME") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("PRICE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("WRN_PRICE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("WRN_PRD") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("SALE_STS") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("CRT_DATE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("CLS_MNTH") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("PNT_NO") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("CUST_NAME") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("ZIP1") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("ZIP2") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("ADRS1") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("ADRS2") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("SEX") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("BRTH_DATE") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("TEL_NO") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("CNT_NO") & "'"
        strSQL = strSQL & ", '" & DtView2(i)("IMPT_FILE") & "')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub Master_ADD()            'マスタデータ追加
        'M_category
        If RTrim(strDATA(8)) <> Nothing Then
            DsList1.Clear()
            strSQL = "SELECT CAT_CODE, CAT_NAME FROM M_category WHERE (CAT_CODE  = '" & RTrim(strDATA(7)) & "')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            SqlCmd1.CommandTimeout = 600
            DaList1.Fill(DsList1, "M_category")
            DtView1 = New DataView(DsList1.Tables("M_category"), "", "", DataViewRowState.CurrentRows)
            If DtView1.Count = 0 Then
                strSQL = "INSERT INTO M_category"
                strSQL = strSQL & " (CAT_CODE, CAT_NAME)"
                strSQL = strSQL & " VALUES ('" & RTrim(strDATA(7)) & "'"
                strSQL = strSQL & ", '" & RTrim(strDATA(8)) & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.CommandTimeout = 600
                SqlCmd1.ExecuteNonQuery()
            Else
                If RTrim(strDATA(8)) <> RTrim(DtView1(0)("CAT_NAME")) Then
                    strSQL = "UPDATE M_category"
                    strSQL = strSQL & " SET CAT_NAME = '" & RTrim(strDATA(8)) & "'"
                    strSQL = strSQL & " WHERE (CAT_CODE = '" & RTrim(strDATA(7)) & "')"
                    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                    SqlCmd1.CommandTimeout = 600
                    SqlCmd1.ExecuteNonQuery()
                End If
            End If
        End If

        'M_maker
        If RTrim(strDATA(10)) <> Nothing Then
            DsList1.Clear()
            strSQL = "SELECT MKR_CODE, MKR_NAME FROM M_maker WHERE (MKR_CODE  = '" & RTrim(strDATA(9)) & "')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            SqlCmd1.CommandTimeout = 600
            DaList1.Fill(DsList1, "M_maker")
            DtView1 = New DataView(DsList1.Tables("M_maker"), "", "", DataViewRowState.CurrentRows)
            If DtView1.Count = 0 Then
                strSQL = "INSERT INTO M_maker"
                strSQL = strSQL & " (MKR_CODE, MKR_NAME)"
                strSQL = strSQL & " VALUES ('" & RTrim(strDATA(9)) & "'"
                strSQL = strSQL & ", '" & RTrim(strDATA(10)) & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.CommandTimeout = 600
                SqlCmd1.ExecuteNonQuery()
            Else
                If RTrim(strDATA(10)) <> RTrim(DtView1(0)("MKR_NAME")) Then
                    strSQL = "UPDATE M_maker"
                    strSQL = strSQL & " SET MKR_NAME = '" & RTrim(strDATA(10)) & "'"
                    strSQL = strSQL & " WHERE (MKR_CODE = '" & RTrim(strDATA(9)) & "')"
                    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                    SqlCmd1.CommandTimeout = 600
                    SqlCmd1.ExecuteNonQuery()
                End If
            End If
        End If

        'M_item
        If RTrim(strDATA(6)) <> Nothing Then
            DsList1.Clear()
            strSQL = "SELECT ITEM_CODE, MODEL FROM M_item WHERE (ITEM_CODE   = '" & RTrim(strDATA(5)) & "')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            SqlCmd1.CommandTimeout = 600
            DaList1.Fill(DsList1, "M_item")
            DtView1 = New DataView(DsList1.Tables("M_item"), "", "", DataViewRowState.CurrentRows)
            If DtView1.Count = 0 Then
                strSQL = "INSERT INTO M_item"
                strSQL = strSQL & " (ITEM_CODE, MODEL)"
                strSQL = strSQL & " VALUES ('" & RTrim(strDATA(5)) & "'"
                strSQL = strSQL & ", '" & RTrim(strDATA(6)) & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.CommandTimeout = 600
                SqlCmd1.ExecuteNonQuery()
            Else
                If RTrim(strDATA(6)) <> RTrim(DtView1(0)("MODEL")) Then
                    strSQL = "UPDATE M_item"
                    strSQL = strSQL & " SET MODEL = '" & RTrim(strDATA(6)) & "'"
                    strSQL = strSQL & " WHERE (ITEM_CODE = '" & RTrim(strDATA(5)) & "')"
                    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                    SqlCmd1.CommandTimeout = 600
                    SqlCmd1.ExecuteNonQuery()
                End If
            End If
        End If

    End Sub

    Sub Inport_Log()            'インポートLOG
        strSQL = "INSERT INTO Inport_log (Inport_File) VALUES ('" & filename & "')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub LAST_IMPORT_FILE()      '最終取込みファイル名保存
        strSQL = "UPDATE LAST_IMPORT_FILE SET Inport_File = '" & filename & "'"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub

    Sub Err_output(ByVal kbn As String, ByVal cnt1 As Integer, ByVal cnt2 As Integer)
        strSQL = "INSERT INTO RCV_ERR (err_date, kbn, cnt1, cnt2, mnt_flg)"
        strSQL = strSQL & " VALUES ('" & Format(WK_date, "yyyy/MM/dd") & "'"
        strSQL = strSQL & ", '" & kbn & "'"
        strSQL = strSQL & ", " & cnt1 & ""
        strSQL = strSQL & ", " & cnt2 & ""
        strSQL = strSQL & ", 0)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        SqlCmd1.CommandTimeout = 600
        SqlCmd1.ExecuteNonQuery()
    End Sub
End Class
