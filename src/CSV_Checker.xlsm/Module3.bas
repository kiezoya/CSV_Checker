Attribute VB_Name = "Module3"
Sub csv_analysis()

    Call Select_Folder
    Call FileList
    Call GetCsvData
    Call wksheet
    
    
    
End Sub


Sub Select_Folder()

'フォルダ選択
'    Const cnsDR = "\*.CSV"
'    Dim FldNm
'    Dim CsvNm
'    Dim cntRw 'ファイル数
'
'
'    FldNm = ThisWorkbook.Sheets(1).Cells(3, 12)
'    cntRw = 1
'
'    'ファイル名取得
'    CsvNm = Dir(FldNm & cnsDR, vbNormal)
'    Do While CsvNm <> ""
'        ThisWorkbook.Sheets(1).Cells(cntRw, 25).Value = CsvNm
'        cntRw = cntRw + 1
'        CsvNm = Dir()
'
'    Loop
'
'    Debug.Print CsvNm

    With Application.FileDialog(msoFileDialogFolderPicker)

        If .Show = True Then
            ThisWorkbook.Sheets(1).Cells(3, 12) = .SelectedItems(1) 'ダイアログで選択したフォルダ名をセルに書込み
        End If
    End With
End Sub

Sub FileList()
'************************
'   対象ファイルリスト
'************************
    Const cnsDIR = "\*.csv"
    Dim FldNm '試験データフォルダ
    Dim CsvNm
    Dim cntRw 'フォルダ内ファイル数
    
    '初期処理
    FldNm = ThisWorkbook.Sheets(1).Cells(3, 12)
        
    'フォルダ先指定なしの場合処理なし
    If FldNm = "" Then
        Exit Sub
    End If
    
    '指定フォルダ先にファイルがない場合
    If Dir(FldNm, vbDirectory) = "" Then
        MsgBox "指定のフォルダは存在しません。", vbExclamation, cnsTitle
        Exit Sub
    End If

    cntRw = 1
    ' 先頭のファイル名取得
    CsvNm = Dir(FldNm & cnsDIR, vbNormal)

    'ファイルが見つからなくなるまで繰返す
    Do While CsvNm <> ""
        '取得ファイル名をセルに書込み
        ThisWorkbook.Sheets(1).Cells(cntRw, 25).Value = CsvNm
        cntRw = cntRw + 1 ' 行を加算
        CsvNm = Dir()     ' 次のファイル名を取得
    Loop
    'セル幅オートフィット
    ThisWorkbook.Sheets(1).Columns(2).EntireColumn.AutoFit

End Sub
Sub GetCsvData()
'***********************
'   CSVデータ抽出処理
'***********************
 '初期処理
    Application.ScreenUpdating = False
    Dim cntFile '読み込みCSVファイルカウント数
    Dim sid As Variant '見出し1番目　サンプルID
    Dim opmode '見出し2番目　動作モード
    Dim strMd      '見出し3番目　変調
    Dim ori     '見出し4番目　測定方向
    Dim Pol     '見出し6番目　アンテナ偏波
    Dim siken_no '見出し5番目　試験番号
    Dim freq '周波数
    Dim RwTyp
    Dim f_data '数値データ識別フラグ
    Dim f_setClm 'CSV見出し列番号フラグ
    Dim stRw '出力シート書き込み行番号
    Dim posClm 'データ書き込み列番号
    Dim cntMHz '{MHz}データ行読み込み回数カウンタ
    Dim stClm   '開始列番号
    Dim maxClm  '見出し書き込み済み最大列番号
    Dim strKeyL '見出しKEY情報
    Dim maxGrRw '書き込み列最終行番号
    Dim dblMHz  'csv読み込み周波数
    Dim ClmName 'アルファベット列名
    Dim frmRw   '見出し7番目情報
    Dim startFq 'start 周波数
    Dim stopFq 'stop周波数
    
    
    f_setClm = 0
    
    '----------------------------------------------
    '列番号からアルファベット列名に変換
    '----------------------------------------------
   On Error Resume Next
   ClmName = Split(Cells(1, stClm).Address, "$")(1)
   '-----------------------------------------------
   
    
    '読込CSVファイル数をカウント
    cntFile = WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Columns("Y:Y"))
    stRw = 1  '出力先シート書込行に開始行をセット
    '出力シートクリア
'    ThisWorkbook.Sheets(2).Rows("2:1048576").Clear
    ThisWorkbook.Sheets("wk").Cells.Clear   'バンド帯エリアデータ一時コピー用シート
    
    'CSVファイル読込処理
    For i = 1 To cntFile
        '読み込んだCSVファイル名をフルパスに変換
        With ThisWorkbook.Sheets(1)
          CsvNm = .Cells(3, 12) & "\" & .Cells(i, 25)
        End With
        'CSVファイルオープン処理
        If Dir(CsvNm) <> "" Then
            Workbooks.Open CsvNm
        Else
            MsgBox "ファイルが存在しません。", vbExclamation
            Exit Sub
        End If
        'CSV最終行を格納
        maxRw = ActiveWorkbook.Sheets(1).Cells(1048576, 2).End(xlUp).Row

        '変数初期化
        strPol = ""
        flgLvCmp = 0 '書き出しフラグをオフ
        cntfg = 0
        f_data = 0 '数値データ識別フラグをオフ
        posClm = 0 'データ書き込み列番号リセット
        cntMHz = 0 'MHzデータ行読み込み回数カウンタリセット
        strMd = "" '変調情報クリア
        stRw = 1
        sid = ""      '見出し1番目情報クリア
        opmode = ""   '見出し2番目情報クリア
        strMd = ""   '見出し3番目情報クリア
        ori = ""   '見出し4番目情報クリア
        Pol = ""    '見出し6番目情報クリア
        siken_no = "" '見出し5番目クリア
        strKeyL = ""  '見出しKEY情報クリア
        
        
        
        stClm = 1 '開始列番号
        
      '  posClm = cntFile + 1
            '----------見出し情報リスト-----------
            '見出し1番目---サンプルID
            '見出し2番目---動作モード
            '見出し3番目---変調
            '見出し4番目---測定方向
            '見出し5番目---試験番号
            '見出し6番目---アンテナ偏波
            '見出し7番目---周波数範囲
            '-------------------------------------
    
      '---情報取得-----------------------------------------------------
        With ActiveWorkbook.Sheets(1)
            sid = Right(.Cells(4, 1), 11)
            opmode = Mid(.Cells(5, 1), 16, 14)
            strMd = .Cells(21, 1)
            ori = Right(.Cells(6, 1), 1)
            siken_no = Right(.Cells(10, 1), 6)
           ' freq = .Cells(j, 2)
            Pol = Trim(.Cells(1048576, 1).End(xlUp).Offset(-10, 0))
            
        End With
        
        'strKeyL = sid & opmode & ori & siken_no  '見出しKEY情報取得
        strKeyL = siken_no
        
        '------------------------------------------------------------
        'wkシート1列目から右側に1列ずつ見出し項目を出力する
        '------------------------------------------------------------
        maxClm = ThisWorkbook.Sheets("wk").Range("XFD1").End(xlToLeft).Column
         
        
        If maxClm < stClm Then
            maxClm = stClm
        End If
        '-----------------------------------------------------------------
        '見出し書き込み済最大列番号に空白列を1列足した列まで処理を繰り返す
        '-----------------------------------------------------------------
        For k = stClm To maxClm + 1
        
            If ThisWorkbook.Sheets("wk").Cells(1, k) <> "" Then
        '---------------------------------------------------------------
        '読み込んだcsvの見出しKEY情報がすでにwkシートに存在した場合
        'その列番号を保存してforループを抜ける
        '---------------------------------------------------------------
            With ThisWorkbook.Sheets("wk")
                'If strKeyL & strMd = .Cells(1, k) & .Cells(2, k) & .Cells(4, k) & .Cells(5, k) & .Cells(3, k) Then
                If strKeyL = .Cells(5, k) Then
                    posClm = k '列番号を保存
                    Exit For
                End If
            End With
            Else
        '--------------------------------------------------------------------
        '読み込んだcsvの見出しKEY情報がwkシートに存在しなかった場合
        '空白列に見出し項目を書き込む、列番号を保存してforループを抜ける
        '--------------------------------------------------------------------
        ThisWorkbook.Sheets("wk").Cells(1, k) = sid
        ThisWorkbook.Sheets("wk").Cells(2, k) = opmode
        ThisWorkbook.Sheets("wk").Cells(3, k) = strMd
        ThisWorkbook.Sheets("wk").Cells(4, k) = ori
        ThisWorkbook.Sheets("wk").Cells(5, k) = siken_no
        ThisWorkbook.Sheets("wk").Cells(6, k) = Pol
        
        posClm = k  '列番号を保存
         
                Exit For
            End If
       Next k
       
       
          'CSV行データ読込処理 1行目から最終行まで処理する
        For j = 1 To maxRw
            
        '読み込んだCSV行の先頭を変数に確保
        RwTyp = Left(ActiveWorkbook.Sheets(1).Cells(j, 2), 4)
        
        Select Case RwTyp
        
            Case "[MHz"
                   
                    '    flgLvCmp = 1 'レベル算出処理フラグをオン
                        f_setClm = 1 'CSV見出し列番号フラグをオン　csv見出し列番号を取得するタイミングを示すフラグ
                        f_data = 1 '数値データ識別フラグをオン　次行から数値データが始まる
                        cntMHz = cntMHz + 1 '[MHz]データ読み込み回数カウンタを1増やす
                       'If RwTyp + 1 <> "" Then
                  '-----------------------------------------------------
                  '[MHz]データ初回読み込み時
                  '-----------------------------------------------------
                If cntMHz = 1 Then
                
                    strMd = Left(ActiveWorkbook.Sheets(1).Cells(j - 3, 1), 2)
                    ThisWorkbook.Sheets("wk").Cells(3, posClm) = strMd
                    
                    '----------------------------------
                    '数値データ識別フラグがオンの場合
                    '----------------------------------
                    If ActiveWorkbook.Sheets(1).Cells(j + 1, 2) <> "" Then
                    
                        startFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2)
                        stopFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2).End(xlDown)
            
                        If f_data = 1 Then
                            ThisWorkbook.Sheets("wk").Cells(8, posClm) = startFq & "-" & stopFq
                        End If
                    End If
                   
                    
                    
    
                '----------------------------------------------------------------------------
                '変調分割　[MHz]行2回目以降は右に見出し一致列がなければ1列挿入し見出しを出力
                '次の行が空白だったら出力しない
                '----------------------------------------------------------------------------
                 ElseIf cntMHz > 1 And ActiveWorkbook.Sheets(1).Cells(j + 1, 2) <> "" Then
                    
                    strMd = Left(ActiveWorkbook.Sheets(1).Cells(j - 3, 1), 2)
                    With ThisWorkbook.Sheets("wk")
                        '----------------------------------------------------------
                        '右に一致列があった場合
                        '----------------------------------------------------------
                        If .Cells(1, posClm) & .Cells(2, posClm) & strMd & .Cells(4, posClm) & .Cells(5, posClm) = _
                            .Cells(1, posClm + 1) & .Cells(2, posClm + 1) & .Cells(3, posClm + 1) & .Cells(4, posClm + 1) & .Cells(5, posClm + 1) Then
                                   
                            posClm = posClm + 1
                            stRw = 1
                                  
                    '----------------------------------
                    '数値データ識別フラグがオンの場合
                    '----------------------------------
                    If ActiveWorkbook.Sheets(1).Cells(j + 1, 2) <> "" Then
                    
                        startFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2)
                        stopFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2).End(xlDown)
    
                        If f_data = 1 Then
                            ThisWorkbook.Sheets("wk").Cells(8, posClm) = startFq & "-" & stopFq
                        End If
                    End If
                    
                         Else
                              
                        '------------------------------------------------------------
                        '右に一致列がなかった場合
                        '------------------------------------------------------------
                        posClm = posClm + 1
                        '右に1列挿入する
                        ThisWorkbook.Sheets("wk").Columns(posClm).Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
                        '見出し項目書き込み
                        With ThisWorkbook.Sheets("wk")
                            .Cells(1, posClm) = .Cells(1, posClm - 1)
                            .Cells(2, posClm) = .Cells(2, posClm - 1)
                            .Cells(3, posClm) = strMd
                            .Cells(4, posClm) = .Cells(4, posClm - 1)
                            .Cells(5, posClm) = .Cells(5, posClm - 1)
                            .Cells(6, posClm) = .Cells(6, posClm - 1)
                        End With
                                               
                                
                            stRw = 1
                        
                        
                    '----------------------------------
                    '数値データ識別フラグがオンの場合
                    '----------------------------------
                    If ActiveWorkbook.Sheets(1).Cells(j + 1, 2) <> "" Then
                    
                        startFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2)
                        stopFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2).End(xlDown)
    
                        If f_data = 1 Then
                            ThisWorkbook.Sheets("wk").Cells(8, posClm) = startFq & "-" & stopFq
                        End If
                    End If
                    
                End If
                    End With
                        
                Else
                            f_data = 0
                       
                    End If
     
         
         Case ""
              f_data = 0
            

        Case Else
          If ActiveWorkbook.Sheets(1).Cells(j + 1, 2) = "5100" Then
                       
                        posClm = posClm + 1
                        '右に1列挿入する
                        ThisWorkbook.Sheets("wk").Columns(posClm).Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
                        
                    
                    End If
                
End Select


   Next j
   

        'CSVファイルクローズ処理
        ActiveWorkbook.Close SaveChanges:=False
    Next i
    '終了処理
    Application.ScreenUpdating = True
    MsgBox ("処理終了")
    

    
End Sub

Sub wksheet()
 
 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
 Dim m As Integer
 
 Dim lastClm
 
  lastClm = ThisWorkbook.Sheets("wk").Cells(1, Columns.Count).End(xlToLeft).Column
 
 

 
    For i = 1 To lastClm
      If ThisWorkbook.Sheets("wk").Cells(8, i) = "" Then
        ThisWorkbook.Sheets("wk").Columns(i).delete
        Else
        End If
        
        Next i
        
        For j = 1 To lastClm
        
        With ThisWorkbook.Sheets("wk")
        
            If .Cells(3, j) = "PM" Then
                .Cells(3, j) = Replace(.Cells(3, j), "PM", "PM1")
            ElseIf .Cells(3, j) = "ﾚｰ" Then
                .Cells(3, j) = Replace(.Cells(3, j), "ﾚｰ", "PM2")
            
            End If

        End With
            
     Next j
    
  For k = 1 To lastClm
    ThisWorkbook.Sheets("wk").Cells(9, k) = Left(ThisWorkbook.Sheets("wk").Cells(1, k), 7)
    ThisWorkbook.Sheets("wk").Cells(10, k) = Right(ThisWorkbook.Sheets("wk").Cells(1, k), 3)
  
  Next k
  
 '  For m = 1 To lstClm + n
 '
 '   If ThisWorkbook.Sheets("wk").Cells(8, m) = "5100-6000" Then
 '        ThisWorkbook.Sheets("wk").Range("m:m").Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
 '
 '
 '   With ThisWorkbook.Sheets("wk")
 '                   .Cells(1, m) = .Cells(1, m - 1)
 '                   .Cells(2, m) = .Cells(2, m - 1)
 '                   .Cells(3, m) = .Cells(3, m - 1)
 '                   .Cells(4, m) = .Cells(4, m - 1)
 '                   .Cells(5, m) = .Cells(5, m - 1)
 '                   .Cells(6, m) = .Cells(6, m - 1)
 '   End With
 '
 '  End If
 '
' Next m



End Sub






