Attribute VB_Name = "Module3"
Sub csv_analysis()

    Call Select_Folder
    Call FileList
    Call GetCsvData
    Call wksheet
    
    
    
End Sub


Sub Select_Folder()

'�t�H���_�I��
'    Const cnsDR = "\*.CSV"
'    Dim FldNm
'    Dim CsvNm
'    Dim cntRw '�t�@�C����
'
'
'    FldNm = ThisWorkbook.Sheets(1).Cells(3, 12)
'    cntRw = 1
'
'    '�t�@�C�����擾
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
            ThisWorkbook.Sheets(1).Cells(3, 12) = .SelectedItems(1) '�_�C�A���O�őI�������t�H���_�����Z���ɏ�����
        End If
    End With
End Sub

Sub FileList()
'************************
'   �Ώۃt�@�C�����X�g
'************************
    Const cnsDIR = "\*.csv"
    Dim FldNm '�����f�[�^�t�H���_
    Dim CsvNm
    Dim cntRw '�t�H���_���t�@�C����
    
    '��������
    FldNm = ThisWorkbook.Sheets(1).Cells(3, 12)
        
    '�t�H���_��w��Ȃ��̏ꍇ�����Ȃ�
    If FldNm = "" Then
        Exit Sub
    End If
    
    '�w��t�H���_��Ƀt�@�C�����Ȃ��ꍇ
    If Dir(FldNm, vbDirectory) = "" Then
        MsgBox "�w��̃t�H���_�͑��݂��܂���B", vbExclamation, cnsTitle
        Exit Sub
    End If

    cntRw = 1
    ' �擪�̃t�@�C�����擾
    CsvNm = Dir(FldNm & cnsDIR, vbNormal)

    '�t�@�C����������Ȃ��Ȃ�܂ŌJ�Ԃ�
    Do While CsvNm <> ""
        '�擾�t�@�C�������Z���ɏ�����
        ThisWorkbook.Sheets(1).Cells(cntRw, 25).Value = CsvNm
        cntRw = cntRw + 1 ' �s�����Z
        CsvNm = Dir()     ' ���̃t�@�C�������擾
    Loop
    '�Z�����I�[�g�t�B�b�g
    ThisWorkbook.Sheets(1).Columns(2).EntireColumn.AutoFit

End Sub
Sub GetCsvData()
'***********************
'   CSV�f�[�^���o����
'***********************
 '��������
    Application.ScreenUpdating = False
    Dim cntFile '�ǂݍ���CSV�t�@�C���J�E���g��
    Dim sid As Variant '���o��1�Ԗځ@�T���v��ID
    Dim opmode '���o��2�Ԗځ@���샂�[�h
    Dim strMd      '���o��3�Ԗځ@�ϒ�
    Dim ori     '���o��4�Ԗځ@�������
    Dim Pol     '���o��6�Ԗځ@�A���e�i�Δg
    Dim siken_no '���o��5�Ԗځ@�����ԍ�
    Dim freq '���g��
    Dim RwTyp
    Dim f_data '���l�f�[�^���ʃt���O
    Dim f_setClm 'CSV���o����ԍ��t���O
    Dim stRw '�o�̓V�[�g�������ݍs�ԍ�
    Dim posClm '�f�[�^�������ݗ�ԍ�
    Dim cntMHz '{MHz}�f�[�^�s�ǂݍ��݉񐔃J�E���^
    Dim stClm   '�J�n��ԍ�
    Dim maxClm  '���o���������ݍςݍő��ԍ�
    Dim strKeyL '���o��KEY���
    Dim maxGrRw '�������ݗ�ŏI�s�ԍ�
    Dim dblMHz  'csv�ǂݍ��ݎ��g��
    Dim ClmName '�A���t�@�x�b�g��
    Dim frmRw   '���o��7�Ԗڏ��
    Dim startFq 'start ���g��
    Dim stopFq 'stop���g��
    
    
    f_setClm = 0
    
    '----------------------------------------------
    '��ԍ�����A���t�@�x�b�g�񖼂ɕϊ�
    '----------------------------------------------
   On Error Resume Next
   ClmName = Split(Cells(1, stClm).Address, "$")(1)
   '-----------------------------------------------
   
    
    '�Ǎ�CSV�t�@�C�������J�E���g
    cntFile = WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Columns("Y:Y"))
    stRw = 1  '�o�͐�V�[�g�����s�ɊJ�n�s���Z�b�g
    '�o�̓V�[�g�N���A
'    ThisWorkbook.Sheets(2).Rows("2:1048576").Clear
    ThisWorkbook.Sheets("wk").Cells.Clear   '�o���h�уG���A�f�[�^�ꎞ�R�s�[�p�V�[�g
    
    'CSV�t�@�C���Ǎ�����
    For i = 1 To cntFile
        '�ǂݍ���CSV�t�@�C�������t���p�X�ɕϊ�
        With ThisWorkbook.Sheets(1)
          CsvNm = .Cells(3, 12) & "\" & .Cells(i, 25)
        End With
        'CSV�t�@�C���I�[�v������
        If Dir(CsvNm) <> "" Then
            Workbooks.Open CsvNm
        Else
            MsgBox "�t�@�C�������݂��܂���B", vbExclamation
            Exit Sub
        End If
        'CSV�ŏI�s���i�[
        maxRw = ActiveWorkbook.Sheets(1).Cells(1048576, 2).End(xlUp).Row

        '�ϐ�������
        strPol = ""
        flgLvCmp = 0 '�����o���t���O���I�t
        cntfg = 0
        f_data = 0 '���l�f�[�^���ʃt���O���I�t
        posClm = 0 '�f�[�^�������ݗ�ԍ����Z�b�g
        cntMHz = 0 'MHz�f�[�^�s�ǂݍ��݉񐔃J�E���^���Z�b�g
        strMd = "" '�ϒ����N���A
        stRw = 1
        sid = ""      '���o��1�Ԗڏ��N���A
        opmode = ""   '���o��2�Ԗڏ��N���A
        strMd = ""   '���o��3�Ԗڏ��N���A
        ori = ""   '���o��4�Ԗڏ��N���A
        Pol = ""    '���o��6�Ԗڏ��N���A
        siken_no = "" '���o��5�ԖڃN���A
        strKeyL = ""  '���o��KEY���N���A
        
        
        
        stClm = 1 '�J�n��ԍ�
        
      '  posClm = cntFile + 1
            '----------���o����񃊃X�g-----------
            '���o��1�Ԗ�---�T���v��ID
            '���o��2�Ԗ�---���샂�[�h
            '���o��3�Ԗ�---�ϒ�
            '���o��4�Ԗ�---�������
            '���o��5�Ԗ�---�����ԍ�
            '���o��6�Ԗ�---�A���e�i�Δg
            '���o��7�Ԗ�---���g���͈�
            '-------------------------------------
    
      '---���擾-----------------------------------------------------
        With ActiveWorkbook.Sheets(1)
            sid = Right(.Cells(4, 1), 11)
            opmode = Mid(.Cells(5, 1), 16, 14)
            strMd = .Cells(21, 1)
            ori = Right(.Cells(6, 1), 1)
            siken_no = Right(.Cells(10, 1), 6)
           ' freq = .Cells(j, 2)
            Pol = Trim(.Cells(1048576, 1).End(xlUp).Offset(-10, 0))
            
        End With
        
        'strKeyL = sid & opmode & ori & siken_no  '���o��KEY���擾
        strKeyL = siken_no
        
        '------------------------------------------------------------
        'wk�V�[�g1��ڂ���E����1�񂸂��o�����ڂ��o�͂���
        '------------------------------------------------------------
        maxClm = ThisWorkbook.Sheets("wk").Range("XFD1").End(xlToLeft).Column
         
        
        If maxClm < stClm Then
            maxClm = stClm
        End If
        '-----------------------------------------------------------------
        '���o���������ݍύő��ԍ��ɋ󔒗��1�񑫂�����܂ŏ������J��Ԃ�
        '-----------------------------------------------------------------
        For k = stClm To maxClm + 1
        
            If ThisWorkbook.Sheets("wk").Cells(1, k) <> "" Then
        '---------------------------------------------------------------
        '�ǂݍ���csv�̌��o��KEY��񂪂��ł�wk�V�[�g�ɑ��݂����ꍇ
        '���̗�ԍ���ۑ�����for���[�v�𔲂���
        '---------------------------------------------------------------
            With ThisWorkbook.Sheets("wk")
                'If strKeyL & strMd = .Cells(1, k) & .Cells(2, k) & .Cells(4, k) & .Cells(5, k) & .Cells(3, k) Then
                If strKeyL = .Cells(5, k) Then
                    posClm = k '��ԍ���ۑ�
                    Exit For
                End If
            End With
            Else
        '--------------------------------------------------------------------
        '�ǂݍ���csv�̌��o��KEY���wk�V�[�g�ɑ��݂��Ȃ������ꍇ
        '�󔒗�Ɍ��o�����ڂ��������ށA��ԍ���ۑ�����for���[�v�𔲂���
        '--------------------------------------------------------------------
        ThisWorkbook.Sheets("wk").Cells(1, k) = sid
        ThisWorkbook.Sheets("wk").Cells(2, k) = opmode
        ThisWorkbook.Sheets("wk").Cells(3, k) = strMd
        ThisWorkbook.Sheets("wk").Cells(4, k) = ori
        ThisWorkbook.Sheets("wk").Cells(5, k) = siken_no
        ThisWorkbook.Sheets("wk").Cells(6, k) = Pol
        
        posClm = k  '��ԍ���ۑ�
         
                Exit For
            End If
       Next k
       
       
          'CSV�s�f�[�^�Ǎ����� 1�s�ڂ���ŏI�s�܂ŏ�������
        For j = 1 To maxRw
            
        '�ǂݍ���CSV�s�̐擪��ϐ��Ɋm��
        RwTyp = Left(ActiveWorkbook.Sheets(1).Cells(j, 2), 4)
        
        Select Case RwTyp
        
            Case "[MHz"
                   
                    '    flgLvCmp = 1 '���x���Z�o�����t���O���I��
                        f_setClm = 1 'CSV���o����ԍ��t���O���I���@csv���o����ԍ����擾����^�C�~���O�������t���O
                        f_data = 1 '���l�f�[�^���ʃt���O���I���@���s���琔�l�f�[�^���n�܂�
                        cntMHz = cntMHz + 1 '[MHz]�f�[�^�ǂݍ��݉񐔃J�E���^��1���₷
                       'If RwTyp + 1 <> "" Then
                  '-----------------------------------------------------
                  '[MHz]�f�[�^����ǂݍ��ݎ�
                  '-----------------------------------------------------
                If cntMHz = 1 Then
                
                    strMd = Left(ActiveWorkbook.Sheets(1).Cells(j - 3, 1), 2)
                    ThisWorkbook.Sheets("wk").Cells(3, posClm) = strMd
                    
                    '----------------------------------
                    '���l�f�[�^���ʃt���O���I���̏ꍇ
                    '----------------------------------
                    If ActiveWorkbook.Sheets(1).Cells(j + 1, 2) <> "" Then
                    
                        startFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2)
                        stopFq = ActiveWorkbook.Sheets(1).Cells(j + 1, 2).End(xlDown)
            
                        If f_data = 1 Then
                            ThisWorkbook.Sheets("wk").Cells(8, posClm) = startFq & "-" & stopFq
                        End If
                    End If
                   
                    
                    
    
                '----------------------------------------------------------------------------
                '�ϒ������@[MHz]�s2��ڈȍ~�͉E�Ɍ��o����v�񂪂Ȃ����1��}�������o�����o��
                '���̍s���󔒂�������o�͂��Ȃ�
                '----------------------------------------------------------------------------
                 ElseIf cntMHz > 1 And ActiveWorkbook.Sheets(1).Cells(j + 1, 2) <> "" Then
                    
                    strMd = Left(ActiveWorkbook.Sheets(1).Cells(j - 3, 1), 2)
                    With ThisWorkbook.Sheets("wk")
                        '----------------------------------------------------------
                        '�E�Ɉ�v�񂪂������ꍇ
                        '----------------------------------------------------------
                        If .Cells(1, posClm) & .Cells(2, posClm) & strMd & .Cells(4, posClm) & .Cells(5, posClm) = _
                            .Cells(1, posClm + 1) & .Cells(2, posClm + 1) & .Cells(3, posClm + 1) & .Cells(4, posClm + 1) & .Cells(5, posClm + 1) Then
                                   
                            posClm = posClm + 1
                            stRw = 1
                                  
                    '----------------------------------
                    '���l�f�[�^���ʃt���O���I���̏ꍇ
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
                        '�E�Ɉ�v�񂪂Ȃ������ꍇ
                        '------------------------------------------------------------
                        posClm = posClm + 1
                        '�E��1��}������
                        ThisWorkbook.Sheets("wk").Columns(posClm).Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
                        '���o�����ڏ�������
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
                    '���l�f�[�^���ʃt���O���I���̏ꍇ
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
                        '�E��1��}������
                        ThisWorkbook.Sheets("wk").Columns(posClm).Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
                        
                    
                    End If
                
End Select


   Next j
   

        'CSV�t�@�C���N���[�Y����
        ActiveWorkbook.Close SaveChanges:=False
    Next i
    '�I������
    Application.ScreenUpdating = True
    MsgBox ("�����I��")
    

    
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
            ElseIf .Cells(3, j) = "ڰ" Then
                .Cells(3, j) = Replace(.Cells(3, j), "ڰ", "PM2")
            
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






