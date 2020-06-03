VERSION 5.00
Begin VB.Form frmAssumptions
  Caption = "Assumptions"
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 720
  ClientWidth = 9075
  ClientHeight = 6120
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraY
    Caption = "In Terms of the Dependent Variable"
    Left = 0
    Top = 4320
    Width = 11775
    Height = 3700
    TabIndex = 1
    BeginProperty Font
      Name = "MS Sans Serif"
      Size = 18
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.Label lblVarY
      Caption = "Equal Variance"
      Left = 240
      Top = 2760
      Width = 2430
      Height = 435
      TabIndex = 9
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblNorY
      Caption = "Normality"
      Left = 240
      Top = 2040
      Width = 1530
      Height = 435
      TabIndex = 8
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblIndY
      Caption = "Independence"
      Left = 240
      Top = 1320
      Width = 2325
      Height = 435
      TabIndex = 7
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblLineY
      Caption = "Linearity"
      Left = 240
      Top = 720
      Width = 1350
      Height = 435
      TabIndex = 6
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
  End
  Begin VB.Frame fraError
    Caption = "In Terms of the Error: Error equals to Y minus the mean of Y | X"
    Left = 0
    Top = 360
    Width = 11775
    Height = 3700
    TabIndex = 0
    BeginProperty Font
      Name = "MS Sans Serif"
      Size = 18
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.Label lblVarE
      Caption = "Equal Variance"
      Left = 240
      Top = 2760
      Width = 2430
      Height = 435
      TabIndex = 5
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblNorE
      Caption = "Normality"
      Left = 240
      Top = 2040
      Width = 1530
      Height = 435
      TabIndex = 4
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblIndE
      Caption = "Independence"
      Left = 240
      Top = 1320
      Width = 2325
      Height = 435
      TabIndex = 3
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblLineE
      Caption = "Linearity"
      Left = 240
      Top = 600
      Width = 1350
      Height = 435
      TabIndex = 2
      AutoSize = -1  'True
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
  End
  Begin VB.Menu mnuInstructions
    Caption = "&Instructions"
  End
  Begin VB.Menu mnuChangeNames
    Caption = "&Change Names"
  End
  Begin VB.Menu mnuGoto
    Caption = "&Go To"
    Begin VB.Menu mnuIntro
      Caption = "I&ntroduction"
    End
    Begin VB.Menu mnustatistics
      Caption = "Statistics and Point Esti&mates"
    End
    Begin VB.Menu mnuQuestions
      Caption = "&Questions"
    End
    Begin VB.Menu mnuInferences
      Caption = "In&ferences"
      Begin VB.Menu mnuSlope
        Caption = "&Slope"
      End
      Begin VB.Menu mnuEstPred
        Caption = "&Estimation and Prediction"
      End
    End
  End
  Begin VB.Menu mnuExit
    Caption = "E&xit"
  End
End

Attribute VB_Name = "frmAssumptions"


Private Sub mnuStatistics_Click() '41B640
  loc_0041B640: push ebp
  loc_0041B641: mov ebp, esp
  loc_0041B643: sub esp, 0000000Ch
  loc_0041B646: push 00401926h ; __vbaExceptHandler
  loc_0041B64B: mov eax, fs:[00000000h]
  loc_0041B651: push eax
  loc_0041B652: mov fs:[00000000h], esp
  loc_0041B659: sub esp, 00000030h
  loc_0041B65C: push ebx
  loc_0041B65D: push esi
  loc_0041B65E: push edi
  loc_0041B65F: mov var_C, esp
  loc_0041B662: mov var_8, 004011A0h
  loc_0041B669: mov eax, Me
  loc_0041B66C: mov ecx, eax
  loc_0041B66E: and ecx, 00000001h
  loc_0041B671: mov var_4, ecx
  loc_0041B674: and al, FEh
  loc_0041B676: push eax
  loc_0041B677: mov Me, eax
  loc_0041B67A: mov edx, [eax]
  loc_0041B67C: call [edx+00000004h]
  loc_0041B67F: mov eax, [0043009Ch]
  loc_0041B684: test eax, eax
  loc_0041B686: jnz 0041B698h
  loc_0041B688: push 0043009Ch
  loc_0041B68D: push 00405FC0h
  loc_0041B692: call [004010D4h] ; __vbaNew2
  loc_0041B698: mov esi, [0043009Ch]
  loc_0041B69E: push esi
  loc_0041B69F: mov eax, [esi]
  loc_0041B6A1: call [eax+000002B4h]
  loc_0041B6A7: test eax, eax
  loc_0041B6A9: fnclex
  loc_0041B6AB: jge 0041B6BFh
  loc_0041B6AD: push 000002B4h
  loc_0041B6B2: push 0040DDE0h
  loc_0041B6B7: push esi
  loc_0041B6B8: push eax
  loc_0041B6B9: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B6BF: mov eax, [004300ECh]
  loc_0041B6C4: test eax, eax
  loc_0041B6C6: jnz 0041B6D8h
  loc_0041B6C8: push 004300ECh
  loc_0041B6CD: push 0040A624h
  loc_0041B6D2: call [004010D4h] ; __vbaNew2
  loc_0041B6D8: sub esp, 00000010h
  loc_0041B6DB: mov ecx, 0000000Ah
  loc_0041B6E0: mov ebx, esp
  loc_0041B6E2: mov var_24, ecx
  loc_0041B6E5: mov eax, 80020004h
  loc_0041B6EA: sub esp, 00000010h
  loc_0041B6ED: mov [ebx], ecx
  loc_0041B6EF: mov ecx, var_30
  loc_0041B6F2: mov edx, eax
  loc_0041B6F4: mov esi, [004300ECh]
  loc_0041B6FA: mov [ebx+00000004h], ecx
  loc_0041B6FD: mov ecx, esp
  loc_0041B6FF: mov edi, [esi]
  loc_0041B701: push esi
  loc_0041B702: mov [ebx+00000008h], eax
  loc_0041B705: mov eax, var_28
  loc_0041B708: mov [ebx+0000000Ch], eax
  loc_0041B70B: mov eax, var_24
  loc_0041B70E: mov [ecx], eax
  loc_0041B710: mov eax, var_20
  loc_0041B713: mov [ecx+00000004h], eax
  loc_0041B716: mov [ecx+00000008h], edx
  loc_0041B719: mov edx, var_18
  loc_0041B71C: mov [ecx+0000000Ch], edx
  loc_0041B71F: call [edi+000002B0h]
  loc_0041B725: test eax, eax
  loc_0041B727: fnclex
  loc_0041B729: jge 0041B73Dh
  loc_0041B72B: push 000002B0h
  loc_0041B730: push 0040ECECh
  loc_0041B735: push esi
  loc_0041B736: push eax
  loc_0041B737: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B73D: mov var_4, 00000000h
  loc_0041B744: mov eax, Me
  loc_0041B747: push eax
  loc_0041B748: mov ecx, [eax]
  loc_0041B74A: call [ecx+00000008h]
  loc_0041B74D: mov eax, var_4
  loc_0041B750: mov ecx, var_14
  loc_0041B753: pop edi
  loc_0041B754: pop esi
  loc_0041B755: mov fs:[00000000h], ecx
  loc_0041B75C: pop ebx
  loc_0041B75D: mov esp, ebp
  loc_0041B75F: pop ebp
  loc_0041B760: retn 0004h
End Sub

Private Sub mnuIntro_Click() '41A8A0
  loc_0041A8A0: push ebp
  loc_0041A8A1: mov ebp, esp
  loc_0041A8A3: sub esp, 0000000Ch
  loc_0041A8A6: push 00401926h ; __vbaExceptHandler
  loc_0041A8AB: mov eax, fs:[00000000h]
  loc_0041A8B1: push eax
  loc_0041A8B2: mov fs:[00000000h], esp
  loc_0041A8B9: sub esp, 00000030h
  loc_0041A8BC: push ebx
  loc_0041A8BD: push esi
  loc_0041A8BE: push edi
  loc_0041A8BF: mov var_C, esp
  loc_0041A8C2: mov var_8, 00401158h
  loc_0041A8C9: mov eax, Me
  loc_0041A8CC: mov ecx, eax
  loc_0041A8CE: and ecx, 00000001h
  loc_0041A8D1: mov var_4, ecx
  loc_0041A8D4: and al, FEh
  loc_0041A8D6: push eax
  loc_0041A8D7: mov Me, eax
  loc_0041A8DA: mov edx, [eax]
  loc_0041A8DC: call [edx+00000004h]
  loc_0041A8DF: mov eax, [00430088h]
  loc_0041A8E4: test eax, eax
  loc_0041A8E6: jnz 0041A8F8h
  loc_0041A8E8: push 00430088h
  loc_0041A8ED: push 004058C0h
  loc_0041A8F2: call [004010D4h] ; __vbaNew2
  loc_0041A8F8: sub esp, 00000010h
  loc_0041A8FB: mov ecx, 0000000Ah
  loc_0041A900: mov ebx, esp
  loc_0041A902: mov var_24, ecx
  loc_0041A905: mov eax, 80020004h
  loc_0041A90A: sub esp, 00000010h
  loc_0041A90D: mov [ebx], ecx
  loc_0041A90F: mov ecx, var_30
  loc_0041A912: mov edx, eax
  loc_0041A914: mov esi, [00430088h]
  loc_0041A91A: mov [ebx+00000004h], ecx
  loc_0041A91D: mov ecx, esp
  loc_0041A91F: mov edi, [esi]
  loc_0041A921: push esi
  loc_0041A922: mov [ebx+00000008h], eax
  loc_0041A925: mov eax, var_28
  loc_0041A928: mov [ebx+0000000Ch], eax
  loc_0041A92B: mov eax, var_24
  loc_0041A92E: mov [ecx], eax
  loc_0041A930: mov eax, var_20
  loc_0041A933: mov [ecx+00000004h], eax
  loc_0041A936: mov [ecx+00000008h], edx
  loc_0041A939: mov edx, var_18
  loc_0041A93C: mov [ecx+0000000Ch], edx
  loc_0041A93F: call [edi+000002B0h]
  loc_0041A945: test eax, eax
  loc_0041A947: fnclex
  loc_0041A949: jge 0041A95Dh
  loc_0041A94B: push 000002B0h
  loc_0041A950: push 0040DB64h
  loc_0041A955: push esi
  loc_0041A956: push eax
  loc_0041A957: call [00401030h] ; __vbaHresultCheckObj
  loc_0041A95D: mov eax, [0043009Ch]
  loc_0041A962: test eax, eax
  loc_0041A964: jnz 0041A976h
  loc_0041A966: push 0043009Ch
  loc_0041A96B: push 00405FC0h
  loc_0041A970: call [004010D4h] ; __vbaNew2
  loc_0041A976: mov esi, [0043009Ch]
  loc_0041A97C: push esi
  loc_0041A97D: mov eax, [esi]
  loc_0041A97F: call [eax+000002B4h]
  loc_0041A985: test eax, eax
  loc_0041A987: fnclex
  loc_0041A989: jge 0041A99Dh
  loc_0041A98B: push 000002B4h
  loc_0041A990: push 0040DDE0h
  loc_0041A995: push esi
  loc_0041A996: push eax
  loc_0041A997: call [00401030h] ; __vbaHresultCheckObj
  loc_0041A99D: mov var_4, 00000000h
  loc_0041A9A4: mov eax, Me
  loc_0041A9A7: push eax
  loc_0041A9A8: mov ecx, [eax]
  loc_0041A9AA: call [ecx+00000008h]
  loc_0041A9AD: mov eax, var_4
  loc_0041A9B0: mov ecx, var_14
  loc_0041A9B3: pop edi
  loc_0041A9B4: pop esi
  loc_0041A9B5: mov fs:[00000000h], ecx
  loc_0041A9BC: pop ebx
  loc_0041A9BD: mov esp, ebp
  loc_0041A9BF: pop ebp
  loc_0041A9C0: retn 0004h
End Sub

Private Sub mnuInstructions_Click() '41A780
  loc_0041A780: push ebp
  loc_0041A781: mov ebp, esp
  loc_0041A783: sub esp, 0000000Ch
  loc_0041A786: push 00401926h ; __vbaExceptHandler
  loc_0041A78B: mov eax, fs:[00000000h]
  loc_0041A791: push eax
  loc_0041A792: mov fs:[00000000h], esp
  loc_0041A799: sub esp, 00000088h
  loc_0041A79F: push ebx
  loc_0041A7A0: push esi
  loc_0041A7A1: push edi
  loc_0041A7A2: mov var_C, esp
  loc_0041A7A5: mov var_8, 00401148h
  loc_0041A7AC: mov eax, Me
  loc_0041A7AF: mov ecx, eax
  loc_0041A7B1: and ecx, 00000001h
  loc_0041A7B4: mov var_4, ecx
  loc_0041A7B7: and al, FEh
  loc_0041A7B9: push eax
  loc_0041A7BA: mov Me, eax
  loc_0041A7BD: mov edx, [eax]
  loc_0041A7BF: call [edx+00000004h]
  loc_0041A7C2: mov edi, [004010F4h] ; __vbaVarDup
  loc_0041A7C8: mov ecx, 80020004h
  loc_0041A7CD: xor esi, esi
  loc_0041A7CF: mov var_4C, ecx
  loc_0041A7D2: mov eax, 0000000Ah
  loc_0041A7D7: mov var_3C, ecx
  loc_0041A7DA: mov ebx, 00000008h
  loc_0041A7DF: mov var_44, esi
  loc_0041A7E2: mov var_54, esi
  loc_0041A7E5: mov var_74, esi
  loc_0041A7E8: lea edx, var_74
  loc_0041A7EB: lea ecx, var_34
  loc_0041A7EE: mov var_24, esi
  loc_0041A7F1: mov var_34, esi
  loc_0041A7F4: mov var_64, esi
  loc_0041A7F7: mov var_54, eax
  loc_0041A7FA: mov var_44, eax
  loc_0041A7FD: mov var_6C, 0040E0D0h ; "Instructions"
  loc_0041A804: mov var_74, ebx
  loc_0041A807: call edi
  loc_0041A809: lea edx, var_64
  loc_0041A80C: lea ecx, var_24
  loc_0041A80F: mov var_5C, 0040E034h ; "Hold mouse over statements to see the statement in context of the example."
  loc_0041A816: mov var_64, ebx
  loc_0041A819: call edi
  loc_0041A81B: lea eax, var_54
  loc_0041A81E: lea ecx, var_44
  loc_0041A821: push eax
  loc_0041A822: lea edx, var_34
  loc_0041A825: push ecx
  loc_0041A826: push edx
  loc_0041A827: lea eax, var_24
  loc_0041A82A: push esi
  loc_0041A82B: push eax
  loc_0041A82C: call [00401038h] ; rtcMsgBox
  loc_0041A832: lea ecx, var_54
  loc_0041A835: lea edx, var_44
  loc_0041A838: push ecx
  loc_0041A839: lea eax, var_34
  loc_0041A83C: push edx
  loc_0041A83D: lea ecx, var_24
  loc_0041A840: push eax
  loc_0041A841: push ecx
  loc_0041A842: push 00000004h
  loc_0041A844: call [00401018h] ; __vbaFreeVarList
  loc_0041A84A: add esp, 00000014h
  loc_0041A84D: mov var_4, esi
  loc_0041A850: push 0041A874h
  loc_0041A855: jmp 0041A873h
  loc_0041A857: lea edx, var_54
  loc_0041A85A: lea eax, var_44
  loc_0041A85D: push edx
  loc_0041A85E: lea ecx, var_34
  loc_0041A861: push eax
  loc_0041A862: lea edx, var_24
  loc_0041A865: push ecx
  loc_0041A866: push edx
  loc_0041A867: push 00000004h
  loc_0041A869: call [00401018h] ; __vbaFreeVarList
  loc_0041A86F: add esp, 00000014h
  loc_0041A872: ret
  loc_0041A873: ret
  loc_0041A874: mov eax, Me
  loc_0041A877: push eax
  loc_0041A878: mov ecx, [eax]
  loc_0041A87A: call [ecx+00000008h]
  loc_0041A87D: mov eax, var_4
  loc_0041A880: mov ecx, var_14
  loc_0041A883: pop edi
  loc_0041A884: pop esi
  loc_0041A885: mov fs:[00000000h], ecx
  loc_0041A88C: pop ebx
  loc_0041A88D: mov esp, ebp
  loc_0041A88F: pop ebp
  loc_0041A890: retn 0004h
End Sub

Private Sub mnuExit_Click() '41B3D0
  loc_0041B3D0: push ebp
  loc_0041B3D1: mov ebp, esp
  loc_0041B3D3: sub esp, 0000000Ch
  loc_0041B3D6: push 00401926h ; __vbaExceptHandler
  loc_0041B3DB: mov eax, fs:[00000000h]
  loc_0041B3E1: push eax
  loc_0041B3E2: mov fs:[00000000h], esp
  loc_0041B3E9: sub esp, 00000018h
  loc_0041B3EC: push ebx
  loc_0041B3ED: push esi
  loc_0041B3EE: push edi
  loc_0041B3EF: mov var_C, esp
  loc_0041B3F2: mov var_8, 00401180h
  loc_0041B3F9: mov edi, Me
  loc_0041B3FC: mov eax, edi
  loc_0041B3FE: and eax, 00000001h
  loc_0041B401: mov var_4, eax
  loc_0041B404: and edi, FFFFFFFEh
  loc_0041B407: push edi
  loc_0041B408: mov Me, edi
  loc_0041B40B: mov ecx, [edi]
  loc_0041B40D: call [ecx+00000004h]
  loc_0041B410: mov eax, [00430A24h]
  loc_0041B415: xor ebx, ebx
  loc_0041B417: cmp eax, ebx
  loc_0041B419: mov var_18, ebx
  loc_0041B41C: jnz 0041B42Eh
  loc_0041B41E: push 00430A24h
  loc_0041B423: push 0040EC7Ch
  loc_0041B428: call [004010D4h] ; __vbaNew2
  loc_0041B42E: mov esi, [00430A24h]
  loc_0041B434: lea eax, var_18
  loc_0041B437: push edi
  loc_0041B438: push eax
  loc_0041B439: mov edx, [esi]
  loc_0041B43B: mov var_2C, edx
  loc_0041B43E: call [00401044h] ; __vbaObjSetAddref
  loc_0041B444: mov ecx, var_2C
  loc_0041B447: push eax
  loc_0041B448: push esi
  loc_0041B449: call [ecx+00000010h]
  loc_0041B44C: cmp eax, ebx
  loc_0041B44E: fnclex
  loc_0041B450: jge 0041B461h
  loc_0041B452: push 00000010h
  loc_0041B454: push 0040EC6Ch
  loc_0041B459: push esi
  loc_0041B45A: push eax
  loc_0041B45B: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B461: lea ecx, var_18
  loc_0041B464: call [00401114h] ; __vbaFreeObj
  loc_0041B46A: mov var_4, ebx
  loc_0041B46D: push 0041B47Fh
  loc_0041B472: jmp 0041B47Eh
  loc_0041B474: lea ecx, var_18
  loc_0041B477: call [00401114h] ; __vbaFreeObj
  loc_0041B47D: ret
  loc_0041B47E: ret
  loc_0041B47F: mov eax, Me
  loc_0041B482: push eax
  loc_0041B483: mov edx, [eax]
  loc_0041B485: call [edx+00000008h]
  loc_0041B488: mov eax, var_4
  loc_0041B48B: mov ecx, var_14
  loc_0041B48E: pop edi
  loc_0041B48F: pop esi
  loc_0041B490: mov fs:[00000000h], ecx
  loc_0041B497: pop ebx
  loc_0041B498: mov esp, ebp
  loc_0041B49A: pop ebp
  loc_0041B49B: retn 0004h
End Sub

Private Sub Form_Load() '41A5E0
  loc_0041A5E0: push ebp
  loc_0041A5E1: mov ebp, esp
  loc_0041A5E3: sub esp, 0000000Ch
  loc_0041A5E6: push 00401926h ; __vbaExceptHandler
  loc_0041A5EB: mov eax, fs:[00000000h]
  loc_0041A5F1: push eax
  loc_0041A5F2: mov fs:[00000000h], esp
  loc_0041A5F9: sub esp, 00000008h
  loc_0041A5FC: push ebx
  loc_0041A5FD: push esi
  loc_0041A5FE: push edi
  loc_0041A5FF: mov var_C, esp
  loc_0041A602: mov var_8, 00401138h
  loc_0041A609: mov esi, Me
  loc_0041A60C: mov eax, esi
  loc_0041A60E: and eax, 00000001h
  loc_0041A611: mov var_4, eax
  loc_0041A614: and esi, FFFFFFFEh
  loc_0041A617: push esi
  loc_0041A618: mov Me, esi
  loc_0041A61B: mov ecx, [esi]
  loc_0041A61D: call [ecx+00000004h]
  loc_0041A620: mov edx, [esi]
  loc_0041A622: push esi
  loc_0041A623: call [edx+00000714h]
  loc_0041A629: mov var_4, 00000000h
  loc_0041A630: mov eax, Me
  loc_0041A633: push eax
  loc_0041A634: mov ecx, [eax]
  loc_0041A636: call [ecx+00000008h]
  loc_0041A639: mov eax, var_4
  loc_0041A63C: mov ecx, var_14
  loc_0041A63F: pop edi
  loc_0041A640: pop esi
  loc_0041A641: mov fs:[00000000h], ecx
  loc_0041A648: pop ebx
  loc_0041A649: mov esp, ebp
  loc_0041A64B: pop ebp
  loc_0041A64C: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '41B4A0
  loc_0041B4A0: push ebp
  loc_0041B4A1: mov ebp, esp
  loc_0041B4A3: sub esp, 0000000Ch
  loc_0041B4A6: push 00401926h ; __vbaExceptHandler
  loc_0041B4AB: mov eax, fs:[00000000h]
  loc_0041B4B1: push eax
  loc_0041B4B2: mov fs:[00000000h], esp
  loc_0041B4B9: sub esp, 0000009Ch
  loc_0041B4BF: push ebx
  loc_0041B4C0: push esi
  loc_0041B4C1: push edi
  loc_0041B4C2: mov var_C, esp
  loc_0041B4C5: mov var_8, 00401190h
  loc_0041B4CC: mov edi, Me
  loc_0041B4CF: mov eax, edi
  loc_0041B4D1: and eax, 00000001h
  loc_0041B4D4: mov var_4, eax
  loc_0041B4D7: and edi, FFFFFFFEh
  loc_0041B4DA: push edi
  loc_0041B4DB: mov Me, edi
  loc_0041B4DE: mov ecx, [edi]
  loc_0041B4E0: call [ecx+00000004h]
  loc_0041B4E3: mov ebx, [004010F4h] ; __vbaVarDup
  loc_0041B4E9: mov ecx, 80020004h
  loc_0041B4EE: xor esi, esi
  loc_0041B4F0: mov var_54, ecx
  loc_0041B4F3: mov eax, 0000000Ah
  loc_0041B4F8: mov var_44, ecx
  loc_0041B4FB: mov var_4C, esi
  loc_0041B4FE: mov var_5C, esi
  loc_0041B501: mov var_7C, esi
  loc_0041B504: lea edx, var_7C
  loc_0041B507: lea ecx, var_3C
  loc_0041B50A: mov var_1C, esi
  loc_0041B50D: mov var_2C, esi
  loc_0041B510: mov var_3C, esi
  loc_0041B513: mov var_6C, esi
  loc_0041B516: mov var_5C, eax
  loc_0041B519: mov var_4C, eax
  loc_0041B51C: mov var_74, 0040ECD4h ; "Exit Check"
  loc_0041B523: mov var_7C, 00000008h
  loc_0041B52A: call ebx
  loc_0041B52C: lea edx, var_6C
  loc_0041B52F: lea ecx, var_2C
  loc_0041B532: mov var_64, 0040EC90h ; "Are you sure you want to exit?"
  loc_0041B539: mov var_6C, 00000008h
  loc_0041B540: call ebx
  loc_0041B542: lea edx, var_5C
  loc_0041B545: lea eax, var_4C
  loc_0041B548: push edx
  loc_0041B549: lea ecx, var_3C
  loc_0041B54C: push eax
  loc_0041B54D: push ecx
  loc_0041B54E: lea edx, var_2C
  loc_0041B551: push 00000004h
  loc_0041B553: push edx
  loc_0041B554: call [00401038h] ; rtcMsgBox
  loc_0041B55A: mov ecx, eax
  loc_0041B55C: call [00401078h] ; __vbaI2I4
  loc_0041B562: mov ebx, eax
  loc_0041B564: lea eax, var_5C
  loc_0041B567: lea ecx, var_4C
  loc_0041B56A: push eax
  loc_0041B56B: lea edx, var_3C
  loc_0041B56E: push ecx
  loc_0041B56F: lea eax, var_2C
  loc_0041B572: push edx
  loc_0041B573: push eax
  loc_0041B574: push 00000004h
  loc_0041B576: call [00401018h] ; __vbaFreeVarList
  loc_0041B57C: add esp, 00000014h
  loc_0041B57F: cmp bx, 0007h
  loc_0041B583: jnz 0041B58Fh
  loc_0041B585: mov ecx, Cancel
  loc_0041B588: mov [ecx], FFFFFFh
  loc_0041B58D: jmp 0041B5EFh
  loc_0041B58F: cmp [00430A24h], esi
  loc_0041B595: jnz 0041B5A7h
  loc_0041B597: push 00430A24h
  loc_0041B59C: push 0040EC7Ch
  loc_0041B5A1: call [004010D4h] ; __vbaNew2
  loc_0041B5A7: mov ebx, [00430A24h]
  loc_0041B5AD: lea eax, var_1C
  loc_0041B5B0: push edi
  loc_0041B5B1: push eax
  loc_0041B5B2: mov edx, [ebx]
  loc_0041B5B4: mov var_B0, edx
  loc_0041B5BA: call [00401044h] ; __vbaObjSetAddref
  loc_0041B5C0: mov ecx, var_B0
  loc_0041B5C6: push eax
  loc_0041B5C7: push ebx
  loc_0041B5C8: call [ecx+00000010h]
  loc_0041B5CB: cmp eax, esi
  loc_0041B5CD: fnclex
  loc_0041B5CF: jge 0041B5E0h
  loc_0041B5D1: push 00000010h
  loc_0041B5D3: push 0040EC6Ch
  loc_0041B5D8: push ebx
  loc_0041B5D9: push eax
  loc_0041B5DA: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B5E0: lea ecx, var_1C
  loc_0041B5E3: call [00401114h] ; __vbaFreeObj
  loc_0041B5E9: call [0040101Ch] ; __vbaEnd
  loc_0041B5EF: mov var_4, esi
  loc_0041B5F2: push 0041B61Fh
  loc_0041B5F7: jmp 0041B61Eh
  loc_0041B5F9: lea ecx, var_1C
  loc_0041B5FC: call [00401114h] ; __vbaFreeObj
  loc_0041B602: lea edx, var_5C
  loc_0041B605: lea eax, var_4C
  loc_0041B608: push edx
  loc_0041B609: lea ecx, var_3C
  loc_0041B60C: push eax
  loc_0041B60D: lea edx, var_2C
  loc_0041B610: push ecx
  loc_0041B611: push edx
  loc_0041B612: push 00000004h
  loc_0041B614: call [00401018h] ; __vbaFreeVarList
  loc_0041B61A: add esp, 00000014h
  loc_0041B61D: ret
  loc_0041B61E: ret
  loc_0041B61F: mov eax, Me
  loc_0041B622: push eax
  loc_0041B623: mov ecx, [eax]
  loc_0041B625: call [ecx+00000008h]
  loc_0041B628: mov eax, var_4
  loc_0041B62B: mov ecx, var_14
  loc_0041B62E: pop edi
  loc_0041B62F: pop esi
  loc_0041B630: mov fs:[00000000h], ecx
  loc_0041B637: pop ebx
  loc_0041B638: mov esp, ebp
  loc_0041B63A: pop ebp
  loc_0041B63B: retn 000Ch
End Sub

Private Sub Form_Activate() '41A570
  loc_0041A570: push ebp
  loc_0041A571: mov ebp, esp
  loc_0041A573: sub esp, 0000000Ch
  loc_0041A576: push 00401926h ; __vbaExceptHandler
  loc_0041A57B: mov eax, fs:[00000000h]
  loc_0041A581: push eax
  loc_0041A582: mov fs:[00000000h], esp
  loc_0041A589: sub esp, 00000008h
  loc_0041A58C: push ebx
  loc_0041A58D: push esi
  loc_0041A58E: push edi
  loc_0041A58F: mov var_C, esp
  loc_0041A592: mov var_8, 00401130h
  loc_0041A599: mov esi, Me
  loc_0041A59C: mov eax, esi
  loc_0041A59E: and eax, 00000001h
  loc_0041A5A1: mov var_4, eax
  loc_0041A5A4: and esi, FFFFFFFEh
  loc_0041A5A7: push esi
  loc_0041A5A8: mov Me, esi
  loc_0041A5AB: mov ecx, [esi]
  loc_0041A5AD: call [ecx+00000004h]
  loc_0041A5B0: mov edx, [esi]
  loc_0041A5B2: push esi
  loc_0041A5B3: call [edx+00000714h]
  loc_0041A5B9: mov var_4, 00000000h
  loc_0041A5C0: mov eax, Me
  loc_0041A5C3: push eax
  loc_0041A5C4: mov ecx, [eax]
  loc_0041A5C6: call [ecx+00000008h]
  loc_0041A5C9: mov eax, var_4
  loc_0041A5CC: mov ecx, var_14
  loc_0041A5CF: pop edi
  loc_0041A5D0: pop esi
  loc_0041A5D1: mov fs:[00000000h], ecx
  loc_0041A5D8: pop ebx
  loc_0041A5D9: mov esp, ebp
  loc_0041A5DB: pop ebp
  loc_0041A5DC: retn 0004h
End Sub

Private Sub mnuSlope_Click() '41A9D0
  loc_0041A9D0: push ebp
  loc_0041A9D1: mov ebp, esp
  loc_0041A9D3: sub esp, 0000000Ch
  loc_0041A9D6: push 00401926h ; __vbaExceptHandler
  loc_0041A9DB: mov eax, fs:[00000000h]
  loc_0041A9E1: push eax
  loc_0041A9E2: mov fs:[00000000h], esp
  loc_0041A9E9: sub esp, 00000030h
  loc_0041A9EC: push ebx
  loc_0041A9ED: push esi
  loc_0041A9EE: push edi
  loc_0041A9EF: mov var_C, esp
  loc_0041A9F2: mov var_8, 00401160h
  loc_0041A9F9: mov eax, Me
  loc_0041A9FC: mov ecx, eax
  loc_0041A9FE: and ecx, 00000001h
  loc_0041AA01: mov var_4, ecx
  loc_0041AA04: and al, FEh
  loc_0041AA06: push eax
  loc_0041AA07: mov Me, eax
  loc_0041AA0A: mov edx, [eax]
  loc_0041AA0C: call [edx+00000004h]
  loc_0041AA0F: mov eax, [0043009Ch]
  loc_0041AA14: test eax, eax
  loc_0041AA16: jnz 0041AA28h
  loc_0041AA18: push 0043009Ch
  loc_0041AA1D: push 00405FC0h
  loc_0041AA22: call [004010D4h] ; __vbaNew2
  loc_0041AA28: mov esi, [0043009Ch]
  loc_0041AA2E: push esi
  loc_0041AA2F: mov eax, [esi]
  loc_0041AA31: call [eax+000002B4h]
  loc_0041AA37: test eax, eax
  loc_0041AA39: fnclex
  loc_0041AA3B: jge 0041AA4Fh
  loc_0041AA3D: push 000002B4h
  loc_0041AA42: push 0040DDE0h
  loc_0041AA47: push esi
  loc_0041AA48: push eax
  loc_0041AA49: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AA4F: mov eax, [004300C4h]
  loc_0041AA54: test eax, eax
  loc_0041AA56: jnz 0041AA68h
  loc_0041AA58: push 004300C4h
  loc_0041AA5D: push 00409228h
  loc_0041AA62: call [004010D4h] ; __vbaNew2
  loc_0041AA68: sub esp, 00000010h
  loc_0041AA6B: mov ecx, 0000000Ah
  loc_0041AA70: mov ebx, esp
  loc_0041AA72: mov var_24, ecx
  loc_0041AA75: mov eax, 80020004h
  loc_0041AA7A: sub esp, 00000010h
  loc_0041AA7D: mov [ebx], ecx
  loc_0041AA7F: mov ecx, var_30
  loc_0041AA82: mov edx, eax
  loc_0041AA84: mov esi, [004300C4h]
  loc_0041AA8A: mov [ebx+00000004h], ecx
  loc_0041AA8D: mov ecx, esp
  loc_0041AA8F: mov edi, [esi]
  loc_0041AA91: push esi
  loc_0041AA92: mov [ebx+00000008h], eax
  loc_0041AA95: mov eax, var_28
  loc_0041AA98: mov [ebx+0000000Ch], eax
  loc_0041AA9B: mov eax, var_24
  loc_0041AA9E: mov [ecx], eax
  loc_0041AAA0: mov eax, var_20
  loc_0041AAA3: mov [ecx+00000004h], eax
  loc_0041AAA6: mov [ecx+00000008h], edx
  loc_0041AAA9: mov edx, var_18
  loc_0041AAAC: mov [ecx+0000000Ch], edx
  loc_0041AAAF: call [edi+000002B0h]
  loc_0041AAB5: test eax, eax
  loc_0041AAB7: fnclex
  loc_0041AAB9: jge 0041AACDh
  loc_0041AABB: push 000002B0h
  loc_0041AAC0: push 0040E0ECh
  loc_0041AAC5: push esi
  loc_0041AAC6: push eax
  loc_0041AAC7: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AACD: mov var_4, 00000000h
  loc_0041AAD4: mov eax, Me
  loc_0041AAD7: push eax
  loc_0041AAD8: mov ecx, [eax]
  loc_0041AADA: call [ecx+00000008h]
  loc_0041AADD: mov eax, var_4
  loc_0041AAE0: mov ecx, var_14
  loc_0041AAE3: pop edi
  loc_0041AAE4: pop esi
  loc_0041AAE5: mov fs:[00000000h], ecx
  loc_0041AAEC: pop ebx
  loc_0041AAED: mov esp, ebp
  loc_0041AAEF: pop ebp
  loc_0041AAF0: retn 0004h
End Sub

Private Sub mnuChangeNames_Click() '41AB00
  loc_0041AB00: push ebp
  loc_0041AB01: mov ebp, esp
  loc_0041AB03: sub esp, 0000000Ch
  loc_0041AB06: push 00401926h ; __vbaExceptHandler
  loc_0041AB0B: mov eax, fs:[00000000h]
  loc_0041AB11: push eax
  loc_0041AB12: mov fs:[00000000h], esp
  loc_0041AB19: sub esp, 00000030h
  loc_0041AB1C: push ebx
  loc_0041AB1D: push esi
  loc_0041AB1E: push edi
  loc_0041AB1F: mov var_C, esp
  loc_0041AB22: mov var_8, 00401168h
  loc_0041AB29: mov eax, Me
  loc_0041AB2C: mov ecx, eax
  loc_0041AB2E: and ecx, 00000001h
  loc_0041AB31: mov var_4, ecx
  loc_0041AB34: and al, FEh
  loc_0041AB36: push eax
  loc_0041AB37: mov Me, eax
  loc_0041AB3A: mov edx, [eax]
  loc_0041AB3C: call [edx+00000004h]
  loc_0041AB3F: mov eax, [004300D8h]
  loc_0041AB44: test eax, eax
  loc_0041AB46: jnz 0041AB58h
  loc_0041AB48: push 004300D8h
  loc_0041AB4D: push 00402E04h
  loc_0041AB52: call [004010D4h] ; __vbaNew2
  loc_0041AB58: sub esp, 00000010h
  loc_0041AB5B: mov ecx, 0000000Ah
  loc_0041AB60: mov ebx, esp
  loc_0041AB62: mov var_24, ecx
  loc_0041AB65: mov eax, 80020004h
  loc_0041AB6A: sub esp, 00000010h
  loc_0041AB6D: mov [ebx], ecx
  loc_0041AB6F: mov ecx, var_30
  loc_0041AB72: mov edx, eax
  loc_0041AB74: mov esi, [004300D8h]
  loc_0041AB7A: mov [ebx+00000004h], ecx
  loc_0041AB7D: mov ecx, esp
  loc_0041AB7F: mov edi, [esi]
  loc_0041AB81: push esi
  loc_0041AB82: mov [ebx+00000008h], eax
  loc_0041AB85: mov eax, var_28
  loc_0041AB88: mov [ebx+0000000Ch], eax
  loc_0041AB8B: mov eax, var_24
  loc_0041AB8E: mov [ecx], eax
  loc_0041AB90: mov eax, var_20
  loc_0041AB93: mov [ecx+00000004h], eax
  loc_0041AB96: mov [ecx+00000008h], edx
  loc_0041AB99: mov edx, var_18
  loc_0041AB9C: mov [ecx+0000000Ch], edx
  loc_0041AB9F: call [edi+000002B0h]
  loc_0041ABA5: test eax, eax
  loc_0041ABA7: fnclex
  loc_0041ABA9: jge 0041ABBDh
  loc_0041ABAB: push 000002B0h
  loc_0041ABB0: push 0040E260h
  loc_0041ABB5: push esi
  loc_0041ABB6: push eax
  loc_0041ABB7: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ABBD: mov eax, Me
  loc_0041ABC0: push eax
  loc_0041ABC1: mov ecx, [eax]
  loc_0041ABC3: call [ecx+00000714h]
  loc_0041ABC9: mov var_4, 00000000h
  loc_0041ABD0: mov eax, Me
  loc_0041ABD3: push eax
  loc_0041ABD4: mov edx, [eax]
  loc_0041ABD6: call [edx+00000008h]
  loc_0041ABD9: mov eax, var_4
  loc_0041ABDC: mov ecx, var_14
  loc_0041ABDF: pop edi
  loc_0041ABE0: pop esi
  loc_0041ABE1: mov fs:[00000000h], ecx
  loc_0041ABE8: pop ebx
  loc_0041ABE9: mov esp, ebp
  loc_0041ABEB: pop ebp
  loc_0041ABEC: retn 0004h
End Sub

Private Sub mnuEstPred_Click() '41A650
  loc_0041A650: push ebp
  loc_0041A651: mov ebp, esp
  loc_0041A653: sub esp, 0000000Ch
  loc_0041A656: push 00401926h ; __vbaExceptHandler
  loc_0041A65B: mov eax, fs:[00000000h]
  loc_0041A661: push eax
  loc_0041A662: mov fs:[00000000h], esp
  loc_0041A669: sub esp, 00000030h
  loc_0041A66C: push ebx
  loc_0041A66D: push esi
  loc_0041A66E: push edi
  loc_0041A66F: mov var_C, esp
  loc_0041A672: mov var_8, 00401140h
  loc_0041A679: mov eax, Me
  loc_0041A67C: mov ecx, eax
  loc_0041A67E: and ecx, 00000001h
  loc_0041A681: mov var_4, ecx
  loc_0041A684: and al, FEh
  loc_0041A686: push eax
  loc_0041A687: mov Me, eax
  loc_0041A68A: mov edx, [eax]
  loc_0041A68C: call [edx+00000004h]
  loc_0041A68F: mov eax, [004300B0h]
  loc_0041A694: test eax, eax
  loc_0041A696: jnz 0041A6A8h
  loc_0041A698: push 004300B0h
  loc_0041A69D: push 00407528h
  loc_0041A6A2: call [004010D4h] ; __vbaNew2
  loc_0041A6A8: sub esp, 00000010h
  loc_0041A6AB: mov ecx, 0000000Ah
  loc_0041A6B0: mov ebx, esp
  loc_0041A6B2: mov var_24, ecx
  loc_0041A6B5: mov eax, 80020004h
  loc_0041A6BA: sub esp, 00000010h
  loc_0041A6BD: mov [ebx], ecx
  loc_0041A6BF: mov ecx, var_30
  loc_0041A6C2: mov edx, eax
  loc_0041A6C4: mov esi, [004300B0h]
  loc_0041A6CA: mov [ebx+00000004h], ecx
  loc_0041A6CD: mov ecx, esp
  loc_0041A6CF: mov edi, [esi]
  loc_0041A6D1: push esi
  loc_0041A6D2: mov [ebx+00000008h], eax
  loc_0041A6D5: mov eax, var_28
  loc_0041A6D8: mov [ebx+0000000Ch], eax
  loc_0041A6DB: mov eax, var_24
  loc_0041A6DE: mov [ecx], eax
  loc_0041A6E0: mov eax, var_20
  loc_0041A6E3: mov [ecx+00000004h], eax
  loc_0041A6E6: mov [ecx+00000008h], edx
  loc_0041A6E9: mov edx, var_18
  loc_0041A6EC: mov [ecx+0000000Ch], edx
  loc_0041A6EF: call [edi+000002B0h]
  loc_0041A6F5: test eax, eax
  loc_0041A6F7: fnclex
  loc_0041A6F9: jge 0041A70Dh
  loc_0041A6FB: push 000002B0h
  loc_0041A700: push 0040DE98h
  loc_0041A705: push esi
  loc_0041A706: push eax
  loc_0041A707: call [00401030h] ; __vbaHresultCheckObj
  loc_0041A70D: mov eax, [0043009Ch]
  loc_0041A712: test eax, eax
  loc_0041A714: jnz 0041A726h
  loc_0041A716: push 0043009Ch
  loc_0041A71B: push 00405FC0h
  loc_0041A720: call [004010D4h] ; __vbaNew2
  loc_0041A726: mov esi, [0043009Ch]
  loc_0041A72C: push esi
  loc_0041A72D: mov eax, [esi]
  loc_0041A72F: call [eax+000002B4h]
  loc_0041A735: test eax, eax
  loc_0041A737: fnclex
  loc_0041A739: jge 0041A74Dh
  loc_0041A73B: push 000002B4h
  loc_0041A740: push 0040DDE0h
  loc_0041A745: push esi
  loc_0041A746: push eax
  loc_0041A747: call [00401030h] ; __vbaHresultCheckObj
  loc_0041A74D: mov var_4, 00000000h
  loc_0041A754: mov eax, Me
  loc_0041A757: push eax
  loc_0041A758: mov ecx, [eax]
  loc_0041A75A: call [ecx+00000008h]
  loc_0041A75D: mov eax, var_4
  loc_0041A760: mov ecx, var_14
  loc_0041A763: pop edi
  loc_0041A764: pop esi
  loc_0041A765: mov fs:[00000000h], ecx
  loc_0041A76C: pop ebx
  loc_0041A76D: mov esp, ebp
  loc_0041A76F: pop ebp
  loc_0041A770: retn 0004h
End Sub

Private Sub Proc_1_10_41ABF0
  loc_0041ABF0: push ebp
  loc_0041ABF1: mov ebp, esp
  loc_0041ABF3: sub esp, 00000008h
  loc_0041ABF6: push 00401926h ; __vbaExceptHandler
  loc_0041ABFB: mov eax, fs:[00000000h]
  loc_0041AC01: push eax
  loc_0041AC02: mov fs:[00000000h], esp
  loc_0041AC09: sub esp, 00000044h
  loc_0041AC0C: push ebx
  loc_0041AC0D: push esi
  loc_0041AC0E: push edi
  loc_0041AC0F: mov var_8, esp
  loc_0041AC12: mov var_4, 00401170h
  loc_0041AC19: mov esi, Me
  loc_0041AC1C: xor eax, eax
  loc_0041AC1E: mov var_14, eax
  loc_0041AC21: mov var_18, eax
  loc_0041AC24: mov var_1C, eax
  loc_0041AC27: mov var_20, eax
  loc_0041AC2A: mov var_24, eax
  loc_0041AC2D: mov var_28, eax
  loc_0041AC30: mov var_2C, eax
  loc_0041AC33: mov eax, [esi]
  loc_0041AC35: push esi
  loc_0041AC36: call [eax+00000320h]
  loc_0041AC3C: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041AC42: lea ecx, var_2C
  loc_0041AC45: push eax
  loc_0041AC46: push ecx
  loc_0041AC47: call ebx
  loc_0041AC49: mov edi, eax
  loc_0041AC4B: push 0040E31Ch ; "The mean of the errors is zero for any given value of X."
  loc_0041AC50: push edi
  loc_0041AC51: mov edx, [edi]
  loc_0041AC53: call [edx+00000054h]
  loc_0041AC56: test eax, eax
  loc_0041AC58: fnclex
  loc_0041AC5A: jge 0041AC6Bh
  loc_0041AC5C: push 00000054h
  loc_0041AC5E: push 0040E390h
  loc_0041AC63: push edi
  loc_0041AC64: push eax
  loc_0041AC65: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AC6B: lea ecx, var_2C
  loc_0041AC6E: call [00401114h] ; __vbaFreeObj
  loc_0041AC74: mov eax, [esi]
  loc_0041AC76: push esi
  loc_0041AC77: call [eax+0000031Ch]
  loc_0041AC7D: lea ecx, var_2C
  loc_0041AC80: push eax
  loc_0041AC81: push ecx
  loc_0041AC82: call ebx
  loc_0041AC84: mov edi, eax
  loc_0041AC86: push 0040E3A4h ; "The errors are independent."
  loc_0041AC8B: push edi
  loc_0041AC8C: mov edx, [edi]
  loc_0041AC8E: call [edx+00000054h]
  loc_0041AC91: test eax, eax
  loc_0041AC93: fnclex
  loc_0041AC95: jge 0041ACA6h
  loc_0041AC97: push 00000054h
  loc_0041AC99: push 0040E390h
  loc_0041AC9E: push edi
  loc_0041AC9F: push eax
  loc_0041ACA0: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ACA6: lea ecx, var_2C
  loc_0041ACA9: call [00401114h] ; __vbaFreeObj
  loc_0041ACAF: mov eax, [esi]
  loc_0041ACB1: push esi
  loc_0041ACB2: call [eax+00000318h]
  loc_0041ACB8: lea ecx, var_2C
  loc_0041ACBB: push eax
  loc_0041ACBC: push ecx
  loc_0041ACBD: call ebx
  loc_0041ACBF: mov edi, eax
  loc_0041ACC1: push 0040E3E0h ; "The errors are normally distributed for any given value of X."
  loc_0041ACC6: push edi
  loc_0041ACC7: mov edx, [edi]
  loc_0041ACC9: call [edx+00000054h]
  loc_0041ACCC: test eax, eax
  loc_0041ACCE: fnclex
  loc_0041ACD0: jge 0041ACE1h
  loc_0041ACD2: push 00000054h
  loc_0041ACD4: push 0040E390h
  loc_0041ACD9: push edi
  loc_0041ACDA: push eax
  loc_0041ACDB: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ACE1: lea ecx, var_2C
  loc_0041ACE4: call [00401114h] ; __vbaFreeObj
  loc_0041ACEA: mov eax, [esi]
  loc_0041ACEC: push esi
  loc_0041ACED: call [eax+00000314h]
  loc_0041ACF3: lea ecx, var_2C
  loc_0041ACF6: push eax
  loc_0041ACF7: push ecx
  loc_0041ACF8: call ebx
  loc_0041ACFA: mov edi, eax
  loc_0041ACFC: push 0040E488h ; "The variance of the errors is a constant for any given value of ."
  loc_0041AD01: push edi
  loc_0041AD02: mov edx, [edi]
  loc_0041AD04: call [edx+00000054h]
  loc_0041AD07: test eax, eax
  loc_0041AD09: fnclex
  loc_0041AD0B: jge 0041AD1Ch
  loc_0041AD0D: push 00000054h
  loc_0041AD0F: push 0040E390h
  loc_0041AD14: push edi
  loc_0041AD15: push eax
  loc_0041AD16: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AD1C: lea ecx, var_2C
  loc_0041AD1F: call [00401114h] ; __vbaFreeObj
  loc_0041AD25: mov eax, [esi]
  loc_0041AD27: push esi
  loc_0041AD28: call [eax+00000320h]
  loc_0041AD2E: lea ecx, var_2C
  loc_0041AD31: push eax
  loc_0041AD32: push ecx
  loc_0041AD33: call ebx
  loc_0041AD35: mov edx, [00430018h]
  loc_0041AD3B: mov esi, [0040102Ch] ; __vbaStrCat
  loc_0041AD41: mov ebx, [eax]
  loc_0041AD43: push 0040E510h ; "The mean of the errors is zero for any value of "
  loc_0041AD48: push edx
  loc_0041AD49: mov var_30, eax
  loc_0041AD4C: call __vbaStrCat
  loc_0041AD4E: mov edi, [004010FCh] ; __vbaStrMove
  loc_0041AD54: mov edx, eax
  loc_0041AD56: lea ecx, var_14
  loc_0041AD59: call edi
  loc_0041AD5B: push eax
  loc_0041AD5C: push 0040DD3Ch ; "."
  loc_0041AD61: call __vbaStrCat
  loc_0041AD63: mov edx, eax
  loc_0041AD65: lea ecx, var_18
  loc_0041AD68: call edi
  loc_0041AD6A: mov var_3C, ebx
  loc_0041AD6D: mov ebx, var_30
  loc_0041AD70: push eax
  loc_0041AD71: mov eax, var_3C
  loc_0041AD74: push ebx
  loc_0041AD75: call [eax+0000019Ch]
  loc_0041AD7B: test eax, eax
  loc_0041AD7D: fnclex
  loc_0041AD7F: jge 0041AD93h
  loc_0041AD81: push 0000019Ch
  loc_0041AD86: push 0040E390h
  loc_0041AD8B: push ebx
  loc_0041AD8C: push eax
  loc_0041AD8D: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AD93: lea ecx, var_18
  loc_0041AD96: lea edx, var_14
  loc_0041AD99: push ecx
  loc_0041AD9A: push edx
  loc_0041AD9B: push 00000002h
  loc_0041AD9D: call [004010E4h] ; __vbaFreeStrList
  loc_0041ADA3: add esp, 0000000Ch
  loc_0041ADA6: lea ecx, var_2C
  loc_0041ADA9: call [00401114h] ; __vbaFreeObj
  loc_0041ADAF: mov eax, Me
  loc_0041ADB2: push eax
  loc_0041ADB3: mov ecx, [eax]
  loc_0041ADB5: call [ecx+0000031Ch]
  loc_0041ADBB: lea edx, var_2C
  loc_0041ADBE: push eax
  loc_0041ADBF: push edx
  loc_0041ADC0: call [0040103Ch] ; __vbaObjSet
  loc_0041ADC6: mov ebx, eax
  loc_0041ADC8: push 0040E3A4h ; "The errors are independent."
  loc_0041ADCD: push ebx
  loc_0041ADCE: mov eax, [ebx]
  loc_0041ADD0: call [eax+0000019Ch]
  loc_0041ADD6: test eax, eax
  loc_0041ADD8: fnclex
  loc_0041ADDA: jge 0041ADEEh
  loc_0041ADDC: push 0000019Ch
  loc_0041ADE1: push 0040E390h
  loc_0041ADE6: push ebx
  loc_0041ADE7: push eax
  loc_0041ADE8: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ADEE: lea ecx, var_2C
  loc_0041ADF1: call [00401114h] ; __vbaFreeObj
  loc_0041ADF7: mov eax, Me
  loc_0041ADFA: push eax
  loc_0041ADFB: mov ecx, [eax]
  loc_0041ADFD: call [ecx+00000318h]
  loc_0041AE03: lea edx, var_2C
  loc_0041AE06: push eax
  loc_0041AE07: push edx
  loc_0041AE08: call [0040103Ch] ; __vbaObjSet
  loc_0041AE0E: mov ebx, [eax]
  loc_0041AE10: mov var_30, eax
  loc_0041AE13: mov eax, [00430018h]
  loc_0041AE18: push 0040E578h ; "The errors are normally distributed for any value of "
  loc_0041AE1D: push eax
  loc_0041AE1E: call __vbaStrCat
  loc_0041AE20: mov edx, eax
  loc_0041AE22: lea ecx, var_14
  loc_0041AE25: call edi
  loc_0041AE27: push eax
  loc_0041AE28: push 0040DD3Ch ; "."
  loc_0041AE2D: call __vbaStrCat
  loc_0041AE2F: mov edx, eax
  loc_0041AE31: lea ecx, var_18
  loc_0041AE34: call edi
  loc_0041AE36: mov ecx, ebx
  loc_0041AE38: mov ebx, var_30
  loc_0041AE3B: push eax
  loc_0041AE3C: push ebx
  loc_0041AE3D: call [ecx+0000019Ch]
  loc_0041AE43: test eax, eax
  loc_0041AE45: fnclex
  loc_0041AE47: jge 0041AE5Bh
  loc_0041AE49: push 0000019Ch
  loc_0041AE4E: push 0040E390h
  loc_0041AE53: push ebx
  loc_0041AE54: push eax
  loc_0041AE55: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AE5B: lea edx, var_18
  loc_0041AE5E: lea eax, var_14
  loc_0041AE61: push edx
  loc_0041AE62: push eax
  loc_0041AE63: push 00000002h
  loc_0041AE65: call [004010E4h] ; __vbaFreeStrList
  loc_0041AE6B: add esp, 0000000Ch
  loc_0041AE6E: lea ecx, var_2C
  loc_0041AE71: call [00401114h] ; __vbaFreeObj
  loc_0041AE77: mov eax, Me
  loc_0041AE7A: push eax
  loc_0041AE7B: mov ecx, [eax]
  loc_0041AE7D: call [ecx+00000314h]
  loc_0041AE83: lea edx, var_2C
  loc_0041AE86: push eax
  loc_0041AE87: push edx
  loc_0041AE88: call [0040103Ch] ; __vbaObjSet
  loc_0041AE8E: mov ebx, [eax]
  loc_0041AE90: mov var_30, eax
  loc_0041AE93: mov eax, [00430018h]
  loc_0041AE98: push 0040E5E8h ; "The variance of the errors is a constant for any value of "
  loc_0041AE9D: push eax
  loc_0041AE9E: call __vbaStrCat
  loc_0041AEA0: mov edx, eax
  loc_0041AEA2: lea ecx, var_14
  loc_0041AEA5: call edi
  loc_0041AEA7: push eax
  loc_0041AEA8: push 0040DD3Ch ; "."
  loc_0041AEAD: call __vbaStrCat
  loc_0041AEAF: mov edx, eax
  loc_0041AEB1: lea ecx, var_18
  loc_0041AEB4: call edi
  loc_0041AEB6: mov ecx, ebx
  loc_0041AEB8: mov ebx, var_30
  loc_0041AEBB: push eax
  loc_0041AEBC: push ebx
  loc_0041AEBD: call [ecx+0000019Ch]
  loc_0041AEC3: test eax, eax
  loc_0041AEC5: fnclex
  loc_0041AEC7: jge 0041AEDBh
  loc_0041AEC9: push 0000019Ch
  loc_0041AECE: push 0040E390h
  loc_0041AED3: push ebx
  loc_0041AED4: push eax
  loc_0041AED5: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AEDB: lea edx, var_18
  loc_0041AEDE: lea eax, var_14
  loc_0041AEE1: push edx
  loc_0041AEE2: push eax
  loc_0041AEE3: push 00000002h
  loc_0041AEE5: call [004010E4h] ; __vbaFreeStrList
  loc_0041AEEB: add esp, 0000000Ch
  loc_0041AEEE: lea ecx, var_2C
  loc_0041AEF1: call [00401114h] ; __vbaFreeObj
  loc_0041AEF7: mov eax, Me
  loc_0041AEFA: push eax
  loc_0041AEFB: mov ecx, [eax]
  loc_0041AEFD: call [ecx+00000310h]
  loc_0041AF03: lea edx, var_2C
  loc_0041AF06: push eax
  loc_0041AF07: push edx
  loc_0041AF08: call [0040103Ch] ; __vbaObjSet
  loc_0041AF0E: mov ebx, [eax]
  loc_0041AF10: mov var_30, eax
  loc_0041AF13: mov eax, [00430010h]
  loc_0041AF18: push 0040E674h ; " Error is the difference between "
  loc_0041AF1D: push eax
  loc_0041AF1E: call __vbaStrCat
  loc_0041AF20: mov edx, eax
  loc_0041AF22: lea ecx, var_14
  loc_0041AF25: call edi
  loc_0041AF27: push eax
  loc_0041AF28: push 0040E6BCh ; " and the expected value of "
  loc_0041AF2D: call __vbaStrCat
  loc_0041AF2F: mov edx, eax
  loc_0041AF31: lea ecx, var_18
  loc_0041AF34: call edi
  loc_0041AF36: mov ecx, [00430010h]
  loc_0041AF3C: push eax
  loc_0041AF3D: push ecx
  loc_0041AF3E: call __vbaStrCat
  loc_0041AF40: mov edx, eax
  loc_0041AF42: lea ecx, var_1C
  loc_0041AF45: call edi
  loc_0041AF47: push eax
  loc_0041AF48: push 0040E6F8h ; " for a given value of "
  loc_0041AF4D: call __vbaStrCat
  loc_0041AF4F: mov edx, eax
  loc_0041AF51: lea ecx, var_20
  loc_0041AF54: call edi
  loc_0041AF56: mov edx, [00430018h]
  loc_0041AF5C: push eax
  loc_0041AF5D: push edx
  loc_0041AF5E: call __vbaStrCat
  loc_0041AF60: mov edx, eax
  loc_0041AF62: lea ecx, var_24
  loc_0041AF65: call edi
  loc_0041AF67: push eax
  loc_0041AF68: push 0040DD3Ch ; "."
  loc_0041AF6D: call __vbaStrCat
  loc_0041AF6F: mov edx, eax
  loc_0041AF71: lea ecx, var_28
  loc_0041AF74: call edi
  loc_0041AF76: mov var_48, ebx
  loc_0041AF79: mov ebx, var_30
  loc_0041AF7C: push eax
  loc_0041AF7D: mov eax, var_48
  loc_0041AF80: push ebx
  loc_0041AF81: call [eax+0000014Ch]
  loc_0041AF87: test eax, eax
  loc_0041AF89: fnclex
  loc_0041AF8B: jge 0041AF9Fh
  loc_0041AF8D: push 0000014Ch
  loc_0041AF92: push 0040E728h
  loc_0041AF97: push ebx
  loc_0041AF98: push eax
  loc_0041AF99: call [00401030h] ; __vbaHresultCheckObj
  loc_0041AF9F: lea ecx, var_28
  loc_0041AFA2: lea edx, var_24
  loc_0041AFA5: push ecx
  loc_0041AFA6: lea eax, var_20
  loc_0041AFA9: push edx
  loc_0041AFAA: lea ecx, var_1C
  loc_0041AFAD: push eax
  loc_0041AFAE: lea edx, var_18
  loc_0041AFB1: push ecx
  loc_0041AFB2: lea eax, var_14
  loc_0041AFB5: push edx
  loc_0041AFB6: push eax
  loc_0041AFB7: push 00000006h
  loc_0041AFB9: call [004010E4h] ; __vbaFreeStrList
  loc_0041AFBF: add esp, 0000001Ch
  loc_0041AFC2: lea ecx, var_2C
  loc_0041AFC5: call [00401114h] ; __vbaFreeObj
  loc_0041AFCB: mov ebx, Me
  loc_0041AFCE: push ebx
  loc_0041AFCF: mov ecx, [ebx]
  loc_0041AFD1: call [ecx+0000030Ch]
  loc_0041AFD7: lea edx, var_2C
  loc_0041AFDA: push eax
  loc_0041AFDB: push edx
  loc_0041AFDC: call [0040103Ch] ; __vbaObjSet
  loc_0041AFE2: mov ecx, [eax]
  loc_0041AFE4: push 0040E73Ch ; "L - Linearity: The mean Y is a linear function of X."
  loc_0041AFE9: push eax
  loc_0041AFEA: mov var_30, eax
  loc_0041AFED: call [ecx+00000054h]
  loc_0041AFF0: test eax, eax
  loc_0041AFF2: fnclex
  loc_0041AFF4: jge 0041B008h
  loc_0041AFF6: mov edx, var_30
  loc_0041AFF9: push 00000054h
  loc_0041AFFB: push 0040E390h
  loc_0041B000: push edx
  loc_0041B001: push eax
  loc_0041B002: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B008: lea ecx, var_2C
  loc_0041B00B: call [00401114h] ; __vbaFreeObj
  loc_0041B011: mov eax, [ebx]
  loc_0041B013: push ebx
  loc_0041B014: call [eax+00000308h]
  loc_0041B01A: lea ecx, var_2C
  loc_0041B01D: push eax
  loc_0041B01E: push ecx
  loc_0041B01F: call [0040103Ch] ; __vbaObjSet
  loc_0041B025: mov edx, [eax]
  loc_0041B027: push 0040E7ACh ; "I - Independence: The values of Y are independent for any value of X."
  loc_0041B02C: push eax
  loc_0041B02D: mov var_30, eax
  loc_0041B030: call [edx+00000054h]
  loc_0041B033: test eax, eax
  loc_0041B035: fnclex
  loc_0041B037: jge 0041B04Bh
  loc_0041B039: mov ecx, var_30
  loc_0041B03C: push 00000054h
  loc_0041B03E: push 0040E390h
  loc_0041B043: push ecx
  loc_0041B044: push eax
  loc_0041B045: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B04B: lea ecx, var_2C
  loc_0041B04E: call [00401114h] ; __vbaFreeObj
  loc_0041B054: mov edx, [ebx]
  loc_0041B056: push ebx
  loc_0041B057: call [edx+00000304h]
  loc_0041B05D: push eax
  loc_0041B05E: lea eax, var_2C
  loc_0041B061: push eax
  loc_0041B062: call [0040103Ch] ; __vbaObjSet
  loc_0041B068: mov ecx, [eax]
  loc_0041B06A: push 0040E854h ; "N - Normality: The values of Y are normally distributed for any value of X."
  loc_0041B06F: push eax
  loc_0041B070: mov var_30, eax
  loc_0041B073: call [ecx+00000054h]
  loc_0041B076: test eax, eax
  loc_0041B078: fnclex
  loc_0041B07A: jge 0041B08Eh
  loc_0041B07C: mov edx, var_30
  loc_0041B07F: push 00000054h
  loc_0041B081: push 0040E390h
  loc_0041B086: push edx
  loc_0041B087: push eax
  loc_0041B088: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B08E: lea ecx, var_2C
  loc_0041B091: call [00401114h] ; __vbaFreeObj
  loc_0041B097: mov eax, [ebx]
  loc_0041B099: push ebx
  loc_0041B09A: call [eax+00000300h]
  loc_0041B0A0: lea ecx, var_2C
  loc_0041B0A3: push eax
  loc_0041B0A4: push ecx
  loc_0041B0A5: call [0040103Ch] ; __vbaObjSet
  loc_0041B0AB: mov edx, [eax]
  loc_0041B0AD: push 0040E8F0h ; "E - Equal Variance: The variance of the values of Y is the same for any X"
  loc_0041B0B2: push eax
  loc_0041B0B3: mov var_30, eax
  loc_0041B0B6: call [edx+00000054h]
  loc_0041B0B9: test eax, eax
  loc_0041B0BB: fnclex
  loc_0041B0BD: jge 0041B0D1h
  loc_0041B0BF: mov ecx, var_30
  loc_0041B0C2: push 00000054h
  loc_0041B0C4: push 0040E390h
  loc_0041B0C9: push ecx
  loc_0041B0CA: push eax
  loc_0041B0CB: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B0D1: lea ecx, var_2C
  loc_0041B0D4: call [00401114h] ; __vbaFreeObj
  loc_0041B0DA: mov edx, [ebx]
  loc_0041B0DC: push ebx
  loc_0041B0DD: call [edx+0000030Ch]
  loc_0041B0E3: push eax
  loc_0041B0E4: lea eax, var_2C
  loc_0041B0E7: push eax
  loc_0041B0E8: call [0040103Ch] ; __vbaObjSet
  loc_0041B0EE: mov ecx, [00430010h]
  loc_0041B0F4: mov ebx, [eax]
  loc_0041B0F6: push 0040E988h ; "L - Linearity: The mean "
  loc_0041B0FB: push ecx
  loc_0041B0FC: mov var_30, eax
  loc_0041B0FF: call __vbaStrCat
  loc_0041B101: mov edx, eax
  loc_0041B103: lea ecx, var_14
  loc_0041B106: call edi
  loc_0041B108: push eax
  loc_0041B109: push 0040E9C0h ; "  is a linear function of "
  loc_0041B10E: call __vbaStrCat
  loc_0041B110: mov edx, eax
  loc_0041B112: lea ecx, var_18
  loc_0041B115: call edi
  loc_0041B117: mov edx, [00430018h]
  loc_0041B11D: push eax
  loc_0041B11E: push edx
  loc_0041B11F: call __vbaStrCat
  loc_0041B121: mov edx, eax
  loc_0041B123: lea ecx, var_1C
  loc_0041B126: call edi
  loc_0041B128: push eax
  loc_0041B129: push 0040DD3Ch ; "."
  loc_0041B12E: call __vbaStrCat
  loc_0041B130: mov edx, eax
  loc_0041B132: lea ecx, var_20
  loc_0041B135: call edi
  loc_0041B137: mov var_4C, ebx
  loc_0041B13A: mov ebx, var_30
  loc_0041B13D: push eax
  loc_0041B13E: mov eax, var_4C
  loc_0041B141: push ebx
  loc_0041B142: call [eax+0000019Ch]
  loc_0041B148: test eax, eax
  loc_0041B14A: fnclex
  loc_0041B14C: jge 0041B160h
  loc_0041B14E: push 0000019Ch
  loc_0041B153: push 0040E390h
  loc_0041B158: push ebx
  loc_0041B159: push eax
  loc_0041B15A: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B160: lea ecx, var_20
  loc_0041B163: lea edx, var_1C
  loc_0041B166: push ecx
  loc_0041B167: lea eax, var_18
  loc_0041B16A: push edx
  loc_0041B16B: lea ecx, var_14
  loc_0041B16E: push eax
  loc_0041B16F: push ecx
  loc_0041B170: push 00000004h
  loc_0041B172: call [004010E4h] ; __vbaFreeStrList
  loc_0041B178: add esp, 00000014h
  loc_0041B17B: lea ecx, var_2C
  loc_0041B17E: call [00401114h] ; __vbaFreeObj
  loc_0041B184: mov eax, Me
  loc_0041B187: push eax
  loc_0041B188: mov edx, [eax]
  loc_0041B18A: call [edx+00000308h]
  loc_0041B190: push eax
  loc_0041B191: lea eax, var_2C
  loc_0041B194: push eax
  loc_0041B195: call [0040103Ch] ; __vbaObjSet
  loc_0041B19B: mov ecx, [00430010h]
  loc_0041B1A1: mov ebx, [eax]
  loc_0041B1A3: push 0040EA30h ; "I - Independence: The values of "
  loc_0041B1A8: push ecx
  loc_0041B1A9: mov var_30, eax
  loc_0041B1AC: call __vbaStrCat
  loc_0041B1AE: mov edx, eax
  loc_0041B1B0: lea ecx, var_14
  loc_0041B1B3: call edi
  loc_0041B1B5: push eax
  loc_0041B1B6: push 0040EA78h ; " are independent for any given value of "
  loc_0041B1BB: call __vbaStrCat
  loc_0041B1BD: mov edx, eax
  loc_0041B1BF: lea ecx, var_18
  loc_0041B1C2: call edi
  loc_0041B1C4: mov edx, [00430018h]
  loc_0041B1CA: push eax
  loc_0041B1CB: push edx
  loc_0041B1CC: call __vbaStrCat
  loc_0041B1CE: mov edx, eax
  loc_0041B1D0: lea ecx, var_1C
  loc_0041B1D3: call edi
  loc_0041B1D5: push eax
  loc_0041B1D6: push 0040DD3Ch ; "."
  loc_0041B1DB: call __vbaStrCat
  loc_0041B1DD: mov edx, eax
  loc_0041B1DF: lea ecx, var_20
  loc_0041B1E2: call edi
  loc_0041B1E4: mov var_50, ebx
  loc_0041B1E7: mov ebx, var_30
  loc_0041B1EA: push eax
  loc_0041B1EB: mov eax, var_50
  loc_0041B1EE: push ebx
  loc_0041B1EF: call [eax+0000019Ch]
  loc_0041B1F5: test eax, eax
  loc_0041B1F7: fnclex
  loc_0041B1F9: jge 0041B20Dh
  loc_0041B1FB: push 0000019Ch
  loc_0041B200: push 0040E390h
  loc_0041B205: push ebx
  loc_0041B206: push eax
  loc_0041B207: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B20D: lea ecx, var_20
  loc_0041B210: lea edx, var_1C
  loc_0041B213: push ecx
  loc_0041B214: lea eax, var_18
  loc_0041B217: push edx
  loc_0041B218: lea ecx, var_14
  loc_0041B21B: push eax
  loc_0041B21C: push ecx
  loc_0041B21D: push 00000004h
  loc_0041B21F: call [004010E4h] ; __vbaFreeStrList
  loc_0041B225: add esp, 00000014h
  loc_0041B228: lea ecx, var_2C
  loc_0041B22B: call [00401114h] ; __vbaFreeObj
  loc_0041B231: mov eax, Me
  loc_0041B234: push eax
  loc_0041B235: mov edx, [eax]
  loc_0041B237: call [edx+00000304h]
  loc_0041B23D: push eax
  loc_0041B23E: lea eax, var_2C
  loc_0041B241: push eax
  loc_0041B242: call [0040103Ch] ; __vbaObjSet
  loc_0041B248: mov ecx, [00430010h]
  loc_0041B24E: mov ebx, [eax]
  loc_0041B250: push 0040EAD0h ; "N - Normality: The values of "
  loc_0041B255: push ecx
  loc_0041B256: mov var_30, eax
  loc_0041B259: call __vbaStrCat
  loc_0041B25B: mov edx, eax
  loc_0041B25D: lea ecx, var_14
  loc_0041B260: call edi
  loc_0041B262: push eax
  loc_0041B263: push 0040EB10h ; " are normally distributed for any given value of "
  loc_0041B268: call __vbaStrCat
  loc_0041B26A: mov edx, eax
  loc_0041B26C: lea ecx, var_18
  loc_0041B26F: call edi
  loc_0041B271: mov edx, [00430018h]
  loc_0041B277: push eax
  loc_0041B278: push edx
  loc_0041B279: call __vbaStrCat
  loc_0041B27B: mov edx, eax
  loc_0041B27D: lea ecx, var_1C
  loc_0041B280: call edi
  loc_0041B282: push eax
  loc_0041B283: push 0040DD3Ch ; "."
  loc_0041B288: call __vbaStrCat
  loc_0041B28A: mov edx, eax
  loc_0041B28C: lea ecx, var_20
  loc_0041B28F: call edi
  loc_0041B291: mov var_54, ebx
  loc_0041B294: mov ebx, var_30
  loc_0041B297: push eax
  loc_0041B298: mov eax, var_54
  loc_0041B29B: push ebx
  loc_0041B29C: call [eax+0000019Ch]
  loc_0041B2A2: test eax, eax
  loc_0041B2A4: fnclex
  loc_0041B2A6: jge 0041B2BAh
  loc_0041B2A8: push 0000019Ch
  loc_0041B2AD: push 0040E390h
  loc_0041B2B2: push ebx
  loc_0041B2B3: push eax
  loc_0041B2B4: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B2BA: lea ecx, var_20
  loc_0041B2BD: lea edx, var_1C
  loc_0041B2C0: push ecx
  loc_0041B2C1: lea eax, var_18
  loc_0041B2C4: push edx
  loc_0041B2C5: lea ecx, var_14
  loc_0041B2C8: push eax
  loc_0041B2C9: push ecx
  loc_0041B2CA: push 00000004h
  loc_0041B2CC: call [004010E4h] ; __vbaFreeStrList
  loc_0041B2D2: add esp, 00000014h
  loc_0041B2D5: lea ecx, var_2C
  loc_0041B2D8: call [00401114h] ; __vbaFreeObj
  loc_0041B2DE: mov eax, Me
  loc_0041B2E1: push eax
  loc_0041B2E2: mov edx, [eax]
  loc_0041B2E4: call [edx+00000300h]
  loc_0041B2EA: push eax
  loc_0041B2EB: lea eax, var_2C
  loc_0041B2EE: push eax
  loc_0041B2EF: call [0040103Ch] ; __vbaObjSet
  loc_0041B2F5: mov ecx, [00430010h]
  loc_0041B2FB: mov ebx, [eax]
  loc_0041B2FD: push 0040EB78h ; "E - Equal Variance: The variance of the values of "
  loc_0041B302: push ecx
  loc_0041B303: mov var_30, eax
  loc_0041B306: call __vbaStrCat
  loc_0041B308: mov edx, eax
  loc_0041B30A: lea ecx, var_14
  loc_0041B30D: call edi
  loc_0041B30F: push eax
  loc_0041B310: push 0040EC0Ch ; " is a constant for any given value of "
  loc_0041B315: call __vbaStrCat
  loc_0041B317: mov edx, eax
  loc_0041B319: lea ecx, var_18
  loc_0041B31C: call edi
  loc_0041B31E: mov edx, [00430018h]
  loc_0041B324: push eax
  loc_0041B325: push edx
  loc_0041B326: call __vbaStrCat
  loc_0041B328: mov edx, eax
  loc_0041B32A: lea ecx, var_1C
  loc_0041B32D: call edi
  loc_0041B32F: push eax
  loc_0041B330: push 0040DD3Ch ; "."
  loc_0041B335: call __vbaStrCat
  loc_0041B337: mov edx, eax
  loc_0041B339: lea ecx, var_20
  loc_0041B33C: call edi
  loc_0041B33E: mov esi, var_30
  loc_0041B341: push eax
  loc_0041B342: push esi
  loc_0041B343: call [ebx+0000019Ch]
  loc_0041B349: test eax, eax
  loc_0041B34B: fnclex
  loc_0041B34D: jge 0041B361h
  loc_0041B34F: push 0000019Ch
  loc_0041B354: push 0040E390h
  loc_0041B359: push esi
  loc_0041B35A: push eax
  loc_0041B35B: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B361: lea eax, var_20
  loc_0041B364: lea ecx, var_1C
  loc_0041B367: push eax
  loc_0041B368: lea edx, var_18
  loc_0041B36B: push ecx
  loc_0041B36C: lea eax, var_14
  loc_0041B36F: push edx
  loc_0041B370: push eax
  loc_0041B371: push 00000004h
  loc_0041B373: call [004010E4h] ; __vbaFreeStrList
  loc_0041B379: add esp, 00000014h
  loc_0041B37C: lea ecx, var_2C
  loc_0041B37F: call [00401114h] ; __vbaFreeObj
  loc_0041B385: push 0041B3BAh
  loc_0041B38A: jmp 0041B3B9h
  loc_0041B38C: lea ecx, var_28
  loc_0041B38F: lea edx, var_24
  loc_0041B392: push ecx
  loc_0041B393: lea eax, var_20
  loc_0041B396: push edx
  loc_0041B397: lea ecx, var_1C
  loc_0041B39A: push eax
  loc_0041B39B: lea edx, var_18
  loc_0041B39E: push ecx
  loc_0041B39F: lea eax, var_14
  loc_0041B3A2: push edx
  loc_0041B3A3: push eax
  loc_0041B3A4: push 00000006h
  loc_0041B3A6: call [004010E4h] ; __vbaFreeStrList
  loc_0041B3AC: add esp, 0000001Ch
  loc_0041B3AF: lea ecx, var_2C
  loc_0041B3B2: call [00401114h] ; __vbaFreeObj
  loc_0041B3B8: ret
  loc_0041B3B9: ret
  loc_0041B3BA: mov ecx, var_10
  loc_0041B3BD: pop edi
  loc_0041B3BE: pop esi
  loc_0041B3BF: xor eax, eax
  loc_0041B3C1: mov fs:[00000000h], ecx
  loc_0041B3C8: pop ebx
  loc_0041B3C9: mov esp, ebp
  loc_0041B3CB: pop ebp
  loc_0041B3CC: retn 0004h
End Sub
