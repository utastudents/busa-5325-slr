VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "TABCTL32.OCX"
Begin VB.Form frmEstimationPrediction
  Caption = "Estimation and Prediction Intervals"
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 510
  ClientWidth = 8595
  ClientHeight = 5055
  StartUpPosition = 2 'CenterScreen
  Begin TabDlg.SSTab SSTab1
    Left = 120
    Top = 480
    Width = 11655
    Height = 7215
    TabIndex = 0
    OleObjectBlob = "frmEstimationPrediction.frx":0000
    Begin VB.Frame Frame6
      Caption = "Conclusion:"
      Left = -74760
      Top = 4680
      Width = 11175
      Height = 2175
      TabIndex = 11
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      Begin VB.Label lblPredCon
        Left = 120
        Top = 480
        Width = 10935
        Height = 1455
        TabIndex = 12
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
    Begin VB.Frame Frame5
      Caption = "Substitution"
      Left = -74760
      Top = 2640
      Width = 11175
      Height = 1815
      TabIndex = 9
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      Begin VB.Label lblPredSub
        Left = 240
        Top = 480
        Width = 10695
        Height = 1095
        TabIndex = 10
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
    Begin VB.Frame Frame4
      Caption = "Formula"
      Left = -74760
      Top = 840
      Width = 11175
      Height = 1575
      TabIndex = 7
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      Begin VB.Label Label2
        Caption = "(Point Estimate of Y | X) ±  t(n-2, 1-alpha/2) (Estimate of Standard Error of the Y Prediction)"
        Left = 120
        Top = 480
        Width = 10815
        Height = 975
        TabIndex = 8
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
    Begin VB.Frame Frame1
      Caption = "Formula"
      Left = 120
      Top = 840
      Width = 11295
      Height = 1455
      TabIndex = 5
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      Begin VB.Label Label1
        Caption = "(Point Estimate of the mean) ±  t(n-2, 1-alpha/2) (Estimate of Standard Error of Mean Estimate)"
        Left = 120
        Top = 480
        Width = 10935
        Height = 855
        TabIndex = 6
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
    Begin VB.Frame Frame2
      Caption = "Substitution"
      Left = 120
      Top = 2400
      Width = 11295
      Height = 1695
      TabIndex = 3
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      Begin VB.Label lblEstSub
        Left = 120
        Top = 360
        Width = 10695
        Height = 1095
        TabIndex = 4
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
    Begin VB.Frame Frame3
      Caption = "Conclusion:"
      Left = 120
      Top = 4200
      Width = 11295
      Height = 2775
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
      Begin VB.Label lblEstCon
        Caption = "With 100(1-alpha/2)% confidence, we can say that X takes on the value of Xg, the value of Y falls in the range  ______ to _____"
        Left = 120
        Top = 600
        Width = 11055
        Height = 2055
        TabIndex = 2
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
  End
  Begin VB.Menu mnuChange
    Caption = "&Change"
    Begin VB.Menu mnuChangeNames
      Caption = "&Variable Names"
    End
    Begin VB.Menu mnuChangeUnits
      Caption = "&Units"
    End
    Begin VB.Menu mnuChangettable
      Caption = "Alpha and &T-table Values"
    End
    Begin VB.Menu mnuChangePredEst
      Caption = "&Prediction and Estimation Values"
    End
  End
  Begin VB.Menu mnuGoTo
    Caption = "&Go To"
    Begin VB.Menu mnuIntroduction
      Caption = "&Introduction"
    End
    Begin VB.Menu mnuStatistics
      Caption = "Statistics and Point Esti&mates"
    End
    Begin VB.Menu mnuQuestions
      Caption = "&Questions"
    End
    Begin VB.Menu mnuSlope
      Caption = "&Inferences on the Slope"
    End
    Begin VB.Menu mnuAssuptions
      Caption = "&Assumptions"
    End
  End
  Begin VB.Menu mnuExit
    Caption = "&Exit"
  End
End

Attribute VB_Name = "frmEstimationPrediction"


Private Sub mnuExit_Click() '4248E0
  loc_004248E0: push ebp
  loc_004248E1: mov ebp, esp
  loc_004248E3: sub esp, 0000000Ch
  loc_004248E6: push 00401926h ; __vbaExceptHandler
  loc_004248EB: mov eax, fs:[00000000h]
  loc_004248F1: push eax
  loc_004248F2: mov fs:[00000000h], esp
  loc_004248F9: sub esp, 00000018h
  loc_004248FC: push ebx
  loc_004248FD: push esi
  loc_004248FE: push edi
  loc_004248FF: mov var_C, esp
  loc_00424902: mov var_8, 00401560h
  loc_00424909: mov edi, Me
  loc_0042490C: mov eax, edi
  loc_0042490E: and eax, 00000001h
  loc_00424911: mov var_4, eax
  loc_00424914: and edi, FFFFFFFEh
  loc_00424917: push edi
  loc_00424918: mov Me, edi
  loc_0042491B: mov ecx, [edi]
  loc_0042491D: call [ecx+00000004h]
  loc_00424920: mov eax, [00430A24h]
  loc_00424925: xor ebx, ebx
  loc_00424927: cmp eax, ebx
  loc_00424929: mov var_18, ebx
  loc_0042492C: jnz 0042493Eh
  loc_0042492E: push 00430A24h
  loc_00424933: push 0040EC7Ch
  loc_00424938: call [004010D4h] ; __vbaNew2
  loc_0042493E: mov esi, [00430A24h]
  loc_00424944: lea eax, var_18
  loc_00424947: push edi
  loc_00424948: push eax
  loc_00424949: mov edx, [esi]
  loc_0042494B: mov var_2C, edx
  loc_0042494E: call [00401044h] ; __vbaObjSetAddref
  loc_00424954: mov ecx, var_2C
  loc_00424957: push eax
  loc_00424958: push esi
  loc_00424959: call [ecx+00000010h]
  loc_0042495C: cmp eax, ebx
  loc_0042495E: fnclex
  loc_00424960: jge 00424971h
  loc_00424962: push 00000010h
  loc_00424964: push 0040EC6Ch
  loc_00424969: push esi
  loc_0042496A: push eax
  loc_0042496B: call [00401030h] ; __vbaHresultCheckObj
  loc_00424971: lea ecx, var_18
  loc_00424974: call [00401114h] ; __vbaFreeObj
  loc_0042497A: mov var_4, ebx
  loc_0042497D: push 0042498Fh
  loc_00424982: jmp 0042498Eh
  loc_00424984: lea ecx, var_18
  loc_00424987: call [00401114h] ; __vbaFreeObj
  loc_0042498D: ret
  loc_0042498E: ret
  loc_0042498F: mov eax, Me
  loc_00424992: push eax
  loc_00424993: mov edx, [eax]
  loc_00424995: call [edx+00000008h]
  loc_00424998: mov eax, var_4
  loc_0042499B: mov ecx, var_14
  loc_0042499E: pop edi
  loc_0042499F: pop esi
  loc_004249A0: mov fs:[00000000h], ecx
  loc_004249A7: pop ebx
  loc_004249A8: mov esp, ebp
  loc_004249AA: pop ebp
  loc_004249AB: retn 0004h
End Sub

Private Sub mnuQuestions_Click() '424C80
  loc_00424C80: push ebp
  loc_00424C81: mov ebp, esp
  loc_00424C83: sub esp, 0000000Ch
  loc_00424C86: push 00401926h ; __vbaExceptHandler
  loc_00424C8B: mov eax, fs:[00000000h]
  loc_00424C91: push eax
  loc_00424C92: mov fs:[00000000h], esp
  loc_00424C99: sub esp, 00000030h
  loc_00424C9C: push ebx
  loc_00424C9D: push esi
  loc_00424C9E: push edi
  loc_00424C9F: mov var_C, esp
  loc_00424CA2: mov var_8, 00401588h
  loc_00424CA9: mov eax, Me
  loc_00424CAC: mov ecx, eax
  loc_00424CAE: and ecx, 00000001h
  loc_00424CB1: mov var_4, ecx
  loc_00424CB4: and al, FEh
  loc_00424CB6: push eax
  loc_00424CB7: mov Me, eax
  loc_00424CBA: mov edx, [eax]
  loc_00424CBC: call [edx+00000004h]
  loc_00424CBF: mov eax, [00430164h]
  loc_00424CC4: test eax, eax
  loc_00424CC6: jnz 00424CD8h
  loc_00424CC8: push 00430164h
  loc_00424CCD: push 00408108h
  loc_00424CD2: call [004010D4h] ; __vbaNew2
  loc_00424CD8: sub esp, 00000010h
  loc_00424CDB: mov ecx, 0000000Ah
  loc_00424CE0: mov ebx, esp
  loc_00424CE2: mov var_24, ecx
  loc_00424CE5: mov eax, 80020004h
  loc_00424CEA: sub esp, 00000010h
  loc_00424CED: mov [ebx], ecx
  loc_00424CEF: mov ecx, var_30
  loc_00424CF2: mov edx, eax
  loc_00424CF4: mov esi, [00430164h]
  loc_00424CFA: mov [ebx+00000004h], ecx
  loc_00424CFD: mov ecx, esp
  loc_00424CFF: mov edi, [esi]
  loc_00424D01: push esi
  loc_00424D02: mov [ebx+00000008h], eax
  loc_00424D05: mov eax, var_28
  loc_00424D08: mov [ebx+0000000Ch], eax
  loc_00424D0B: mov eax, var_24
  loc_00424D0E: mov [ecx], eax
  loc_00424D10: mov eax, var_20
  loc_00424D13: mov [ecx+00000004h], eax
  loc_00424D16: mov [ecx+00000008h], edx
  loc_00424D19: mov edx, var_18
  loc_00424D1C: mov [ecx+0000000Ch], edx
  loc_00424D1F: call [edi+000002B0h]
  loc_00424D25: test eax, eax
  loc_00424D27: fnclex
  loc_00424D29: jge 00424D3Dh
  loc_00424D2B: push 000002B0h
  loc_00424D30: push 0040FF58h
  loc_00424D35: push esi
  loc_00424D36: push eax
  loc_00424D37: call [00401030h] ; __vbaHresultCheckObj
  loc_00424D3D: mov eax, [004300B0h]
  loc_00424D42: test eax, eax
  loc_00424D44: jnz 00424D56h
  loc_00424D46: push 004300B0h
  loc_00424D4B: push 00407528h
  loc_00424D50: call [004010D4h] ; __vbaNew2
  loc_00424D56: mov esi, [004300B0h]
  loc_00424D5C: push esi
  loc_00424D5D: mov eax, [esi]
  loc_00424D5F: call [eax+000002B4h]
  loc_00424D65: test eax, eax
  loc_00424D67: fnclex
  loc_00424D69: jge 00424D7Dh
  loc_00424D6B: push 000002B4h
  loc_00424D70: push 0040DE98h
  loc_00424D75: push esi
  loc_00424D76: push eax
  loc_00424D77: call [00401030h] ; __vbaHresultCheckObj
  loc_00424D7D: mov var_4, 00000000h
  loc_00424D84: mov eax, Me
  loc_00424D87: push eax
  loc_00424D88: mov ecx, [eax]
  loc_00424D8A: call [ecx+00000008h]
  loc_00424D8D: mov eax, var_4
  loc_00424D90: mov ecx, var_14
  loc_00424D93: pop edi
  loc_00424D94: pop esi
  loc_00424D95: mov fs:[00000000h], ecx
  loc_00424D9C: pop ebx
  loc_00424D9D: mov esp, ebp
  loc_00424D9F: pop ebp
  loc_00424DA0: retn 0004h
End Sub

Private Sub mnuChangettable_Click() '424520
  loc_00424520: push ebp
  loc_00424521: mov ebp, esp
  loc_00424523: sub esp, 0000000Ch
  loc_00424526: push 00401926h ; __vbaExceptHandler
  loc_0042452B: mov eax, fs:[00000000h]
  loc_00424531: push eax
  loc_00424532: mov fs:[00000000h], esp
  loc_00424539: sub esp, 00000030h
  loc_0042453C: push ebx
  loc_0042453D: push esi
  loc_0042453E: push edi
  loc_0042453F: mov var_C, esp
  loc_00424542: mov var_8, 00401540h
  loc_00424549: mov eax, Me
  loc_0042454C: mov ecx, eax
  loc_0042454E: and ecx, 00000001h
  loc_00424551: mov var_4, ecx
  loc_00424554: and al, FEh
  loc_00424556: push eax
  loc_00424557: mov Me, eax
  loc_0042455A: mov edx, [eax]
  loc_0042455C: call [edx+00000004h]
  loc_0042455F: mov eax, [00430178h]
  loc_00424564: test eax, eax
  loc_00424566: jnz 00424578h
  loc_00424568: push 00430178h
  loc_0042456D: push 004051D0h
  loc_00424572: call [004010D4h] ; __vbaNew2
  loc_00424578: sub esp, 00000010h
  loc_0042457B: mov ecx, 0000000Ah
  loc_00424580: mov ebx, esp
  loc_00424582: mov var_24, ecx
  loc_00424585: mov eax, 80020004h
  loc_0042458A: sub esp, 00000010h
  loc_0042458D: mov [ebx], ecx
  loc_0042458F: mov ecx, var_30
  loc_00424592: mov edx, eax
  loc_00424594: mov esi, [00430178h]
  loc_0042459A: mov [ebx+00000004h], ecx
  loc_0042459D: mov ecx, esp
  loc_0042459F: mov edi, [esi]
  loc_004245A1: push esi
  loc_004245A2: mov [ebx+00000008h], eax
  loc_004245A5: mov eax, var_28
  loc_004245A8: mov [ebx+0000000Ch], eax
  loc_004245AB: mov eax, var_24
  loc_004245AE: mov [ecx], eax
  loc_004245B0: mov eax, var_20
  loc_004245B3: mov [ecx+00000004h], eax
  loc_004245B6: mov [ecx+00000008h], edx
  loc_004245B9: mov edx, var_18
  loc_004245BC: mov [ecx+0000000Ch], edx
  loc_004245BF: call [edi+000002B0h]
  loc_004245C5: test eax, eax
  loc_004245C7: fnclex
  loc_004245C9: jge 004245DDh
  loc_004245CB: push 000002B0h
  loc_004245D0: push 004100A4h
  loc_004245D5: push esi
  loc_004245D6: push eax
  loc_004245D7: call [00401030h] ; __vbaHresultCheckObj
  loc_004245DD: mov var_4, 00000000h
  loc_004245E4: mov eax, Me
  loc_004245E7: push eax
  loc_004245E8: mov ecx, [eax]
  loc_004245EA: call [ecx+00000008h]
  loc_004245ED: mov eax, var_4
  loc_004245F0: mov ecx, var_14
  loc_004245F3: pop edi
  loc_004245F4: pop esi
  loc_004245F5: mov fs:[00000000h], ecx
  loc_004245FC: pop ebx
  loc_004245FD: mov esp, ebp
  loc_004245FF: pop ebp
  loc_00424600: retn 0004h
End Sub

Private Sub Form_Load() '423B30
  loc_00423B30: push ebp
  loc_00423B31: mov ebp, esp
  loc_00423B33: sub esp, 0000000Ch
  loc_00423B36: push 00401926h ; __vbaExceptHandler
  loc_00423B3B: mov eax, fs:[00000000h]
  loc_00423B41: push eax
  loc_00423B42: mov fs:[00000000h], esp
  loc_00423B49: sub esp, 00000008h
  loc_00423B4C: push ebx
  loc_00423B4D: push esi
  loc_00423B4E: push edi
  loc_00423B4F: mov var_C, esp
  loc_00423B52: mov var_8, 00401518h
  loc_00423B59: mov esi, Me
  loc_00423B5C: mov eax, esi
  loc_00423B5E: and eax, 00000001h
  loc_00423B61: mov var_4, eax
  loc_00423B64: and esi, FFFFFFFEh
  loc_00423B67: push esi
  loc_00423B68: mov Me, esi
  loc_00423B6B: mov ecx, [esi]
  loc_00423B6D: call [ecx+00000004h]
  loc_00423B70: mov edx, [esi]
  loc_00423B72: push esi
  loc_00423B73: call [edx+00000700h]
  loc_00423B79: mov var_4, 00000000h
  loc_00423B80: mov eax, Me
  loc_00423B83: push eax
  loc_00423B84: mov ecx, [eax]
  loc_00423B86: call [ecx+00000008h]
  loc_00423B89: mov eax, var_4
  loc_00423B8C: mov ecx, var_14
  loc_00423B8F: pop edi
  loc_00423B90: pop esi
  loc_00423B91: mov fs:[00000000h], ecx
  loc_00423B98: pop ebx
  loc_00423B99: mov esp, ebp
  loc_00423B9B: pop ebp
  loc_00423B9C: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '4249B0
  loc_004249B0: push ebp
  loc_004249B1: mov ebp, esp
  loc_004249B3: sub esp, 0000000Ch
  loc_004249B6: push 00401926h ; __vbaExceptHandler
  loc_004249BB: mov eax, fs:[00000000h]
  loc_004249C1: push eax
  loc_004249C2: mov fs:[00000000h], esp
  loc_004249C9: sub esp, 0000009Ch
  loc_004249CF: push ebx
  loc_004249D0: push esi
  loc_004249D1: push edi
  loc_004249D2: mov var_C, esp
  loc_004249D5: mov var_8, 00401570h
  loc_004249DC: mov edi, Me
  loc_004249DF: mov eax, edi
  loc_004249E1: and eax, 00000001h
  loc_004249E4: mov var_4, eax
  loc_004249E7: and edi, FFFFFFFEh
  loc_004249EA: push edi
  loc_004249EB: mov Me, edi
  loc_004249EE: mov ecx, [edi]
  loc_004249F0: call [ecx+00000004h]
  loc_004249F3: mov ebx, [004010F4h] ; __vbaVarDup
  loc_004249F9: mov ecx, 80020004h
  loc_004249FE: xor esi, esi
  loc_00424A00: mov var_54, ecx
  loc_00424A03: mov eax, 0000000Ah
  loc_00424A08: mov var_44, ecx
  loc_00424A0B: mov var_4C, esi
  loc_00424A0E: mov var_5C, esi
  loc_00424A11: mov var_7C, esi
  loc_00424A14: lea edx, var_7C
  loc_00424A17: lea ecx, var_3C
  loc_00424A1A: mov var_1C, esi
  loc_00424A1D: mov var_2C, esi
  loc_00424A20: mov var_3C, esi
  loc_00424A23: mov var_6C, esi
  loc_00424A26: mov var_5C, eax
  loc_00424A29: mov var_4C, eax
  loc_00424A2C: mov var_74, 0040ECD4h ; "Exit Check"
  loc_00424A33: mov var_7C, 00000008h
  loc_00424A3A: call ebx
  loc_00424A3C: lea edx, var_6C
  loc_00424A3F: lea ecx, var_2C
  loc_00424A42: mov var_64, 0040EC90h ; "Are you sure you want to exit?"
  loc_00424A49: mov var_6C, 00000008h
  loc_00424A50: call ebx
  loc_00424A52: lea edx, var_5C
  loc_00424A55: lea eax, var_4C
  loc_00424A58: push edx
  loc_00424A59: lea ecx, var_3C
  loc_00424A5C: push eax
  loc_00424A5D: push ecx
  loc_00424A5E: lea edx, var_2C
  loc_00424A61: push 00000004h
  loc_00424A63: push edx
  loc_00424A64: call [00401038h] ; rtcMsgBox
  loc_00424A6A: mov ecx, eax
  loc_00424A6C: call [00401078h] ; __vbaI2I4
  loc_00424A72: mov ebx, eax
  loc_00424A74: lea eax, var_5C
  loc_00424A77: lea ecx, var_4C
  loc_00424A7A: push eax
  loc_00424A7B: lea edx, var_3C
  loc_00424A7E: push ecx
  loc_00424A7F: lea eax, var_2C
  loc_00424A82: push edx
  loc_00424A83: push eax
  loc_00424A84: push 00000004h
  loc_00424A86: call [00401018h] ; __vbaFreeVarList
  loc_00424A8C: add esp, 00000014h
  loc_00424A8F: cmp bx, 0007h
  loc_00424A93: jnz 00424A9Fh
  loc_00424A95: mov ecx, Cancel
  loc_00424A98: mov [ecx], FFFFFFh
  loc_00424A9D: jmp 00424AFFh
  loc_00424A9F: cmp [00430A24h], esi
  loc_00424AA5: jnz 00424AB7h
  loc_00424AA7: push 00430A24h
  loc_00424AAC: push 0040EC7Ch
  loc_00424AB1: call [004010D4h] ; __vbaNew2
  loc_00424AB7: mov ebx, [00430A24h]
  loc_00424ABD: lea eax, var_1C
  loc_00424AC0: push edi
  loc_00424AC1: push eax
  loc_00424AC2: mov edx, [ebx]
  loc_00424AC4: mov var_B0, edx
  loc_00424ACA: call [00401044h] ; __vbaObjSetAddref
  loc_00424AD0: mov ecx, var_B0
  loc_00424AD6: push eax
  loc_00424AD7: push ebx
  loc_00424AD8: call [ecx+00000010h]
  loc_00424ADB: cmp eax, esi
  loc_00424ADD: fnclex
  loc_00424ADF: jge 00424AF0h
  loc_00424AE1: push 00000010h
  loc_00424AE3: push 0040EC6Ch
  loc_00424AE8: push ebx
  loc_00424AE9: push eax
  loc_00424AEA: call [00401030h] ; __vbaHresultCheckObj
  loc_00424AF0: lea ecx, var_1C
  loc_00424AF3: call [00401114h] ; __vbaFreeObj
  loc_00424AF9: call [0040101Ch] ; __vbaEnd
  loc_00424AFF: mov var_4, esi
  loc_00424B02: push 00424B2Fh
  loc_00424B07: jmp 00424B2Eh
  loc_00424B09: lea ecx, var_1C
  loc_00424B0C: call [00401114h] ; __vbaFreeObj
  loc_00424B12: lea edx, var_5C
  loc_00424B15: lea eax, var_4C
  loc_00424B18: push edx
  loc_00424B19: lea ecx, var_3C
  loc_00424B1C: push eax
  loc_00424B1D: lea edx, var_2C
  loc_00424B20: push ecx
  loc_00424B21: push edx
  loc_00424B22: push 00000004h
  loc_00424B24: call [00401018h] ; __vbaFreeVarList
  loc_00424B2A: add esp, 00000014h
  loc_00424B2D: ret
  loc_00424B2E: ret
  loc_00424B2F: mov eax, Me
  loc_00424B32: push eax
  loc_00424B33: mov ecx, [eax]
  loc_00424B35: call [ecx+00000008h]
  loc_00424B38: mov eax, var_4
  loc_00424B3B: mov ecx, var_14
  loc_00424B3E: pop edi
  loc_00424B3F: pop esi
  loc_00424B40: mov fs:[00000000h], ecx
  loc_00424B47: pop ebx
  loc_00424B48: mov esp, ebp
  loc_00424B4A: pop ebp
  loc_00424B4B: retn 000Ch
End Sub

Private Sub Form_Activate() '423AC0
  loc_00423AC0: push ebp
  loc_00423AC1: mov ebp, esp
  loc_00423AC3: sub esp, 0000000Ch
  loc_00423AC6: push 00401926h ; __vbaExceptHandler
  loc_00423ACB: mov eax, fs:[00000000h]
  loc_00423AD1: push eax
  loc_00423AD2: mov fs:[00000000h], esp
  loc_00423AD9: sub esp, 00000008h
  loc_00423ADC: push ebx
  loc_00423ADD: push esi
  loc_00423ADE: push edi
  loc_00423ADF: mov var_C, esp
  loc_00423AE2: mov var_8, 00401510h
  loc_00423AE9: mov esi, Me
  loc_00423AEC: mov eax, esi
  loc_00423AEE: and eax, 00000001h
  loc_00423AF1: mov var_4, eax
  loc_00423AF4: and esi, FFFFFFFEh
  loc_00423AF7: push esi
  loc_00423AF8: mov Me, esi
  loc_00423AFB: mov ecx, [esi]
  loc_00423AFD: call [ecx+00000004h]
  loc_00423B00: mov edx, [esi]
  loc_00423B02: push esi
  loc_00423B03: call [edx+00000700h]
  loc_00423B09: mov var_4, 00000000h
  loc_00423B10: mov eax, Me
  loc_00423B13: push eax
  loc_00423B14: mov ecx, [eax]
  loc_00423B16: call [ecx+00000008h]
  loc_00423B19: mov eax, var_4
  loc_00423B1C: mov ecx, var_14
  loc_00423B1F: pop edi
  loc_00423B20: pop esi
  loc_00423B21: mov fs:[00000000h], ecx
  loc_00423B28: pop ebx
  loc_00423B29: mov esp, ebp
  loc_00423B2B: pop ebp
  loc_00423B2C: retn 0004h
End Sub

Private Sub mnuChangeUnits_Click() '4247F0
  loc_004247F0: push ebp
  loc_004247F1: mov ebp, esp
  loc_004247F3: sub esp, 0000000Ch
  loc_004247F6: push 00401926h ; __vbaExceptHandler
  loc_004247FB: mov eax, fs:[00000000h]
  loc_00424801: push eax
  loc_00424802: mov fs:[00000000h], esp
  loc_00424809: sub esp, 00000030h
  loc_0042480C: push ebx
  loc_0042480D: push esi
  loc_0042480E: push edi
  loc_0042480F: mov var_C, esp
  loc_00424812: mov var_8, 00401558h
  loc_00424819: mov eax, Me
  loc_0042481C: mov ecx, eax
  loc_0042481E: and ecx, 00000001h
  loc_00424821: mov var_4, ecx
  loc_00424824: and al, FEh
  loc_00424826: push eax
  loc_00424827: mov Me, eax
  loc_0042482A: mov edx, [eax]
  loc_0042482C: call [edx+00000004h]
  loc_0042482F: mov eax, [004301A0h]
  loc_00424834: test eax, eax
  loc_00424836: jnz 00424848h
  loc_00424838: push 004301A0h
  loc_0042483D: push 004033C0h
  loc_00424842: call [004010D4h] ; __vbaNew2
  loc_00424848: sub esp, 00000010h
  loc_0042484B: mov ecx, 0000000Ah
  loc_00424850: mov ebx, esp
  loc_00424852: mov var_24, ecx
  loc_00424855: mov eax, 80020004h
  loc_0042485A: sub esp, 00000010h
  loc_0042485D: mov [ebx], ecx
  loc_0042485F: mov ecx, var_30
  loc_00424862: mov edx, eax
  loc_00424864: mov esi, [004301A0h]
  loc_0042486A: mov [ebx+00000004h], ecx
  loc_0042486D: mov ecx, esp
  loc_0042486F: mov edi, [esi]
  loc_00424871: push esi
  loc_00424872: mov [ebx+00000008h], eax
  loc_00424875: mov eax, var_28
  loc_00424878: mov [ebx+0000000Ch], eax
  loc_0042487B: mov eax, var_24
  loc_0042487E: mov [ecx], eax
  loc_00424880: mov eax, var_20
  loc_00424883: mov [ecx+00000004h], eax
  loc_00424886: mov [ecx+00000008h], edx
  loc_00424889: mov edx, var_18
  loc_0042488C: mov [ecx+0000000Ch], edx
  loc_0042488F: call [edi+000002B0h]
  loc_00424895: test eax, eax
  loc_00424897: fnclex
  loc_00424899: jge 004248ADh
  loc_0042489B: push 000002B0h
  loc_004248A0: push 00410454h
  loc_004248A5: push esi
  loc_004248A6: push eax
  loc_004248A7: call [00401030h] ; __vbaHresultCheckObj
  loc_004248AD: mov var_4, 00000000h
  loc_004248B4: mov eax, Me
  loc_004248B7: push eax
  loc_004248B8: mov ecx, [eax]
  loc_004248BA: call [ecx+00000008h]
  loc_004248BD: mov eax, var_4
  loc_004248C0: mov ecx, var_14
  loc_004248C3: pop edi
  loc_004248C4: pop esi
  loc_004248C5: mov fs:[00000000h], ecx
  loc_004248CC: pop ebx
  loc_004248CD: mov esp, ebp
  loc_004248CF: pop ebp
  loc_004248D0: retn 0004h
End Sub

Private Sub mnuSlope_Click() '424DB0
  loc_00424DB0: push ebp
  loc_00424DB1: mov ebp, esp
  loc_00424DB3: sub esp, 0000000Ch
  loc_00424DB6: push 00401926h ; __vbaExceptHandler
  loc_00424DBB: mov eax, fs:[00000000h]
  loc_00424DC1: push eax
  loc_00424DC2: mov fs:[00000000h], esp
  loc_00424DC9: sub esp, 00000030h
  loc_00424DCC: push ebx
  loc_00424DCD: push esi
  loc_00424DCE: push edi
  loc_00424DCF: mov var_C, esp
  loc_00424DD2: mov var_8, 00401590h
  loc_00424DD9: mov eax, Me
  loc_00424DDC: mov ecx, eax
  loc_00424DDE: and ecx, 00000001h
  loc_00424DE1: mov var_4, ecx
  loc_00424DE4: and al, FEh
  loc_00424DE6: push eax
  loc_00424DE7: mov Me, eax
  loc_00424DEA: mov edx, [eax]
  loc_00424DEC: call [edx+00000004h]
  loc_00424DEF: mov eax, [004300C4h]
  loc_00424DF4: test eax, eax
  loc_00424DF6: jnz 00424E08h
  loc_00424DF8: push 004300C4h
  loc_00424DFD: push 00409228h
  loc_00424E02: call [004010D4h] ; __vbaNew2
  loc_00424E08: sub esp, 00000010h
  loc_00424E0B: mov ecx, 0000000Ah
  loc_00424E10: mov ebx, esp
  loc_00424E12: mov var_24, ecx
  loc_00424E15: mov eax, 80020004h
  loc_00424E1A: sub esp, 00000010h
  loc_00424E1D: mov [ebx], ecx
  loc_00424E1F: mov ecx, var_30
  loc_00424E22: mov edx, eax
  loc_00424E24: mov esi, [004300C4h]
  loc_00424E2A: mov [ebx+00000004h], ecx
  loc_00424E2D: mov ecx, esp
  loc_00424E2F: mov edi, [esi]
  loc_00424E31: push esi
  loc_00424E32: mov [ebx+00000008h], eax
  loc_00424E35: mov eax, var_28
  loc_00424E38: mov [ebx+0000000Ch], eax
  loc_00424E3B: mov eax, var_24
  loc_00424E3E: mov [ecx], eax
  loc_00424E40: mov eax, var_20
  loc_00424E43: mov [ecx+00000004h], eax
  loc_00424E46: mov [ecx+00000008h], edx
  loc_00424E49: mov edx, var_18
  loc_00424E4C: mov [ecx+0000000Ch], edx
  loc_00424E4F: call [edi+000002B0h]
  loc_00424E55: test eax, eax
  loc_00424E57: fnclex
  loc_00424E59: jge 00424E6Dh
  loc_00424E5B: push 000002B0h
  loc_00424E60: push 0040E0ECh
  loc_00424E65: push esi
  loc_00424E66: push eax
  loc_00424E67: call [00401030h] ; __vbaHresultCheckObj
  loc_00424E6D: mov eax, [004300B0h]
  loc_00424E72: test eax, eax
  loc_00424E74: jnz 00424E86h
  loc_00424E76: push 004300B0h
  loc_00424E7B: push 00407528h
  loc_00424E80: call [004010D4h] ; __vbaNew2
  loc_00424E86: mov esi, [004300B0h]
  loc_00424E8C: push esi
  loc_00424E8D: mov eax, [esi]
  loc_00424E8F: call [eax+000002B4h]
  loc_00424E95: test eax, eax
  loc_00424E97: fnclex
  loc_00424E99: jge 00424EADh
  loc_00424E9B: push 000002B4h
  loc_00424EA0: push 0040DE98h
  loc_00424EA5: push esi
  loc_00424EA6: push eax
  loc_00424EA7: call [00401030h] ; __vbaHresultCheckObj
  loc_00424EAD: mov var_4, 00000000h
  loc_00424EB4: mov eax, Me
  loc_00424EB7: push eax
  loc_00424EB8: mov ecx, [eax]
  loc_00424EBA: call [ecx+00000008h]
  loc_00424EBD: mov eax, var_4
  loc_00424EC0: mov ecx, var_14
  loc_00424EC3: pop edi
  loc_00424EC4: pop esi
  loc_00424EC5: mov fs:[00000000h], ecx
  loc_00424ECC: pop ebx
  loc_00424ECD: mov esp, ebp
  loc_00424ECF: pop ebp
  loc_00424ED0: retn 0004h
End Sub

Private Sub mnuChangeNames_Click() '424610
  loc_00424610: push ebp
  loc_00424611: mov ebp, esp
  loc_00424613: sub esp, 0000000Ch
  loc_00424616: push 00401926h ; __vbaExceptHandler
  loc_0042461B: mov eax, fs:[00000000h]
  loc_00424621: push eax
  loc_00424622: mov fs:[00000000h], esp
  loc_00424629: sub esp, 00000030h
  loc_0042462C: push ebx
  loc_0042462D: push esi
  loc_0042462E: push edi
  loc_0042462F: mov var_C, esp
  loc_00424632: mov var_8, 00401548h
  loc_00424639: mov eax, Me
  loc_0042463C: mov ecx, eax
  loc_0042463E: and ecx, 00000001h
  loc_00424641: mov var_4, ecx
  loc_00424644: and al, FEh
  loc_00424646: push eax
  loc_00424647: mov Me, eax
  loc_0042464A: mov edx, [eax]
  loc_0042464C: call [edx+00000004h]
  loc_0042464F: mov eax, [004300D8h]
  loc_00424654: test eax, eax
  loc_00424656: jnz 00424668h
  loc_00424658: push 004300D8h
  loc_0042465D: push 00402E04h
  loc_00424662: call [004010D4h] ; __vbaNew2
  loc_00424668: sub esp, 00000010h
  loc_0042466B: mov ecx, 0000000Ah
  loc_00424670: mov ebx, esp
  loc_00424672: mov var_24, ecx
  loc_00424675: mov eax, 80020004h
  loc_0042467A: sub esp, 00000010h
  loc_0042467D: mov [ebx], ecx
  loc_0042467F: mov ecx, var_30
  loc_00424682: mov edx, eax
  loc_00424684: mov esi, [004300D8h]
  loc_0042468A: mov [ebx+00000004h], ecx
  loc_0042468D: mov ecx, esp
  loc_0042468F: mov edi, [esi]
  loc_00424691: push esi
  loc_00424692: mov [ebx+00000008h], eax
  loc_00424695: mov eax, var_28
  loc_00424698: mov [ebx+0000000Ch], eax
  loc_0042469B: mov eax, var_24
  loc_0042469E: mov [ecx], eax
  loc_004246A0: mov eax, var_20
  loc_004246A3: mov [ecx+00000004h], eax
  loc_004246A6: mov [ecx+00000008h], edx
  loc_004246A9: mov edx, var_18
  loc_004246AC: mov [ecx+0000000Ch], edx
  loc_004246AF: call [edi+000002B0h]
  loc_004246B5: test eax, eax
  loc_004246B7: fnclex
  loc_004246B9: jge 004246CDh
  loc_004246BB: push 000002B0h
  loc_004246C0: push 0040E260h
  loc_004246C5: push esi
  loc_004246C6: push eax
  loc_004246C7: call [00401030h] ; __vbaHresultCheckObj
  loc_004246CD: mov var_4, 00000000h
  loc_004246D4: mov eax, Me
  loc_004246D7: push eax
  loc_004246D8: mov ecx, [eax]
  loc_004246DA: call [ecx+00000008h]
  loc_004246DD: mov eax, var_4
  loc_004246E0: mov ecx, var_14
  loc_004246E3: pop edi
  loc_004246E4: pop esi
  loc_004246E5: mov fs:[00000000h], ecx
  loc_004246EC: pop ebx
  loc_004246ED: mov esp, ebp
  loc_004246EF: pop ebp
  loc_004246F0: retn 0004h
End Sub

Private Sub mnuAssuptions_Click() '4243F0
  loc_004243F0: push ebp
  loc_004243F1: mov ebp, esp
  loc_004243F3: sub esp, 0000000Ch
  loc_004243F6: push 00401926h ; __vbaExceptHandler
  loc_004243FB: mov eax, fs:[00000000h]
  loc_00424401: push eax
  loc_00424402: mov fs:[00000000h], esp
  loc_00424409: sub esp, 00000030h
  loc_0042440C: push ebx
  loc_0042440D: push esi
  loc_0042440E: push edi
  loc_0042440F: mov var_C, esp
  loc_00424412: mov var_8, 00401538h
  loc_00424419: mov eax, Me
  loc_0042441C: mov ecx, eax
  loc_0042441E: and ecx, 00000001h
  loc_00424421: mov var_4, ecx
  loc_00424424: and al, FEh
  loc_00424426: push eax
  loc_00424427: mov Me, eax
  loc_0042442A: mov edx, [eax]
  loc_0042442C: call [edx+00000004h]
  loc_0042442F: mov eax, [0043009Ch]
  loc_00424434: test eax, eax
  loc_00424436: jnz 00424448h
  loc_00424438: push 0043009Ch
  loc_0042443D: push 00405FC0h
  loc_00424442: call [004010D4h] ; __vbaNew2
  loc_00424448: sub esp, 00000010h
  loc_0042444B: mov ecx, 0000000Ah
  loc_00424450: mov ebx, esp
  loc_00424452: mov var_24, ecx
  loc_00424455: mov eax, 80020004h
  loc_0042445A: sub esp, 00000010h
  loc_0042445D: mov [ebx], ecx
  loc_0042445F: mov ecx, var_30
  loc_00424462: mov edx, eax
  loc_00424464: mov esi, [0043009Ch]
  loc_0042446A: mov [ebx+00000004h], ecx
  loc_0042446D: mov ecx, esp
  loc_0042446F: mov edi, [esi]
  loc_00424471: push esi
  loc_00424472: mov [ebx+00000008h], eax
  loc_00424475: mov eax, var_28
  loc_00424478: mov [ebx+0000000Ch], eax
  loc_0042447B: mov eax, var_24
  loc_0042447E: mov [ecx], eax
  loc_00424480: mov eax, var_20
  loc_00424483: mov [ecx+00000004h], eax
  loc_00424486: mov [ecx+00000008h], edx
  loc_00424489: mov edx, var_18
  loc_0042448C: mov [ecx+0000000Ch], edx
  loc_0042448F: call [edi+000002B0h]
  loc_00424495: test eax, eax
  loc_00424497: fnclex
  loc_00424499: jge 004244ADh
  loc_0042449B: push 000002B0h
  loc_004244A0: push 0040DDE0h
  loc_004244A5: push esi
  loc_004244A6: push eax
  loc_004244A7: call [00401030h] ; __vbaHresultCheckObj
  loc_004244AD: mov eax, [004300B0h]
  loc_004244B2: test eax, eax
  loc_004244B4: jnz 004244C6h
  loc_004244B6: push 004300B0h
  loc_004244BB: push 00407528h
  loc_004244C0: call [004010D4h] ; __vbaNew2
  loc_004244C6: mov esi, [004300B0h]
  loc_004244CC: push esi
  loc_004244CD: mov eax, [esi]
  loc_004244CF: call [eax+000002B4h]
  loc_004244D5: test eax, eax
  loc_004244D7: fnclex
  loc_004244D9: jge 004244EDh
  loc_004244DB: push 000002B4h
  loc_004244E0: push 0040DE98h
  loc_004244E5: push esi
  loc_004244E6: push eax
  loc_004244E7: call [00401030h] ; __vbaHresultCheckObj
  loc_004244ED: mov var_4, 00000000h
  loc_004244F4: mov eax, Me
  loc_004244F7: push eax
  loc_004244F8: mov ecx, [eax]
  loc_004244FA: call [ecx+00000008h]
  loc_004244FD: mov eax, var_4
  loc_00424500: mov ecx, var_14
  loc_00424503: pop edi
  loc_00424504: pop esi
  loc_00424505: mov fs:[00000000h], ecx
  loc_0042450C: pop ebx
  loc_0042450D: mov esp, ebp
  loc_0042450F: pop ebp
  loc_00424510: retn 0004h
End Sub

Private Sub mnuIntroduction_Click() '424B50
  loc_00424B50: push ebp
  loc_00424B51: mov ebp, esp
  loc_00424B53: sub esp, 0000000Ch
  loc_00424B56: push 00401926h ; __vbaExceptHandler
  loc_00424B5B: mov eax, fs:[00000000h]
  loc_00424B61: push eax
  loc_00424B62: mov fs:[00000000h], esp
  loc_00424B69: sub esp, 00000030h
  loc_00424B6C: push ebx
  loc_00424B6D: push esi
  loc_00424B6E: push edi
  loc_00424B6F: mov var_C, esp
  loc_00424B72: mov var_8, 00401580h
  loc_00424B79: mov eax, Me
  loc_00424B7C: mov ecx, eax
  loc_00424B7E: and ecx, 00000001h
  loc_00424B81: mov var_4, ecx
  loc_00424B84: and al, FEh
  loc_00424B86: push eax
  loc_00424B87: mov Me, eax
  loc_00424B8A: mov edx, [eax]
  loc_00424B8C: call [edx+00000004h]
  loc_00424B8F: mov eax, [00430088h]
  loc_00424B94: test eax, eax
  loc_00424B96: jnz 00424BA8h
  loc_00424B98: push 00430088h
  loc_00424B9D: push 004058C0h
  loc_00424BA2: call [004010D4h] ; __vbaNew2
  loc_00424BA8: sub esp, 00000010h
  loc_00424BAB: mov ecx, 0000000Ah
  loc_00424BB0: mov ebx, esp
  loc_00424BB2: mov var_24, ecx
  loc_00424BB5: mov eax, 80020004h
  loc_00424BBA: sub esp, 00000010h
  loc_00424BBD: mov [ebx], ecx
  loc_00424BBF: mov ecx, var_30
  loc_00424BC2: mov edx, eax
  loc_00424BC4: mov esi, [00430088h]
  loc_00424BCA: mov [ebx+00000004h], ecx
  loc_00424BCD: mov ecx, esp
  loc_00424BCF: mov edi, [esi]
  loc_00424BD1: push esi
  loc_00424BD2: mov [ebx+00000008h], eax
  loc_00424BD5: mov eax, var_28
  loc_00424BD8: mov [ebx+0000000Ch], eax
  loc_00424BDB: mov eax, var_24
  loc_00424BDE: mov [ecx], eax
  loc_00424BE0: mov eax, var_20
  loc_00424BE3: mov [ecx+00000004h], eax
  loc_00424BE6: mov [ecx+00000008h], edx
  loc_00424BE9: mov edx, var_18
  loc_00424BEC: mov [ecx+0000000Ch], edx
  loc_00424BEF: call [edi+000002B0h]
  loc_00424BF5: test eax, eax
  loc_00424BF7: fnclex
  loc_00424BF9: jge 00424C0Dh
  loc_00424BFB: push 000002B0h
  loc_00424C00: push 0040DB64h
  loc_00424C05: push esi
  loc_00424C06: push eax
  loc_00424C07: call [00401030h] ; __vbaHresultCheckObj
  loc_00424C0D: mov eax, [004300B0h]
  loc_00424C12: test eax, eax
  loc_00424C14: jnz 00424C26h
  loc_00424C16: push 004300B0h
  loc_00424C1B: push 00407528h
  loc_00424C20: call [004010D4h] ; __vbaNew2
  loc_00424C26: mov esi, [004300B0h]
  loc_00424C2C: push esi
  loc_00424C2D: mov eax, [esi]
  loc_00424C2F: call [eax+000002B4h]
  loc_00424C35: test eax, eax
  loc_00424C37: fnclex
  loc_00424C39: jge 00424C4Dh
  loc_00424C3B: push 000002B4h
  loc_00424C40: push 0040DE98h
  loc_00424C45: push esi
  loc_00424C46: push eax
  loc_00424C47: call [00401030h] ; __vbaHresultCheckObj
  loc_00424C4D: mov var_4, 00000000h
  loc_00424C54: mov eax, Me
  loc_00424C57: push eax
  loc_00424C58: mov ecx, [eax]
  loc_00424C5A: call [ecx+00000008h]
  loc_00424C5D: mov eax, var_4
  loc_00424C60: mov ecx, var_14
  loc_00424C63: pop edi
  loc_00424C64: pop esi
  loc_00424C65: mov fs:[00000000h], ecx
  loc_00424C6C: pop ebx
  loc_00424C6D: mov esp, ebp
  loc_00424C6F: pop ebp
  loc_00424C70: retn 0004h
End Sub

Private Sub mnuChangePredEst_Click() '424700
  loc_00424700: push ebp
  loc_00424701: mov ebp, esp
  loc_00424703: sub esp, 0000000Ch
  loc_00424706: push 00401926h ; __vbaExceptHandler
  loc_0042470B: mov eax, fs:[00000000h]
  loc_00424711: push eax
  loc_00424712: mov fs:[00000000h], esp
  loc_00424719: sub esp, 00000030h
  loc_0042471C: push ebx
  loc_0042471D: push esi
  loc_0042471E: push edi
  loc_0042471F: mov var_C, esp
  loc_00424722: mov var_8, 00401550h
  loc_00424729: mov eax, Me
  loc_0042472C: mov ecx, eax
  loc_0042472E: and ecx, 00000001h
  loc_00424731: mov var_4, ecx
  loc_00424734: and al, FEh
  loc_00424736: push eax
  loc_00424737: mov Me, eax
  loc_0042473A: mov edx, [eax]
  loc_0042473C: call [edx+00000004h]
  loc_0042473F: mov eax, [00430128h]
  loc_00424744: test eax, eax
  loc_00424746: jnz 00424758h
  loc_00424748: push 00430128h
  loc_0042474D: push 00404AE0h
  loc_00424752: call [004010D4h] ; __vbaNew2
  loc_00424758: sub esp, 00000010h
  loc_0042475B: mov ecx, 0000000Ah
  loc_00424760: mov ebx, esp
  loc_00424762: mov var_24, ecx
  loc_00424765: mov eax, 80020004h
  loc_0042476A: sub esp, 00000010h
  loc_0042476D: mov [ebx], ecx
  loc_0042476F: mov ecx, var_30
  loc_00424772: mov edx, eax
  loc_00424774: mov esi, [00430128h]
  loc_0042477A: mov [ebx+00000004h], ecx
  loc_0042477D: mov ecx, esp
  loc_0042477F: mov edi, [esi]
  loc_00424781: push esi
  loc_00424782: mov [ebx+00000008h], eax
  loc_00424785: mov eax, var_28
  loc_00424788: mov [ebx+0000000Ch], eax
  loc_0042478B: mov eax, var_24
  loc_0042478E: mov [ecx], eax
  loc_00424790: mov eax, var_20
  loc_00424793: mov [ecx+00000004h], eax
  loc_00424796: mov [ecx+00000008h], edx
  loc_00424799: mov edx, var_18
  loc_0042479C: mov [ecx+0000000Ch], edx
  loc_0042479F: call [edi+000002B0h]
  loc_004247A5: test eax, eax
  loc_004247A7: fnclex
  loc_004247A9: jge 004247BDh
  loc_004247AB: push 000002B0h
  loc_004247B0: push 0040F8B4h
  loc_004247B5: push esi
  loc_004247B6: push eax
  loc_004247B7: call [00401030h] ; __vbaHresultCheckObj
  loc_004247BD: mov var_4, 00000000h
  loc_004247C4: mov eax, Me
  loc_004247C7: push eax
  loc_004247C8: mov ecx, [eax]
  loc_004247CA: call [ecx+00000008h]
  loc_004247CD: mov eax, var_4
  loc_004247D0: mov ecx, var_14
  loc_004247D3: pop edi
  loc_004247D4: pop esi
  loc_004247D5: mov fs:[00000000h], ecx
  loc_004247DC: pop ebx
  loc_004247DD: mov esp, ebp
  loc_004247DF: pop ebp
  loc_004247E0: retn 0004h
End Sub

Private Sub mnuStatistics_Click() '424EE0
  loc_00424EE0: push ebp
  loc_00424EE1: mov ebp, esp
  loc_00424EE3: sub esp, 0000000Ch
  loc_00424EE6: push 00401926h ; __vbaExceptHandler
  loc_00424EEB: mov eax, fs:[00000000h]
  loc_00424EF1: push eax
  loc_00424EF2: mov fs:[00000000h], esp
  loc_00424EF9: sub esp, 00000030h
  loc_00424EFC: push ebx
  loc_00424EFD: push esi
  loc_00424EFE: push edi
  loc_00424EFF: mov var_C, esp
  loc_00424F02: mov var_8, 00401598h
  loc_00424F09: mov eax, Me
  loc_00424F0C: mov ecx, eax
  loc_00424F0E: and ecx, 00000001h
  loc_00424F11: mov var_4, ecx
  loc_00424F14: and al, FEh
  loc_00424F16: push eax
  loc_00424F17: mov Me, eax
  loc_00424F1A: mov edx, [eax]
  loc_00424F1C: call [edx+00000004h]
  loc_00424F1F: mov eax, [004300ECh]
  loc_00424F24: test eax, eax
  loc_00424F26: jnz 00424F38h
  loc_00424F28: push 004300ECh
  loc_00424F2D: push 0040A624h
  loc_00424F32: call [004010D4h] ; __vbaNew2
  loc_00424F38: sub esp, 00000010h
  loc_00424F3B: mov ecx, 0000000Ah
  loc_00424F40: mov ebx, esp
  loc_00424F42: mov var_24, ecx
  loc_00424F45: mov eax, 80020004h
  loc_00424F4A: sub esp, 00000010h
  loc_00424F4D: mov [ebx], ecx
  loc_00424F4F: mov ecx, var_30
  loc_00424F52: mov edx, eax
  loc_00424F54: mov esi, [004300ECh]
  loc_00424F5A: mov [ebx+00000004h], ecx
  loc_00424F5D: mov ecx, esp
  loc_00424F5F: mov edi, [esi]
  loc_00424F61: push esi
  loc_00424F62: mov [ebx+00000008h], eax
  loc_00424F65: mov eax, var_28
  loc_00424F68: mov [ebx+0000000Ch], eax
  loc_00424F6B: mov eax, var_24
  loc_00424F6E: mov [ecx], eax
  loc_00424F70: mov eax, var_20
  loc_00424F73: mov [ecx+00000004h], eax
  loc_00424F76: mov [ecx+00000008h], edx
  loc_00424F79: mov edx, var_18
  loc_00424F7C: mov [ecx+0000000Ch], edx
  loc_00424F7F: call [edi+000002B0h]
  loc_00424F85: test eax, eax
  loc_00424F87: fnclex
  loc_00424F89: jge 00424F9Dh
  loc_00424F8B: push 000002B0h
  loc_00424F90: push 0040ECECh
  loc_00424F95: push esi
  loc_00424F96: push eax
  loc_00424F97: call [00401030h] ; __vbaHresultCheckObj
  loc_00424F9D: mov eax, [004300B0h]
  loc_00424FA2: test eax, eax
  loc_00424FA4: jnz 00424FB6h
  loc_00424FA6: push 004300B0h
  loc_00424FAB: push 00407528h
  loc_00424FB0: call [004010D4h] ; __vbaNew2
  loc_00424FB6: mov esi, [004300B0h]
  loc_00424FBC: push esi
  loc_00424FBD: mov eax, [esi]
  loc_00424FBF: call [eax+000002B4h]
  loc_00424FC5: test eax, eax
  loc_00424FC7: fnclex
  loc_00424FC9: jge 00424FDDh
  loc_00424FCB: push 000002B4h
  loc_00424FD0: push 0040DE98h
  loc_00424FD5: push esi
  loc_00424FD6: push eax
  loc_00424FD7: call [00401030h] ; __vbaHresultCheckObj
  loc_00424FDD: mov var_4, 00000000h
  loc_00424FE4: mov eax, Me
  loc_00424FE7: push eax
  loc_00424FE8: mov ecx, [eax]
  loc_00424FEA: call [ecx+00000008h]
  loc_00424FED: mov eax, var_4
  loc_00424FF0: mov ecx, var_14
  loc_00424FF3: pop edi
  loc_00424FF4: pop esi
  loc_00424FF5: mov fs:[00000000h], ecx
  loc_00424FFC: pop ebx
  loc_00424FFD: mov esp, ebp
  loc_00424FFF: pop ebp
  loc_00425000: retn 0004h
End Sub

Private Sub Proc_12_13_423BA0
  loc_00423BA0: push ebp
  loc_00423BA1: mov ebp, esp
  loc_00423BA3: sub esp, 00000008h
  loc_00423BA6: push 00401926h ; __vbaExceptHandler
  loc_00423BAB: mov eax, fs:[00000000h]
  loc_00423BB1: push eax
  loc_00423BB2: mov fs:[00000000h], esp
  loc_00423BB9: sub esp, 00000060h
  loc_00423BBC: push ebx
  loc_00423BBD: push esi
  loc_00423BBE: push edi
  loc_00423BBF: mov var_8, esp
  loc_00423BC2: mov var_4, 00401528h
  loc_00423BC9: fld real4 ptr [0043003Ch]
  loc_00423BCF: fmul st0, real4 ptr [00430028h]
  loc_00423BD5: xor eax, eax
  loc_00423BD7: mov ebx, [0040107Ch] ; __vbaStrR4
  loc_00423BDD: mov var_14, eax
  loc_00423BE0: mov var_24, eax
  loc_00423BE3: fadd st0, real4 ptr [00430038h]
  loc_00423BE9: mov var_2C, eax
  loc_00423BEC: mov var_30, eax
  loc_00423BEF: mov var_38, eax
  loc_00423BF2: mov var_3C, eax
  loc_00423BF5: mov var_40, eax
  loc_00423BF8: mov var_44, eax
  loc_00423BFB: mov var_48, eax
  loc_00423BFE: mov var_4C, eax
  loc_00423C01: mov var_50, eax
  loc_00423C04: mov var_54, eax
  loc_00423C07: mov var_58, eax
  loc_00423C0A: mov var_5C, eax
  loc_00423C0D: mov var_60, eax
  loc_00423C10: mov var_64, eax
  loc_00423C13: fstp real4 ptr var_28
  loc_00423C16: fnstsw ax
  loc_00423C18: test al, 0Dh
  loc_00423C1A: jnz 004243E6h
  loc_00423C20: fld real4 ptr [0043003Ch]
  loc_00423C26: fmul st0, real4 ptr [00430028h]
  loc_00423C2C: push 00410694h ; "( "
  loc_00423C31: fsubr st0, real4 ptr [00430038h]
  loc_00423C37: fstp real4 ptr var_1C
  loc_00423C3A: fnstsw ax
  loc_00423C3C: test al, 0Dh
  loc_00423C3E: jnz 004243E6h
  loc_00423C44: mov eax, [00430038h]
  loc_00423C49: push eax
  loc_00423C4A: call ebx
  loc_00423C4C: mov esi, [004010FCh] ; __vbaStrMove
  loc_00423C52: mov edx, eax
  loc_00423C54: lea ecx, var_38
  loc_00423C57: call __vbaStrMove
  loc_00423C59: mov edi, [0040102Ch] ; __vbaStrCat
  loc_00423C5F: push eax
  loc_00423C60: call edi
  loc_00423C62: mov edx, eax
  loc_00423C64: lea ecx, var_3C
  loc_00423C67: call __vbaStrMove
  loc_00423C69: push eax
  loc_00423C6A: push 004106A0h ; " )"
  loc_00423C6F: call edi
  loc_00423C71: mov edx, eax
  loc_00423C73: lea ecx, var_40
  loc_00423C76: call __vbaStrMove
  loc_00423C78: push eax
  loc_00423C79: push 004106ACh ; " ± "
  loc_00423C7E: call edi
  loc_00423C80: mov edx, eax
  loc_00423C82: lea ecx, var_44
  loc_00423C85: call __vbaStrMove
  loc_00423C87: push eax
  loc_00423C88: push 00410694h ; "( "
  loc_00423C8D: call edi
  loc_00423C8F: mov edx, eax
  loc_00423C91: lea ecx, var_48
  loc_00423C94: call __vbaStrMove
  loc_00423C96: mov ecx, [00430028h]
  loc_00423C9C: push eax
  loc_00423C9D: push ecx
  loc_00423C9E: call ebx
  loc_00423CA0: mov edx, eax
  loc_00423CA2: lea ecx, var_4C
  loc_00423CA5: call __vbaStrMove
  loc_00423CA7: push eax
  loc_00423CA8: call edi
  loc_00423CAA: mov edx, eax
  loc_00423CAC: lea ecx, var_50
  loc_00423CAF: call __vbaStrMove
  loc_00423CB1: push eax
  loc_00423CB2: push 004106B8h ; " ) * ( "
  loc_00423CB7: call edi
  loc_00423CB9: mov edx, eax
  loc_00423CBB: lea ecx, var_54
  loc_00423CBE: call __vbaStrMove
  loc_00423CC0: mov edx, [0043003Ch]
  loc_00423CC6: push eax
  loc_00423CC7: push edx
  loc_00423CC8: call ebx
  loc_00423CCA: mov edx, eax
  loc_00423CCC: lea ecx, var_58
  loc_00423CCF: call __vbaStrMove
  loc_00423CD1: push eax
  loc_00423CD2: call edi
  loc_00423CD4: mov edx, eax
  loc_00423CD6: lea ecx, var_5C
  loc_00423CD9: call __vbaStrMove
  loc_00423CDB: push eax
  loc_00423CDC: push 004106A0h ; " )"
  loc_00423CE1: call edi
  loc_00423CE3: mov edx, eax
  loc_00423CE5: lea ecx, var_30
  loc_00423CE8: call __vbaStrMove
  loc_00423CEA: lea eax, var_5C
  loc_00423CED: lea ecx, var_58
  loc_00423CF0: push eax
  loc_00423CF1: lea edx, var_54
  loc_00423CF4: push ecx
  loc_00423CF5: lea eax, var_50
  loc_00423CF8: push edx
  loc_00423CF9: lea ecx, var_4C
  loc_00423CFC: push eax
  loc_00423CFD: lea edx, var_48
  loc_00423D00: push ecx
  loc_00423D01: lea eax, var_44
  loc_00423D04: push edx
  loc_00423D05: lea ecx, var_40
  loc_00423D08: push eax
  loc_00423D09: lea edx, var_3C
  loc_00423D0C: push ecx
  loc_00423D0D: lea eax, var_38
  loc_00423D10: push edx
  loc_00423D11: push eax
  loc_00423D12: push 0000000Ah
  loc_00423D14: call [004010E4h] ; __vbaFreeStrList
  loc_00423D1A: mov ecx, var_30
  loc_00423D1D: add esp, 0000002Ch
  loc_00423D20: push ecx
  loc_00423D21: push 004106CCh ; " = "
  loc_00423D26: call edi
  loc_00423D28: mov edx, eax
  loc_00423D2A: lea ecx, var_38
  loc_00423D2D: call __vbaStrMove
  loc_00423D2F: mov edx, var_1C
  loc_00423D32: push eax
  loc_00423D33: push edx
  loc_00423D34: call ebx
  loc_00423D36: mov edx, eax
  loc_00423D38: lea ecx, var_3C
  loc_00423D3B: call __vbaStrMove
  loc_00423D3D: push eax
  loc_00423D3E: call edi
  loc_00423D40: mov edx, eax
  loc_00423D42: lea ecx, var_40
  loc_00423D45: call __vbaStrMove
  loc_00423D47: push eax
  loc_00423D48: push 004106D8h ; " to "
  loc_00423D4D: call edi
  loc_00423D4F: mov edx, eax
  loc_00423D51: lea ecx, var_44
  loc_00423D54: call __vbaStrMove
  loc_00423D56: push eax
  loc_00423D57: mov eax, var_28
  loc_00423D5A: push eax
  loc_00423D5B: call ebx
  loc_00423D5D: mov edx, eax
  loc_00423D5F: lea ecx, var_48
  loc_00423D62: call __vbaStrMove
  loc_00423D64: push eax
  loc_00423D65: call edi
  loc_00423D67: mov edx, eax
  loc_00423D69: lea ecx, var_30
  loc_00423D6C: call __vbaStrMove
  loc_00423D6E: lea ecx, var_48
  loc_00423D71: lea edx, var_44
  loc_00423D74: push ecx
  loc_00423D75: lea eax, var_40
  loc_00423D78: push edx
  loc_00423D79: lea ecx, var_3C
  loc_00423D7C: push eax
  loc_00423D7D: lea edx, var_38
  loc_00423D80: push ecx
  loc_00423D81: push edx
  loc_00423D82: push 00000005h
  loc_00423D84: call [004010E4h] ; __vbaFreeStrList
  loc_00423D8A: mov eax, Me
  loc_00423D8D: add esp, 00000018h
  loc_00423D90: mov ecx, [eax]
  loc_00423D92: push eax
  loc_00423D93: call [ecx+00000320h]
  loc_00423D99: lea edx, var_64
  loc_00423D9C: push eax
  loc_00423D9D: push edx
  loc_00423D9E: call [0040103Ch] ; __vbaObjSet
  loc_00423DA4: mov edx, var_30
  loc_00423DA7: mov ecx, [eax]
  loc_00423DA9: push edx
  loc_00423DAA: push eax
  loc_00423DAB: mov var_68, eax
  loc_00423DAE: call [ecx+00000054h]
  loc_00423DB1: test eax, eax
  loc_00423DB3: fnclex
  loc_00423DB5: jge 00423DC9h
  loc_00423DB7: mov ecx, var_68
  loc_00423DBA: push 00000054h
  loc_00423DBC: push 0040E390h
  loc_00423DC1: push ecx
  loc_00423DC2: push eax
  loc_00423DC3: call [00401030h] ; __vbaHresultCheckObj
  loc_00423DC9: lea ecx, var_64
  loc_00423DCC: call [00401114h] ; __vbaFreeObj
  loc_00423DD2: fld real4 ptr [00430020h]
  loc_00423DD8: fsubr st0, real8 ptr [004011F8h]
  loc_00423DDE: push 004106E8h ; "With "
  loc_00423DE3: fmul st0, real8 ptr [00401520h]
  loc_00423DE9: fstp real4 ptr var_20
  loc_00423DEC: fnstsw ax
  loc_00423DEE: test al, 0Dh
  loc_00423DF0: jnz 004243E6h
  loc_00423DF6: mov edx, var_20
  loc_00423DF9: push edx
  loc_00423DFA: call ebx
  loc_00423DFC: mov edx, eax
  loc_00423DFE: lea ecx, var_38
  loc_00423E01: call __vbaStrMove
  loc_00423E03: push eax
  loc_00423E04: call edi
  loc_00423E06: mov edx, eax
  loc_00423E08: lea ecx, var_3C
  loc_00423E0B: call __vbaStrMove
  loc_00423E0D: push eax
  loc_00423E0E: push 004106F8h ; Chr(37) & " confidence we can say that when "
  loc_00423E13: call edi
  loc_00423E15: mov edx, eax
  loc_00423E17: lea ecx, var_40
  loc_00423E1A: call __vbaStrMove
  loc_00423E1C: push eax
  loc_00423E1D: mov eax, [00430018h]
  loc_00423E22: push eax
  loc_00423E23: call edi
  loc_00423E25: mov edx, eax
  loc_00423E27: lea ecx, var_44
  loc_00423E2A: call __vbaStrMove
  loc_00423E2C: push eax
  loc_00423E2D: push 00410744h ; " equals to "
  loc_00423E32: call edi
  loc_00423E34: mov edx, eax
  loc_00423E36: lea ecx, var_48
  loc_00423E39: call __vbaStrMove
  loc_00423E3B: mov ecx, [0043002Ch]
  loc_00423E41: push eax
  loc_00423E42: push ecx
  loc_00423E43: call edi
  loc_00423E45: mov edx, eax
  loc_00423E47: lea ecx, var_4C
  loc_00423E4A: call __vbaStrMove
  loc_00423E4C: push eax
  loc_00423E4D: push 0040F028h
  loc_00423E52: call edi
  loc_00423E54: mov edx, eax
  loc_00423E56: lea ecx, var_50
  loc_00423E59: call __vbaStrMove
  loc_00423E5B: mov edx, [0043001Ch]
  loc_00423E61: push eax
  loc_00423E62: push edx
  loc_00423E63: call edi
  loc_00423E65: mov edx, eax
  loc_00423E67: lea ecx, var_2C
  loc_00423E6A: call __vbaStrMove
  loc_00423E6C: lea eax, var_50
  loc_00423E6F: lea ecx, var_4C
  loc_00423E72: push eax
  loc_00423E73: lea edx, var_48
  loc_00423E76: push ecx
  loc_00423E77: lea eax, var_44
  loc_00423E7A: push edx
  loc_00423E7B: lea ecx, var_40
  loc_00423E7E: push eax
  loc_00423E7F: lea edx, var_3C
  loc_00423E82: push ecx
  loc_00423E83: lea eax, var_38
  loc_00423E86: push edx
  loc_00423E87: push eax
  loc_00423E88: push 00000007h
  loc_00423E8A: call [004010E4h] ; __vbaFreeStrList
  loc_00423E90: mov ecx, var_2C
  loc_00423E93: add esp, 00000020h
  loc_00423E96: push ecx
  loc_00423E97: push 00410760h ; ", the mean value of "
  loc_00423E9C: call edi
  loc_00423E9E: mov edx, eax
  loc_00423EA0: lea ecx, var_38
  loc_00423EA3: call __vbaStrMove
  loc_00423EA5: mov edx, [00430010h]
  loc_00423EAB: push eax
  loc_00423EAC: push edx
  loc_00423EAD: call edi
  loc_00423EAF: mov edx, eax
  loc_00423EB1: lea ecx, var_3C
  loc_00423EB4: call __vbaStrMove
  loc_00423EB6: push eax
  loc_00423EB7: push 00410790h ; " falls in the range "
  loc_00423EBC: call edi
  loc_00423EBE: mov edx, eax
  loc_00423EC0: lea ecx, var_40
  loc_00423EC3: call __vbaStrMove
  loc_00423EC5: push eax
  loc_00423EC6: mov eax, var_1C
  loc_00423EC9: push eax
  loc_00423ECA: call ebx
  loc_00423ECC: mov edx, eax
  loc_00423ECE: lea ecx, var_44
  loc_00423ED1: call __vbaStrMove
  loc_00423ED3: push eax
  loc_00423ED4: call edi
  loc_00423ED6: mov edx, eax
  loc_00423ED8: lea ecx, var_48
  loc_00423EDB: call __vbaStrMove
  loc_00423EDD: push eax
  loc_00423EDE: push 0040F028h
  loc_00423EE3: call edi
  loc_00423EE5: mov edx, eax
  loc_00423EE7: lea ecx, var_4C
  loc_00423EEA: call __vbaStrMove
  loc_00423EEC: mov ecx, [00430014h]
  loc_00423EF2: push eax
  loc_00423EF3: push ecx
  loc_00423EF4: call edi
  loc_00423EF6: mov edx, eax
  loc_00423EF8: lea ecx, var_50
  loc_00423EFB: call __vbaStrMove
  loc_00423EFD: push eax
  loc_00423EFE: push 004106D8h ; " to "
  loc_00423F03: call edi
  loc_00423F05: mov edx, eax
  loc_00423F07: lea ecx, var_54
  loc_00423F0A: call __vbaStrMove
  loc_00423F0C: mov edx, var_28
  loc_00423F0F: push eax
  loc_00423F10: push edx
  loc_00423F11: call ebx
  loc_00423F13: mov edx, eax
  loc_00423F15: lea ecx, var_58
  loc_00423F18: call __vbaStrMove
  loc_00423F1A: push eax
  loc_00423F1B: call edi
  loc_00423F1D: mov edx, eax
  loc_00423F1F: lea ecx, var_5C
  loc_00423F22: call __vbaStrMove
  loc_00423F24: push eax
  loc_00423F25: push 0040F028h
  loc_00423F2A: call edi
  loc_00423F2C: mov edx, eax
  loc_00423F2E: lea ecx, var_60
  loc_00423F31: call __vbaStrMove
  loc_00423F33: push eax
  loc_00423F34: mov eax, [00430014h]
  loc_00423F39: push eax
  loc_00423F3A: call edi
  loc_00423F3C: mov edx, eax
  loc_00423F3E: lea ecx, var_2C
  loc_00423F41: call __vbaStrMove
  loc_00423F43: lea ecx, var_60
  loc_00423F46: lea edx, var_5C
  loc_00423F49: push ecx
  loc_00423F4A: lea eax, var_58
  loc_00423F4D: push edx
  loc_00423F4E: lea ecx, var_54
  loc_00423F51: push eax
  loc_00423F52: push ecx
  loc_00423F53: lea edx, var_50
  loc_00423F56: lea eax, var_4C
  loc_00423F59: push edx
  loc_00423F5A: lea ecx, var_48
  loc_00423F5D: push eax
  loc_00423F5E: lea edx, var_44
  loc_00423F61: push ecx
  loc_00423F62: lea eax, var_40
  loc_00423F65: push edx
  loc_00423F66: lea ecx, var_3C
  loc_00423F69: push eax
  loc_00423F6A: lea edx, var_38
  loc_00423F6D: push ecx
  loc_00423F6E: push edx
  loc_00423F6F: push 0000000Bh
  loc_00423F71: call [004010E4h] ; __vbaFreeStrList
  loc_00423F77: mov eax, Me
  loc_00423F7A: add esp, 00000030h
  loc_00423F7D: mov ecx, [eax]
  loc_00423F7F: push eax
  loc_00423F80: call [ecx+00000328h]
  loc_00423F86: lea edx, var_64
  loc_00423F89: push eax
  loc_00423F8A: push edx
  loc_00423F8B: call [0040103Ch] ; __vbaObjSet
  loc_00423F91: mov edx, var_2C
  loc_00423F94: mov ecx, [eax]
  loc_00423F96: push edx
  loc_00423F97: push eax
  loc_00423F98: mov var_68, eax
  loc_00423F9B: call [ecx+00000054h]
  loc_00423F9E: test eax, eax
  loc_00423FA0: fnclex
  loc_00423FA2: jge 00423FB6h
  loc_00423FA4: mov ecx, var_68
  loc_00423FA7: push 00000054h
  loc_00423FA9: push 0040E390h
  loc_00423FAE: push ecx
  loc_00423FAF: push eax
  loc_00423FB0: call [00401030h] ; __vbaHresultCheckObj
  loc_00423FB6: lea ecx, var_64
  loc_00423FB9: call [00401114h] ; __vbaFreeObj
  loc_00423FBF: fld real4 ptr [00430034h]
  loc_00423FC5: fmul st0, real4 ptr [00430028h]
  loc_00423FCB: mov edx, [00430038h]
  loc_00423FD1: push 00410694h ; "( "
  loc_00423FD6: push edx
  loc_00423FD7: fadd st0, real4 ptr [00430038h]
  loc_00423FDD: fstp real4 ptr var_34
  loc_00423FE0: fnstsw ax
  loc_00423FE2: test al, 0Dh
  loc_00423FE4: jnz 004243E6h
  loc_00423FEA: fld real4 ptr [00430034h]
  loc_00423FF0: fmul st0, real4 ptr [00430028h]
  loc_00423FF6: fsubr st0, real4 ptr [00430038h]
  loc_00423FFC: fstp real4 ptr var_18
  loc_00423FFF: fnstsw ax
  loc_00424001: test al, 0Dh
  loc_00424003: jnz 004243E6h
  loc_00424009: call ebx
  loc_0042400B: mov edx, eax
  loc_0042400D: lea ecx, var_38
  loc_00424010: call __vbaStrMove
  loc_00424012: push eax
  loc_00424013: call edi
  loc_00424015: mov edx, eax
  loc_00424017: lea ecx, var_3C
  loc_0042401A: call __vbaStrMove
  loc_0042401C: push eax
  loc_0042401D: push 004106A0h ; " )"
  loc_00424022: call edi
  loc_00424024: mov edx, eax
  loc_00424026: lea ecx, var_40
  loc_00424029: call __vbaStrMove
  loc_0042402B: push eax
  loc_0042402C: push 004106ACh ; " ± "
  loc_00424031: call edi
  loc_00424033: mov edx, eax
  loc_00424035: lea ecx, var_44
  loc_00424038: call __vbaStrMove
  loc_0042403A: push eax
  loc_0042403B: push 00410694h ; "( "
  loc_00424040: call edi
  loc_00424042: mov edx, eax
  loc_00424044: lea ecx, var_48
  loc_00424047: call __vbaStrMove
  loc_00424049: push eax
  loc_0042404A: mov eax, [00430028h]
  loc_0042404F: push eax
  loc_00424050: call ebx
  loc_00424052: mov edx, eax
  loc_00424054: lea ecx, var_4C
  loc_00424057: call __vbaStrMove
  loc_00424059: push eax
  loc_0042405A: call edi
  loc_0042405C: mov edx, eax
  loc_0042405E: lea ecx, var_50
  loc_00424061: call __vbaStrMove
  loc_00424063: push eax
  loc_00424064: push 004106B8h ; " ) * ( "
  loc_00424069: call edi
  loc_0042406B: mov edx, eax
  loc_0042406D: lea ecx, var_54
  loc_00424070: call __vbaStrMove
  loc_00424072: mov ecx, [00430034h]
  loc_00424078: push eax
  loc_00424079: push ecx
  loc_0042407A: call ebx
  loc_0042407C: mov edx, eax
  loc_0042407E: lea ecx, var_58
  loc_00424081: call __vbaStrMove
  loc_00424083: push eax
  loc_00424084: call edi
  loc_00424086: mov edx, eax
  loc_00424088: lea ecx, var_5C
  loc_0042408B: call __vbaStrMove
  loc_0042408D: push eax
  loc_0042408E: push 004106A0h ; " )"
  loc_00424093: call edi
  loc_00424095: mov edx, eax
  loc_00424097: lea ecx, var_14
  loc_0042409A: call __vbaStrMove
  loc_0042409C: lea edx, var_5C
  loc_0042409F: lea eax, var_58
  loc_004240A2: push edx
  loc_004240A3: lea ecx, var_54
  loc_004240A6: push eax
  loc_004240A7: lea edx, var_50
  loc_004240AA: push ecx
  loc_004240AB: lea eax, var_4C
  loc_004240AE: push edx
  loc_004240AF: lea ecx, var_48
  loc_004240B2: push eax
  loc_004240B3: lea edx, var_44
  loc_004240B6: push ecx
  loc_004240B7: lea eax, var_40
  loc_004240BA: push edx
  loc_004240BB: lea ecx, var_3C
  loc_004240BE: push eax
  loc_004240BF: lea edx, var_38
  loc_004240C2: push ecx
  loc_004240C3: push edx
  loc_004240C4: push 0000000Ah
  loc_004240C6: call [004010E4h] ; __vbaFreeStrList
  loc_004240CC: mov eax, var_14
  loc_004240CF: add esp, 0000002Ch
  loc_004240D2: push eax
  loc_004240D3: push 004106CCh ; " = "
  loc_004240D8: call edi
  loc_004240DA: mov edx, eax
  loc_004240DC: lea ecx, var_38
  loc_004240DF: call __vbaStrMove
  loc_004240E1: mov ecx, var_18
  loc_004240E4: push eax
  loc_004240E5: push ecx
  loc_004240E6: call ebx
  loc_004240E8: mov edx, eax
  loc_004240EA: lea ecx, var_3C
  loc_004240ED: call __vbaStrMove
  loc_004240EF: push eax
  loc_004240F0: call edi
  loc_004240F2: mov edx, eax
  loc_004240F4: lea ecx, var_40
  loc_004240F7: call __vbaStrMove
  loc_004240F9: push eax
  loc_004240FA: push 004106D8h ; " to "
  loc_004240FF: call edi
  loc_00424101: mov edx, eax
  loc_00424103: lea ecx, var_44
  loc_00424106: call __vbaStrMove
  loc_00424108: mov edx, var_34
  loc_0042410B: push eax
  loc_0042410C: push edx
  loc_0042410D: call ebx
  loc_0042410F: mov edx, eax
  loc_00424111: lea ecx, var_48
  loc_00424114: call __vbaStrMove
  loc_00424116: push eax
  loc_00424117: call edi
  loc_00424119: mov edx, eax
  loc_0042411B: lea ecx, var_14
  loc_0042411E: call __vbaStrMove
  loc_00424120: lea eax, var_48
  loc_00424123: lea ecx, var_44
  loc_00424126: push eax
  loc_00424127: lea edx, var_40
  loc_0042412A: push ecx
  loc_0042412B: lea eax, var_3C
  loc_0042412E: push edx
  loc_0042412F: lea ecx, var_38
  loc_00424132: push eax
  loc_00424133: push ecx
  loc_00424134: push 00000005h
  loc_00424136: call [004010E4h] ; __vbaFreeStrList
  loc_0042413C: mov eax, Me
  loc_0042413F: add esp, 00000018h
  loc_00424142: mov edx, [eax]
  loc_00424144: push eax
  loc_00424145: call [edx+00000308h]
  loc_0042414B: push eax
  loc_0042414C: lea eax, var_64
  loc_0042414F: push eax
  loc_00424150: call [0040103Ch] ; __vbaObjSet
  loc_00424156: mov edx, var_14
  loc_00424159: mov ecx, [eax]
  loc_0042415B: push edx
  loc_0042415C: push eax
  loc_0042415D: mov var_68, eax
  loc_00424160: call [ecx+00000054h]
  loc_00424163: test eax, eax
  loc_00424165: fnclex
  loc_00424167: jge 0042417Bh
  loc_00424169: mov ecx, var_68
  loc_0042416C: push 00000054h
  loc_0042416E: push 0040E390h
  loc_00424173: push ecx
  loc_00424174: push eax
  loc_00424175: call [00401030h] ; __vbaHresultCheckObj
  loc_0042417B: lea ecx, var_64
  loc_0042417E: call [00401114h] ; __vbaFreeObj
  loc_00424184: fld real4 ptr [00430020h]
  loc_0042418A: fsubr st0, real8 ptr [004011F8h]
  loc_00424190: push 004106E8h ; "With "
  loc_00424195: fmul st0, real8 ptr [00401520h]
  loc_0042419B: fstp real4 ptr var_20
  loc_0042419E: fnstsw ax
  loc_004241A0: test al, 0Dh
  loc_004241A2: jnz 004243E6h
  loc_004241A8: mov edx, var_20
  loc_004241AB: push edx
  loc_004241AC: call ebx
  loc_004241AE: mov edx, eax
  loc_004241B0: lea ecx, var_38
  loc_004241B3: call __vbaStrMove
  loc_004241B5: push eax
  loc_004241B6: call edi
  loc_004241B8: mov edx, eax
  loc_004241BA: lea ecx, var_3C
  loc_004241BD: call __vbaStrMove
  loc_004241BF: push eax
  loc_004241C0: push 004106F8h ; Chr(37) & " confidence we can say that when "
  loc_004241C5: call edi
  loc_004241C7: mov edx, eax
  loc_004241C9: lea ecx, var_40
  loc_004241CC: call __vbaStrMove
  loc_004241CE: push eax
  loc_004241CF: mov eax, [00430018h]
  loc_004241D4: push eax
  loc_004241D5: call edi
  loc_004241D7: mov edx, eax
  loc_004241D9: lea ecx, var_44
  loc_004241DC: call __vbaStrMove
  loc_004241DE: push eax
  loc_004241DF: push 00410744h ; " equals to "
  loc_004241E4: call edi
  loc_004241E6: mov edx, eax
  loc_004241E8: lea ecx, var_48
  loc_004241EB: call __vbaStrMove
  loc_004241ED: mov ecx, [0043002Ch]
  loc_004241F3: push eax
  loc_004241F4: push ecx
  loc_004241F5: call edi
  loc_004241F7: mov edx, eax
  loc_004241F9: lea ecx, var_4C
  loc_004241FC: call __vbaStrMove
  loc_004241FE: push eax
  loc_004241FF: push 0040F028h
  loc_00424204: call edi
  loc_00424206: mov edx, eax
  loc_00424208: lea ecx, var_50
  loc_0042420B: call __vbaStrMove
  loc_0042420D: mov edx, [0043001Ch]
  loc_00424213: push eax
  loc_00424214: push edx
  loc_00424215: call edi
  loc_00424217: mov edx, eax
  loc_00424219: lea ecx, var_24
  loc_0042421C: call __vbaStrMove
  loc_0042421E: lea eax, var_50
  loc_00424221: lea ecx, var_4C
  loc_00424224: push eax
  loc_00424225: lea edx, var_48
  loc_00424228: push ecx
  loc_00424229: lea eax, var_44
  loc_0042422C: push edx
  loc_0042422D: lea ecx, var_40
  loc_00424230: push eax
  loc_00424231: lea edx, var_3C
  loc_00424234: push ecx
  loc_00424235: lea eax, var_38
  loc_00424238: push edx
  loc_00424239: push eax
  loc_0042423A: push 00000007h
  loc_0042423C: call [004010E4h] ; __vbaFreeStrList
  loc_00424242: mov ecx, var_24
  loc_00424245: add esp, 00000020h
  loc_00424248: push ecx
  loc_00424249: push 004107C0h ; ", the value of "
  loc_0042424E: call edi
  loc_00424250: mov edx, eax
  loc_00424252: lea ecx, var_38
  loc_00424255: call __vbaStrMove
  loc_00424257: mov edx, [00430010h]
  loc_0042425D: push eax
  loc_0042425E: push edx
  loc_0042425F: call edi
  loc_00424261: mov edx, eax
  loc_00424263: lea ecx, var_3C
  loc_00424266: call __vbaStrMove
  loc_00424268: push eax
  loc_00424269: push 00410790h ; " falls in the range "
  loc_0042426E: call edi
  loc_00424270: mov edx, eax
  loc_00424272: lea ecx, var_40
  loc_00424275: call __vbaStrMove
  loc_00424277: push eax
  loc_00424278: mov eax, var_18
  loc_0042427B: push eax
  loc_0042427C: call ebx
  loc_0042427E: mov edx, eax
  loc_00424280: lea ecx, var_44
  loc_00424283: call __vbaStrMove
  loc_00424285: push eax
  loc_00424286: call edi
  loc_00424288: mov edx, eax
  loc_0042428A: lea ecx, var_48
  loc_0042428D: call __vbaStrMove
  loc_0042428F: push eax
  loc_00424290: push 0040F028h
  loc_00424295: call edi
  loc_00424297: mov edx, eax
  loc_00424299: lea ecx, var_4C
  loc_0042429C: call __vbaStrMove
  loc_0042429E: mov ecx, [00430014h]
  loc_004242A4: push eax
  loc_004242A5: push ecx
  loc_004242A6: call edi
  loc_004242A8: mov edx, eax
  loc_004242AA: lea ecx, var_50
  loc_004242AD: call __vbaStrMove
  loc_004242AF: push eax
  loc_004242B0: push 004106D8h ; " to "
  loc_004242B5: call edi
  loc_004242B7: mov edx, eax
  loc_004242B9: lea ecx, var_54
  loc_004242BC: call __vbaStrMove
  loc_004242BE: mov edx, var_34
  loc_004242C1: push eax
  loc_004242C2: push edx
  loc_004242C3: call ebx
  loc_004242C5: mov edx, eax
  loc_004242C7: lea ecx, var_58
  loc_004242CA: call __vbaStrMove
  loc_004242CC: push eax
  loc_004242CD: call edi
  loc_004242CF: mov edx, eax
  loc_004242D1: lea ecx, var_5C
  loc_004242D4: call __vbaStrMove
  loc_004242D6: push eax
  loc_004242D7: push 0040F028h
  loc_004242DC: call edi
  loc_004242DE: mov edx, eax
  loc_004242E0: lea ecx, var_60
  loc_004242E3: call __vbaStrMove
  loc_004242E5: push eax
  loc_004242E6: mov eax, [00430014h]
  loc_004242EB: push eax
  loc_004242EC: call edi
  loc_004242EE: mov edx, eax
  loc_004242F0: lea ecx, var_24
  loc_004242F3: call __vbaStrMove
  loc_004242F5: lea ecx, var_60
  loc_004242F8: lea edx, var_5C
  loc_004242FB: push ecx
  loc_004242FC: lea eax, var_58
  loc_004242FF: push edx
  loc_00424300: lea ecx, var_54
  loc_00424303: push eax
  loc_00424304: push ecx
  loc_00424305: lea edx, var_50
  loc_00424308: lea eax, var_4C
  loc_0042430B: push edx
  loc_0042430C: lea ecx, var_48
  loc_0042430F: push eax
  loc_00424310: lea edx, var_44
  loc_00424313: push ecx
  loc_00424314: lea eax, var_40
  loc_00424317: push edx
  loc_00424318: lea ecx, var_3C
  loc_0042431B: push eax
  loc_0042431C: lea edx, var_38
  loc_0042431F: push ecx
  loc_00424320: push edx
  loc_00424321: push 0000000Bh
  loc_00424323: call [004010E4h] ; __vbaFreeStrList
  loc_00424329: mov eax, Me
  loc_0042432C: add esp, 00000030h
  loc_0042432F: mov ecx, [eax]
  loc_00424331: push eax
  loc_00424332: call [ecx+00000300h]
  loc_00424338: lea edx, var_64
  loc_0042433B: push eax
  loc_0042433C: push edx
  loc_0042433D: call [0040103Ch] ; __vbaObjSet
  loc_00424343: mov ecx, var_24
  loc_00424346: mov esi, eax
  loc_00424348: push ecx
  loc_00424349: push esi
  loc_0042434A: mov eax, [esi]
  loc_0042434C: call [eax+00000054h]
  loc_0042434F: test eax, eax
  loc_00424351: fnclex
  loc_00424353: jge 00424364h
  loc_00424355: push 00000054h
  loc_00424357: push 0040E390h
  loc_0042435C: push esi
  loc_0042435D: push eax
  loc_0042435E: call [00401030h] ; __vbaHresultCheckObj
  loc_00424364: lea ecx, var_64
  loc_00424367: call [00401114h] ; __vbaFreeObj
  loc_0042436D: fwait
  loc_0042436E: push 004243D1h
  loc_00424373: jmp 004243B6h
  loc_00424375: lea edx, var_60
  loc_00424378: lea eax, var_5C
  loc_0042437B: push edx
  loc_0042437C: lea ecx, var_58
  loc_0042437F: push eax
  loc_00424380: lea edx, var_54
  loc_00424383: push ecx
  loc_00424384: lea eax, var_50
  loc_00424387: push edx
  loc_00424388: lea ecx, var_4C
  loc_0042438B: push eax
  loc_0042438C: lea edx, var_48
  loc_0042438F: push ecx
  loc_00424390: lea eax, var_44
  loc_00424393: push edx
  loc_00424394: lea ecx, var_40
  loc_00424397: push eax
  loc_00424398: lea edx, var_3C
  loc_0042439B: push ecx
  loc_0042439C: lea eax, var_38
  loc_0042439F: push edx
  loc_004243A0: push eax
  loc_004243A1: push 0000000Bh
  loc_004243A3: call [004010E4h] ; __vbaFreeStrList
  loc_004243A9: add esp, 00000030h
  loc_004243AC: lea ecx, var_64
  loc_004243AF: call [00401114h] ; __vbaFreeObj
  loc_004243B5: ret
  loc_004243B6: mov esi, [00401110h] ; __vbaFreeStr
  loc_004243BC: lea ecx, var_14
  loc_004243BF: call __vbaFreeStr
  loc_004243C1: lea ecx, var_24
  loc_004243C4: call __vbaFreeStr
  loc_004243C6: lea ecx, var_2C
  loc_004243C9: call __vbaFreeStr
  loc_004243CB: lea ecx, var_30
  loc_004243CE: call __vbaFreeStr
  loc_004243D0: ret
  loc_004243D1: mov ecx, var_10
  loc_004243D4: pop edi
  loc_004243D5: pop esi
  loc_004243D6: xor eax, eax
  loc_004243D8: mov fs:[00000000h], ecx
  loc_004243DF: pop ebx
  loc_004243E0: mov esp, ebp
  loc_004243E2: pop ebp
  loc_004243E3: retn 0004h
End Sub
