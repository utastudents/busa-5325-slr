VERSION 5.00
Begin VB.Form frmIntro
  Caption = "Simple Linear Regression - Introduction"
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmIntro.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 510
  ClientWidth = 8400
  ClientHeight = 5145
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraExample
    Caption = "Example"
    Left = 720
    Top = 3600
    Width = 10695
    Height = 2295
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
    Begin VB.Label lblExample
      Left = 240
      Top = 480
      Width = 10335
      Height = 1575
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
    End
  End
  Begin VB.Frame fraSituation
    Caption = "Situation"
    Left = 720
    Top = 600
    Width = 10695
    Height = 2175
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
    Begin VB.Label lblSituation
      Left = 120
      Top = 480
      Width = 10215
      Height = 1455
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
    End
  End
  Begin VB.Menu mnuInstructions
    Caption = "&Explanations"
  End
  Begin VB.Menu mnuVariables
    Caption = "&Change Names"
  End
  Begin VB.Menu mnuGoTo
    Caption = "&Go To"
    Begin VB.Menu mnuStatistics
      Caption = "Statistics and Point Estimates"
    End
    Begin VB.Menu mnuInferences
      Caption = "&Inferences"
      Begin VB.Menu mnuSlope
        Caption = "&Slope"
      End
      Begin VB.Menu mnuPrediction
        Caption = "&Prediction and Estimation Intervals"
      End
    End
    Begin VB.Menu mnuQuestions
      Caption = "&Questions"
    End
    Begin VB.Menu mnuAssumptions
      Caption = "&Assumptions"
    End
  End
  Begin VB.Menu mnuExit
    Caption = "E&xit"
  End
  Begin VB.Menu mnuAbout
    Caption = "&About"
  End
End

Attribute VB_Name = "frmIntro"


Private Sub Form_Load() '42EBE0
  loc_0042EBE0: push ebp
  loc_0042EBE1: mov ebp, esp
  loc_0042EBE3: sub esp, 0000000Ch
  loc_0042EBE6: push 00401926h ; __vbaExceptHandler
  loc_0042EBEB: mov eax, fs:[00000000h]
  loc_0042EBF1: push eax
  loc_0042EBF2: mov fs:[00000000h], esp
  loc_0042EBF9: sub esp, 00000014h
  loc_0042EBFC: push ebx
  loc_0042EBFD: push esi
  loc_0042EBFE: push edi
  loc_0042EBFF: mov var_C, esp
  loc_0042EC02: mov var_8, 00401890h
  loc_0042EC09: mov esi, Me
  loc_0042EC0C: mov eax, esi
  loc_0042EC0E: and eax, 00000001h
  loc_0042EC11: mov var_4, eax
  loc_0042EC14: and esi, FFFFFFFEh
  loc_0042EC17: push esi
  loc_0042EC18: mov Me, esi
  loc_0042EC1B: mov ecx, [esi]
  loc_0042EC1D: call [ecx+00000004h]
  loc_0042EC20: mov edx, [esi]
  loc_0042EC22: xor ebx, ebx
  loc_0042EC24: push esi
  loc_0042EC25: mov var_18, ebx
  loc_0042EC28: call [edx+00000308h]
  loc_0042EC2E: push eax
  loc_0042EC2F: lea eax, var_18
  loc_0042EC32: push eax
  loc_0042EC33: call [0040103Ch] ; __vbaObjSet
  loc_0042EC39: mov edi, eax
  loc_0042EC3B: push 00412D14h ; "You wish to determine if X has an effect of Y and then use this relationship for inferences about Y."
  loc_0042EC40: push edi
  loc_0042EC41: mov ecx, [edi]
  loc_0042EC43: call [ecx+00000054h]
  loc_0042EC46: cmp eax, ebx
  loc_0042EC48: fnclex
  loc_0042EC4A: jge 0042EC5Bh
  loc_0042EC4C: push 00000054h
  loc_0042EC4E: push 0040E390h
  loc_0042EC53: push edi
  loc_0042EC54: push eax
  loc_0042EC55: call [00401030h] ; __vbaHresultCheckObj
  loc_0042EC5B: lea ecx, var_18
  loc_0042EC5E: call [00401114h] ; __vbaFreeObj
  loc_0042EC64: mov edx, [esi]
  loc_0042EC66: push esi
  loc_0042EC67: call [edx+00000728h]
  loc_0042EC6D: mov var_4, ebx
  loc_0042EC70: push 0042EC82h
  loc_0042EC75: jmp 0042EC81h
  loc_0042EC77: lea ecx, var_18
  loc_0042EC7A: call [00401114h] ; __vbaFreeObj
  loc_0042EC80: ret
  loc_0042EC81: ret
  loc_0042EC82: mov eax, Me
  loc_0042EC85: push eax
  loc_0042EC86: mov ecx, [eax]
  loc_0042EC88: call [ecx+00000008h]
  loc_0042EC8B: mov eax, var_4
  loc_0042EC8E: mov ecx, var_14
  loc_0042EC91: pop edi
  loc_0042EC92: pop esi
  loc_0042EC93: mov fs:[00000000h], ecx
  loc_0042EC9A: pop ebx
  loc_0042EC9B: mov esp, ebp
  loc_0042EC9D: pop ebp
  loc_0042EC9E: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '42ECB0
  loc_0042ECB0: push ebp
  loc_0042ECB1: mov ebp, esp
  loc_0042ECB3: sub esp, 0000000Ch
  loc_0042ECB6: push 00401926h ; __vbaExceptHandler
  loc_0042ECBB: mov eax, fs:[00000000h]
  loc_0042ECC1: push eax
  loc_0042ECC2: mov fs:[00000000h], esp
  loc_0042ECC9: sub esp, 0000009Ch
  loc_0042ECCF: push ebx
  loc_0042ECD0: push esi
  loc_0042ECD1: push edi
  loc_0042ECD2: mov var_C, esp
  loc_0042ECD5: mov var_8, 004018A0h
  loc_0042ECDC: mov edi, Me
  loc_0042ECDF: mov eax, edi
  loc_0042ECE1: and eax, 00000001h
  loc_0042ECE4: mov var_4, eax
  loc_0042ECE7: and edi, FFFFFFFEh
  loc_0042ECEA: push edi
  loc_0042ECEB: mov Me, edi
  loc_0042ECEE: mov ecx, [edi]
  loc_0042ECF0: call [ecx+00000004h]
  loc_0042ECF3: mov ebx, [004010F4h] ; __vbaVarDup
  loc_0042ECF9: mov ecx, 80020004h
  loc_0042ECFE: xor esi, esi
  loc_0042ED00: mov var_54, ecx
  loc_0042ED03: mov eax, 0000000Ah
  loc_0042ED08: mov var_44, ecx
  loc_0042ED0B: mov var_4C, esi
  loc_0042ED0E: mov var_5C, esi
  loc_0042ED11: mov var_7C, esi
  loc_0042ED14: lea edx, var_7C
  loc_0042ED17: lea ecx, var_3C
  loc_0042ED1A: mov var_1C, esi
  loc_0042ED1D: mov var_2C, esi
  loc_0042ED20: mov var_3C, esi
  loc_0042ED23: mov var_6C, esi
  loc_0042ED26: mov var_5C, eax
  loc_0042ED29: mov var_4C, eax
  loc_0042ED2C: mov var_74, 0040ECD4h ; "Exit Check"
  loc_0042ED33: mov var_7C, 00000008h
  loc_0042ED3A: call ebx
  loc_0042ED3C: lea edx, var_6C
  loc_0042ED3F: lea ecx, var_2C
  loc_0042ED42: mov var_64, 0040EC90h ; "Are you sure you want to exit?"
  loc_0042ED49: mov var_6C, 00000008h
  loc_0042ED50: call ebx
  loc_0042ED52: lea edx, var_5C
  loc_0042ED55: lea eax, var_4C
  loc_0042ED58: push edx
  loc_0042ED59: lea ecx, var_3C
  loc_0042ED5C: push eax
  loc_0042ED5D: push ecx
  loc_0042ED5E: lea edx, var_2C
  loc_0042ED61: push 00000004h
  loc_0042ED63: push edx
  loc_0042ED64: call [00401038h] ; rtcMsgBox
  loc_0042ED6A: mov ecx, eax
  loc_0042ED6C: call [00401078h] ; __vbaI2I4
  loc_0042ED72: mov ebx, eax
  loc_0042ED74: lea eax, var_5C
  loc_0042ED77: lea ecx, var_4C
  loc_0042ED7A: push eax
  loc_0042ED7B: lea edx, var_3C
  loc_0042ED7E: push ecx
  loc_0042ED7F: lea eax, var_2C
  loc_0042ED82: push edx
  loc_0042ED83: push eax
  loc_0042ED84: push 00000004h
  loc_0042ED86: call [00401018h] ; __vbaFreeVarList
  loc_0042ED8C: add esp, 00000014h
  loc_0042ED8F: cmp bx, 0007h
  loc_0042ED93: jnz 0042ED9Fh
  loc_0042ED95: mov ecx, Cancel
  loc_0042ED98: mov [ecx], FFFFFFh
  loc_0042ED9D: jmp 0042EDFFh
  loc_0042ED9F: cmp [00430A24h], esi
  loc_0042EDA5: jnz 0042EDB7h
  loc_0042EDA7: push 00430A24h
  loc_0042EDAC: push 0040EC7Ch
  loc_0042EDB1: call [004010D4h] ; __vbaNew2
  loc_0042EDB7: mov ebx, [00430A24h]
  loc_0042EDBD: lea eax, var_1C
  loc_0042EDC0: push edi
  loc_0042EDC1: push eax
  loc_0042EDC2: mov edx, [ebx]
  loc_0042EDC4: mov var_B0, edx
  loc_0042EDCA: call [00401044h] ; __vbaObjSetAddref
  loc_0042EDD0: mov ecx, var_B0
  loc_0042EDD6: push eax
  loc_0042EDD7: push ebx
  loc_0042EDD8: call [ecx+00000010h]
  loc_0042EDDB: cmp eax, esi
  loc_0042EDDD: fnclex
  loc_0042EDDF: jge 0042EDF0h
  loc_0042EDE1: push 00000010h
  loc_0042EDE3: push 0040EC6Ch
  loc_0042EDE8: push ebx
  loc_0042EDE9: push eax
  loc_0042EDEA: call [00401030h] ; __vbaHresultCheckObj
  loc_0042EDF0: lea ecx, var_1C
  loc_0042EDF3: call [00401114h] ; __vbaFreeObj
  loc_0042EDF9: call [0040101Ch] ; __vbaEnd
  loc_0042EDFF: mov var_4, esi
  loc_0042EE02: push 0042EE2Fh
  loc_0042EE07: jmp 0042EE2Eh
  loc_0042EE09: lea ecx, var_1C
  loc_0042EE0C: call [00401114h] ; __vbaFreeObj
  loc_0042EE12: lea edx, var_5C
  loc_0042EE15: lea eax, var_4C
  loc_0042EE18: push edx
  loc_0042EE19: lea ecx, var_3C
  loc_0042EE1C: push eax
  loc_0042EE1D: lea edx, var_2C
  loc_0042EE20: push ecx
  loc_0042EE21: push edx
  loc_0042EE22: push 00000004h
  loc_0042EE24: call [00401018h] ; __vbaFreeVarList
  loc_0042EE2A: add esp, 00000014h
  loc_0042EE2D: ret
  loc_0042EE2E: ret
  loc_0042EE2F: mov eax, Me
  loc_0042EE32: push eax
  loc_0042EE33: mov ecx, [eax]
  loc_0042EE35: call [ecx+00000008h]
  loc_0042EE38: mov eax, var_4
  loc_0042EE3B: mov ecx, var_14
  loc_0042EE3E: pop edi
  loc_0042EE3F: pop esi
  loc_0042EE40: mov fs:[00000000h], ecx
  loc_0042EE47: pop ebx
  loc_0042EE48: mov esp, ebp
  loc_0042EE4A: pop ebp
  loc_0042EE4B: retn 000Ch
End Sub

Private Sub Form_Activate() '42EB70
  loc_0042EB70: push ebp
  loc_0042EB71: mov ebp, esp
  loc_0042EB73: sub esp, 0000000Ch
  loc_0042EB76: push 00401926h ; __vbaExceptHandler
  loc_0042EB7B: mov eax, fs:[00000000h]
  loc_0042EB81: push eax
  loc_0042EB82: mov fs:[00000000h], esp
  loc_0042EB89: sub esp, 00000008h
  loc_0042EB8C: push ebx
  loc_0042EB8D: push esi
  loc_0042EB8E: push edi
  loc_0042EB8F: mov var_C, esp
  loc_0042EB92: mov var_8, 00401888h
  loc_0042EB99: mov esi, Me
  loc_0042EB9C: mov eax, esi
  loc_0042EB9E: and eax, 00000001h
  loc_0042EBA1: mov var_4, eax
  loc_0042EBA4: and esi, FFFFFFFEh
  loc_0042EBA7: push esi
  loc_0042EBA8: mov Me, esi
  loc_0042EBAB: mov ecx, [esi]
  loc_0042EBAD: call [ecx+00000004h]
  loc_0042EBB0: mov edx, [esi]
  loc_0042EBB2: push esi
  loc_0042EBB3: call [edx+00000728h]
  loc_0042EBB9: mov var_4, 00000000h
  loc_0042EBC0: mov eax, Me
  loc_0042EBC3: push eax
  loc_0042EBC4: mov ecx, [eax]
  loc_0042EBC6: call [ecx+00000008h]
  loc_0042EBC9: mov eax, var_4
  loc_0042EBCC: mov ecx, var_14
  loc_0042EBCF: pop edi
  loc_0042EBD0: pop esi
  loc_0042EBD1: mov fs:[00000000h], ecx
  loc_0042EBD8: pop ebx
  loc_0042EBD9: mov esp, ebp
  loc_0042EBDB: pop ebp
  loc_0042EBDC: retn 0004h
End Sub

Private Sub mnuExit_Click() '42F090
  loc_0042F090: push ebp
  loc_0042F091: mov ebp, esp
  loc_0042F093: sub esp, 0000000Ch
  loc_0042F096: push 00401926h ; __vbaExceptHandler
  loc_0042F09B: mov eax, fs:[00000000h]
  loc_0042F0A1: push eax
  loc_0042F0A2: mov fs:[00000000h], esp
  loc_0042F0A9: sub esp, 00000018h
  loc_0042F0AC: push ebx
  loc_0042F0AD: push esi
  loc_0042F0AE: push edi
  loc_0042F0AF: mov var_C, esp
  loc_0042F0B2: mov var_8, 004018C8h
  loc_0042F0B9: mov edi, Me
  loc_0042F0BC: mov eax, edi
  loc_0042F0BE: and eax, 00000001h
  loc_0042F0C1: mov var_4, eax
  loc_0042F0C4: and edi, FFFFFFFEh
  loc_0042F0C7: push edi
  loc_0042F0C8: mov Me, edi
  loc_0042F0CB: mov ecx, [edi]
  loc_0042F0CD: call [ecx+00000004h]
  loc_0042F0D0: mov eax, [00430A24h]
  loc_0042F0D5: xor ebx, ebx
  loc_0042F0D7: cmp eax, ebx
  loc_0042F0D9: mov var_18, ebx
  loc_0042F0DC: jnz 0042F0EEh
  loc_0042F0DE: push 00430A24h
  loc_0042F0E3: push 0040EC7Ch
  loc_0042F0E8: call [004010D4h] ; __vbaNew2
  loc_0042F0EE: mov esi, [00430A24h]
  loc_0042F0F4: lea eax, var_18
  loc_0042F0F7: push edi
  loc_0042F0F8: push eax
  loc_0042F0F9: mov edx, [esi]
  loc_0042F0FB: mov var_2C, edx
  loc_0042F0FE: call [00401044h] ; __vbaObjSetAddref
  loc_0042F104: mov ecx, var_2C
  loc_0042F107: push eax
  loc_0042F108: push esi
  loc_0042F109: call [ecx+00000010h]
  loc_0042F10C: cmp eax, ebx
  loc_0042F10E: fnclex
  loc_0042F110: jge 0042F121h
  loc_0042F112: push 00000010h
  loc_0042F114: push 0040EC6Ch
  loc_0042F119: push esi
  loc_0042F11A: push eax
  loc_0042F11B: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F121: lea ecx, var_18
  loc_0042F124: call [00401114h] ; __vbaFreeObj
  loc_0042F12A: mov var_4, ebx
  loc_0042F12D: push 0042F13Fh
  loc_0042F132: jmp 0042F13Eh
  loc_0042F134: lea ecx, var_18
  loc_0042F137: call [00401114h] ; __vbaFreeObj
  loc_0042F13D: ret
  loc_0042F13E: ret
  loc_0042F13F: mov eax, Me
  loc_0042F142: push eax
  loc_0042F143: mov edx, [eax]
  loc_0042F145: call [edx+00000008h]
  loc_0042F148: mov eax, var_4
  loc_0042F14B: mov ecx, var_14
  loc_0042F14E: pop edi
  loc_0042F14F: pop esi
  loc_0042F150: mov fs:[00000000h], ecx
  loc_0042F157: pop ebx
  loc_0042F158: mov esp, ebp
  loc_0042F15A: pop ebp
  loc_0042F15B: retn 0004h
End Sub

Private Sub mnuSlope_Click() '42F560
  loc_0042F560: push ebp
  loc_0042F561: mov ebp, esp
  loc_0042F563: sub esp, 0000000Ch
  loc_0042F566: push 00401926h ; __vbaExceptHandler
  loc_0042F56B: mov eax, fs:[00000000h]
  loc_0042F571: push eax
  loc_0042F572: mov fs:[00000000h], esp
  loc_0042F579: sub esp, 00000030h
  loc_0042F57C: push ebx
  loc_0042F57D: push esi
  loc_0042F57E: push edi
  loc_0042F57F: mov var_C, esp
  loc_0042F582: mov var_8, 004018F8h
  loc_0042F589: mov eax, Me
  loc_0042F58C: mov ecx, eax
  loc_0042F58E: and ecx, 00000001h
  loc_0042F591: mov var_4, ecx
  loc_0042F594: and al, FEh
  loc_0042F596: push eax
  loc_0042F597: mov Me, eax
  loc_0042F59A: mov edx, [eax]
  loc_0042F59C: call [edx+00000004h]
  loc_0042F59F: mov eax, [00430088h]
  loc_0042F5A4: test eax, eax
  loc_0042F5A6: jnz 0042F5B8h
  loc_0042F5A8: push 00430088h
  loc_0042F5AD: push 004058C0h
  loc_0042F5B2: call [004010D4h] ; __vbaNew2
  loc_0042F5B8: mov esi, [00430088h]
  loc_0042F5BE: push esi
  loc_0042F5BF: mov eax, [esi]
  loc_0042F5C1: call [eax+000002B4h]
  loc_0042F5C7: test eax, eax
  loc_0042F5C9: fnclex
  loc_0042F5CB: jge 0042F5DFh
  loc_0042F5CD: push 000002B4h
  loc_0042F5D2: push 0040DB64h
  loc_0042F5D7: push esi
  loc_0042F5D8: push eax
  loc_0042F5D9: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F5DF: mov eax, [004300C4h]
  loc_0042F5E4: test eax, eax
  loc_0042F5E6: jnz 0042F5F8h
  loc_0042F5E8: push 004300C4h
  loc_0042F5ED: push 00409228h
  loc_0042F5F2: call [004010D4h] ; __vbaNew2
  loc_0042F5F8: sub esp, 00000010h
  loc_0042F5FB: mov ecx, 0000000Ah
  loc_0042F600: mov ebx, esp
  loc_0042F602: mov var_24, ecx
  loc_0042F605: mov eax, 80020004h
  loc_0042F60A: sub esp, 00000010h
  loc_0042F60D: mov [ebx], ecx
  loc_0042F60F: mov ecx, var_30
  loc_0042F612: mov edx, eax
  loc_0042F614: mov esi, [004300C4h]
  loc_0042F61A: mov [ebx+00000004h], ecx
  loc_0042F61D: mov ecx, esp
  loc_0042F61F: mov edi, [esi]
  loc_0042F621: push esi
  loc_0042F622: mov [ebx+00000008h], eax
  loc_0042F625: mov eax, var_28
  loc_0042F628: mov [ebx+0000000Ch], eax
  loc_0042F62B: mov eax, var_24
  loc_0042F62E: mov [ecx], eax
  loc_0042F630: mov eax, var_20
  loc_0042F633: mov [ecx+00000004h], eax
  loc_0042F636: mov [ecx+00000008h], edx
  loc_0042F639: mov edx, var_18
  loc_0042F63C: mov [ecx+0000000Ch], edx
  loc_0042F63F: call [edi+000002B0h]
  loc_0042F645: test eax, eax
  loc_0042F647: fnclex
  loc_0042F649: jge 0042F65Dh
  loc_0042F64B: push 000002B0h
  loc_0042F650: push 0040E0ECh
  loc_0042F655: push esi
  loc_0042F656: push eax
  loc_0042F657: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F65D: mov var_4, 00000000h
  loc_0042F664: mov eax, Me
  loc_0042F667: push eax
  loc_0042F668: mov ecx, [eax]
  loc_0042F66A: call [ecx+00000008h]
  loc_0042F66D: mov eax, var_4
  loc_0042F670: mov ecx, var_14
  loc_0042F673: pop edi
  loc_0042F674: pop esi
  loc_0042F675: mov fs:[00000000h], ecx
  loc_0042F67C: pop ebx
  loc_0042F67D: mov esp, ebp
  loc_0042F67F: pop ebp
  loc_0042F680: retn 0004h
End Sub

Private Sub mnuAssumptions_Click() '42EF60
  loc_0042EF60: push ebp
  loc_0042EF61: mov ebp, esp
  loc_0042EF63: sub esp, 0000000Ch
  loc_0042EF66: push 00401926h ; __vbaExceptHandler
  loc_0042EF6B: mov eax, fs:[00000000h]
  loc_0042EF71: push eax
  loc_0042EF72: mov fs:[00000000h], esp
  loc_0042EF79: sub esp, 00000030h
  loc_0042EF7C: push ebx
  loc_0042EF7D: push esi
  loc_0042EF7E: push edi
  loc_0042EF7F: mov var_C, esp
  loc_0042EF82: mov var_8, 004018C0h
  loc_0042EF89: mov eax, Me
  loc_0042EF8C: mov ecx, eax
  loc_0042EF8E: and ecx, 00000001h
  loc_0042EF91: mov var_4, ecx
  loc_0042EF94: and al, FEh
  loc_0042EF96: push eax
  loc_0042EF97: mov Me, eax
  loc_0042EF9A: mov edx, [eax]
  loc_0042EF9C: call [edx+00000004h]
  loc_0042EF9F: mov eax, [0043009Ch]
  loc_0042EFA4: test eax, eax
  loc_0042EFA6: jnz 0042EFB8h
  loc_0042EFA8: push 0043009Ch
  loc_0042EFAD: push 00405FC0h
  loc_0042EFB2: call [004010D4h] ; __vbaNew2
  loc_0042EFB8: sub esp, 00000010h
  loc_0042EFBB: mov ecx, 0000000Ah
  loc_0042EFC0: mov ebx, esp
  loc_0042EFC2: mov var_24, ecx
  loc_0042EFC5: mov eax, 80020004h
  loc_0042EFCA: sub esp, 00000010h
  loc_0042EFCD: mov [ebx], ecx
  loc_0042EFCF: mov ecx, var_30
  loc_0042EFD2: mov edx, eax
  loc_0042EFD4: mov esi, [0043009Ch]
  loc_0042EFDA: mov [ebx+00000004h], ecx
  loc_0042EFDD: mov ecx, esp
  loc_0042EFDF: mov edi, [esi]
  loc_0042EFE1: push esi
  loc_0042EFE2: mov [ebx+00000008h], eax
  loc_0042EFE5: mov eax, var_28
  loc_0042EFE8: mov [ebx+0000000Ch], eax
  loc_0042EFEB: mov eax, var_24
  loc_0042EFEE: mov [ecx], eax
  loc_0042EFF0: mov eax, var_20
  loc_0042EFF3: mov [ecx+00000004h], eax
  loc_0042EFF6: mov [ecx+00000008h], edx
  loc_0042EFF9: mov edx, var_18
  loc_0042EFFC: mov [ecx+0000000Ch], edx
  loc_0042EFFF: call [edi+000002B0h]
  loc_0042F005: test eax, eax
  loc_0042F007: fnclex
  loc_0042F009: jge 0042F01Dh
  loc_0042F00B: push 000002B0h
  loc_0042F010: push 0040DDE0h
  loc_0042F015: push esi
  loc_0042F016: push eax
  loc_0042F017: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F01D: mov eax, [00430088h]
  loc_0042F022: test eax, eax
  loc_0042F024: jnz 0042F036h
  loc_0042F026: push 00430088h
  loc_0042F02B: push 004058C0h
  loc_0042F030: call [004010D4h] ; __vbaNew2
  loc_0042F036: mov esi, [00430088h]
  loc_0042F03C: push esi
  loc_0042F03D: mov eax, [esi]
  loc_0042F03F: call [eax+000002B4h]
  loc_0042F045: test eax, eax
  loc_0042F047: fnclex
  loc_0042F049: jge 0042F05Dh
  loc_0042F04B: push 000002B4h
  loc_0042F050: push 0040DB64h
  loc_0042F055: push esi
  loc_0042F056: push eax
  loc_0042F057: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F05D: mov var_4, 00000000h
  loc_0042F064: mov eax, Me
  loc_0042F067: push eax
  loc_0042F068: mov ecx, [eax]
  loc_0042F06A: call [ecx+00000008h]
  loc_0042F06D: mov eax, var_4
  loc_0042F070: mov ecx, var_14
  loc_0042F073: pop edi
  loc_0042F074: pop esi
  loc_0042F075: mov fs:[00000000h], ecx
  loc_0042F07C: pop ebx
  loc_0042F07D: mov esp, ebp
  loc_0042F07F: pop ebp
  loc_0042F080: retn 0004h
End Sub

Private Sub mnuQuestions_Click() '42F430
  loc_0042F430: push ebp
  loc_0042F431: mov ebp, esp
  loc_0042F433: sub esp, 0000000Ch
  loc_0042F436: push 00401926h ; __vbaExceptHandler
  loc_0042F43B: mov eax, fs:[00000000h]
  loc_0042F441: push eax
  loc_0042F442: mov fs:[00000000h], esp
  loc_0042F449: sub esp, 00000030h
  loc_0042F44C: push ebx
  loc_0042F44D: push esi
  loc_0042F44E: push edi
  loc_0042F44F: mov var_C, esp
  loc_0042F452: mov var_8, 004018F0h
  loc_0042F459: mov eax, Me
  loc_0042F45C: mov ecx, eax
  loc_0042F45E: and ecx, 00000001h
  loc_0042F461: mov var_4, ecx
  loc_0042F464: and al, FEh
  loc_0042F466: push eax
  loc_0042F467: mov Me, eax
  loc_0042F46A: mov edx, [eax]
  loc_0042F46C: call [edx+00000004h]
  loc_0042F46F: mov eax, [00430088h]
  loc_0042F474: test eax, eax
  loc_0042F476: jnz 0042F488h
  loc_0042F478: push 00430088h
  loc_0042F47D: push 004058C0h
  loc_0042F482: call [004010D4h] ; __vbaNew2
  loc_0042F488: mov esi, [00430088h]
  loc_0042F48E: push esi
  loc_0042F48F: mov eax, [esi]
  loc_0042F491: call [eax+000002B4h]
  loc_0042F497: test eax, eax
  loc_0042F499: fnclex
  loc_0042F49B: jge 0042F4AFh
  loc_0042F49D: push 000002B4h
  loc_0042F4A2: push 0040DB64h
  loc_0042F4A7: push esi
  loc_0042F4A8: push eax
  loc_0042F4A9: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F4AF: mov eax, [00430164h]
  loc_0042F4B4: test eax, eax
  loc_0042F4B6: jnz 0042F4C8h
  loc_0042F4B8: push 00430164h
  loc_0042F4BD: push 00408108h
  loc_0042F4C2: call [004010D4h] ; __vbaNew2
  loc_0042F4C8: sub esp, 00000010h
  loc_0042F4CB: mov ecx, 0000000Ah
  loc_0042F4D0: mov ebx, esp
  loc_0042F4D2: mov var_24, ecx
  loc_0042F4D5: mov eax, 80020004h
  loc_0042F4DA: sub esp, 00000010h
  loc_0042F4DD: mov [ebx], ecx
  loc_0042F4DF: mov ecx, var_30
  loc_0042F4E2: mov edx, eax
  loc_0042F4E4: mov esi, [00430164h]
  loc_0042F4EA: mov [ebx+00000004h], ecx
  loc_0042F4ED: mov ecx, esp
  loc_0042F4EF: mov edi, [esi]
  loc_0042F4F1: push esi
  loc_0042F4F2: mov [ebx+00000008h], eax
  loc_0042F4F5: mov eax, var_28
  loc_0042F4F8: mov [ebx+0000000Ch], eax
  loc_0042F4FB: mov eax, var_24
  loc_0042F4FE: mov [ecx], eax
  loc_0042F500: mov eax, var_20
  loc_0042F503: mov [ecx+00000004h], eax
  loc_0042F506: mov [ecx+00000008h], edx
  loc_0042F509: mov edx, var_18
  loc_0042F50C: mov [ecx+0000000Ch], edx
  loc_0042F50F: call [edi+000002B0h]
  loc_0042F515: test eax, eax
  loc_0042F517: fnclex
  loc_0042F519: jge 0042F52Dh
  loc_0042F51B: push 000002B0h
  loc_0042F520: push 0040FF58h
  loc_0042F525: push esi
  loc_0042F526: push eax
  loc_0042F527: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F52D: mov var_4, 00000000h
  loc_0042F534: mov eax, Me
  loc_0042F537: push eax
  loc_0042F538: mov ecx, [eax]
  loc_0042F53A: call [ecx+00000008h]
  loc_0042F53D: mov eax, var_4
  loc_0042F540: mov ecx, var_14
  loc_0042F543: pop edi
  loc_0042F544: pop esi
  loc_0042F545: mov fs:[00000000h], ecx
  loc_0042F54C: pop ebx
  loc_0042F54D: mov esp, ebp
  loc_0042F54F: pop ebp
  loc_0042F550: retn 0004h
End Sub

Private Sub mnuStatistics_Click() '42F690
  loc_0042F690: push ebp
  loc_0042F691: mov ebp, esp
  loc_0042F693: sub esp, 0000000Ch
  loc_0042F696: push 00401926h ; __vbaExceptHandler
  loc_0042F69B: mov eax, fs:[00000000h]
  loc_0042F6A1: push eax
  loc_0042F6A2: mov fs:[00000000h], esp
  loc_0042F6A9: sub esp, 00000030h
  loc_0042F6AC: push ebx
  loc_0042F6AD: push esi
  loc_0042F6AE: push edi
  loc_0042F6AF: mov var_C, esp
  loc_0042F6B2: mov var_8, 00401900h
  loc_0042F6B9: mov eax, Me
  loc_0042F6BC: mov ecx, eax
  loc_0042F6BE: and ecx, 00000001h
  loc_0042F6C1: mov var_4, ecx
  loc_0042F6C4: and al, FEh
  loc_0042F6C6: push eax
  loc_0042F6C7: mov Me, eax
  loc_0042F6CA: mov edx, [eax]
  loc_0042F6CC: call [edx+00000004h]
  loc_0042F6CF: mov eax, [004300ECh]
  loc_0042F6D4: test eax, eax
  loc_0042F6D6: jnz 0042F6E8h
  loc_0042F6D8: push 004300ECh
  loc_0042F6DD: push 0040A624h
  loc_0042F6E2: call [004010D4h] ; __vbaNew2
  loc_0042F6E8: sub esp, 00000010h
  loc_0042F6EB: mov ecx, 0000000Ah
  loc_0042F6F0: mov ebx, esp
  loc_0042F6F2: mov var_24, ecx
  loc_0042F6F5: mov eax, 80020004h
  loc_0042F6FA: sub esp, 00000010h
  loc_0042F6FD: mov [ebx], ecx
  loc_0042F6FF: mov ecx, var_30
  loc_0042F702: mov edx, eax
  loc_0042F704: mov esi, [004300ECh]
  loc_0042F70A: mov [ebx+00000004h], ecx
  loc_0042F70D: mov ecx, esp
  loc_0042F70F: mov edi, [esi]
  loc_0042F711: push esi
  loc_0042F712: mov [ebx+00000008h], eax
  loc_0042F715: mov eax, var_28
  loc_0042F718: mov [ebx+0000000Ch], eax
  loc_0042F71B: mov eax, var_24
  loc_0042F71E: mov [ecx], eax
  loc_0042F720: mov eax, var_20
  loc_0042F723: mov [ecx+00000004h], eax
  loc_0042F726: mov [ecx+00000008h], edx
  loc_0042F729: mov edx, var_18
  loc_0042F72C: mov [ecx+0000000Ch], edx
  loc_0042F72F: call [edi+000002B0h]
  loc_0042F735: test eax, eax
  loc_0042F737: fnclex
  loc_0042F739: jge 0042F74Dh
  loc_0042F73B: push 000002B0h
  loc_0042F740: push 0040ECECh
  loc_0042F745: push esi
  loc_0042F746: push eax
  loc_0042F747: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F74D: mov eax, [00430088h]
  loc_0042F752: test eax, eax
  loc_0042F754: jnz 0042F766h
  loc_0042F756: push 00430088h
  loc_0042F75B: push 004058C0h
  loc_0042F760: call [004010D4h] ; __vbaNew2
  loc_0042F766: mov esi, [00430088h]
  loc_0042F76C: push esi
  loc_0042F76D: mov eax, [esi]
  loc_0042F76F: call [eax+000002B4h]
  loc_0042F775: test eax, eax
  loc_0042F777: fnclex
  loc_0042F779: jge 0042F78Dh
  loc_0042F77B: push 000002B4h
  loc_0042F780: push 0040DB64h
  loc_0042F785: push esi
  loc_0042F786: push eax
  loc_0042F787: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F78D: mov var_4, 00000000h
  loc_0042F794: mov eax, Me
  loc_0042F797: push eax
  loc_0042F798: mov ecx, [eax]
  loc_0042F79A: call [ecx+00000008h]
  loc_0042F79D: mov eax, var_4
  loc_0042F7A0: mov ecx, var_14
  loc_0042F7A3: pop edi
  loc_0042F7A4: pop esi
  loc_0042F7A5: mov fs:[00000000h], ecx
  loc_0042F7AC: pop ebx
  loc_0042F7AD: mov esp, ebp
  loc_0042F7AF: pop ebp
  loc_0042F7B0: retn 0004h
End Sub

Private Sub mnuVariables_Click() '42F7C0
  loc_0042F7C0: push ebp
  loc_0042F7C1: mov ebp, esp
  loc_0042F7C3: sub esp, 0000000Ch
  loc_0042F7C6: push 00401926h ; __vbaExceptHandler
  loc_0042F7CB: mov eax, fs:[00000000h]
  loc_0042F7D1: push eax
  loc_0042F7D2: mov fs:[00000000h], esp
  loc_0042F7D9: sub esp, 00000030h
  loc_0042F7DC: push ebx
  loc_0042F7DD: push esi
  loc_0042F7DE: push edi
  loc_0042F7DF: mov var_C, esp
  loc_0042F7E2: mov var_8, 00401908h
  loc_0042F7E9: mov eax, Me
  loc_0042F7EC: mov ecx, eax
  loc_0042F7EE: and ecx, 00000001h
  loc_0042F7F1: mov var_4, ecx
  loc_0042F7F4: and al, FEh
  loc_0042F7F6: push eax
  loc_0042F7F7: mov Me, eax
  loc_0042F7FA: mov edx, [eax]
  loc_0042F7FC: call [edx+00000004h]
  loc_0042F7FF: mov eax, [004300D8h]
  loc_0042F804: test eax, eax
  loc_0042F806: jnz 0042F818h
  loc_0042F808: push 004300D8h
  loc_0042F80D: push 00402E04h
  loc_0042F812: call [004010D4h] ; __vbaNew2
  loc_0042F818: sub esp, 00000010h
  loc_0042F81B: mov ecx, 0000000Ah
  loc_0042F820: mov ebx, esp
  loc_0042F822: mov var_24, ecx
  loc_0042F825: mov eax, 80020004h
  loc_0042F82A: sub esp, 00000010h
  loc_0042F82D: mov [ebx], ecx
  loc_0042F82F: mov ecx, var_30
  loc_0042F832: mov edx, eax
  loc_0042F834: mov esi, [004300D8h]
  loc_0042F83A: mov [ebx+00000004h], ecx
  loc_0042F83D: mov ecx, esp
  loc_0042F83F: mov edi, [esi]
  loc_0042F841: push esi
  loc_0042F842: mov [ebx+00000008h], eax
  loc_0042F845: mov eax, var_28
  loc_0042F848: mov [ebx+0000000Ch], eax
  loc_0042F84B: mov eax, var_24
  loc_0042F84E: mov [ecx], eax
  loc_0042F850: mov eax, var_20
  loc_0042F853: mov [ecx+00000004h], eax
  loc_0042F856: mov [ecx+00000008h], edx
  loc_0042F859: mov edx, var_18
  loc_0042F85C: mov [ecx+0000000Ch], edx
  loc_0042F85F: call [edi+000002B0h]
  loc_0042F865: test eax, eax
  loc_0042F867: fnclex
  loc_0042F869: jge 0042F87Dh
  loc_0042F86B: push 000002B0h
  loc_0042F870: push 0040E260h
  loc_0042F875: push esi
  loc_0042F876: push eax
  loc_0042F877: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F87D: mov var_4, 00000000h
  loc_0042F884: mov eax, Me
  loc_0042F887: push eax
  loc_0042F888: mov ecx, [eax]
  loc_0042F88A: call [ecx+00000008h]
  loc_0042F88D: mov eax, var_4
  loc_0042F890: mov ecx, var_14
  loc_0042F893: pop edi
  loc_0042F894: pop esi
  loc_0042F895: mov fs:[00000000h], ecx
  loc_0042F89C: pop ebx
  loc_0042F89D: mov esp, ebp
  loc_0042F89F: pop ebp
  loc_0042F8A0: retn 0004h
End Sub

Private Sub mnuAbout_Click() '42EE50
  loc_0042EE50: push ebp
  loc_0042EE51: mov ebp, esp
  loc_0042EE53: sub esp, 0000000Ch
  loc_0042EE56: push 00401926h ; __vbaExceptHandler
  loc_0042EE5B: mov eax, fs:[00000000h]
  loc_0042EE61: push eax
  loc_0042EE62: mov fs:[00000000h], esp
  loc_0042EE69: sub esp, 00000088h
  loc_0042EE6F: push ebx
  loc_0042EE70: push esi
  loc_0042EE71: push edi
  loc_0042EE72: mov var_C, esp
  loc_0042EE75: mov var_8, 004018B0h
  loc_0042EE7C: mov eax, Me
  loc_0042EE7F: mov ecx, eax
  loc_0042EE81: and ecx, 00000001h
  loc_0042EE84: mov var_4, ecx
  loc_0042EE87: and al, FEh
  loc_0042EE89: push eax
  loc_0042EE8A: mov Me, eax
  loc_0042EE8D: mov edx, [eax]
  loc_0042EE8F: call [edx+00000004h]
  loc_0042EE92: mov ecx, 80020004h
  loc_0042EE97: xor esi, esi
  loc_0042EE99: mov var_4C, ecx
  loc_0042EE9C: mov eax, 0000000Ah
  loc_0042EEA1: mov var_3C, ecx
  loc_0042EEA4: mov var_2C, ecx
  loc_0042EEA7: mov var_34, esi
  loc_0042EEAA: mov var_44, esi
  loc_0042EEAD: mov var_54, esi
  loc_0042EEB0: mov var_64, esi
  loc_0042EEB3: lea edx, var_64
  loc_0042EEB6: lea ecx, var_24
  loc_0042EEB9: mov var_24, esi
  loc_0042EEBC: mov var_54, eax
  loc_0042EEBF: mov var_44, eax
  loc_0042EEC2: mov var_34, eax
  loc_0042EEC5: mov var_5C, 00412EDCh ; "Created by Dr. Mark Eakin, December, 1999"
  loc_0042EECC: mov var_64, 00000008h
  loc_0042EED3: call [004010F4h] ; __vbaVarDup
  loc_0042EED9: lea eax, var_54
  loc_0042EEDC: lea ecx, var_44
  loc_0042EEDF: push eax
  loc_0042EEE0: lea edx, var_34
  loc_0042EEE3: push ecx
  loc_0042EEE4: push edx
  loc_0042EEE5: lea eax, var_24
  loc_0042EEE8: push esi
  loc_0042EEE9: push eax
  loc_0042EEEA: call [00401038h] ; rtcMsgBox
  loc_0042EEF0: lea ecx, var_54
  loc_0042EEF3: lea edx, var_44
  loc_0042EEF6: push ecx
  loc_0042EEF7: lea eax, var_34
  loc_0042EEFA: push edx
  loc_0042EEFB: lea ecx, var_24
  loc_0042EEFE: push eax
  loc_0042EEFF: push ecx
  loc_0042EF00: push 00000004h
  loc_0042EF02: call [00401018h] ; __vbaFreeVarList
  loc_0042EF08: add esp, 00000014h
  loc_0042EF0B: mov var_4, esi
  loc_0042EF0E: push 0042EF32h
  loc_0042EF13: jmp 0042EF31h
  loc_0042EF15: lea edx, var_54
  loc_0042EF18: lea eax, var_44
  loc_0042EF1B: push edx
  loc_0042EF1C: lea ecx, var_34
  loc_0042EF1F: push eax
  loc_0042EF20: lea edx, var_24
  loc_0042EF23: push ecx
  loc_0042EF24: push edx
  loc_0042EF25: push 00000004h
  loc_0042EF27: call [00401018h] ; __vbaFreeVarList
  loc_0042EF2D: add esp, 00000014h
  loc_0042EF30: ret
  loc_0042EF31: ret
  loc_0042EF32: mov eax, Me
  loc_0042EF35: push eax
  loc_0042EF36: mov ecx, [eax]
  loc_0042EF38: call [ecx+00000008h]
  loc_0042EF3B: mov eax, var_4
  loc_0042EF3E: mov ecx, var_14
  loc_0042EF41: pop edi
  loc_0042EF42: pop esi
  loc_0042EF43: mov fs:[00000000h], ecx
  loc_0042EF4A: pop ebx
  loc_0042EF4B: mov esp, ebp
  loc_0042EF4D: pop ebp
  loc_0042EF4E: retn 0004h
End Sub

Private Sub mnuPrediction_Click() '42F300
  loc_0042F300: push ebp
  loc_0042F301: mov ebp, esp
  loc_0042F303: sub esp, 0000000Ch
  loc_0042F306: push 00401926h ; __vbaExceptHandler
  loc_0042F30B: mov eax, fs:[00000000h]
  loc_0042F311: push eax
  loc_0042F312: mov fs:[00000000h], esp
  loc_0042F319: sub esp, 00000030h
  loc_0042F31C: push ebx
  loc_0042F31D: push esi
  loc_0042F31E: push edi
  loc_0042F31F: mov var_C, esp
  loc_0042F322: mov var_8, 004018E8h
  loc_0042F329: mov eax, Me
  loc_0042F32C: mov ecx, eax
  loc_0042F32E: and ecx, 00000001h
  loc_0042F331: mov var_4, ecx
  loc_0042F334: and al, FEh
  loc_0042F336: push eax
  loc_0042F337: mov Me, eax
  loc_0042F33A: mov edx, [eax]
  loc_0042F33C: call [edx+00000004h]
  loc_0042F33F: mov eax, [004300B0h]
  loc_0042F344: test eax, eax
  loc_0042F346: jnz 0042F358h
  loc_0042F348: push 004300B0h
  loc_0042F34D: push 00407528h
  loc_0042F352: call [004010D4h] ; __vbaNew2
  loc_0042F358: sub esp, 00000010h
  loc_0042F35B: mov ecx, 0000000Ah
  loc_0042F360: mov ebx, esp
  loc_0042F362: mov var_24, ecx
  loc_0042F365: mov eax, 80020004h
  loc_0042F36A: sub esp, 00000010h
  loc_0042F36D: mov [ebx], ecx
  loc_0042F36F: mov ecx, var_30
  loc_0042F372: mov edx, eax
  loc_0042F374: mov esi, [004300B0h]
  loc_0042F37A: mov [ebx+00000004h], ecx
  loc_0042F37D: mov ecx, esp
  loc_0042F37F: mov edi, [esi]
  loc_0042F381: push esi
  loc_0042F382: mov [ebx+00000008h], eax
  loc_0042F385: mov eax, var_28
  loc_0042F388: mov [ebx+0000000Ch], eax
  loc_0042F38B: mov eax, var_24
  loc_0042F38E: mov [ecx], eax
  loc_0042F390: mov eax, var_20
  loc_0042F393: mov [ecx+00000004h], eax
  loc_0042F396: mov [ecx+00000008h], edx
  loc_0042F399: mov edx, var_18
  loc_0042F39C: mov [ecx+0000000Ch], edx
  loc_0042F39F: call [edi+000002B0h]
  loc_0042F3A5: test eax, eax
  loc_0042F3A7: fnclex
  loc_0042F3A9: jge 0042F3BDh
  loc_0042F3AB: push 000002B0h
  loc_0042F3B0: push 0040DE98h
  loc_0042F3B5: push esi
  loc_0042F3B6: push eax
  loc_0042F3B7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F3BD: mov eax, [00430088h]
  loc_0042F3C2: test eax, eax
  loc_0042F3C4: jnz 0042F3D6h
  loc_0042F3C6: push 00430088h
  loc_0042F3CB: push 004058C0h
  loc_0042F3D0: call [004010D4h] ; __vbaNew2
  loc_0042F3D6: mov esi, [00430088h]
  loc_0042F3DC: push esi
  loc_0042F3DD: mov eax, [esi]
  loc_0042F3DF: call [eax+000002B4h]
  loc_0042F3E5: test eax, eax
  loc_0042F3E7: fnclex
  loc_0042F3E9: jge 0042F3FDh
  loc_0042F3EB: push 000002B4h
  loc_0042F3F0: push 0040DB64h
  loc_0042F3F5: push esi
  loc_0042F3F6: push eax
  loc_0042F3F7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F3FD: mov var_4, 00000000h
  loc_0042F404: mov eax, Me
  loc_0042F407: push eax
  loc_0042F408: mov ecx, [eax]
  loc_0042F40A: call [ecx+00000008h]
  loc_0042F40D: mov eax, var_4
  loc_0042F410: mov ecx, var_14
  loc_0042F413: pop edi
  loc_0042F414: pop esi
  loc_0042F415: mov fs:[00000000h], ecx
  loc_0042F41C: pop ebx
  loc_0042F41D: mov esp, ebp
  loc_0042F41F: pop ebp
  loc_0042F420: retn 0004h
End Sub

Private Sub mnuInstructions_Click() '42F160
  loc_0042F160: push ebp
  loc_0042F161: mov ebp, esp
  loc_0042F163: sub esp, 0000000Ch
  loc_0042F166: push 00401926h ; __vbaExceptHandler
  loc_0042F16B: mov eax, fs:[00000000h]
  loc_0042F171: push eax
  loc_0042F172: mov fs:[00000000h], esp
  loc_0042F179: sub esp, 00000080h
  loc_0042F17F: push ebx
  loc_0042F180: push esi
  loc_0042F181: push edi
  loc_0042F182: mov var_C, esp
  loc_0042F185: mov var_8, 004018D8h
  loc_0042F18C: mov eax, Me
  loc_0042F18F: mov ecx, eax
  loc_0042F191: and ecx, 00000001h
  loc_0042F194: mov var_4, ecx
  loc_0042F197: and al, FEh
  loc_0042F199: push eax
  loc_0042F19A: mov Me, eax
  loc_0042F19D: mov edx, [eax]
  loc_0042F19F: call [edx+00000004h]
  loc_0042F1A2: mov esi, [0040102Ch] ; __vbaStrCat
  loc_0042F1A8: xor ebx, ebx
  loc_0042F1AA: push 00412F34h ; "CHANGE NAMES changes the names of the variables in the example."
  loc_0042F1AF: push 0040F720h ; vbCrLf
  loc_0042F1B4: mov var_18, ebx
  loc_0042F1B7: mov var_1C, ebx
  loc_0042F1BA: mov var_2C, ebx
  loc_0042F1BD: mov var_3C, ebx
  loc_0042F1C0: mov var_4C, ebx
  loc_0042F1C3: mov var_5C, ebx
  loc_0042F1C6: mov var_6C, ebx
  loc_0042F1C9: call __vbaStrCat
  loc_0042F1CB: mov edi, [004010FCh] ; __vbaStrMove
  loc_0042F1D1: mov edx, eax
  loc_0042F1D3: lea ecx, var_18
  loc_0042F1D6: call edi
  loc_0042F1D8: mov eax, var_18
  loc_0042F1DB: push eax
  loc_0042F1DC: push 00412FB8h ; "INFERENCES gives examples of solved hypothesis test and confidence intervals."
  loc_0042F1E1: call __vbaStrCat
  loc_0042F1E3: mov edx, eax
  loc_0042F1E5: lea ecx, var_1C
  loc_0042F1E8: call edi
  loc_0042F1EA: push eax
  loc_0042F1EB: push 0040F720h ; vbCrLf
  loc_0042F1F0: call __vbaStrCat
  loc_0042F1F2: mov edx, eax
  loc_0042F1F4: lea ecx, var_18
  loc_0042F1F7: call edi
  loc_0042F1F9: lea ecx, var_1C
  loc_0042F1FC: call [00401110h] ; __vbaFreeStr
  loc_0042F202: mov ecx, var_18
  loc_0042F205: push ecx
  loc_0042F206: push 00413058h ; "QUESTIONS gives examples of managerial questions that lead to specific inferences."
  loc_0042F20B: call __vbaStrCat
  loc_0042F20D: mov edx, eax
  loc_0042F20F: lea ecx, var_1C
  loc_0042F212: call edi
  loc_0042F214: push eax
  loc_0042F215: push 0040F720h ; vbCrLf
  loc_0042F21A: call __vbaStrCat
  loc_0042F21C: mov edx, eax
  loc_0042F21E: lea ecx, var_18
  loc_0042F221: call edi
  loc_0042F223: lea ecx, var_1C
  loc_0042F226: call [00401110h] ; __vbaFreeStr
  loc_0042F22C: mov edx, var_18
  loc_0042F22F: push edx
  loc_0042F230: push 00413104h ; "EXIT exits the program."
  loc_0042F235: call __vbaStrCat
  loc_0042F237: mov edx, eax
  loc_0042F239: lea ecx, var_18
  loc_0042F23C: call edi
  loc_0042F23E: mov ecx, 80020004h
  loc_0042F243: mov eax, 0000000Ah
  loc_0042F248: mov var_44, ecx
  loc_0042F24B: mov var_34, ecx
  loc_0042F24E: lea edx, var_6C
  loc_0042F251: lea ecx, var_2C
  loc_0042F254: mov var_4C, eax
  loc_0042F257: mov var_3C, eax
  loc_0042F25A: mov var_64, 00413138h ; "Menu Explanations"
  loc_0042F261: mov var_6C, 00000008h
  loc_0042F268: call [004010F4h] ; __vbaVarDup
  loc_0042F26E: lea eax, var_18
  loc_0042F271: mov var_5C, 00004008h
  loc_0042F278: mov var_54, eax
  loc_0042F27B: lea ecx, var_4C
  loc_0042F27E: lea edx, var_3C
  loc_0042F281: push ecx
  loc_0042F282: lea eax, var_2C
  loc_0042F285: push edx
  loc_0042F286: push eax
  loc_0042F287: lea ecx, var_5C
  loc_0042F28A: push ebx
  loc_0042F28B: push ecx
  loc_0042F28C: call [00401038h] ; rtcMsgBox
  loc_0042F292: lea edx, var_4C
  loc_0042F295: lea eax, var_3C
  loc_0042F298: push edx
  loc_0042F299: lea ecx, var_2C
  loc_0042F29C: push eax
  loc_0042F29D: push ecx
  loc_0042F29E: push 00000003h
  loc_0042F2A0: call [00401018h] ; __vbaFreeVarList
  loc_0042F2A6: add esp, 00000010h
  loc_0042F2A9: mov var_4, ebx
  loc_0042F2AC: push 0042F2DEh
  loc_0042F2B1: jmp 0042F2D4h
  loc_0042F2B3: lea ecx, var_1C
  loc_0042F2B6: call [00401110h] ; __vbaFreeStr
  loc_0042F2BC: lea edx, var_4C
  loc_0042F2BF: lea eax, var_3C
  loc_0042F2C2: push edx
  loc_0042F2C3: lea ecx, var_2C
  loc_0042F2C6: push eax
  loc_0042F2C7: push ecx
  loc_0042F2C8: push 00000003h
  loc_0042F2CA: call [00401018h] ; __vbaFreeVarList
  loc_0042F2D0: add esp, 00000010h
  loc_0042F2D3: ret
  loc_0042F2D4: lea ecx, var_18
  loc_0042F2D7: call [00401110h] ; __vbaFreeStr
  loc_0042F2DD: ret
  loc_0042F2DE: mov eax, Me
  loc_0042F2E1: push eax
  loc_0042F2E2: mov edx, [eax]
  loc_0042F2E4: call [edx+00000008h]
  loc_0042F2E7: mov eax, var_4
  loc_0042F2EA: mov ecx, var_14
  loc_0042F2ED: pop edi
  loc_0042F2EE: pop esi
  loc_0042F2EF: mov fs:[00000000h], ecx
  loc_0042F2F6: pop ebx
  loc_0042F2F7: mov esp, ebp
  loc_0042F2F9: pop ebp
  loc_0042F2FA: retn 0004h
End Sub

Private Sub Proc_17_12_42F8B0
  loc_0042F8B0: push ebp
  loc_0042F8B1: mov ebp, esp
  loc_0042F8B3: sub esp, 00000008h
  loc_0042F8B6: push 00401926h ; __vbaExceptHandler
  loc_0042F8BB: mov eax, fs:[00000000h]
  loc_0042F8C1: push eax
  loc_0042F8C2: mov fs:[00000000h], esp
  loc_0042F8C9: sub esp, 00000028h
  loc_0042F8CC: push ebx
  loc_0042F8CD: push esi
  loc_0042F8CE: push edi
  loc_0042F8CF: mov var_8, esp
  loc_0042F8D2: mov var_4, 00401910h
  loc_0042F8D9: xor eax, eax
  loc_0042F8DB: mov var_14, eax
  loc_0042F8DE: mov var_18, eax
  loc_0042F8E1: mov var_1C, eax
  loc_0042F8E4: mov var_20, eax
  loc_0042F8E7: mov var_24, eax
  loc_0042F8EA: mov var_28, eax
  loc_0042F8ED: mov var_2C, eax
  loc_0042F8F0: mov eax, Me
  loc_0042F8F3: push eax
  loc_0042F8F4: mov ecx, [eax]
  loc_0042F8F6: call [ecx+00000300h]
  loc_0042F8FC: lea edx, var_2C
  loc_0042F8FF: push eax
  loc_0042F900: push edx
  loc_0042F901: call [0040103Ch] ; __vbaObjSet
  loc_0042F907: mov ebx, [eax]
  loc_0042F909: mov esi, [0040102Ch] ; __vbaStrCat
  loc_0042F90F: mov var_30, eax
  loc_0042F912: mov eax, [00430018h]
  loc_0042F917: push 00412E1Ch ; "You wish to investigate the effect of"
  loc_0042F91C: push eax
  loc_0042F91D: call __vbaStrCat
  loc_0042F91F: mov edi, [004010FCh] ; __vbaStrMove
  loc_0042F925: mov edx, eax
  loc_0042F927: lea ecx, var_14
  loc_0042F92A: call edi
  loc_0042F92C: push eax
  loc_0042F92D: push 00412E6Ch ; " on"
  loc_0042F932: call __vbaStrCat
  loc_0042F934: mov edx, eax
  loc_0042F936: lea ecx, var_18
  loc_0042F939: call edi
  loc_0042F93B: mov ecx, [00430010h]
  loc_0042F941: push eax
  loc_0042F942: push ecx
  loc_0042F943: call __vbaStrCat
  loc_0042F945: mov edx, eax
  loc_0042F947: lea ecx, var_1C
  loc_0042F94A: call edi
  loc_0042F94C: push eax
  loc_0042F94D: push 00412E78h ; " and use this relationship to infer about"
  loc_0042F952: call __vbaStrCat
  loc_0042F954: mov edx, eax
  loc_0042F956: lea ecx, var_20
  loc_0042F959: call edi
  loc_0042F95B: mov edx, [00430010h]
  loc_0042F961: push eax
  loc_0042F962: push edx
  loc_0042F963: call __vbaStrCat
  loc_0042F965: mov edx, eax
  loc_0042F967: lea ecx, var_24
  loc_0042F96A: call edi
  loc_0042F96C: push eax
  loc_0042F96D: push 0040DD3Ch ; "."
  loc_0042F972: call __vbaStrCat
  loc_0042F974: mov edx, eax
  loc_0042F976: lea ecx, var_28
  loc_0042F979: call edi
  loc_0042F97B: mov esi, var_30
  loc_0042F97E: push eax
  loc_0042F97F: push esi
  loc_0042F980: call [ebx+00000054h]
  loc_0042F983: test eax, eax
  loc_0042F985: fnclex
  loc_0042F987: jge 0042F998h
  loc_0042F989: push 00000054h
  loc_0042F98B: push 0040E390h
  loc_0042F990: push esi
  loc_0042F991: push eax
  loc_0042F992: call [00401030h] ; __vbaHresultCheckObj
  loc_0042F998: lea eax, var_28
  loc_0042F99B: lea ecx, var_24
  loc_0042F99E: push eax
  loc_0042F99F: lea edx, var_20
  loc_0042F9A2: push ecx
  loc_0042F9A3: lea eax, var_1C
  loc_0042F9A6: push edx
  loc_0042F9A7: lea ecx, var_18
  loc_0042F9AA: push eax
  loc_0042F9AB: lea edx, var_14
  loc_0042F9AE: push ecx
  loc_0042F9AF: push edx
  loc_0042F9B0: push 00000006h
  loc_0042F9B2: call [004010E4h] ; __vbaFreeStrList
  loc_0042F9B8: add esp, 0000001Ch
  loc_0042F9BB: lea ecx, var_2C
  loc_0042F9BE: call [00401114h] ; __vbaFreeObj
  loc_0042F9C4: push 0042F9F9h
  loc_0042F9C9: jmp 0042F9F8h
  loc_0042F9CB: lea eax, var_28
  loc_0042F9CE: lea ecx, var_24
  loc_0042F9D1: push eax
  loc_0042F9D2: lea edx, var_20
  loc_0042F9D5: push ecx
  loc_0042F9D6: lea eax, var_1C
  loc_0042F9D9: push edx
  loc_0042F9DA: lea ecx, var_18
  loc_0042F9DD: push eax
  loc_0042F9DE: lea edx, var_14
  loc_0042F9E1: push ecx
  loc_0042F9E2: push edx
  loc_0042F9E3: push 00000006h
  loc_0042F9E5: call [004010E4h] ; __vbaFreeStrList
  loc_0042F9EB: add esp, 0000001Ch
  loc_0042F9EE: lea ecx, var_2C
  loc_0042F9F1: call [00401114h] ; __vbaFreeObj
  loc_0042F9F7: ret
  loc_0042F9F8: ret
  loc_0042F9F9: mov ecx, var_10
  loc_0042F9FC: pop edi
  loc_0042F9FD: pop esi
  loc_0042F9FE: xor eax, eax
  loc_0042FA00: mov fs:[00000000h], ecx
  loc_0042FA07: pop ebx
  loc_0042FA08: mov esp, ebp
  loc_0042FA0A: pop ebp
  loc_0042FA0B: retn 0004h
End Sub
