VERSION 5.00
Begin VB.Form frmChangeQuantiles
  Caption = "Change Alpha or T-Table "
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 5580
  ClientHeight = 4875
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Distribution Values"
    Left = 0
    Top = 0
    Width = 5445
    Height = 4815
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
    Begin VB.Frame Frame3
      Caption = "One-Sided T-Table Value"
      Left = 360
      Top = 2610
      Width = 4755
      Height = 1035
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
      Begin VB.TextBox txt1SidedT
        Left = 240
        Top = 420
        Width = 1845
        Height = 495
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
      End
    End
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 360
      Top = 3870
      Width = 855
      Height = 705
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
    End
    Begin VB.CommandButton cmdClearChange
      Caption = "Clear"
      Left = 1320
      Top = 3870
      Width = 1095
      Height = 705
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
    Begin VB.CommandButton cmdRestore
      Caption = "Reset Defaults"
      Left = 2520
      Top = 3870
      Width = 2655
      Height = 705
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
    End
    Begin VB.Frame Frame1
      Caption = "Value of Alpha"
      Left = 360
      Top = 600
      Width = 4755
      Height = 945
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
      Begin VB.TextBox txtAlpha
        Left = 270
        Top = 420
        Width = 1815
        Height = 435
        Text = "1."
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
    Begin VB.Frame Frame2
      Caption = "Two-Sided T-Table Value"
      Left = 360
      Top = 1590
      Width = 4755
      Height = 975
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
      Begin VB.TextBox txtTtable
        Left = 240
        Top = 420
        Width = 1815
        Height = 465
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
End

Attribute VB_Name = "frmChangeQuantiles"


Private Sub cmdChange_Click() '420900
  loc_00420900: push ebp
  loc_00420901: mov ebp, esp
  loc_00420903: sub esp, 0000000Ch
  loc_00420906: push 00401926h ; __vbaExceptHandler
  loc_0042090B: mov eax, fs:[00000000h]
  loc_00420911: push eax
  loc_00420912: mov fs:[00000000h], esp
  loc_00420919: sub esp, 000000A8h
  loc_0042091F: push ebx
  loc_00420920: push esi
  loc_00420921: push edi
  loc_00420922: mov var_C, esp
  loc_00420925: mov var_8, 004013D0h
  loc_0042092C: mov esi, Me
  loc_0042092F: mov eax, esi
  loc_00420931: and eax, 00000001h
  loc_00420934: mov var_4, eax
  loc_00420937: and esi, FFFFFFFEh
  loc_0042093A: push esi
  loc_0042093B: mov Me, esi
  loc_0042093E: mov ecx, [esi]
  loc_00420940: call [ecx+00000004h]
  loc_00420943: mov edx, [esi]
  loc_00420945: xor eax, eax
  loc_00420947: push esi
  loc_00420948: mov var_18, eax
  loc_0042094B: mov var_1C, eax
  loc_0042094E: mov var_2C, eax
  loc_00420951: mov var_3C, eax
  loc_00420954: mov var_4C, eax
  loc_00420957: mov var_5C, eax
  loc_0042095A: mov var_6C, eax
  loc_0042095D: mov var_7C, eax
  loc_00420960: call [edx+00000318h]
  loc_00420966: mov ebx, [00401050h] ; rtcTrimVar
  loc_0042096C: mov var_24, eax
  loc_0042096F: lea eax, var_2C
  loc_00420972: lea ecx, var_3C
  loc_00420975: push eax
  loc_00420976: push ecx
  loc_00420977: mov var_2C, 00000009h
  loc_0042097E: call ebx
  loc_00420980: lea edx, var_3C
  loc_00420983: lea eax, var_6C
  loc_00420986: push edx
  loc_00420987: push eax
  loc_00420988: mov var_64, 0040F040h
  loc_0042098F: mov var_6C, 00008008h
  loc_00420996: call [00401070h] ; __vbaVarTstEq
  loc_0042099C: mov edi, [00401018h] ; __vbaFreeVarList
  loc_004209A2: lea ecx, var_3C
  loc_004209A5: lea edx, var_2C
  loc_004209A8: push ecx
  loc_004209A9: push edx
  loc_004209AA: push 00000002h
  loc_004209AC: mov var_A0, ax
  loc_004209B3: call edi
  loc_004209B5: add esp, 0000000Ch
  loc_004209B8: cmp var_A0, 0000h
  loc_004209C0: jz 00420A41h
  loc_004209C2: mov eax, [esi]
  loc_004209C4: push esi
  loc_004209C5: call [eax+00000318h]
  loc_004209CB: lea ecx, var_1C
  loc_004209CE: push eax
  loc_004209CF: push ecx
  loc_004209D0: call [0040103Ch] ; __vbaObjSet
  loc_004209D6: mov esi, eax
  loc_004209D8: push esi
  loc_004209D9: mov edx, [esi]
  loc_004209DB: call [edx+00000204h]
  loc_004209E1: test eax, eax
  loc_004209E3: fnclex
  loc_004209E5: jge 004209F9h
  loc_004209E7: push 00000204h
  loc_004209EC: push 0040F02Ch
  loc_004209F1: push esi
  loc_004209F2: push eax
  loc_004209F3: call [00401030h] ; __vbaHresultCheckObj
  loc_004209F9: lea ecx, var_1C
  loc_004209FC: call [00401114h] ; __vbaFreeObj
  loc_00420A02: mov esi, [004010F4h] ; __vbaVarDup
  loc_00420A08: mov ecx, 80020004h
  loc_00420A0D: mov var_54, ecx
  loc_00420A10: mov eax, 0000000Ah
  loc_00420A15: mov var_44, ecx
  loc_00420A18: mov ebx, 00000008h
  loc_00420A1D: lea edx, var_7C
  loc_00420A20: lea ecx, var_3C
  loc_00420A23: mov var_5C, eax
  loc_00420A26: mov var_4C, eax
  loc_00420A29: mov var_74, 00410138h ; "Check alpha"
  loc_00420A30: mov var_7C, ebx
  loc_00420A33: call __vbaVarDup
  loc_00420A35: mov var_64, 004100F4h ; "Please enter a value for alpha."
  loc_00420A3C: jmp 00420B18h
  loc_00420A41: mov edx, [esi]
  loc_00420A43: push esi
  loc_00420A44: call [edx+00000304h]
  loc_00420A4A: mov var_24, eax
  loc_00420A4D: lea eax, var_2C
  loc_00420A50: lea ecx, var_3C
  loc_00420A53: push eax
  loc_00420A54: push ecx
  loc_00420A55: mov var_2C, 00000009h
  loc_00420A5C: call ebx
  loc_00420A5E: lea edx, var_3C
  loc_00420A61: lea eax, var_6C
  loc_00420A64: push edx
  loc_00420A65: push eax
  loc_00420A66: mov var_64, 0040F040h
  loc_00420A6D: mov var_6C, 00008008h
  loc_00420A74: call [00401070h] ; __vbaVarTstEq
  loc_00420A7A: lea ecx, var_3C
  loc_00420A7D: lea edx, var_2C
  loc_00420A80: push ecx
  loc_00420A81: push edx
  loc_00420A82: push 00000002h
  loc_00420A84: mov var_A0, ax
  loc_00420A8B: call edi
  loc_00420A8D: add esp, 0000000Ch
  loc_00420A90: cmp var_A0, 0000h
  loc_00420A98: jz 00420B57h
  loc_00420A9E: mov eax, [esi]
  loc_00420AA0: push esi
  loc_00420AA1: call [eax+00000304h]
  loc_00420AA7: lea ecx, var_1C
  loc_00420AAA: push eax
  loc_00420AAB: push ecx
  loc_00420AAC: call [0040103Ch] ; __vbaObjSet
  loc_00420AB2: mov esi, eax
  loc_00420AB4: push esi
  loc_00420AB5: mov edx, [esi]
  loc_00420AB7: call [edx+00000204h]
  loc_00420ABD: test eax, eax
  loc_00420ABF: fnclex
  loc_00420AC1: jge 00420AD5h
  loc_00420AC3: push 00000204h
  loc_00420AC8: push 0040F02Ch
  loc_00420ACD: push esi
  loc_00420ACE: push eax
  loc_00420ACF: call [00401030h] ; __vbaHresultCheckObj
  loc_00420AD5: lea ecx, var_1C
  loc_00420AD8: call [00401114h] ; __vbaFreeObj
  loc_00420ADE: mov esi, [004010F4h] ; __vbaVarDup
  loc_00420AE4: mov ecx, 80020004h
  loc_00420AE9: mov var_54, ecx
  loc_00420AEC: mov eax, 0000000Ah
  loc_00420AF1: mov var_44, ecx
  loc_00420AF4: mov ebx, 00000008h
  loc_00420AF9: lea edx, var_7C
  loc_00420AFC: lea ecx, var_3C
  loc_00420AFF: mov var_5C, eax
  loc_00420B02: mov var_4C, eax
  loc_00420B05: mov var_74, 0040F100h ; "Check Name"
  loc_00420B0C: mov var_7C, ebx
  loc_00420B0F: call __vbaVarDup
  loc_00420B11: mov var_64, 00410154h ; "Please enter a value from the t-table for a right-sided test."
  loc_00420B18: lea edx, var_6C
  loc_00420B1B: lea ecx, var_2C
  loc_00420B1E: mov var_6C, ebx
  loc_00420B21: call __vbaVarDup
  loc_00420B23: lea eax, var_5C
  loc_00420B26: lea ecx, var_4C
  loc_00420B29: push eax
  loc_00420B2A: lea edx, var_3C
  loc_00420B2D: push ecx
  loc_00420B2E: push edx
  loc_00420B2F: lea eax, var_2C
  loc_00420B32: push 00000000h
  loc_00420B34: push eax
  loc_00420B35: call [00401038h] ; rtcMsgBox
  loc_00420B3B: lea ecx, var_5C
  loc_00420B3E: lea edx, var_4C
  loc_00420B41: push ecx
  loc_00420B42: lea eax, var_3C
  loc_00420B45: push edx
  loc_00420B46: lea ecx, var_2C
  loc_00420B49: push eax
  loc_00420B4A: push ecx
  loc_00420B4B: push 00000004h
  loc_00420B4D: call edi
  loc_00420B4F: add esp, 00000014h
  loc_00420B52: jmp 00420FBCh
  loc_00420B57: mov edx, [esi]
  loc_00420B59: push esi
  loc_00420B5A: call [edx+00000320h]
  loc_00420B60: mov var_24, eax
  loc_00420B63: lea eax, var_2C
  loc_00420B66: lea ecx, var_3C
  loc_00420B69: push eax
  loc_00420B6A: push ecx
  loc_00420B6B: mov var_2C, 00000009h
  loc_00420B72: call ebx
  loc_00420B74: lea edx, var_3C
  loc_00420B77: lea eax, var_6C
  loc_00420B7A: push edx
  loc_00420B7B: push eax
  loc_00420B7C: mov var_64, 0040F040h
  loc_00420B83: mov var_6C, 00008008h
  loc_00420B8A: call [00401070h] ; __vbaVarTstEq
  loc_00420B90: lea ecx, var_3C
  loc_00420B93: lea edx, var_2C
  loc_00420B96: push ecx
  loc_00420B97: push edx
  loc_00420B98: push 00000002h
  loc_00420B9A: mov var_A0, ax
  loc_00420BA1: call edi
  loc_00420BA3: add esp, 0000000Ch
  loc_00420BA6: cmp var_A0, 0000h
  loc_00420BAE: jz 00420C70h
  loc_00420BB4: mov ebx, [004010F4h] ; __vbaVarDup
  loc_00420BBA: mov ecx, 80020004h
  loc_00420BBF: mov var_54, ecx
  loc_00420BC2: mov eax, 0000000Ah
  loc_00420BC7: mov var_44, ecx
  loc_00420BCA: lea edx, var_7C
  loc_00420BCD: lea ecx, var_3C
  loc_00420BD0: mov var_5C, eax
  loc_00420BD3: mov var_4C, eax
  loc_00420BD6: mov var_74, 00410280h ; "Check ttable"
  loc_00420BDD: mov var_7C, 00000008h
  loc_00420BE4: call ebx
  loc_00420BE6: lea edx, var_6C
  loc_00420BE9: lea ecx, var_2C
  loc_00420BEC: mov var_64, 004101D4h ; "Please enter a value from the t-table for a two-sided test or confidence interval."
  loc_00420BF3: mov var_6C, 00000008h
  loc_00420BFA: call ebx
  loc_00420BFC: lea eax, var_5C
  loc_00420BFF: lea ecx, var_4C
  loc_00420C02: push eax
  loc_00420C03: lea edx, var_3C
  loc_00420C06: push ecx
  loc_00420C07: push edx
  loc_00420C08: lea eax, var_2C
  loc_00420C0B: push 00000000h
  loc_00420C0D: push eax
  loc_00420C0E: call [00401038h] ; rtcMsgBox
  loc_00420C14: lea ecx, var_5C
  loc_00420C17: lea edx, var_4C
  loc_00420C1A: push ecx
  loc_00420C1B: lea eax, var_3C
  loc_00420C1E: push edx
  loc_00420C1F: lea ecx, var_2C
  loc_00420C22: push eax
  loc_00420C23: push ecx
  loc_00420C24: push 00000004h
  loc_00420C26: call edi
  loc_00420C28: mov edx, [esi]
  loc_00420C2A: add esp, 00000014h
  loc_00420C2D: push esi
  loc_00420C2E: call [edx+00000320h]
  loc_00420C34: push eax
  loc_00420C35: lea eax, var_1C
  loc_00420C38: push eax
  loc_00420C39: call [0040103Ch] ; __vbaObjSet
  loc_00420C3F: mov esi, eax
  loc_00420C41: push esi
  loc_00420C42: mov ecx, [esi]
  loc_00420C44: call [ecx+00000204h]
  loc_00420C4A: test eax, eax
  loc_00420C4C: fnclex
  loc_00420C4E: jge 00420C62h
  loc_00420C50: push 00000204h
  loc_00420C55: push 0040F02Ch
  loc_00420C5A: push esi
  loc_00420C5B: push eax
  loc_00420C5C: call [00401030h] ; __vbaHresultCheckObj
  loc_00420C62: lea ecx, var_1C
  loc_00420C65: call [00401114h] ; __vbaFreeObj
  loc_00420C6B: jmp 00420FBCh
  loc_00420C70: mov edx, [esi]
  loc_00420C72: push esi
  loc_00420C73: call [edx+00000318h]
  loc_00420C79: mov var_24, eax
  loc_00420C7C: lea eax, var_2C
  loc_00420C7F: lea ecx, var_3C
  loc_00420C82: push eax
  loc_00420C83: push ecx
  loc_00420C84: mov var_2C, 00000009h
  loc_00420C8B: call ebx
  loc_00420C8D: lea edx, var_3C
  loc_00420C90: push edx
  loc_00420C91: call [004010A0h] ; __vbaR4ErrVar
  loc_00420C97: fcomp real4 ptr [004013C8h]
  loc_00420C9D: mov var_BC, 00000001h
  loc_00420CA7: fnstsw ax
  loc_00420CA9: test ah, 41h
  loc_00420CAC: jz 00420CB8h
  loc_00420CAE: mov var_BC, 00000000h
  loc_00420CB8: lea eax, var_3C
  loc_00420CBB: lea ecx, var_3C
  loc_00420CBE: push eax
  loc_00420CBF: lea edx, var_2C
  loc_00420CC2: push ecx
  loc_00420CC3: push edx
  loc_00420CC4: push 00000003h
  loc_00420CC6: call edi
  loc_00420CC8: mov eax, var_BC
  loc_00420CCE: add esp, 00000010h
  loc_00420CD1: neg eax
  loc_00420CD3: test ax, ax
  loc_00420CD6: jz 00420DD9h
  loc_00420CDC: mov ebx, [004010F4h] ; __vbaVarDup
  loc_00420CE2: mov ecx, 80020004h
  loc_00420CE7: mov var_54, ecx
  loc_00420CEA: mov eax, 0000000Ah
  loc_00420CEF: mov var_44, ecx
  loc_00420CF2: lea edx, var_7C
  loc_00420CF5: lea ecx, var_3C
  loc_00420CF8: mov var_5C, eax
  loc_00420CFB: mov var_4C, eax
  loc_00420CFE: mov var_74, 0041031Ch ; "Check Alpha"
  loc_00420D05: mov var_7C, 00000008h
  loc_00420D0C: call ebx
  loc_00420D0E: lea edx, var_6C
  loc_00420D11: lea ecx, var_2C
  loc_00420D14: mov var_64, 004102D0h ; "Alpha can not be greater than 1.00"
  loc_00420D1B: mov var_6C, 00000008h
  loc_00420D22: call ebx
  loc_00420D24: lea eax, var_5C
  loc_00420D27: lea ecx, var_4C
  loc_00420D2A: push eax
  loc_00420D2B: lea edx, var_3C
  loc_00420D2E: push ecx
  loc_00420D2F: push edx
  loc_00420D30: lea eax, var_2C
  loc_00420D33: push 00000000h
  loc_00420D35: push eax
  loc_00420D36: call [00401038h] ; rtcMsgBox
  loc_00420D3C: lea ecx, var_5C
  loc_00420D3F: lea edx, var_4C
  loc_00420D42: push ecx
  loc_00420D43: lea eax, var_3C
  loc_00420D46: push edx
  loc_00420D47: lea ecx, var_2C
  loc_00420D4A: push eax
  loc_00420D4B: push ecx
  loc_00420D4C: push 00000004h
  loc_00420D4E: call edi
  loc_00420D50: mov edx, [esi]
  loc_00420D52: add esp, 00000014h
  loc_00420D55: push esi
  loc_00420D56: call [edx+00000318h]
  loc_00420D5C: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_00420D62: push eax
  loc_00420D63: lea eax, var_1C
  loc_00420D66: push eax
  loc_00420D67: call ebx
  loc_00420D69: mov edi, eax
  loc_00420D6B: push 0040F040h
  loc_00420D70: push edi
  loc_00420D71: mov ecx, [edi]
  loc_00420D73: call [ecx+000000A4h]
  loc_00420D79: test eax, eax
  loc_00420D7B: fnclex
  loc_00420D7D: jge 00420D91h
  loc_00420D7F: push 000000A4h
  loc_00420D84: push 0040F02Ch
  loc_00420D89: push edi
  loc_00420D8A: push eax
  loc_00420D8B: call [00401030h] ; __vbaHresultCheckObj
  loc_00420D91: mov edi, [00401114h] ; __vbaFreeObj
  loc_00420D97: lea ecx, var_1C
  loc_00420D9A: call edi
  loc_00420D9C: mov edx, [esi]
  loc_00420D9E: push esi
  loc_00420D9F: call [edx+00000318h]
  loc_00420DA5: push eax
  loc_00420DA6: lea eax, var_1C
  loc_00420DA9: push eax
  loc_00420DAA: call ebx
  loc_00420DAC: mov esi, eax
  loc_00420DAE: push esi
  loc_00420DAF: mov ecx, [esi]
  loc_00420DB1: call [ecx+00000204h]
  loc_00420DB7: test eax, eax
  loc_00420DB9: fnclex
  loc_00420DBB: jge 00420DCFh
  loc_00420DBD: push 00000204h
  loc_00420DC2: push 0040F02Ch
  loc_00420DC7: push esi
  loc_00420DC8: push eax
  loc_00420DC9: call [00401030h] ; __vbaHresultCheckObj
  loc_00420DCF: lea ecx, var_1C
  loc_00420DD2: call edi
  loc_00420DD4: jmp 00420FBCh
  loc_00420DD9: mov edx, [esi]
  loc_00420DDB: push esi
  loc_00420DDC: call [edx+00000318h]
  loc_00420DE2: push eax
  loc_00420DE3: lea eax, var_1C
  loc_00420DE6: push eax
  loc_00420DE7: call [0040103Ch] ; __vbaObjSet
  loc_00420DED: mov ecx, [eax]
  loc_00420DEF: lea edx, var_18
  loc_00420DF2: push edx
  loc_00420DF3: push eax
  loc_00420DF4: mov var_A0, eax
  loc_00420DFA: call [ecx+000000A0h]
  loc_00420E00: test eax, eax
  loc_00420E02: fnclex
  loc_00420E04: jge 00420E1Eh
  loc_00420E06: mov ecx, var_A0
  loc_00420E0C: push 000000A0h
  loc_00420E11: push 0040F02Ch
  loc_00420E16: push ecx
  loc_00420E17: push eax
  loc_00420E18: call [00401030h] ; __vbaHresultCheckObj
  loc_00420E1E: mov eax, var_18
  loc_00420E21: lea edx, var_2C
  loc_00420E24: mov var_24, eax
  loc_00420E27: lea eax, var_3C
  loc_00420E2A: push edx
  loc_00420E2B: push eax
  loc_00420E2C: mov var_18, 00000000h
  loc_00420E33: mov var_2C, 00000008h
  loc_00420E3A: call ebx
  loc_00420E3C: lea ecx, var_3C
  loc_00420E3F: push ecx
  loc_00420E40: call [004010A0h] ; __vbaR4ErrVar
  loc_00420E46: fstp real4 ptr [00430020h]
  loc_00420E4C: lea ecx, var_1C
  loc_00420E4F: call [00401114h] ; __vbaFreeObj
  loc_00420E55: lea edx, var_3C
  loc_00420E58: lea eax, var_3C
  loc_00420E5B: push edx
  loc_00420E5C: lea ecx, var_2C
  loc_00420E5F: push eax
  loc_00420E60: push ecx
  loc_00420E61: push 00000003h
  loc_00420E63: call edi
  loc_00420E65: mov edx, [esi]
  loc_00420E67: add esp, 00000010h
  loc_00420E6A: push esi
  loc_00420E6B: call [edx+00000320h]
  loc_00420E71: push eax
  loc_00420E72: lea eax, var_1C
  loc_00420E75: push eax
  loc_00420E76: call [0040103Ch] ; __vbaObjSet
  loc_00420E7C: mov ecx, [eax]
  loc_00420E7E: lea edx, var_18
  loc_00420E81: push edx
  loc_00420E82: push eax
  loc_00420E83: mov var_A0, eax
  loc_00420E89: call [ecx+000000A0h]
  loc_00420E8F: test eax, eax
  loc_00420E91: fnclex
  loc_00420E93: jge 00420EADh
  loc_00420E95: mov ecx, var_A0
  loc_00420E9B: push 000000A0h
  loc_00420EA0: push 0040F02Ch
  loc_00420EA5: push ecx
  loc_00420EA6: push eax
  loc_00420EA7: call [00401030h] ; __vbaHresultCheckObj
  loc_00420EAD: mov eax, var_18
  loc_00420EB0: lea edx, var_2C
  loc_00420EB3: mov var_24, eax
  loc_00420EB6: lea eax, var_3C
  loc_00420EB9: push edx
  loc_00420EBA: push eax
  loc_00420EBB: mov var_18, 00000000h
  loc_00420EC2: mov var_2C, 00000008h
  loc_00420EC9: call ebx
  loc_00420ECB: lea ecx, var_3C
  loc_00420ECE: push ecx
  loc_00420ECF: call [004010A0h] ; __vbaR4ErrVar
  loc_00420ED5: fstp real4 ptr [00430028h]
  loc_00420EDB: lea ecx, var_1C
  loc_00420EDE: call [00401114h] ; __vbaFreeObj
  loc_00420EE4: lea edx, var_3C
  loc_00420EE7: lea eax, var_3C
  loc_00420EEA: push edx
  loc_00420EEB: lea ecx, var_2C
  loc_00420EEE: push eax
  loc_00420EEF: push ecx
  loc_00420EF0: push 00000003h
  loc_00420EF2: call edi
  loc_00420EF4: mov edx, [esi]
  loc_00420EF6: add esp, 00000010h
  loc_00420EF9: push esi
  loc_00420EFA: call [edx+00000304h]
  loc_00420F00: push eax
  loc_00420F01: lea eax, var_1C
  loc_00420F04: push eax
  loc_00420F05: call [0040103Ch] ; __vbaObjSet
  loc_00420F0B: mov esi, eax
  loc_00420F0D: lea edx, var_18
  loc_00420F10: push edx
  loc_00420F11: push esi
  loc_00420F12: mov ecx, [esi]
  loc_00420F14: call [ecx+000000A0h]
  loc_00420F1A: test eax, eax
  loc_00420F1C: fnclex
  loc_00420F1E: jge 00420F32h
  loc_00420F20: push 000000A0h
  loc_00420F25: push 0040F02Ch
  loc_00420F2A: push esi
  loc_00420F2B: push eax
  loc_00420F2C: call [00401030h] ; __vbaHresultCheckObj
  loc_00420F32: mov eax, var_18
  loc_00420F35: lea ecx, var_3C
  loc_00420F38: mov var_24, eax
  loc_00420F3B: lea eax, var_2C
  loc_00420F3E: push eax
  loc_00420F3F: push ecx
  loc_00420F40: mov var_18, 00000000h
  loc_00420F47: mov var_2C, 00000008h
  loc_00420F4E: call ebx
  loc_00420F50: lea edx, var_3C
  loc_00420F53: push edx
  loc_00420F54: call [004010A0h] ; __vbaR4ErrVar
  loc_00420F5A: fstp real4 ptr [00430024h]
  loc_00420F60: lea ecx, var_1C
  loc_00420F63: call [00401114h] ; __vbaFreeObj
  loc_00420F69: lea eax, var_3C
  loc_00420F6C: lea ecx, var_3C
  loc_00420F6F: push eax
  loc_00420F70: lea edx, var_2C
  loc_00420F73: push ecx
  loc_00420F74: push edx
  loc_00420F75: push 00000003h
  loc_00420F77: call edi
  loc_00420F79: mov eax, [00430178h]
  loc_00420F7E: add esp, 00000010h
  loc_00420F81: test eax, eax
  loc_00420F83: jnz 00420F95h
  loc_00420F85: push 00430178h
  loc_00420F8A: push 004051D0h
  loc_00420F8F: call [004010D4h] ; __vbaNew2
  loc_00420F95: mov esi, [00430178h]
  loc_00420F9B: push esi
  loc_00420F9C: mov eax, [esi]
  loc_00420F9E: call [eax+000002B4h]
  loc_00420FA4: test eax, eax
  loc_00420FA6: fnclex
  loc_00420FA8: jge 00420FBCh
  loc_00420FAA: push 000002B4h
  loc_00420FAF: push 004100A4h
  loc_00420FB4: push esi
  loc_00420FB5: push eax
  loc_00420FB6: call [00401030h] ; __vbaHresultCheckObj
  loc_00420FBC: mov var_4, 00000000h
  loc_00420FC3: fwait
  loc_00420FC4: push 00420FFAh
  loc_00420FC9: jmp 00420FF9h
  loc_00420FCB: lea ecx, var_18
  loc_00420FCE: call [00401110h] ; __vbaFreeStr
  loc_00420FD4: lea ecx, var_1C
  loc_00420FD7: call [00401114h] ; __vbaFreeObj
  loc_00420FDD: lea ecx, var_5C
  loc_00420FE0: lea edx, var_4C
  loc_00420FE3: push ecx
  loc_00420FE4: lea eax, var_3C
  loc_00420FE7: push edx
  loc_00420FE8: lea ecx, var_2C
  loc_00420FEB: push eax
  loc_00420FEC: push ecx
  loc_00420FED: push 00000004h
  loc_00420FEF: call [00401018h] ; __vbaFreeVarList
  loc_00420FF5: add esp, 00000014h
  loc_00420FF8: ret
  loc_00420FF9: ret
  loc_00420FFA: mov eax, Me
  loc_00420FFD: push eax
  loc_00420FFE: mov edx, [eax]
  loc_00421000: call [edx+00000008h]
  loc_00421003: mov eax, var_4
  loc_00421006: mov ecx, var_14
  loc_00421009: pop edi
  loc_0042100A: pop esi
  loc_0042100B: mov fs:[00000000h], ecx
  loc_00421012: pop ebx
  loc_00421013: mov esp, ebp
  loc_00421015: pop ebp
  loc_00421016: retn 0004h
End Sub

Private Sub Form_Activate() '421020
  loc_00421020: push ebp
  loc_00421021: mov ebp, esp
  loc_00421023: sub esp, 0000000Ch
  loc_00421026: push 00401926h ; __vbaExceptHandler
  loc_0042102B: mov eax, fs:[00000000h]
  loc_00421031: push eax
  loc_00421032: mov fs:[00000000h], esp
  loc_00421039: sub esp, 00000018h
  loc_0042103C: push ebx
  loc_0042103D: push esi
  loc_0042103E: push edi
  loc_0042103F: mov var_C, esp
  loc_00421042: mov var_8, 004013E0h
  loc_00421049: mov esi, Me
  loc_0042104C: mov eax, esi
  loc_0042104E: and eax, 00000001h
  loc_00421051: mov var_4, eax
  loc_00421054: and esi, FFFFFFFEh
  loc_00421057: push esi
  loc_00421058: mov Me, esi
  loc_0042105B: mov ecx, [esi]
  loc_0042105D: call [ecx+00000004h]
  loc_00421060: mov edx, [esi]
  loc_00421062: xor eax, eax
  loc_00421064: push esi
  loc_00421065: mov var_18, eax
  loc_00421068: mov var_1C, eax
  loc_0042106B: call [edx+00000318h]
  loc_00421071: push eax
  loc_00421072: lea eax, var_1C
  loc_00421075: push eax
  loc_00421076: call [0040103Ch] ; __vbaObjSet
  loc_0042107C: mov ecx, [00430020h]
  loc_00421082: mov edi, eax
  loc_00421084: push ecx
  loc_00421085: mov ebx, [edi]
  loc_00421087: call [0040107Ch] ; __vbaStrR4
  loc_0042108D: mov edx, eax
  loc_0042108F: lea ecx, var_18
  loc_00421092: call [004010FCh] ; __vbaStrMove
  loc_00421098: push eax
  loc_00421099: push edi
  loc_0042109A: call [ebx+000000A4h]
  loc_004210A0: test eax, eax
  loc_004210A2: fnclex
  loc_004210A4: jge 004210B8h
  loc_004210A6: push 000000A4h
  loc_004210AB: push 0040F02Ch
  loc_004210B0: push edi
  loc_004210B1: push eax
  loc_004210B2: call [00401030h] ; __vbaHresultCheckObj
  loc_004210B8: lea ecx, var_18
  loc_004210BB: call [00401110h] ; __vbaFreeStr
  loc_004210C1: lea ecx, var_1C
  loc_004210C4: call [00401114h] ; __vbaFreeObj
  loc_004210CA: mov edx, [esi]
  loc_004210CC: push esi
  loc_004210CD: call [edx+00000320h]
  loc_004210D3: push eax
  loc_004210D4: lea eax, var_1C
  loc_004210D7: push eax
  loc_004210D8: call [0040103Ch] ; __vbaObjSet
  loc_004210DE: mov ecx, [00430028h]
  loc_004210E4: mov edi, eax
  loc_004210E6: push ecx
  loc_004210E7: mov ebx, [edi]
  loc_004210E9: call [0040107Ch] ; __vbaStrR4
  loc_004210EF: mov edx, eax
  loc_004210F1: lea ecx, var_18
  loc_004210F4: call [004010FCh] ; __vbaStrMove
  loc_004210FA: push eax
  loc_004210FB: push edi
  loc_004210FC: call [ebx+000000A4h]
  loc_00421102: test eax, eax
  loc_00421104: fnclex
  loc_00421106: jge 0042111Ah
  loc_00421108: push 000000A4h
  loc_0042110D: push 0040F02Ch
  loc_00421112: push edi
  loc_00421113: push eax
  loc_00421114: call [00401030h] ; __vbaHresultCheckObj
  loc_0042111A: lea ecx, var_18
  loc_0042111D: call [00401110h] ; __vbaFreeStr
  loc_00421123: lea ecx, var_1C
  loc_00421126: call [00401114h] ; __vbaFreeObj
  loc_0042112C: mov edx, [esi]
  loc_0042112E: push esi
  loc_0042112F: call [edx+00000304h]
  loc_00421135: push eax
  loc_00421136: lea eax, var_1C
  loc_00421139: push eax
  loc_0042113A: call [0040103Ch] ; __vbaObjSet
  loc_00421140: mov ecx, [00430024h]
  loc_00421146: mov edi, eax
  loc_00421148: push ecx
  loc_00421149: mov ebx, [edi]
  loc_0042114B: call [0040107Ch] ; __vbaStrR4
  loc_00421151: mov edx, eax
  loc_00421153: lea ecx, var_18
  loc_00421156: call [004010FCh] ; __vbaStrMove
  loc_0042115C: push eax
  loc_0042115D: push edi
  loc_0042115E: call [ebx+000000A4h]
  loc_00421164: test eax, eax
  loc_00421166: fnclex
  loc_00421168: jge 0042117Ch
  loc_0042116A: push 000000A4h
  loc_0042116F: push 0040F02Ch
  loc_00421174: push edi
  loc_00421175: push eax
  loc_00421176: call [00401030h] ; __vbaHresultCheckObj
  loc_0042117C: lea ecx, var_18
  loc_0042117F: call [00401110h] ; __vbaFreeStr
  loc_00421185: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042118B: lea ecx, var_1C
  loc_0042118E: call edi
  loc_00421190: mov edx, [esi]
  loc_00421192: push esi
  loc_00421193: call [edx+00000318h]
  loc_00421199: push eax
  loc_0042119A: lea eax, var_1C
  loc_0042119D: push eax
  loc_0042119E: call [0040103Ch] ; __vbaObjSet
  loc_004211A4: mov esi, eax
  loc_004211A6: push esi
  loc_004211A7: mov ecx, [esi]
  loc_004211A9: call [ecx+00000204h]
  loc_004211AF: test eax, eax
  loc_004211B1: fnclex
  loc_004211B3: jge 004211C7h
  loc_004211B5: push 00000204h
  loc_004211BA: push 0040F02Ch
  loc_004211BF: push esi
  loc_004211C0: push eax
  loc_004211C1: call [00401030h] ; __vbaHresultCheckObj
  loc_004211C7: lea ecx, var_1C
  loc_004211CA: call edi
  loc_004211CC: mov var_4, 00000000h
  loc_004211D3: fwait
  loc_004211D4: push 004211EFh
  loc_004211D9: jmp 004211EEh
  loc_004211DB: lea ecx, var_18
  loc_004211DE: call [00401110h] ; __vbaFreeStr
  loc_004211E4: lea ecx, var_1C
  loc_004211E7: call [00401114h] ; __vbaFreeObj
  loc_004211ED: ret
  loc_004211EE: ret
  loc_004211EF: mov eax, Me
  loc_004211F2: push eax
  loc_004211F3: mov edx, [eax]
  loc_004211F5: call [edx+00000008h]
  loc_004211F8: mov eax, var_4
  loc_004211FB: mov ecx, var_14
  loc_004211FE: pop edi
  loc_004211FF: pop esi
  loc_00421200: mov fs:[00000000h], ecx
  loc_00421207: pop ebx
  loc_00421208: mov esp, ebp
  loc_0042120A: pop ebp
  loc_0042120B: retn 0004h
End Sub

Private Sub cmdRestore_Click() '420730
  loc_00420730: push ebp
  loc_00420731: mov ebp, esp
  loc_00420733: sub esp, 0000000Ch
  loc_00420736: push 00401926h ; __vbaExceptHandler
  loc_0042073B: mov eax, fs:[00000000h]
  loc_00420741: push eax
  loc_00420742: mov fs:[00000000h], esp
  loc_00420749: sub esp, 00000018h
  loc_0042074C: push ebx
  loc_0042074D: push esi
  loc_0042074E: push edi
  loc_0042074F: mov var_C, esp
  loc_00420752: mov var_8, 004013B8h
  loc_00420759: mov esi, Me
  loc_0042075C: mov eax, esi
  loc_0042075E: and eax, 00000001h
  loc_00420761: mov var_4, eax
  loc_00420764: and esi, FFFFFFFEh
  loc_00420767: push esi
  loc_00420768: mov Me, esi
  loc_0042076B: mov ecx, [esi]
  loc_0042076D: call [ecx+00000004h]
  loc_00420770: mov [00430020h], 3D4CCCCDh
  loc_0042077A: mov [00430028h], 3FFAE148h
  loc_00420784: mov [00430024h], 3FD28F5Ch
  loc_0042078E: mov edx, [esi]
  loc_00420790: xor eax, eax
  loc_00420792: push esi
  loc_00420793: mov var_18, eax
  loc_00420796: mov var_1C, eax
  loc_00420799: call [edx+00000318h]
  loc_0042079F: push eax
  loc_004207A0: lea eax, var_1C
  loc_004207A3: push eax
  loc_004207A4: call [0040103Ch] ; __vbaObjSet
  loc_004207AA: mov ecx, [00430020h]
  loc_004207B0: mov edi, eax
  loc_004207B2: push ecx
  loc_004207B3: mov ebx, [edi]
  loc_004207B5: call [0040107Ch] ; __vbaStrR4
  loc_004207BB: mov edx, eax
  loc_004207BD: lea ecx, var_18
  loc_004207C0: call [004010FCh] ; __vbaStrMove
  loc_004207C6: push eax
  loc_004207C7: push edi
  loc_004207C8: call [ebx+000000A4h]
  loc_004207CE: test eax, eax
  loc_004207D0: fnclex
  loc_004207D2: jge 004207E6h
  loc_004207D4: push 000000A4h
  loc_004207D9: push 0040F02Ch
  loc_004207DE: push edi
  loc_004207DF: push eax
  loc_004207E0: call [00401030h] ; __vbaHresultCheckObj
  loc_004207E6: lea ecx, var_18
  loc_004207E9: call [00401110h] ; __vbaFreeStr
  loc_004207EF: lea ecx, var_1C
  loc_004207F2: call [00401114h] ; __vbaFreeObj
  loc_004207F8: mov edx, [esi]
  loc_004207FA: push esi
  loc_004207FB: call [edx+00000320h]
  loc_00420801: push eax
  loc_00420802: lea eax, var_1C
  loc_00420805: push eax
  loc_00420806: call [0040103Ch] ; __vbaObjSet
  loc_0042080C: mov ecx, [00430028h]
  loc_00420812: mov edi, eax
  loc_00420814: push ecx
  loc_00420815: mov ebx, [edi]
  loc_00420817: call [0040107Ch] ; __vbaStrR4
  loc_0042081D: mov edx, eax
  loc_0042081F: lea ecx, var_18
  loc_00420822: call [004010FCh] ; __vbaStrMove
  loc_00420828: push eax
  loc_00420829: push edi
  loc_0042082A: call [ebx+000000A4h]
  loc_00420830: test eax, eax
  loc_00420832: fnclex
  loc_00420834: jge 00420848h
  loc_00420836: push 000000A4h
  loc_0042083B: push 0040F02Ch
  loc_00420840: push edi
  loc_00420841: push eax
  loc_00420842: call [00401030h] ; __vbaHresultCheckObj
  loc_00420848: mov edi, [00401110h] ; __vbaFreeStr
  loc_0042084E: lea ecx, var_18
  loc_00420851: call edi
  loc_00420853: lea ecx, var_1C
  loc_00420856: call [00401114h] ; __vbaFreeObj
  loc_0042085C: mov edx, [esi]
  loc_0042085E: push esi
  loc_0042085F: call [edx+00000304h]
  loc_00420865: push eax
  loc_00420866: lea eax, var_1C
  loc_00420869: push eax
  loc_0042086A: call [0040103Ch] ; __vbaObjSet
  loc_00420870: mov ecx, [00430024h]
  loc_00420876: mov esi, eax
  loc_00420878: push ecx
  loc_00420879: mov ebx, [esi]
  loc_0042087B: call [0040107Ch] ; __vbaStrR4
  loc_00420881: mov edx, eax
  loc_00420883: lea ecx, var_18
  loc_00420886: call [004010FCh] ; __vbaStrMove
  loc_0042088C: push eax
  loc_0042088D: push esi
  loc_0042088E: call [ebx+000000A4h]
  loc_00420894: test eax, eax
  loc_00420896: fnclex
  loc_00420898: jge 004208ACh
  loc_0042089A: push 000000A4h
  loc_0042089F: push 0040F02Ch
  loc_004208A4: push esi
  loc_004208A5: push eax
  loc_004208A6: call [00401030h] ; __vbaHresultCheckObj
  loc_004208AC: lea ecx, var_18
  loc_004208AF: call edi
  loc_004208B1: lea ecx, var_1C
  loc_004208B4: call [00401114h] ; __vbaFreeObj
  loc_004208BA: mov var_4, 00000000h
  loc_004208C1: fwait
  loc_004208C2: push 004208DDh
  loc_004208C7: jmp 004208DCh
  loc_004208C9: lea ecx, var_18
  loc_004208CC: call [00401110h] ; __vbaFreeStr
  loc_004208D2: lea ecx, var_1C
  loc_004208D5: call [00401114h] ; __vbaFreeObj
  loc_004208DB: ret
  loc_004208DC: ret
  loc_004208DD: mov eax, Me
  loc_004208E0: push eax
  loc_004208E1: mov edx, [eax]
  loc_004208E3: call [edx+00000008h]
  loc_004208E6: mov eax, var_4
  loc_004208E9: mov ecx, var_14
  loc_004208EC: pop edi
  loc_004208ED: pop esi
  loc_004208EE: mov fs:[00000000h], ecx
  loc_004208F5: pop ebx
  loc_004208F6: mov esp, ebp
  loc_004208F8: pop ebp
  loc_004208F9: retn 0004h
End Sub

Private Sub txtAlpha_KeyPress(KeyAscii As Integer) '421210
  loc_00421210: push ebp
  loc_00421211: mov ebp, esp
  loc_00421213: sub esp, 0000000Ch
  loc_00421216: push 00401926h ; __vbaExceptHandler
  loc_0042121B: mov eax, fs:[00000000h]
  loc_00421221: push eax
  loc_00421222: mov fs:[00000000h], esp
  loc_00421229: sub esp, 0000008Ch
  loc_0042122F: push ebx
  loc_00421230: push esi
  loc_00421231: push edi
  loc_00421232: mov var_C, esp
  loc_00421235: mov var_8, 004013F0h
  loc_0042123C: mov esi, Me
  loc_0042123F: mov eax, esi
  loc_00421241: and eax, 00000001h
  loc_00421244: mov var_4, eax
  loc_00421247: and esi, FFFFFFFEh
  loc_0042124A: push esi
  loc_0042124B: mov Me, esi
  loc_0042124E: mov ecx, [esi]
  loc_00421250: call [ecx+00000004h]
  loc_00421253: mov edi, KeyAscii
  loc_00421256: lea edx, var_24
  loc_00421259: xor ebx, ebx
  loc_0042125B: push edi
  loc_0042125C: push edx
  loc_0042125D: mov var_24, ebx
  loc_00421260: mov var_34, ebx
  loc_00421263: mov var_44, ebx
  loc_00421266: mov var_54, ebx
  loc_00421269: mov var_64, ebx
  loc_0042126C: mov var_74, ebx
  loc_0042126F: mov var_84, ebx
  loc_00421275: call 0041A480h
  loc_0042127A: lea eax, var_24
  loc_0042127D: push eax
  loc_0042127E: call [004010C4h] ; __vbaI2Var
  loc_00421284: lea ecx, var_24
  loc_00421287: mov [edi], ax
  loc_0042128A: call [00401010h] ; __vbaFreeVar
  loc_00421290: push 0040DD3Ch ; "."
  loc_00421295: call [00401024h] ; rtcAnsiValueBstr
  loc_0042129B: mov edx, [esi]
  loc_0042129D: xor ecx, ecx
  loc_0042129F: cmp [edi], ax
  loc_004212A2: push esi
  loc_004212A3: mov var_84, 0000000Bh
  loc_004212AD: setz cl
  loc_004212B0: neg ecx
  loc_004212B2: mov var_7C, cx
  loc_004212B6: call [edx+00000318h]
  loc_004212BC: mov var_1C, eax
  loc_004212BF: lea eax, var_84
  loc_004212C5: push eax
  loc_004212C6: lea ecx, var_24
  loc_004212C9: push 00000001h
  loc_004212CB: lea edx, var_64
  loc_004212CE: push ecx
  loc_004212CF: push edx
  loc_004212D0: lea eax, var_34
  loc_004212D3: push ebx
  loc_004212D4: push eax
  loc_004212D5: mov var_24, 00000009h
  loc_004212DC: mov var_5C, 0040DD3Ch ; "."
  loc_004212E3: mov var_64, 00000008h
  loc_004212EA: mov var_6C, ebx
  loc_004212ED: mov var_74, 00008002h
  loc_004212F4: call [004010B4h] ; __vbaInStrVar
  loc_004212FA: lea ecx, var_74
  loc_004212FD: push eax
  loc_004212FE: lea edx, var_44
  loc_00421301: push ecx
  loc_00421302: push edx
  loc_00421303: call [00401060h] ; __vbaVarCmpGt
  loc_00421309: push eax
  loc_0042130A: lea eax, var_54
  loc_0042130D: push eax
  loc_0042130E: call [00401094h] ; __vbaVarAnd
  loc_00421314: push eax
  loc_00421315: call [00401058h] ; __vbaBoolVarNull
  loc_0042131B: lea ecx, var_84
  loc_00421321: mov esi, eax
  loc_00421323: lea edx, var_34
  loc_00421326: push ecx
  loc_00421327: lea eax, var_24
  loc_0042132A: push edx
  loc_0042132B: push eax
  loc_0042132C: push 00000003h
  loc_0042132E: call [00401018h] ; __vbaFreeVarList
  loc_00421334: add esp, 00000010h
  loc_00421337: cmp si, bx
  loc_0042133A: jz 0042133Fh
  loc_0042133C: mov [edi], bx
  loc_0042133F: push 0040DD44h ; "-"
  loc_00421344: call [00401024h] ; rtcAnsiValueBstr
  loc_0042134A: cmp [edi], ax
  loc_0042134D: jnz 004213BAh
  loc_0042134F: mov ecx, 80020004h
  loc_00421354: mov eax, 0000000Ah
  loc_00421359: mov var_4C, ecx
  loc_0042135C: mov var_3C, ecx
  loc_0042135F: mov var_2C, ecx
  loc_00421362: lea edx, var_64
  loc_00421365: lea ecx, var_24
  loc_00421368: mov var_54, eax
  loc_0042136B: mov var_44, eax
  loc_0042136E: mov var_34, eax
  loc_00421371: mov var_5C, 00410338h ; "Alpha is always positive."
  loc_00421378: mov var_64, 00000008h
  loc_0042137F: call [004010F4h] ; __vbaVarDup
  loc_00421385: lea ecx, var_54
  loc_00421388: lea edx, var_44
  loc_0042138B: push ecx
  loc_0042138C: lea eax, var_34
  loc_0042138F: push edx
  loc_00421390: push eax
  loc_00421391: lea ecx, var_24
  loc_00421394: push ebx
  loc_00421395: push ecx
  loc_00421396: call [00401038h] ; rtcMsgBox
  loc_0042139C: lea edx, var_54
  loc_0042139F: lea eax, var_44
  loc_004213A2: push edx
  loc_004213A3: lea ecx, var_34
  loc_004213A6: push eax
  loc_004213A7: lea edx, var_24
  loc_004213AA: push ecx
  loc_004213AB: push edx
  loc_004213AC: push 00000004h
  loc_004213AE: call [00401018h] ; __vbaFreeVarList
  loc_004213B4: add esp, 00000014h
  loc_004213B7: mov [edi], bx
  loc_004213BA: mov var_4, ebx
  loc_004213BD: push 004213E1h
  loc_004213C2: jmp 004213E0h
  loc_004213C4: lea eax, var_54
  loc_004213C7: lea ecx, var_44
  loc_004213CA: push eax
  loc_004213CB: lea edx, var_34
  loc_004213CE: push ecx
  loc_004213CF: lea eax, var_24
  loc_004213D2: push edx
  loc_004213D3: push eax
  loc_004213D4: push 00000004h
  loc_004213D6: call [00401018h] ; __vbaFreeVarList
  loc_004213DC: add esp, 00000014h
  loc_004213DF: ret
  loc_004213E0: ret
  loc_004213E1: mov eax, Me
  loc_004213E4: push eax
  loc_004213E5: mov ecx, [eax]
  loc_004213E7: call [ecx+00000008h]
  loc_004213EA: mov eax, var_4
  loc_004213ED: mov ecx, var_14
  loc_004213F0: pop edi
  loc_004213F1: pop esi
  loc_004213F2: mov fs:[00000000h], ecx
  loc_004213F9: pop ebx
  loc_004213FA: mov esp, ebp
  loc_004213FC: pop ebp
  loc_004213FD: retn 0008h
End Sub

Private Sub txtTtable_KeyPress(KeyAscii As Integer) '421400
  loc_00421400: push ebp
  loc_00421401: mov ebp, esp
  loc_00421403: sub esp, 0000000Ch
  loc_00421406: push 00401926h ; __vbaExceptHandler
  loc_0042140B: mov eax, fs:[00000000h]
  loc_00421411: push eax
  loc_00421412: mov fs:[00000000h], esp
  loc_00421419: sub esp, 0000007Ch
  loc_0042141C: push ebx
  loc_0042141D: push esi
  loc_0042141E: push edi
  loc_0042141F: mov var_C, esp
  loc_00421422: mov var_8, 00401400h
  loc_00421429: mov esi, Me
  loc_0042142C: mov eax, esi
  loc_0042142E: and eax, 00000001h
  loc_00421431: mov var_4, eax
  loc_00421434: and esi, FFFFFFFEh
  loc_00421437: push esi
  loc_00421438: mov Me, esi
  loc_0042143B: mov ecx, [esi]
  loc_0042143D: call [ecx+00000004h]
  loc_00421440: mov ebx, KeyAscii
  loc_00421443: lea edx, var_24
  loc_00421446: xor edi, edi
  loc_00421448: push ebx
  loc_00421449: push edx
  loc_0042144A: mov var_24, edi
  loc_0042144D: mov var_34, edi
  loc_00421450: mov var_44, edi
  loc_00421453: mov var_54, edi
  loc_00421456: mov var_64, edi
  loc_00421459: mov var_74, edi
  loc_0042145C: mov var_84, edi
  loc_00421462: call 0041A480h
  loc_00421467: lea eax, var_24
  loc_0042146A: push eax
  loc_0042146B: call [004010C4h] ; __vbaI2Var
  loc_00421471: lea ecx, var_24
  loc_00421474: mov [ebx], ax
  loc_00421477: call [00401010h] ; __vbaFreeVar
  loc_0042147D: push 0040DD3Ch ; "."
  loc_00421482: call [00401024h] ; rtcAnsiValueBstr
  loc_00421488: mov edx, [esi]
  loc_0042148A: xor ecx, ecx
  loc_0042148C: cmp [ebx], ax
  loc_0042148F: push esi
  loc_00421490: mov var_84, 0000000Bh
  loc_0042149A: setz cl
  loc_0042149D: neg ecx
  loc_0042149F: mov var_7C, cx
  loc_004214A3: call [edx+00000320h]
  loc_004214A9: mov var_1C, eax
  loc_004214AC: lea eax, var_84
  loc_004214B2: push eax
  loc_004214B3: lea ecx, var_24
  loc_004214B6: push 00000001h
  loc_004214B8: lea edx, var_64
  loc_004214BB: push ecx
  loc_004214BC: push edx
  loc_004214BD: lea eax, var_34
  loc_004214C0: push edi
  loc_004214C1: push eax
  loc_004214C2: mov var_24, 00000009h
  loc_004214C9: mov var_5C, 0040DD3Ch ; "."
  loc_004214D0: mov var_64, 00000008h
  loc_004214D7: mov var_6C, edi
  loc_004214DA: mov var_74, 00008002h
  loc_004214E1: call [004010B4h] ; __vbaInStrVar
  loc_004214E7: lea ecx, var_74
  loc_004214EA: push eax
  loc_004214EB: lea edx, var_44
  loc_004214EE: push ecx
  loc_004214EF: push edx
  loc_004214F0: call [00401060h] ; __vbaVarCmpGt
  loc_004214F6: push eax
  loc_004214F7: lea eax, var_54
  loc_004214FA: push eax
  loc_004214FB: call [00401094h] ; __vbaVarAnd
  loc_00421501: push eax
  loc_00421502: call [00401058h] ; __vbaBoolVarNull
  loc_00421508: lea ecx, var_84
  loc_0042150E: mov esi, eax
  loc_00421510: lea edx, var_34
  loc_00421513: push ecx
  loc_00421514: lea eax, var_24
  loc_00421517: push edx
  loc_00421518: push eax
  loc_00421519: push 00000003h
  loc_0042151B: call [00401018h] ; __vbaFreeVarList
  loc_00421521: add esp, 00000010h
  loc_00421524: cmp si, di
  loc_00421527: jz 0042152Ch
  loc_00421529: mov [ebx], di
  loc_0042152C: mov var_4, edi
  loc_0042152F: push 00421553h
  loc_00421534: jmp 00421552h
  loc_00421536: lea ecx, var_54
  loc_00421539: lea edx, var_44
  loc_0042153C: push ecx
  loc_0042153D: lea eax, var_34
  loc_00421540: push edx
  loc_00421541: lea ecx, var_24
  loc_00421544: push eax
  loc_00421545: push ecx
  loc_00421546: push 00000004h
  loc_00421548: call [00401018h] ; __vbaFreeVarList
  loc_0042154E: add esp, 00000014h
  loc_00421551: ret
  loc_00421552: ret
  loc_00421553: mov eax, Me
  loc_00421556: push eax
  loc_00421557: mov edx, [eax]
  loc_00421559: call [edx+00000008h]
  loc_0042155C: mov eax, var_4
  loc_0042155F: mov ecx, var_14
  loc_00421562: pop edi
  loc_00421563: pop esi
  loc_00421564: mov fs:[00000000h], ecx
  loc_0042156B: pop ebx
  loc_0042156C: mov esp, ebp
  loc_0042156E: pop ebp
  loc_0042156F: retn 0008h
End Sub

Private Sub txt1SidedT_KeyPress(KeyAscii As Integer) '421580
  loc_00421580: push ebp
  loc_00421581: mov ebp, esp
  loc_00421583: sub esp, 0000000Ch
  loc_00421586: push 00401926h ; __vbaExceptHandler
  loc_0042158B: mov eax, fs:[00000000h]
  loc_00421591: push eax
  loc_00421592: mov fs:[00000000h], esp
  loc_00421599: sub esp, 0000008Ch
  loc_0042159F: push ebx
  loc_004215A0: push esi
  loc_004215A1: push edi
  loc_004215A2: mov var_C, esp
  loc_004215A5: mov var_8, 00401410h
  loc_004215AC: mov esi, Me
  loc_004215AF: mov eax, esi
  loc_004215B1: and eax, 00000001h
  loc_004215B4: mov var_4, eax
  loc_004215B7: and esi, FFFFFFFEh
  loc_004215BA: push esi
  loc_004215BB: mov Me, esi
  loc_004215BE: mov ecx, [esi]
  loc_004215C0: call [ecx+00000004h]
  loc_004215C3: mov edi, KeyAscii
  loc_004215C6: lea edx, var_24
  loc_004215C9: xor ebx, ebx
  loc_004215CB: push edi
  loc_004215CC: push edx
  loc_004215CD: mov var_24, ebx
  loc_004215D0: mov var_34, ebx
  loc_004215D3: mov var_44, ebx
  loc_004215D6: mov var_54, ebx
  loc_004215D9: mov var_64, ebx
  loc_004215DC: mov var_74, ebx
  loc_004215DF: mov var_84, ebx
  loc_004215E5: call 0041A480h
  loc_004215EA: lea eax, var_24
  loc_004215ED: push eax
  loc_004215EE: call [004010C4h] ; __vbaI2Var
  loc_004215F4: lea ecx, var_24
  loc_004215F7: mov [edi], ax
  loc_004215FA: call [00401010h] ; __vbaFreeVar
  loc_00421600: push 0040DD3Ch ; "."
  loc_00421605: call [00401024h] ; rtcAnsiValueBstr
  loc_0042160B: mov edx, [esi]
  loc_0042160D: xor ecx, ecx
  loc_0042160F: cmp [edi], ax
  loc_00421612: push esi
  loc_00421613: mov var_84, 0000000Bh
  loc_0042161D: setz cl
  loc_00421620: neg ecx
  loc_00421622: mov var_7C, cx
  loc_00421626: call [edx+00000304h]
  loc_0042162C: mov var_1C, eax
  loc_0042162F: lea eax, var_84
  loc_00421635: push eax
  loc_00421636: lea ecx, var_24
  loc_00421639: push 00000001h
  loc_0042163B: lea edx, var_64
  loc_0042163E: push ecx
  loc_0042163F: push edx
  loc_00421640: lea eax, var_34
  loc_00421643: push ebx
  loc_00421644: push eax
  loc_00421645: mov var_24, 00000009h
  loc_0042164C: mov var_5C, 0040DD3Ch ; "."
  loc_00421653: mov var_64, 00000008h
  loc_0042165A: mov var_6C, ebx
  loc_0042165D: mov var_74, 00008002h
  loc_00421664: call [004010B4h] ; __vbaInStrVar
  loc_0042166A: lea ecx, var_74
  loc_0042166D: push eax
  loc_0042166E: lea edx, var_44
  loc_00421671: push ecx
  loc_00421672: push edx
  loc_00421673: call [00401060h] ; __vbaVarCmpGt
  loc_00421679: push eax
  loc_0042167A: lea eax, var_54
  loc_0042167D: push eax
  loc_0042167E: call [00401094h] ; __vbaVarAnd
  loc_00421684: push eax
  loc_00421685: call [00401058h] ; __vbaBoolVarNull
  loc_0042168B: lea ecx, var_84
  loc_00421691: mov esi, eax
  loc_00421693: lea edx, var_34
  loc_00421696: push ecx
  loc_00421697: lea eax, var_24
  loc_0042169A: push edx
  loc_0042169B: push eax
  loc_0042169C: push 00000003h
  loc_0042169E: call [00401018h] ; __vbaFreeVarList
  loc_004216A4: add esp, 00000010h
  loc_004216A7: cmp si, bx
  loc_004216AA: jz 004216AFh
  loc_004216AC: mov [edi], bx
  loc_004216AF: push 0040DD44h ; "-"
  loc_004216B4: call [00401024h] ; rtcAnsiValueBstr
  loc_004216BA: cmp [edi], ax
  loc_004216BD: jnz 0042172Ah
  loc_004216BF: mov ecx, 80020004h
  loc_004216C4: mov eax, 0000000Ah
  loc_004216C9: mov var_4C, ecx
  loc_004216CC: mov var_3C, ecx
  loc_004216CF: mov var_2C, ecx
  loc_004216D2: lea edx, var_64
  loc_004216D5: lea ecx, var_24
  loc_004216D8: mov var_54, eax
  loc_004216DB: mov var_44, eax
  loc_004216DE: mov var_34, eax
  loc_004216E1: mov var_5C, 00410370h ; "Please enter the right-sided (or positive) t-table value."
  loc_004216E8: mov var_64, 00000008h
  loc_004216EF: call [004010F4h] ; __vbaVarDup
  loc_004216F5: lea ecx, var_54
  loc_004216F8: lea edx, var_44
  loc_004216FB: push ecx
  loc_004216FC: lea eax, var_34
  loc_004216FF: push edx
  loc_00421700: push eax
  loc_00421701: lea ecx, var_24
  loc_00421704: push ebx
  loc_00421705: push ecx
  loc_00421706: call [00401038h] ; rtcMsgBox
  loc_0042170C: lea edx, var_54
  loc_0042170F: lea eax, var_44
  loc_00421712: push edx
  loc_00421713: lea ecx, var_34
  loc_00421716: push eax
  loc_00421717: lea edx, var_24
  loc_0042171A: push ecx
  loc_0042171B: push edx
  loc_0042171C: push 00000004h
  loc_0042171E: call [00401018h] ; __vbaFreeVarList
  loc_00421724: add esp, 00000014h
  loc_00421727: mov [edi], bx
  loc_0042172A: mov var_4, ebx
  loc_0042172D: push 00421751h
  loc_00421732: jmp 00421750h
  loc_00421734: lea eax, var_54
  loc_00421737: lea ecx, var_44
  loc_0042173A: push eax
  loc_0042173B: lea edx, var_34
  loc_0042173E: push ecx
  loc_0042173F: lea eax, var_24
  loc_00421742: push edx
  loc_00421743: push eax
  loc_00421744: push 00000004h
  loc_00421746: call [00401018h] ; __vbaFreeVarList
  loc_0042174C: add esp, 00000014h
  loc_0042174F: ret
  loc_00421750: ret
  loc_00421751: mov eax, Me
  loc_00421754: push eax
  loc_00421755: mov ecx, [eax]
  loc_00421757: call [ecx+00000008h]
  loc_0042175A: mov eax, var_4
  loc_0042175D: mov ecx, var_14
  loc_00421760: pop edi
  loc_00421761: pop esi
  loc_00421762: mov fs:[00000000h], ecx
  loc_00421769: pop ebx
  loc_0042176A: mov esp, ebp
  loc_0042176C: pop ebp
  loc_0042176D: retn 0008h
End Sub

Private Sub cmdClearChange_Click() '4205A0
  loc_004205A0: push ebp
  loc_004205A1: mov ebp, esp
  loc_004205A3: sub esp, 0000000Ch
  loc_004205A6: push 00401926h ; __vbaExceptHandler
  loc_004205AB: mov eax, fs:[00000000h]
  loc_004205B1: push eax
  loc_004205B2: mov fs:[00000000h], esp
  loc_004205B9: sub esp, 00000014h
  loc_004205BC: push ebx
  loc_004205BD: push esi
  loc_004205BE: push edi
  loc_004205BF: mov var_C, esp
  loc_004205C2: mov var_8, 004013A8h
  loc_004205C9: mov esi, Me
  loc_004205CC: mov eax, esi
  loc_004205CE: and eax, 00000001h
  loc_004205D1: mov var_4, eax
  loc_004205D4: and esi, FFFFFFFEh
  loc_004205D7: push esi
  loc_004205D8: mov Me, esi
  loc_004205DB: mov ecx, [esi]
  loc_004205DD: call [ecx+00000004h]
  loc_004205E0: mov edx, [esi]
  loc_004205E2: push esi
  loc_004205E3: mov var_18, 00000000h
  loc_004205EA: call [edx+00000318h]
  loc_004205F0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_004205F6: push eax
  loc_004205F7: lea eax, var_18
  loc_004205FA: push eax
  loc_004205FB: call ebx
  loc_004205FD: mov edi, eax
  loc_004205FF: push 0040F040h
  loc_00420604: push edi
  loc_00420605: mov ecx, [edi]
  loc_00420607: call [ecx+000000A4h]
  loc_0042060D: test eax, eax
  loc_0042060F: fnclex
  loc_00420611: jge 00420625h
  loc_00420613: push 000000A4h
  loc_00420618: push 0040F02Ch
  loc_0042061D: push edi
  loc_0042061E: push eax
  loc_0042061F: call [00401030h] ; __vbaHresultCheckObj
  loc_00420625: lea ecx, var_18
  loc_00420628: call [00401114h] ; __vbaFreeObj
  loc_0042062E: mov edx, [esi]
  loc_00420630: push esi
  loc_00420631: call [edx+00000320h]
  loc_00420637: push eax
  loc_00420638: lea eax, var_18
  loc_0042063B: push eax
  loc_0042063C: call ebx
  loc_0042063E: mov edi, eax
  loc_00420640: push 0040F040h
  loc_00420645: push edi
  loc_00420646: mov ecx, [edi]
  loc_00420648: call [ecx+000000A4h]
  loc_0042064E: test eax, eax
  loc_00420650: fnclex
  loc_00420652: jge 00420666h
  loc_00420654: push 000000A4h
  loc_00420659: push 0040F02Ch
  loc_0042065E: push edi
  loc_0042065F: push eax
  loc_00420660: call [00401030h] ; __vbaHresultCheckObj
  loc_00420666: lea ecx, var_18
  loc_00420669: call [00401114h] ; __vbaFreeObj
  loc_0042066F: mov edx, [esi]
  loc_00420671: push esi
  loc_00420672: call [edx+00000304h]
  loc_00420678: push eax
  loc_00420679: lea eax, var_18
  loc_0042067C: push eax
  loc_0042067D: call ebx
  loc_0042067F: mov edi, eax
  loc_00420681: push 0040F040h
  loc_00420686: push edi
  loc_00420687: mov ecx, [edi]
  loc_00420689: call [ecx+000000A4h]
  loc_0042068F: test eax, eax
  loc_00420691: fnclex
  loc_00420693: jge 004206A7h
  loc_00420695: push 000000A4h
  loc_0042069A: push 0040F02Ch
  loc_0042069F: push edi
  loc_004206A0: push eax
  loc_004206A1: call [00401030h] ; __vbaHresultCheckObj
  loc_004206A7: mov edi, [00401114h] ; __vbaFreeObj
  loc_004206AD: lea ecx, var_18
  loc_004206B0: call edi
  loc_004206B2: mov edx, [esi]
  loc_004206B4: push esi
  loc_004206B5: call [edx+00000318h]
  loc_004206BB: push eax
  loc_004206BC: lea eax, var_18
  loc_004206BF: push eax
  loc_004206C0: call ebx
  loc_004206C2: mov esi, eax
  loc_004206C4: push esi
  loc_004206C5: mov ecx, [esi]
  loc_004206C7: call [ecx+00000204h]
  loc_004206CD: test eax, eax
  loc_004206CF: fnclex
  loc_004206D1: jge 004206E5h
  loc_004206D3: push 00000204h
  loc_004206D8: push 0040F02Ch
  loc_004206DD: push esi
  loc_004206DE: push eax
  loc_004206DF: call [00401030h] ; __vbaHresultCheckObj
  loc_004206E5: lea ecx, var_18
  loc_004206E8: call edi
  loc_004206EA: mov var_4, 00000000h
  loc_004206F1: push 00420703h
  loc_004206F6: jmp 00420702h
  loc_004206F8: lea ecx, var_18
  loc_004206FB: call [00401114h] ; __vbaFreeObj
  loc_00420701: ret
  loc_00420702: ret
  loc_00420703: mov eax, Me
  loc_00420706: push eax
  loc_00420707: mov edx, [eax]
  loc_00420709: call [edx+00000008h]
  loc_0042070C: mov eax, var_4
  loc_0042070F: mov ecx, var_14
  loc_00420712: pop edi
  loc_00420713: pop esi
  loc_00420714: mov fs:[00000000h], ecx
  loc_0042071B: pop ebx
  loc_0042071C: mov esp, ebp
  loc_0042071E: pop ebp
  loc_0042071F: retn 0004h
End Sub
