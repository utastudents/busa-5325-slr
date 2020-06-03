VERSION 5.00
Begin VB.Form frmChangeHypothesizedValue
  Caption = "Change "
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 4965
  ClientHeight = 2415
  BeginProperty Font
    Name = "MS Sans Serif"
    Size = 18
    Charset = 0
    Weight = 400
    Underline = 0 'False
    Italic = 0 'False
    Strikethrough = 0 'False
  EndProperty
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Hypothesized Value of Slope"
    Left = 0
    Top = 0
    Width = 4905
    Height = 2325
    TabIndex = 0
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 450
      Top = 1560
      Width = 885
      Height = 555
      TabIndex = 4
    End
    Begin VB.CommandButton cmdClearChange
      Caption = "Clear"
      Left = 1560
      Top = 1560
      Width = 1425
      Height = 555
      TabIndex = 3
    End
    Begin VB.CommandButton cmdRestore
      Caption = "Reset "
      Left = 3150
      Top = 1560
      Width = 1455
      Height = 555
      TabIndex = 2
    End
    Begin VB.TextBox txtHypothesizedValue
      Left = 900
      Top = 660
      Width = 2175
      Height = 495
      Text = "0"
      TabIndex = 1
    End
  End
End

Attribute VB_Name = "frmChangeHypothesizedValue"


Private Sub cmdRestore_Click() '41E6B0
  loc_0041E6B0: push ebp
  loc_0041E6B1: mov ebp, esp
  loc_0041E6B3: sub esp, 0000000Ch
  loc_0041E6B6: push 00401926h ; __vbaExceptHandler
  loc_0041E6BB: mov eax, fs:[00000000h]
  loc_0041E6C1: push eax
  loc_0041E6C2: mov fs:[00000000h], esp
  loc_0041E6C9: sub esp, 00000014h
  loc_0041E6CC: push ebx
  loc_0041E6CD: push esi
  loc_0041E6CE: push edi
  loc_0041E6CF: mov var_C, esp
  loc_0041E6D2: mov var_8, 004012C8h
  loc_0041E6D9: mov esi, Me
  loc_0041E6DC: mov eax, esi
  loc_0041E6DE: and eax, 00000001h
  loc_0041E6E1: mov var_4, eax
  loc_0041E6E4: and esi, FFFFFFFEh
  loc_0041E6E7: push esi
  loc_0041E6E8: mov Me, esi
  loc_0041E6EB: mov ecx, [esi]
  loc_0041E6ED: call [ecx+00000004h]
  loc_0041E6F0: mov [0043005Ch], 42C80000h
  loc_0041E6FA: mov edx, [esi]
  loc_0041E6FC: xor edi, edi
  loc_0041E6FE: push esi
  loc_0041E6FF: mov var_18, edi
  loc_0041E702: call [edx+0000030Ch]
  loc_0041E708: push eax
  loc_0041E709: lea eax, var_18
  loc_0041E70C: push eax
  loc_0041E70D: call [0040103Ch] ; __vbaObjSet
  loc_0041E713: mov esi, eax
  loc_0041E715: push 0040FCA0h ; "100"
  loc_0041E71A: push esi
  loc_0041E71B: mov ecx, [esi]
  loc_0041E71D: call [ecx+000000A4h]
  loc_0041E723: cmp eax, edi
  loc_0041E725: fnclex
  loc_0041E727: jge 0041E73Bh
  loc_0041E729: push 000000A4h
  loc_0041E72E: push 0040F02Ch
  loc_0041E733: push esi
  loc_0041E734: push eax
  loc_0041E735: call [00401030h] ; __vbaHresultCheckObj
  loc_0041E73B: lea ecx, var_18
  loc_0041E73E: call [00401114h] ; __vbaFreeObj
  loc_0041E744: mov var_4, edi
  loc_0041E747: fwait
  loc_0041E748: push 0041E75Ah
  loc_0041E74D: jmp 0041E759h
  loc_0041E74F: lea ecx, var_18
  loc_0041E752: call [00401114h] ; __vbaFreeObj
  loc_0041E758: ret
  loc_0041E759: ret
  loc_0041E75A: mov eax, Me
  loc_0041E75D: push eax
  loc_0041E75E: mov edx, [eax]
  loc_0041E760: call [edx+00000008h]
  loc_0041E763: mov eax, var_4
  loc_0041E766: mov ecx, var_14
  loc_0041E769: pop edi
  loc_0041E76A: pop esi
  loc_0041E76B: mov fs:[00000000h], ecx
  loc_0041E772: pop ebx
  loc_0041E773: mov esp, ebp
  loc_0041E775: pop ebp
  loc_0041E776: retn 0004h
End Sub

Private Sub cmdClearChange_Click() '41E5B0
  loc_0041E5B0: push ebp
  loc_0041E5B1: mov ebp, esp
  loc_0041E5B3: sub esp, 0000000Ch
  loc_0041E5B6: push 00401926h ; __vbaExceptHandler
  loc_0041E5BB: mov eax, fs:[00000000h]
  loc_0041E5C1: push eax
  loc_0041E5C2: mov fs:[00000000h], esp
  loc_0041E5C9: sub esp, 00000014h
  loc_0041E5CC: push ebx
  loc_0041E5CD: push esi
  loc_0041E5CE: push edi
  loc_0041E5CF: mov var_C, esp
  loc_0041E5D2: mov var_8, 004012B8h
  loc_0041E5D9: mov esi, Me
  loc_0041E5DC: mov eax, esi
  loc_0041E5DE: and eax, 00000001h
  loc_0041E5E1: mov var_4, eax
  loc_0041E5E4: and esi, FFFFFFFEh
  loc_0041E5E7: push esi
  loc_0041E5E8: mov Me, esi
  loc_0041E5EB: mov ecx, [esi]
  loc_0041E5ED: call [ecx+00000004h]
  loc_0041E5F0: mov edx, [esi]
  loc_0041E5F2: push esi
  loc_0041E5F3: mov var_18, 00000000h
  loc_0041E5FA: call [edx+0000030Ch]
  loc_0041E600: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041E606: push eax
  loc_0041E607: lea eax, var_18
  loc_0041E60A: push eax
  loc_0041E60B: call ebx
  loc_0041E60D: mov edi, eax
  loc_0041E60F: push 0040F040h
  loc_0041E614: push edi
  loc_0041E615: mov ecx, [edi]
  loc_0041E617: call [ecx+000000A4h]
  loc_0041E61D: test eax, eax
  loc_0041E61F: fnclex
  loc_0041E621: jge 0041E635h
  loc_0041E623: push 000000A4h
  loc_0041E628: push 0040F02Ch
  loc_0041E62D: push edi
  loc_0041E62E: push eax
  loc_0041E62F: call [00401030h] ; __vbaHresultCheckObj
  loc_0041E635: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041E63B: lea ecx, var_18
  loc_0041E63E: call edi
  loc_0041E640: mov edx, [esi]
  loc_0041E642: push esi
  loc_0041E643: call [edx+0000030Ch]
  loc_0041E649: push eax
  loc_0041E64A: lea eax, var_18
  loc_0041E64D: push eax
  loc_0041E64E: call ebx
  loc_0041E650: mov esi, eax
  loc_0041E652: push esi
  loc_0041E653: mov ecx, [esi]
  loc_0041E655: call [ecx+00000204h]
  loc_0041E65B: test eax, eax
  loc_0041E65D: fnclex
  loc_0041E65F: jge 0041E673h
  loc_0041E661: push 00000204h
  loc_0041E666: push 0040F02Ch
  loc_0041E66B: push esi
  loc_0041E66C: push eax
  loc_0041E66D: call [00401030h] ; __vbaHresultCheckObj
  loc_0041E673: lea ecx, var_18
  loc_0041E676: call edi
  loc_0041E678: mov var_4, 00000000h
  loc_0041E67F: push 0041E691h
  loc_0041E684: jmp 0041E690h
  loc_0041E686: lea ecx, var_18
  loc_0041E689: call [00401114h] ; __vbaFreeObj
  loc_0041E68F: ret
  loc_0041E690: ret
  loc_0041E691: mov eax, Me
  loc_0041E694: push eax
  loc_0041E695: mov edx, [eax]
  loc_0041E697: call [edx+00000008h]
  loc_0041E69A: mov eax, var_4
  loc_0041E69D: mov ecx, var_14
  loc_0041E6A0: pop edi
  loc_0041E6A1: pop esi
  loc_0041E6A2: mov fs:[00000000h], ecx
  loc_0041E6A9: pop ebx
  loc_0041E6AA: mov esp, ebp
  loc_0041E6AC: pop ebp
  loc_0041E6AD: retn 0004h
End Sub

Private Sub Form_Activate() '41E9D0
  loc_0041E9D0: push ebp
  loc_0041E9D1: mov ebp, esp
  loc_0041E9D3: sub esp, 0000000Ch
  loc_0041E9D6: push 00401926h ; __vbaExceptHandler
  loc_0041E9DB: mov eax, fs:[00000000h]
  loc_0041E9E1: push eax
  loc_0041E9E2: mov fs:[00000000h], esp
  loc_0041E9E9: sub esp, 00000018h
  loc_0041E9EC: push ebx
  loc_0041E9ED: push esi
  loc_0041E9EE: push edi
  loc_0041E9EF: mov var_C, esp
  loc_0041E9F2: mov var_8, 004012E8h
  loc_0041E9F9: mov esi, Me
  loc_0041E9FC: mov eax, esi
  loc_0041E9FE: and eax, 00000001h
  loc_0041EA01: mov var_4, eax
  loc_0041EA04: and esi, FFFFFFFEh
  loc_0041EA07: push esi
  loc_0041EA08: mov Me, esi
  loc_0041EA0B: mov ecx, [esi]
  loc_0041EA0D: call [ecx+00000004h]
  loc_0041EA10: mov edx, [esi]
  loc_0041EA12: xor eax, eax
  loc_0041EA14: push esi
  loc_0041EA15: mov var_18, eax
  loc_0041EA18: mov var_1C, eax
  loc_0041EA1B: call [edx+0000030Ch]
  loc_0041EA21: push eax
  loc_0041EA22: lea eax, var_1C
  loc_0041EA25: push eax
  loc_0041EA26: call [0040103Ch] ; __vbaObjSet
  loc_0041EA2C: mov ecx, [0043005Ch]
  loc_0041EA32: mov edi, eax
  loc_0041EA34: push ecx
  loc_0041EA35: mov ebx, [edi]
  loc_0041EA37: call [0040107Ch] ; __vbaStrR4
  loc_0041EA3D: mov edx, eax
  loc_0041EA3F: lea ecx, var_18
  loc_0041EA42: call [004010FCh] ; __vbaStrMove
  loc_0041EA48: push eax
  loc_0041EA49: push edi
  loc_0041EA4A: call [ebx+000000A4h]
  loc_0041EA50: test eax, eax
  loc_0041EA52: fnclex
  loc_0041EA54: jge 0041EA68h
  loc_0041EA56: push 000000A4h
  loc_0041EA5B: push 0040F02Ch
  loc_0041EA60: push edi
  loc_0041EA61: push eax
  loc_0041EA62: call [00401030h] ; __vbaHresultCheckObj
  loc_0041EA68: lea ecx, var_18
  loc_0041EA6B: call [00401110h] ; __vbaFreeStr
  loc_0041EA71: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041EA77: lea ecx, var_1C
  loc_0041EA7A: call edi
  loc_0041EA7C: mov edx, [esi]
  loc_0041EA7E: push esi
  loc_0041EA7F: call [edx+0000030Ch]
  loc_0041EA85: push eax
  loc_0041EA86: lea eax, var_1C
  loc_0041EA89: push eax
  loc_0041EA8A: call [0040103Ch] ; __vbaObjSet
  loc_0041EA90: mov esi, eax
  loc_0041EA92: push esi
  loc_0041EA93: mov ecx, [esi]
  loc_0041EA95: call [ecx+00000204h]
  loc_0041EA9B: test eax, eax
  loc_0041EA9D: fnclex
  loc_0041EA9F: jge 0041EAB3h
  loc_0041EAA1: push 00000204h
  loc_0041EAA6: push 0040F02Ch
  loc_0041EAAB: push esi
  loc_0041EAAC: push eax
  loc_0041EAAD: call [00401030h] ; __vbaHresultCheckObj
  loc_0041EAB3: lea ecx, var_1C
  loc_0041EAB6: call edi
  loc_0041EAB8: mov var_4, 00000000h
  loc_0041EABF: fwait
  loc_0041EAC0: push 0041EADBh
  loc_0041EAC5: jmp 0041EADAh
  loc_0041EAC7: lea ecx, var_18
  loc_0041EACA: call [00401110h] ; __vbaFreeStr
  loc_0041EAD0: lea ecx, var_1C
  loc_0041EAD3: call [00401114h] ; __vbaFreeObj
  loc_0041EAD9: ret
  loc_0041EADA: ret
  loc_0041EADB: mov eax, Me
  loc_0041EADE: push eax
  loc_0041EADF: mov edx, [eax]
  loc_0041EAE1: call [edx+00000008h]
  loc_0041EAE4: mov eax, var_4
  loc_0041EAE7: mov ecx, var_14
  loc_0041EAEA: pop edi
  loc_0041EAEB: pop esi
  loc_0041EAEC: mov fs:[00000000h], ecx
  loc_0041EAF3: pop ebx
  loc_0041EAF4: mov esp, ebp
  loc_0041EAF6: pop ebp
  loc_0041EAF7: retn 0004h
End Sub

Private Sub txtHypothesizedValue_KeyPress(KeyAscii As Integer) '41EB00
  loc_0041EB00: push ebp
  loc_0041EB01: mov ebp, esp
  loc_0041EB03: sub esp, 0000000Ch
  loc_0041EB06: push 00401926h ; __vbaExceptHandler
  loc_0041EB0B: mov eax, fs:[00000000h]
  loc_0041EB11: push eax
  loc_0041EB12: mov fs:[00000000h], esp
  loc_0041EB19: sub esp, 0000007Ch
  loc_0041EB1C: push ebx
  loc_0041EB1D: push esi
  loc_0041EB1E: push edi
  loc_0041EB1F: mov var_C, esp
  loc_0041EB22: mov var_8, 004012F8h
  loc_0041EB29: mov esi, Me
  loc_0041EB2C: mov eax, esi
  loc_0041EB2E: and eax, 00000001h
  loc_0041EB31: mov var_4, eax
  loc_0041EB34: and esi, FFFFFFFEh
  loc_0041EB37: push esi
  loc_0041EB38: mov Me, esi
  loc_0041EB3B: mov ecx, [esi]
  loc_0041EB3D: call [ecx+00000004h]
  loc_0041EB40: mov ebx, KeyAscii
  loc_0041EB43: lea edx, var_24
  loc_0041EB46: xor edi, edi
  loc_0041EB48: push ebx
  loc_0041EB49: push edx
  loc_0041EB4A: mov var_24, edi
  loc_0041EB4D: mov var_34, edi
  loc_0041EB50: mov var_44, edi
  loc_0041EB53: mov var_54, edi
  loc_0041EB56: mov var_64, edi
  loc_0041EB59: mov var_74, edi
  loc_0041EB5C: mov var_84, edi
  loc_0041EB62: call 0041A480h
  loc_0041EB67: lea eax, var_24
  loc_0041EB6A: push eax
  loc_0041EB6B: call [004010C4h] ; __vbaI2Var
  loc_0041EB71: lea ecx, var_24
  loc_0041EB74: mov [ebx], ax
  loc_0041EB77: call [00401010h] ; __vbaFreeVar
  loc_0041EB7D: push 0040DD3Ch ; "."
  loc_0041EB82: call [00401024h] ; rtcAnsiValueBstr
  loc_0041EB88: mov edx, [esi]
  loc_0041EB8A: xor ecx, ecx
  loc_0041EB8C: cmp [ebx], ax
  loc_0041EB8F: push esi
  loc_0041EB90: mov var_84, 0000000Bh
  loc_0041EB9A: setz cl
  loc_0041EB9D: neg ecx
  loc_0041EB9F: mov var_7C, cx
  loc_0041EBA3: call [edx+0000030Ch]
  loc_0041EBA9: mov var_1C, eax
  loc_0041EBAC: lea eax, var_84
  loc_0041EBB2: push eax
  loc_0041EBB3: lea ecx, var_24
  loc_0041EBB6: push 00000001h
  loc_0041EBB8: lea edx, var_64
  loc_0041EBBB: push ecx
  loc_0041EBBC: push edx
  loc_0041EBBD: lea eax, var_34
  loc_0041EBC0: push edi
  loc_0041EBC1: push eax
  loc_0041EBC2: mov var_24, 00000009h
  loc_0041EBC9: mov var_5C, 0040DD3Ch ; "."
  loc_0041EBD0: mov var_64, 00000008h
  loc_0041EBD7: mov var_6C, edi
  loc_0041EBDA: mov var_74, 00008002h
  loc_0041EBE1: call [004010B4h] ; __vbaInStrVar
  loc_0041EBE7: lea ecx, var_74
  loc_0041EBEA: push eax
  loc_0041EBEB: lea edx, var_44
  loc_0041EBEE: push ecx
  loc_0041EBEF: push edx
  loc_0041EBF0: call [00401060h] ; __vbaVarCmpGt
  loc_0041EBF6: push eax
  loc_0041EBF7: lea eax, var_54
  loc_0041EBFA: push eax
  loc_0041EBFB: call [00401094h] ; __vbaVarAnd
  loc_0041EC01: push eax
  loc_0041EC02: call [00401058h] ; __vbaBoolVarNull
  loc_0041EC08: lea ecx, var_84
  loc_0041EC0E: mov esi, eax
  loc_0041EC10: lea edx, var_34
  loc_0041EC13: push ecx
  loc_0041EC14: lea eax, var_24
  loc_0041EC17: push edx
  loc_0041EC18: push eax
  loc_0041EC19: push 00000003h
  loc_0041EC1B: call [00401018h] ; __vbaFreeVarList
  loc_0041EC21: add esp, 00000010h
  loc_0041EC24: cmp si, di
  loc_0041EC27: jz 0041EC2Ch
  loc_0041EC29: mov [ebx], di
  loc_0041EC2C: mov var_4, edi
  loc_0041EC2F: push 0041EC53h
  loc_0041EC34: jmp 0041EC52h
  loc_0041EC36: lea ecx, var_54
  loc_0041EC39: lea edx, var_44
  loc_0041EC3C: push ecx
  loc_0041EC3D: lea eax, var_34
  loc_0041EC40: push edx
  loc_0041EC41: lea ecx, var_24
  loc_0041EC44: push eax
  loc_0041EC45: push ecx
  loc_0041EC46: push 00000004h
  loc_0041EC48: call [00401018h] ; __vbaFreeVarList
  loc_0041EC4E: add esp, 00000014h
  loc_0041EC51: ret
  loc_0041EC52: ret
  loc_0041EC53: mov eax, Me
  loc_0041EC56: push eax
  loc_0041EC57: mov edx, [eax]
  loc_0041EC59: call [edx+00000008h]
  loc_0041EC5C: mov eax, var_4
  loc_0041EC5F: mov ecx, var_14
  loc_0041EC62: pop edi
  loc_0041EC63: pop esi
  loc_0041EC64: mov fs:[00000000h], ecx
  loc_0041EC6B: pop ebx
  loc_0041EC6C: mov esp, ebp
  loc_0041EC6E: pop ebp
  loc_0041EC6F: retn 0008h
End Sub

Private Sub cmdChange_Click() '41E780
  loc_0041E780: push ebp
  loc_0041E781: mov ebp, esp
  loc_0041E783: sub esp, 0000000Ch
  loc_0041E786: push 00401926h ; __vbaExceptHandler
  loc_0041E78B: mov eax, fs:[00000000h]
  loc_0041E791: push eax
  loc_0041E792: mov fs:[00000000h], esp
  loc_0041E799: sub esp, 00000094h
  loc_0041E79F: push ebx
  loc_0041E7A0: push esi
  loc_0041E7A1: push edi
  loc_0041E7A2: mov var_C, esp
  loc_0041E7A5: mov var_8, 004012D8h
  loc_0041E7AC: mov esi, Me
  loc_0041E7AF: mov eax, esi
  loc_0041E7B1: and eax, 00000001h
  loc_0041E7B4: mov var_4, eax
  loc_0041E7B7: and esi, FFFFFFFEh
  loc_0041E7BA: push esi
  loc_0041E7BB: mov Me, esi
  loc_0041E7BE: mov ecx, [esi]
  loc_0041E7C0: call [ecx+00000004h]
  loc_0041E7C3: mov edx, [esi]
  loc_0041E7C5: xor edi, edi
  loc_0041E7C7: push esi
  loc_0041E7C8: mov var_18, edi
  loc_0041E7CB: mov var_28, edi
  loc_0041E7CE: mov var_38, edi
  loc_0041E7D1: mov var_48, edi
  loc_0041E7D4: mov var_58, edi
  loc_0041E7D7: mov var_68, edi
  loc_0041E7DA: mov var_78, edi
  loc_0041E7DD: call [edx+0000030Ch]
  loc_0041E7E3: mov var_20, eax
  loc_0041E7E6: lea eax, var_28
  loc_0041E7E9: lea ecx, var_38
  loc_0041E7EC: push eax
  loc_0041E7ED: push ecx
  loc_0041E7EE: mov var_28, 00000009h
  loc_0041E7F5: call [00401050h] ; rtcTrimVar
  loc_0041E7FB: lea edx, var_38
  loc_0041E7FE: push edx
  loc_0041E7FF: call [004010A0h] ; __vbaR4ErrVar
  loc_0041E805: mov ebx, [00401018h] ; __vbaFreeVarList
  loc_0041E80B: lea eax, var_38
  loc_0041E80E: fstp real4 ptr [0043005Ch]
  loc_0041E814: lea ecx, var_38
  loc_0041E817: push eax
  loc_0041E818: lea edx, var_28
  loc_0041E81B: push ecx
  loc_0041E81C: push edx
  loc_0041E81D: push 00000003h
  loc_0041E81F: call ebx
  loc_0041E821: mov eax, [esi]
  loc_0041E823: add esp, 00000010h
  loc_0041E826: push esi
  loc_0041E827: call [eax+0000030Ch]
  loc_0041E82D: lea ecx, var_28
  loc_0041E830: lea edx, var_38
  loc_0041E833: push ecx
  loc_0041E834: push edx
  loc_0041E835: mov var_20, eax
  loc_0041E838: mov var_28, 00000009h
  loc_0041E83F: call [00401050h] ; rtcTrimVar
  loc_0041E845: lea eax, var_38
  loc_0041E848: lea ecx, var_68
  loc_0041E84B: push eax
  loc_0041E84C: push ecx
  loc_0041E84D: mov var_60, 0040F040h
  loc_0041E854: mov var_68, 00008008h
  loc_0041E85B: call [00401070h] ; __vbaVarTstEq
  loc_0041E861: mov var_9C, ax
  loc_0041E868: lea edx, var_38
  loc_0041E86B: lea eax, var_28
  loc_0041E86E: push edx
  loc_0041E86F: push eax
  loc_0041E870: push 00000002h
  loc_0041E872: call ebx
  loc_0041E874: add esp, 0000000Ch
  loc_0041E877: cmp var_9C, di
  loc_0041E87E: jz 0041E93Ch
  loc_0041E884: mov ecx, [esi]
  loc_0041E886: push esi
  loc_0041E887: call [ecx+0000030Ch]
  loc_0041E88D: lea edx, var_18
  loc_0041E890: push eax
  loc_0041E891: push edx
  loc_0041E892: call [0040103Ch] ; __vbaObjSet
  loc_0041E898: mov esi, eax
  loc_0041E89A: push esi
  loc_0041E89B: mov eax, [esi]
  loc_0041E89D: call [eax+00000204h]
  loc_0041E8A3: cmp eax, edi
  loc_0041E8A5: fnclex
  loc_0041E8A7: jge 0041E8BBh
  loc_0041E8A9: push 00000204h
  loc_0041E8AE: push 0040F02Ch
  loc_0041E8B3: push esi
  loc_0041E8B4: push eax
  loc_0041E8B5: call [00401030h] ; __vbaHresultCheckObj
  loc_0041E8BB: lea ecx, var_18
  loc_0041E8BE: call [00401114h] ; __vbaFreeObj
  loc_0041E8C4: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041E8CA: mov ecx, 80020004h
  loc_0041E8CF: mov var_50, ecx
  loc_0041E8D2: mov eax, 0000000Ah
  loc_0041E8D7: mov var_40, ecx
  loc_0041E8DA: lea edx, var_78
  loc_0041E8DD: lea ecx, var_38
  loc_0041E8E0: mov var_58, eax
  loc_0041E8E3: mov var_48, eax
  loc_0041E8E6: mov var_70, 0040FD34h ; "Check Value"
  loc_0041E8ED: mov var_78, 00000008h
  loc_0041E8F4: call __vbaVarDup
  loc_0041E8F6: lea edx, var_68
  loc_0041E8F9: lea ecx, var_28
  loc_0041E8FC: mov var_60, 0040FCACh ; "Please enter a value for the given value of X: HypothesizedValue"
  loc_0041E903: mov var_68, 00000008h
  loc_0041E90A: call __vbaVarDup
  loc_0041E90C: lea ecx, var_58
  loc_0041E90F: lea edx, var_48
  loc_0041E912: push ecx
  loc_0041E913: lea eax, var_38
  loc_0041E916: push edx
  loc_0041E917: push eax
  loc_0041E918: lea ecx, var_28
  loc_0041E91B: push edi
  loc_0041E91C: push ecx
  loc_0041E91D: call [00401038h] ; rtcMsgBox
  loc_0041E923: lea edx, var_58
  loc_0041E926: lea eax, var_48
  loc_0041E929: push edx
  loc_0041E92A: lea ecx, var_38
  loc_0041E92D: push eax
  loc_0041E92E: lea edx, var_28
  loc_0041E931: push ecx
  loc_0041E932: push edx
  loc_0041E933: push 00000004h
  loc_0041E935: call ebx
  loc_0041E937: add esp, 00000014h
  loc_0041E93A: jmp 0041E97Bh
  loc_0041E93C: cmp [0043013Ch], edi
  loc_0041E942: jnz 0041E954h
  loc_0041E944: push 0043013Ch
  loc_0041E949: push 0040204Ch
  loc_0041E94E: call [004010D4h] ; __vbaNew2
  loc_0041E954: mov esi, [0043013Ch]
  loc_0041E95A: push esi
  loc_0041E95B: mov eax, [esi]
  loc_0041E95D: call [eax+000002B4h]
  loc_0041E963: cmp eax, edi
  loc_0041E965: fnclex
  loc_0041E967: jge 0041E97Bh
  loc_0041E969: push 000002B4h
  loc_0041E96E: push 0040FC38h
  loc_0041E973: push esi
  loc_0041E974: push eax
  loc_0041E975: call [00401030h] ; __vbaHresultCheckObj
  loc_0041E97B: mov var_4, edi
  loc_0041E97E: fwait
  loc_0041E97F: push 0041E9ACh
  loc_0041E984: jmp 0041E9ABh
  loc_0041E986: lea ecx, var_18
  loc_0041E989: call [00401114h] ; __vbaFreeObj
  loc_0041E98F: lea ecx, var_58
  loc_0041E992: lea edx, var_48
  loc_0041E995: push ecx
  loc_0041E996: lea eax, var_38
  loc_0041E999: push edx
  loc_0041E99A: lea ecx, var_28
  loc_0041E99D: push eax
  loc_0041E99E: push ecx
  loc_0041E99F: push 00000004h
  loc_0041E9A1: call [00401018h] ; __vbaFreeVarList
  loc_0041E9A7: add esp, 00000014h
  loc_0041E9AA: ret
  loc_0041E9AB: ret
  loc_0041E9AC: mov eax, Me
  loc_0041E9AF: push eax
  loc_0041E9B0: mov edx, [eax]
  loc_0041E9B2: call [edx+00000008h]
  loc_0041E9B5: mov eax, var_4
  loc_0041E9B8: mov ecx, var_14
  loc_0041E9BB: pop edi
  loc_0041E9BC: pop esi
  loc_0041E9BD: mov fs:[00000000h], ecx
  loc_0041E9C4: pop ebx
  loc_0041E9C5: mov esp, ebp
  loc_0041E9C7: pop ebp
  loc_0041E9C8: retn 0004h
End Sub
