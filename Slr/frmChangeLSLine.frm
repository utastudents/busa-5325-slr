VERSION 5.00
Begin VB.Form frmChangeLSLine
  Caption = "Change Least Squares Line"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 6315
  ClientHeight = 4455
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "The Least Squares Line"
    Left = 0
    Top = 0
    Width = 6255
    Height = 4455
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
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 240
      Top = 3720
      Width = 1215
      Height = 495
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
      Left = 1680
      Top = 3720
      Width = 1335
      Height = 495
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
      Left = 3240
      Top = 3720
      Width = 2775
      Height = 495
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
      Caption = "Estimate of Slope"
      Left = 360
      Top = 2280
      Width = 5655
      Height = 1215
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
      Begin VB.TextBox txtBetaHat
        Left = 240
        Top = 480
        Width = 4335
        Height = 615
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
      Caption = "Estimate of Intercept"
      Left = 360
      Top = 840
      Width = 5655
      Height = 1215
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
      Begin VB.TextBox txtbeta0Hat
        Left = 240
        Top = 480
        Width = 4335
        Height = 615
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

Attribute VB_Name = "frmChangeLSLine"


Private Sub cmdClearChange_Click() '41EC80
  loc_0041EC80: push ebp
  loc_0041EC81: mov ebp, esp
  loc_0041EC83: sub esp, 0000000Ch
  loc_0041EC86: push 00401926h ; __vbaExceptHandler
  loc_0041EC8B: mov eax, fs:[00000000h]
  loc_0041EC91: push eax
  loc_0041EC92: mov fs:[00000000h], esp
  loc_0041EC99: sub esp, 00000014h
  loc_0041EC9C: push ebx
  loc_0041EC9D: push esi
  loc_0041EC9E: push edi
  loc_0041EC9F: mov var_C, esp
  loc_0041ECA2: mov var_8, 00401308h
  loc_0041ECA9: mov esi, Me
  loc_0041ECAC: mov eax, esi
  loc_0041ECAE: and eax, 00000001h
  loc_0041ECB1: mov var_4, eax
  loc_0041ECB4: and esi, FFFFFFFEh
  loc_0041ECB7: push esi
  loc_0041ECB8: mov Me, esi
  loc_0041ECBB: mov ecx, [esi]
  loc_0041ECBD: call [ecx+00000004h]
  loc_0041ECC0: mov edx, [esi]
  loc_0041ECC2: push esi
  loc_0041ECC3: mov var_18, 00000000h
  loc_0041ECCA: call [edx+00000310h]
  loc_0041ECD0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041ECD6: push eax
  loc_0041ECD7: lea eax, var_18
  loc_0041ECDA: push eax
  loc_0041ECDB: call ebx
  loc_0041ECDD: mov edi, eax
  loc_0041ECDF: push 0040F040h
  loc_0041ECE4: push edi
  loc_0041ECE5: mov ecx, [edi]
  loc_0041ECE7: call [ecx+000000A4h]
  loc_0041ECED: test eax, eax
  loc_0041ECEF: fnclex
  loc_0041ECF1: jge 0041ED05h
  loc_0041ECF3: push 000000A4h
  loc_0041ECF8: push 0040F02Ch
  loc_0041ECFD: push edi
  loc_0041ECFE: push eax
  loc_0041ECFF: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ED05: lea ecx, var_18
  loc_0041ED08: call [00401114h] ; __vbaFreeObj
  loc_0041ED0E: mov edx, [esi]
  loc_0041ED10: push esi
  loc_0041ED11: call [edx+00000318h]
  loc_0041ED17: push eax
  loc_0041ED18: lea eax, var_18
  loc_0041ED1B: push eax
  loc_0041ED1C: call ebx
  loc_0041ED1E: mov edi, eax
  loc_0041ED20: push 0040F040h
  loc_0041ED25: push edi
  loc_0041ED26: mov ecx, [edi]
  loc_0041ED28: call [ecx+000000A4h]
  loc_0041ED2E: test eax, eax
  loc_0041ED30: fnclex
  loc_0041ED32: jge 0041ED46h
  loc_0041ED34: push 000000A4h
  loc_0041ED39: push 0040F02Ch
  loc_0041ED3E: push edi
  loc_0041ED3F: push eax
  loc_0041ED40: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ED46: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041ED4C: lea ecx, var_18
  loc_0041ED4F: call edi
  loc_0041ED51: mov edx, [esi]
  loc_0041ED53: push esi
  loc_0041ED54: call [edx+00000310h]
  loc_0041ED5A: push eax
  loc_0041ED5B: lea eax, var_18
  loc_0041ED5E: push eax
  loc_0041ED5F: call ebx
  loc_0041ED61: mov esi, eax
  loc_0041ED63: push esi
  loc_0041ED64: mov ecx, [esi]
  loc_0041ED66: call [ecx+00000204h]
  loc_0041ED6C: test eax, eax
  loc_0041ED6E: fnclex
  loc_0041ED70: jge 0041ED84h
  loc_0041ED72: push 00000204h
  loc_0041ED77: push 0040F02Ch
  loc_0041ED7C: push esi
  loc_0041ED7D: push eax
  loc_0041ED7E: call [00401030h] ; __vbaHresultCheckObj
  loc_0041ED84: lea ecx, var_18
  loc_0041ED87: call edi
  loc_0041ED89: mov var_4, 00000000h
  loc_0041ED90: push 0041EDA2h
  loc_0041ED95: jmp 0041EDA1h
  loc_0041ED97: lea ecx, var_18
  loc_0041ED9A: call [00401114h] ; __vbaFreeObj
  loc_0041EDA0: ret
  loc_0041EDA1: ret
  loc_0041EDA2: mov eax, Me
  loc_0041EDA5: push eax
  loc_0041EDA6: mov edx, [eax]
  loc_0041EDA8: call [edx+00000008h]
  loc_0041EDAB: mov eax, var_4
  loc_0041EDAE: mov ecx, var_14
  loc_0041EDB1: pop edi
  loc_0041EDB2: pop esi
  loc_0041EDB3: mov fs:[00000000h], ecx
  loc_0041EDBA: pop ebx
  loc_0041EDBB: mov esp, ebp
  loc_0041EDBD: pop ebp
  loc_0041EDBE: retn 0004h
End Sub

Private Sub txtBetaHat_KeyPress(KeyAscii As Integer) '41FA50
  loc_0041FA50: push ebp
  loc_0041FA51: mov ebp, esp
  loc_0041FA53: sub esp, 0000000Ch
  loc_0041FA56: push 00401926h ; __vbaExceptHandler
  loc_0041FA5B: mov eax, fs:[00000000h]
  loc_0041FA61: push eax
  loc_0041FA62: mov fs:[00000000h], esp
  loc_0041FA69: sub esp, 0000007Ch
  loc_0041FA6C: push ebx
  loc_0041FA6D: push esi
  loc_0041FA6E: push edi
  loc_0041FA6F: mov var_C, esp
  loc_0041FA72: mov var_8, 00401348h
  loc_0041FA79: mov esi, Me
  loc_0041FA7C: mov eax, esi
  loc_0041FA7E: and eax, 00000001h
  loc_0041FA81: mov var_4, eax
  loc_0041FA84: and esi, FFFFFFFEh
  loc_0041FA87: push esi
  loc_0041FA88: mov Me, esi
  loc_0041FA8B: mov ecx, [esi]
  loc_0041FA8D: call [ecx+00000004h]
  loc_0041FA90: mov ebx, KeyAscii
  loc_0041FA93: lea edx, var_24
  loc_0041FA96: xor edi, edi
  loc_0041FA98: push ebx
  loc_0041FA99: push edx
  loc_0041FA9A: mov var_24, edi
  loc_0041FA9D: mov var_34, edi
  loc_0041FAA0: mov var_44, edi
  loc_0041FAA3: mov var_54, edi
  loc_0041FAA6: mov var_64, edi
  loc_0041FAA9: mov var_74, edi
  loc_0041FAAC: mov var_84, edi
  loc_0041FAB2: call 0041A480h
  loc_0041FAB7: lea eax, var_24
  loc_0041FABA: push eax
  loc_0041FABB: call [004010C4h] ; __vbaI2Var
  loc_0041FAC1: lea ecx, var_24
  loc_0041FAC4: mov [ebx], ax
  loc_0041FAC7: call [00401010h] ; __vbaFreeVar
  loc_0041FACD: push 0040DD3Ch ; "."
  loc_0041FAD2: call [00401024h] ; rtcAnsiValueBstr
  loc_0041FAD8: mov edx, [esi]
  loc_0041FADA: xor ecx, ecx
  loc_0041FADC: cmp [ebx], ax
  loc_0041FADF: push esi
  loc_0041FAE0: mov var_84, 0000000Bh
  loc_0041FAEA: setz cl
  loc_0041FAED: neg ecx
  loc_0041FAEF: mov var_7C, cx
  loc_0041FAF3: call [edx+00000310h]
  loc_0041FAF9: mov var_1C, eax
  loc_0041FAFC: lea eax, var_84
  loc_0041FB02: push eax
  loc_0041FB03: lea ecx, var_24
  loc_0041FB06: push 00000001h
  loc_0041FB08: lea edx, var_64
  loc_0041FB0B: push ecx
  loc_0041FB0C: push edx
  loc_0041FB0D: lea eax, var_34
  loc_0041FB10: push edi
  loc_0041FB11: push eax
  loc_0041FB12: mov var_24, 00000009h
  loc_0041FB19: mov var_5C, 0040DD3Ch ; "."
  loc_0041FB20: mov var_64, 00000008h
  loc_0041FB27: mov var_6C, edi
  loc_0041FB2A: mov var_74, 00008002h
  loc_0041FB31: call [004010B4h] ; __vbaInStrVar
  loc_0041FB37: lea ecx, var_74
  loc_0041FB3A: push eax
  loc_0041FB3B: lea edx, var_44
  loc_0041FB3E: push ecx
  loc_0041FB3F: push edx
  loc_0041FB40: call [00401060h] ; __vbaVarCmpGt
  loc_0041FB46: push eax
  loc_0041FB47: lea eax, var_54
  loc_0041FB4A: push eax
  loc_0041FB4B: call [00401094h] ; __vbaVarAnd
  loc_0041FB51: push eax
  loc_0041FB52: call [00401058h] ; __vbaBoolVarNull
  loc_0041FB58: lea ecx, var_84
  loc_0041FB5E: mov esi, eax
  loc_0041FB60: lea edx, var_34
  loc_0041FB63: push ecx
  loc_0041FB64: lea eax, var_24
  loc_0041FB67: push edx
  loc_0041FB68: push eax
  loc_0041FB69: push 00000003h
  loc_0041FB6B: call [00401018h] ; __vbaFreeVarList
  loc_0041FB71: add esp, 00000010h
  loc_0041FB74: cmp si, di
  loc_0041FB77: jz 0041FB7Ch
  loc_0041FB79: mov [ebx], di
  loc_0041FB7C: mov var_4, edi
  loc_0041FB7F: push 0041FBA3h
  loc_0041FB84: jmp 0041FBA2h
  loc_0041FB86: lea ecx, var_54
  loc_0041FB89: lea edx, var_44
  loc_0041FB8C: push ecx
  loc_0041FB8D: lea eax, var_34
  loc_0041FB90: push edx
  loc_0041FB91: lea ecx, var_24
  loc_0041FB94: push eax
  loc_0041FB95: push ecx
  loc_0041FB96: push 00000004h
  loc_0041FB98: call [00401018h] ; __vbaFreeVarList
  loc_0041FB9E: add esp, 00000014h
  loc_0041FBA1: ret
  loc_0041FBA2: ret
  loc_0041FBA3: mov eax, Me
  loc_0041FBA6: push eax
  loc_0041FBA7: mov edx, [eax]
  loc_0041FBA9: call [edx+00000008h]
  loc_0041FBAC: mov eax, var_4
  loc_0041FBAF: mov ecx, var_14
  loc_0041FBB2: pop edi
  loc_0041FBB3: pop esi
  loc_0041FBB4: mov fs:[00000000h], ecx
  loc_0041FBBB: pop ebx
  loc_0041FBBC: mov esp, ebp
  loc_0041FBBE: pop ebp
  loc_0041FBBF: retn 0008h
End Sub

Private Sub cmdChange_Click() '41EF30
  loc_0041EF30: push ebp
  loc_0041EF31: mov ebp, esp
  loc_0041EF33: sub esp, 0000000Ch
  loc_0041EF36: push 00401926h ; __vbaExceptHandler
  loc_0041EF3B: mov eax, fs:[00000000h]
  loc_0041EF41: push eax
  loc_0041EF42: mov fs:[00000000h], esp
  loc_0041EF49: sub esp, 000000D8h
  loc_0041EF4F: push ebx
  loc_0041EF50: push esi
  loc_0041EF51: push edi
  loc_0041EF52: mov var_C, esp
  loc_0041EF55: mov var_8, 00401328h
  loc_0041EF5C: mov ebx, Me
  loc_0041EF5F: mov eax, ebx
  loc_0041EF61: and eax, 00000001h
  loc_0041EF64: mov var_4, eax
  loc_0041EF67: and ebx, FFFFFFFEh
  loc_0041EF6A: push ebx
  loc_0041EF6B: mov Me, ebx
  loc_0041EF6E: mov ecx, [ebx]
  loc_0041EF70: call [ecx+00000004h]
  loc_0041EF73: mov edx, [ebx]
  loc_0041EF75: xor esi, esi
  loc_0041EF77: push ebx
  loc_0041EF78: mov var_1C, esi
  loc_0041EF7B: mov var_20, esi
  loc_0041EF7E: mov var_24, esi
  loc_0041EF81: mov var_28, esi
  loc_0041EF84: mov var_2C, esi
  loc_0041EF87: mov var_30, esi
  loc_0041EF8A: mov var_34, esi
  loc_0041EF8D: mov var_38, esi
  loc_0041EF90: mov var_3C, esi
  loc_0041EF93: mov var_40, esi
  loc_0041EF96: mov var_44, esi
  loc_0041EF99: mov var_48, esi
  loc_0041EF9C: mov var_4C, esi
  loc_0041EF9F: mov var_5C, esi
  loc_0041EFA2: mov var_6C, esi
  loc_0041EFA5: mov var_7C, esi
  loc_0041EFA8: mov var_8C, esi
  loc_0041EFAE: mov var_9C, esi
  loc_0041EFB4: mov var_AC, esi
  loc_0041EFBA: call [edx+00000310h]
  loc_0041EFC0: mov var_54, eax
  loc_0041EFC3: lea eax, var_5C
  loc_0041EFC6: lea ecx, var_6C
  loc_0041EFC9: push eax
  loc_0041EFCA: push ecx
  loc_0041EFCB: mov var_5C, 00000009h
  loc_0041EFD2: call [00401050h] ; rtcTrimVar
  loc_0041EFD8: lea edx, var_6C
  loc_0041EFDB: lea eax, var_9C
  loc_0041EFE1: push edx
  loc_0041EFE2: push eax
  loc_0041EFE3: mov var_94, 0040F040h
  loc_0041EFED: mov var_9C, 00008008h
  loc_0041EFF7: call [00401070h] ; __vbaVarTstEq
  loc_0041EFFD: mov edi, [00401018h] ; __vbaFreeVarList
  loc_0041F003: lea ecx, var_6C
  loc_0041F006: lea edx, var_5C
  loc_0041F009: push ecx
  loc_0041F00A: push edx
  loc_0041F00B: push 00000002h
  loc_0041F00D: mov var_D0, ax
  loc_0041F014: call edi
  loc_0041F016: add esp, 0000000Ch
  loc_0041F019: cmp var_D0, si
  loc_0041F020: jz 0041F0FFh
  loc_0041F026: mov eax, [ebx]
  loc_0041F028: push ebx
  loc_0041F029: call [eax+00000310h]
  loc_0041F02F: lea ecx, var_4C
  loc_0041F032: push eax
  loc_0041F033: push ecx
  loc_0041F034: call [0040103Ch] ; __vbaObjSet
  loc_0041F03A: mov ebx, eax
  loc_0041F03C: push ebx
  loc_0041F03D: mov edx, [ebx]
  loc_0041F03F: call [edx+00000204h]
  loc_0041F045: cmp eax, esi
  loc_0041F047: fnclex
  loc_0041F049: jge 0041F05Dh
  loc_0041F04B: push 00000204h
  loc_0041F050: push 0040F02Ch
  loc_0041F055: push ebx
  loc_0041F056: push eax
  loc_0041F057: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F05D: lea ecx, var_4C
  loc_0041F060: call [00401114h] ; __vbaFreeObj
  loc_0041F066: mov ebx, [004010F4h] ; __vbaVarDup
  loc_0041F06C: mov ecx, 80020004h
  loc_0041F071: mov var_84, ecx
  loc_0041F077: mov eax, 0000000Ah
  loc_0041F07C: mov var_74, ecx
  loc_0041F07F: lea edx, var_AC
  loc_0041F085: lea ecx, var_6C
  loc_0041F088: mov var_8C, eax
  loc_0041F08E: mov var_7C, eax
  loc_0041F091: mov var_A4, 0040F9DCh ; "Check Slope"
  loc_0041F09B: mov var_AC, 00000008h
  loc_0041F0A5: call ebx
  loc_0041F0A7: lea edx, var_9C
  loc_0041F0AD: lea ecx, var_5C
  loc_0041F0B0: mov var_94, 0040F3A0h ; "Please enter a value for the estimate of the slope."
  loc_0041F0BA: mov var_9C, 00000008h
  loc_0041F0C4: call ebx
  loc_0041F0C6: lea eax, var_8C
  loc_0041F0CC: lea ecx, var_7C
  loc_0041F0CF: push eax
  loc_0041F0D0: lea edx, var_6C
  loc_0041F0D3: push ecx
  loc_0041F0D4: push edx
  loc_0041F0D5: lea eax, var_5C
  loc_0041F0D8: push esi
  loc_0041F0D9: push eax
  loc_0041F0DA: call [00401038h] ; rtcMsgBox
  loc_0041F0E0: lea ecx, var_8C
  loc_0041F0E6: lea edx, var_7C
  loc_0041F0E9: push ecx
  loc_0041F0EA: lea eax, var_6C
  loc_0041F0ED: push edx
  loc_0041F0EE: lea ecx, var_5C
  loc_0041F0F1: push eax
  loc_0041F0F2: push ecx
  loc_0041F0F3: push 00000004h
  loc_0041F0F5: call edi
  loc_0041F0F7: add esp, 00000014h
  loc_0041F0FA: jmp 0041F807h
  loc_0041F0FF: mov edx, [ebx]
  loc_0041F101: push ebx
  loc_0041F102: call [edx+00000318h]
  loc_0041F108: mov var_54, eax
  loc_0041F10B: lea eax, var_5C
  loc_0041F10E: lea ecx, var_6C
  loc_0041F111: push eax
  loc_0041F112: push ecx
  loc_0041F113: mov var_5C, 00000009h
  loc_0041F11A: call [00401050h] ; rtcTrimVar
  loc_0041F120: lea edx, var_6C
  loc_0041F123: lea eax, var_9C
  loc_0041F129: push edx
  loc_0041F12A: push eax
  loc_0041F12B: mov var_94, 0040F040h
  loc_0041F135: mov var_9C, 00008008h
  loc_0041F13F: call [00401070h] ; __vbaVarTstEq
  loc_0041F145: lea ecx, var_6C
  loc_0041F148: lea edx, var_5C
  loc_0041F14B: push ecx
  loc_0041F14C: push edx
  loc_0041F14D: push 00000002h
  loc_0041F14F: mov var_D0, ax
  loc_0041F156: call edi
  loc_0041F158: add esp, 0000000Ch
  loc_0041F15B: cmp var_D0, si
  loc_0041F162: jz 0041F243h
  loc_0041F168: mov ecx, 80020004h
  loc_0041F16D: mov eax, 0000000Ah
  loc_0041F172: mov var_84, ecx
  loc_0041F178: mov var_74, ecx
  loc_0041F17B: lea edx, var_AC
  loc_0041F181: lea ecx, var_6C
  loc_0041F184: mov var_8C, eax
  loc_0041F18A: mov var_7C, eax
  loc_0041F18D: mov var_A4, 0040FE20h ; "Check Intercept"
  loc_0041F197: mov var_AC, 00000008h
  loc_0041F1A1: call [004010F4h] ; __vbaVarDup
  loc_0041F1A7: lea edx, var_9C
  loc_0041F1AD: lea ecx, var_5C
  loc_0041F1B0: mov var_94, 0040FDACh ; "Please enter a value for the estimate of the intercept."
  loc_0041F1BA: mov var_9C, 00000008h
  loc_0041F1C4: call [004010F4h] ; __vbaVarDup
  loc_0041F1CA: lea eax, var_8C
  loc_0041F1D0: lea ecx, var_7C
  loc_0041F1D3: push eax
  loc_0041F1D4: lea edx, var_6C
  loc_0041F1D7: push ecx
  loc_0041F1D8: push edx
  loc_0041F1D9: lea eax, var_5C
  loc_0041F1DC: push esi
  loc_0041F1DD: push eax
  loc_0041F1DE: call [00401038h] ; rtcMsgBox
  loc_0041F1E4: lea ecx, var_8C
  loc_0041F1EA: lea edx, var_7C
  loc_0041F1ED: push ecx
  loc_0041F1EE: lea eax, var_6C
  loc_0041F1F1: push edx
  loc_0041F1F2: lea ecx, var_5C
  loc_0041F1F5: push eax
  loc_0041F1F6: push ecx
  loc_0041F1F7: push 00000004h
  loc_0041F1F9: call edi
  loc_0041F1FB: mov edx, [ebx]
  loc_0041F1FD: add esp, 00000014h
  loc_0041F200: push ebx
  loc_0041F201: call [edx+00000318h]
  loc_0041F207: push eax
  loc_0041F208: lea eax, var_4C
  loc_0041F20B: push eax
  loc_0041F20C: call [0040103Ch] ; __vbaObjSet
  loc_0041F212: mov edi, eax
  loc_0041F214: push edi
  loc_0041F215: mov ecx, [edi]
  loc_0041F217: call [ecx+00000204h]
  loc_0041F21D: cmp eax, esi
  loc_0041F21F: fnclex
  loc_0041F221: jge 0041F235h
  loc_0041F223: push 00000204h
  loc_0041F228: push 0040F02Ch
  loc_0041F22D: push edi
  loc_0041F22E: push eax
  loc_0041F22F: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F235: lea ecx, var_4C
  loc_0041F238: call [00401114h] ; __vbaFreeObj
  loc_0041F23E: jmp 0041F807h
  loc_0041F243: mov edx, [ebx]
  loc_0041F245: push ebx
  loc_0041F246: call [edx+00000310h]
  loc_0041F24C: mov var_54, eax
  loc_0041F24F: lea eax, var_5C
  loc_0041F252: lea ecx, var_6C
  loc_0041F255: push eax
  loc_0041F256: push ecx
  loc_0041F257: mov var_5C, 00000009h
  loc_0041F25E: call [00401050h] ; rtcTrimVar
  loc_0041F264: lea edx, var_6C
  loc_0041F267: lea eax, var_1C
  loc_0041F26A: push edx
  loc_0041F26B: push eax
  loc_0041F26C: call [004010B8h] ; __vbaStrVarVal
  loc_0041F272: push eax
  loc_0041F273: call [00401118h] ; rtcR8ValFromBstr
  loc_0041F279: call [00401054h] ; __vbaFpR8
  loc_0041F27F: fcomp real8 ptr [004011F0h]
  loc_0041F285: mov var_E8, 00000001h
  loc_0041F28F: fnstsw ax
  loc_0041F291: test ah, 41h
  loc_0041F294: jz 0041F29Ch
  loc_0041F296: mov var_E8, esi
  loc_0041F29C: fld real4 ptr [0043006Ch]
  loc_0041F2A2: fcomp real4 ptr [004011E8h]
  loc_0041F2A8: fnstsw ax
  loc_0041F2AA: test ah, 01h
  loc_0041F2AD: jz 0041F2B6h
  loc_0041F2AF: mov edi, 00000001h
  loc_0041F2B4: jmp 0041F2B8h
  loc_0041F2B6: xor edi, edi
  loc_0041F2B8: mov ecx, [ebx]
  loc_0041F2BA: push ebx
  loc_0041F2BB: call [ecx+00000310h]
  loc_0041F2C1: mov var_74, eax
  loc_0041F2C4: lea edx, var_7C
  loc_0041F2C7: lea eax, var_8C
  loc_0041F2CD: push edx
  loc_0041F2CE: push eax
  loc_0041F2CF: mov var_7C, 00000009h
  loc_0041F2D6: call [00401050h] ; rtcTrimVar
  loc_0041F2DC: lea ecx, var_8C
  loc_0041F2E2: lea edx, var_20
  loc_0041F2E5: push ecx
  loc_0041F2E6: push edx
  loc_0041F2E7: call [004010B8h] ; __vbaStrVarVal
  loc_0041F2ED: push eax
  loc_0041F2EE: call [00401118h] ; rtcR8ValFromBstr
  loc_0041F2F4: call [00401054h] ; __vbaFpR8
  loc_0041F2FA: fcomp real8 ptr [004011F0h]
  loc_0041F300: mov var_EC, 00000001h
  loc_0041F30A: fnstsw ax
  loc_0041F30C: test ah, 01h
  loc_0041F30F: jnz 0041F317h
  loc_0041F311: mov var_EC, esi
  loc_0041F317: fld real4 ptr [0043006Ch]
  loc_0041F31D: fcomp real4 ptr [004011E8h]
  loc_0041F323: fnstsw ax
  loc_0041F325: test ah, 41h
  loc_0041F328: jnz 0041F32Fh
  loc_0041F32A: mov esi, 00000001h
  loc_0041F32F: lea eax, var_20
  loc_0041F332: lea ecx, var_1C
  loc_0041F335: push eax
  loc_0041F336: push ecx
  loc_0041F337: push 00000002h
  loc_0041F339: call [004010E4h] ; __vbaFreeStrList
  loc_0041F33F: lea edx, var_8C
  loc_0041F345: lea eax, var_7C
  loc_0041F348: push edx
  loc_0041F349: lea ecx, var_6C
  loc_0041F34C: push eax
  loc_0041F34D: lea edx, var_5C
  loc_0041F350: push ecx
  loc_0041F351: push edx
  loc_0041F352: push 00000004h
  loc_0041F354: call [00401018h] ; __vbaFreeVarList
  loc_0041F35A: mov eax, var_EC
  loc_0041F360: add esp, 00000020h
  loc_0041F363: neg esi
  loc_0041F365: neg eax
  loc_0041F367: and esi, eax
  loc_0041F369: mov eax, var_E8
  loc_0041F36F: neg edi
  loc_0041F371: neg eax
  loc_0041F373: and edi, eax
  loc_0041F375: or esi, edi
  loc_0041F377: test si, si
  loc_0041F37A: jz 0041F4EBh
  loc_0041F380: mov ecx, 80020004h
  loc_0041F385: mov eax, 0000000Ah
  loc_0041F38A: mov var_84, ecx
  loc_0041F390: mov var_74, ecx
  loc_0041F393: lea edx, var_9C
  loc_0041F399: lea ecx, var_6C
  loc_0041F39C: mov var_8C, eax
  loc_0041F3A2: mov var_7C, eax
  loc_0041F3A5: mov var_94, 0040F50Ch ; "Check Corr"
  loc_0041F3AF: mov var_9C, 00000008h
  loc_0041F3B9: call [004010F4h] ; __vbaVarDup
  loc_0041F3BF: mov esi, [0040102Ch] ; __vbaStrCat
  loc_0041F3C5: push 0040F674h ; "Warning, the correlation coefficient MUST have the same sign as the estimated slope"
  loc_0041F3CA: push 0040F720h ; vbCrLf
  loc_0041F3CF: call __vbaStrCat
  loc_0041F3D1: mov edi, [004010FCh] ; __vbaStrMove
  loc_0041F3D7: mov edx, eax
  loc_0041F3D9: lea ecx, var_1C
  loc_0041F3DC: call edi
  loc_0041F3DE: push eax
  loc_0041F3DF: push 0040FE44h ; "Currently they do not. The correlation coefficient = "
  loc_0041F3E4: call __vbaStrCat
  loc_0041F3E6: mov edx, eax
  loc_0041F3E8: lea ecx, var_20
  loc_0041F3EB: call edi
  loc_0041F3ED: push eax
  loc_0041F3EE: mov eax, [0043006Ch]
  loc_0041F3F3: push eax
  loc_0041F3F4: call [0040107Ch] ; __vbaStrR4
  loc_0041F3FA: mov edx, eax
  loc_0041F3FC: lea ecx, var_24
  loc_0041F3FF: call edi
  loc_0041F401: push eax
  loc_0041F402: call __vbaStrCat
  loc_0041F404: mov edx, eax
  loc_0041F406: lea ecx, var_28
  loc_0041F409: call edi
  loc_0041F40B: push eax
  loc_0041F40C: push 0040F720h ; vbCrLf
  loc_0041F411: call __vbaStrCat
  loc_0041F413: mov edx, eax
  loc_0041F415: lea ecx, var_2C
  loc_0041F418: call edi
  loc_0041F41A: push eax
  loc_0041F41B: push 0040F720h ; vbCrLf
  loc_0041F420: call __vbaStrCat
  loc_0041F422: mov edx, eax
  loc_0041F424: lea ecx, var_30
  loc_0041F427: call edi
  loc_0041F429: push eax
  loc_0041F42A: push 0040F78Ch ; "Do you wish to continue with this value?"
  loc_0041F42F: call __vbaStrCat
  loc_0041F431: lea ecx, var_8C
  loc_0041F437: mov var_54, eax
  loc_0041F43A: lea edx, var_7C
  loc_0041F43D: push ecx
  loc_0041F43E: lea eax, var_6C
  loc_0041F441: push edx
  loc_0041F442: push eax
  loc_0041F443: lea ecx, var_5C
  loc_0041F446: push 00000004h
  loc_0041F448: push ecx
  loc_0041F449: mov var_5C, 00000008h
  loc_0041F450: call [00401038h] ; rtcMsgBox
  loc_0041F456: mov ecx, eax
  loc_0041F458: call [00401078h] ; __vbaI2I4
  loc_0041F45E: mov var_18, eax
  loc_0041F461: lea edx, var_30
  loc_0041F464: lea eax, var_2C
  loc_0041F467: push edx
  loc_0041F468: lea ecx, var_28
  loc_0041F46B: push eax
  loc_0041F46C: lea edx, var_24
  loc_0041F46F: push ecx
  loc_0041F470: lea eax, var_20
  loc_0041F473: push edx
  loc_0041F474: lea ecx, var_1C
  loc_0041F477: push eax
  loc_0041F478: push ecx
  loc_0041F479: push 00000006h
  loc_0041F47B: call [004010E4h] ; __vbaFreeStrList
  loc_0041F481: lea edx, var_8C
  loc_0041F487: lea eax, var_7C
  loc_0041F48A: push edx
  loc_0041F48B: lea ecx, var_6C
  loc_0041F48E: push eax
  loc_0041F48F: lea edx, var_5C
  loc_0041F492: push ecx
  loc_0041F493: push edx
  loc_0041F494: push 00000004h
  loc_0041F496: call [00401018h] ; __vbaFreeVarList
  loc_0041F49C: add esp, 00000030h
  loc_0041F49F: cmp var_18, 0007h
  loc_0041F4A4: jnz 0041F4F7h
  loc_0041F4A6: mov eax, [ebx]
  loc_0041F4A8: push ebx
  loc_0041F4A9: call [eax+00000310h]
  loc_0041F4AF: lea ecx, var_4C
  loc_0041F4B2: push eax
  loc_0041F4B3: push ecx
  loc_0041F4B4: call [0040103Ch] ; __vbaObjSet
  loc_0041F4BA: mov esi, eax
  loc_0041F4BC: push esi
  loc_0041F4BD: mov edx, [esi]
  loc_0041F4BF: call [edx+00000204h]
  loc_0041F4C5: test eax, eax
  loc_0041F4C7: fnclex
  loc_0041F4C9: jge 0041F4DDh
  loc_0041F4CB: push 00000204h
  loc_0041F4D0: push 0040F02Ch
  loc_0041F4D5: push esi
  loc_0041F4D6: push eax
  loc_0041F4D7: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F4DD: lea ecx, var_4C
  loc_0041F4E0: call [00401114h] ; __vbaFreeObj
  loc_0041F4E6: jmp 0041F805h
  loc_0041F4EB: mov esi, [0040102Ch] ; __vbaStrCat
  loc_0041F4F1: mov edi, [004010FCh] ; __vbaStrMove
  loc_0041F4F7: mov eax, [ebx]
  loc_0041F4F9: push ebx
  loc_0041F4FA: call [eax+00000310h]
  loc_0041F500: lea ecx, var_4C
  loc_0041F503: push eax
  loc_0041F504: push ecx
  loc_0041F505: call [0040103Ch] ; __vbaObjSet
  loc_0041F50B: mov edx, [eax]
  loc_0041F50D: lea ecx, var_1C
  loc_0041F510: push ecx
  loc_0041F511: push eax
  loc_0041F512: mov var_D0, eax
  loc_0041F518: call [edx+000000A0h]
  loc_0041F51E: test eax, eax
  loc_0041F520: fnclex
  loc_0041F522: jge 0041F53Ch
  loc_0041F524: mov edx, var_D0
  loc_0041F52A: push 000000A0h
  loc_0041F52F: push 0040F02Ch
  loc_0041F534: push edx
  loc_0041F535: push eax
  loc_0041F536: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F53C: mov eax, var_1C
  loc_0041F53F: lea ecx, var_6C
  loc_0041F542: mov var_54, eax
  loc_0041F545: lea eax, var_5C
  loc_0041F548: push eax
  loc_0041F549: push ecx
  loc_0041F54A: mov var_1C, 00000000h
  loc_0041F551: mov var_5C, 00000008h
  loc_0041F558: call [00401050h] ; rtcTrimVar
  loc_0041F55E: lea edx, var_6C
  loc_0041F561: push edx
  loc_0041F562: call [004010A0h] ; __vbaR4ErrVar
  loc_0041F568: fstp real4 ptr [00430060h]
  loc_0041F56E: lea ecx, var_4C
  loc_0041F571: call [00401114h] ; __vbaFreeObj
  loc_0041F577: lea eax, var_6C
  loc_0041F57A: lea ecx, var_6C
  loc_0041F57D: push eax
  loc_0041F57E: lea edx, var_5C
  loc_0041F581: push ecx
  loc_0041F582: push edx
  loc_0041F583: push 00000003h
  loc_0041F585: call [00401018h] ; __vbaFreeVarList
  loc_0041F58B: mov eax, [ebx]
  loc_0041F58D: add esp, 00000010h
  loc_0041F590: push ebx
  loc_0041F591: call [eax+00000318h]
  loc_0041F597: lea ecx, var_4C
  loc_0041F59A: push eax
  loc_0041F59B: push ecx
  loc_0041F59C: call [0040103Ch] ; __vbaObjSet
  loc_0041F5A2: mov ebx, eax
  loc_0041F5A4: lea eax, var_1C
  loc_0041F5A7: push eax
  loc_0041F5A8: push ebx
  loc_0041F5A9: mov edx, [ebx]
  loc_0041F5AB: call [edx+000000A0h]
  loc_0041F5B1: test eax, eax
  loc_0041F5B3: fnclex
  loc_0041F5B5: jge 0041F5C9h
  loc_0041F5B7: push 000000A0h
  loc_0041F5BC: push 0040F02Ch
  loc_0041F5C1: push ebx
  loc_0041F5C2: push eax
  loc_0041F5C3: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F5C9: mov eax, var_1C
  loc_0041F5CC: lea ecx, var_5C
  loc_0041F5CF: lea edx, var_6C
  loc_0041F5D2: push ecx
  loc_0041F5D3: push edx
  loc_0041F5D4: mov var_1C, 00000000h
  loc_0041F5DB: mov var_54, eax
  loc_0041F5DE: mov var_5C, 00000008h
  loc_0041F5E5: call [00401050h] ; rtcTrimVar
  loc_0041F5EB: lea eax, var_6C
  loc_0041F5EE: push eax
  loc_0041F5EF: call [004010A0h] ; __vbaR4ErrVar
  loc_0041F5F5: fstp real4 ptr [00430074h]
  loc_0041F5FB: lea ecx, var_4C
  loc_0041F5FE: call [00401114h] ; __vbaFreeObj
  loc_0041F604: mov ebx, [00401018h] ; __vbaFreeVarList
  loc_0041F60A: lea ecx, var_6C
  loc_0041F60D: lea edx, var_6C
  loc_0041F610: push ecx
  loc_0041F611: lea eax, var_5C
  loc_0041F614: push edx
  loc_0041F615: push eax
  loc_0041F616: push 00000003h
  loc_0041F618: call ebx
  loc_0041F61A: mov ecx, [0043002Ch]
  loc_0041F620: add esp, 00000010h
  loc_0041F623: push ecx
  loc_0041F624: call [004010D0h] ; __vbaR8Str
  loc_0041F62A: fmul st0, real4 ptr [00430060h]
  loc_0041F630: fadd st0, real4 ptr [00430074h]
  loc_0041F636: mov ecx, 80020004h
  loc_0041F63B: lea edx, var_9C
  loc_0041F641: fstp real4 ptr [00430038h]
  loc_0041F647: fnstsw ax
  loc_0041F649: test al, 0Dh
  loc_0041F64B: jnz 0041F895h
  loc_0041F651: mov var_84, ecx
  loc_0041F657: mov eax, 0000000Ah
  loc_0041F65C: mov var_74, ecx
  loc_0041F65F: lea ecx, var_6C
  loc_0041F662: mov var_8C, eax
  loc_0041F668: mov var_7C, eax
  loc_0041F66B: mov var_94, 0040FF38h ; "Change in Yhat"
  loc_0041F675: mov var_9C, 00000008h
  loc_0041F67F: call [004010F4h] ; __vbaVarDup
  loc_0041F685: mov edx, [00430010h]
  loc_0041F68B: push 0040FEB4h ; "Note the estimated mean "
  loc_0041F690: push edx
  loc_0041F691: call __vbaStrCat
  loc_0041F693: mov edx, eax
  loc_0041F695: lea ecx, var_1C
  loc_0041F698: call edi
  loc_0041F69A: push eax
  loc_0041F69B: push 0040FEECh ; " is now "
  loc_0041F6A0: call __vbaStrCat
  loc_0041F6A2: mov edx, eax
  loc_0041F6A4: lea ecx, var_20
  loc_0041F6A7: call edi
  loc_0041F6A9: push eax
  loc_0041F6AA: mov eax, [00430038h]
  loc_0041F6AF: push eax
  loc_0041F6B0: call [0040107Ch] ; __vbaStrR4
  loc_0041F6B6: mov edx, eax
  loc_0041F6B8: lea ecx, var_24
  loc_0041F6BB: call edi
  loc_0041F6BD: push eax
  loc_0041F6BE: call __vbaStrCat
  loc_0041F6C0: mov edx, eax
  loc_0041F6C2: lea ecx, var_28
  loc_0041F6C5: call edi
  loc_0041F6C7: push eax
  loc_0041F6C8: push 0040F028h
  loc_0041F6CD: call __vbaStrCat
  loc_0041F6CF: mov edx, eax
  loc_0041F6D1: lea ecx, var_2C
  loc_0041F6D4: call edi
  loc_0041F6D6: mov ecx, [00430014h]
  loc_0041F6DC: push eax
  loc_0041F6DD: push ecx
  loc_0041F6DE: call __vbaStrCat
  loc_0041F6E0: mov edx, eax
  loc_0041F6E2: lea ecx, var_30
  loc_0041F6E5: call edi
  loc_0041F6E7: push eax
  loc_0041F6E8: push 0040FF04h ; " when "
  loc_0041F6ED: call __vbaStrCat
  loc_0041F6EF: mov edx, eax
  loc_0041F6F1: lea ecx, var_34
  loc_0041F6F4: call edi
  loc_0041F6F6: mov edx, [00430018h]
  loc_0041F6FC: push eax
  loc_0041F6FD: push edx
  loc_0041F6FE: call __vbaStrCat
  loc_0041F700: mov edx, eax
  loc_0041F702: lea ecx, var_38
  loc_0041F705: call edi
  loc_0041F707: push eax
  loc_0041F708: push 0040FF18h ; " is equal to "
  loc_0041F70D: call __vbaStrCat
  loc_0041F70F: mov edx, eax
  loc_0041F711: lea ecx, var_3C
  loc_0041F714: call edi
  loc_0041F716: push eax
  loc_0041F717: mov eax, [0043002Ch]
  loc_0041F71C: push eax
  loc_0041F71D: call __vbaStrCat
  loc_0041F71F: mov edx, eax
  loc_0041F721: lea ecx, var_40
  loc_0041F724: call edi
  loc_0041F726: push eax
  loc_0041F727: push 0040F028h
  loc_0041F72C: call __vbaStrCat
  loc_0041F72E: mov edx, eax
  loc_0041F730: lea ecx, var_44
  loc_0041F733: call edi
  loc_0041F735: mov ecx, [0043001Ch]
  loc_0041F73B: push eax
  loc_0041F73C: push ecx
  loc_0041F73D: call __vbaStrCat
  loc_0041F73F: mov edx, eax
  loc_0041F741: lea ecx, var_48
  loc_0041F744: call edi
  loc_0041F746: push eax
  loc_0041F747: push 0040DD3Ch ; "."
  loc_0041F74C: call __vbaStrCat
  loc_0041F74E: mov var_54, eax
  loc_0041F751: lea edx, var_8C
  loc_0041F757: lea eax, var_7C
  loc_0041F75A: push edx
  loc_0041F75B: lea ecx, var_6C
  loc_0041F75E: push eax
  loc_0041F75F: push ecx
  loc_0041F760: lea edx, var_5C
  loc_0041F763: push 00000000h
  loc_0041F765: push edx
  loc_0041F766: mov var_5C, 00000008h
  loc_0041F76D: call [00401038h] ; rtcMsgBox
  loc_0041F773: lea eax, var_48
  loc_0041F776: lea ecx, var_44
  loc_0041F779: push eax
  loc_0041F77A: lea edx, var_40
  loc_0041F77D: push ecx
  loc_0041F77E: lea eax, var_3C
  loc_0041F781: push edx
  loc_0041F782: lea ecx, var_38
  loc_0041F785: push eax
  loc_0041F786: lea edx, var_34
  loc_0041F789: push ecx
  loc_0041F78A: lea eax, var_30
  loc_0041F78D: push edx
  loc_0041F78E: lea ecx, var_2C
  loc_0041F791: push eax
  loc_0041F792: lea edx, var_28
  loc_0041F795: push ecx
  loc_0041F796: lea eax, var_24
  loc_0041F799: push edx
  loc_0041F79A: lea ecx, var_20
  loc_0041F79D: push eax
  loc_0041F79E: lea edx, var_1C
  loc_0041F7A1: push ecx
  loc_0041F7A2: push edx
  loc_0041F7A3: push 0000000Ch
  loc_0041F7A5: call [004010E4h] ; __vbaFreeStrList
  loc_0041F7AB: lea eax, var_8C
  loc_0041F7B1: lea ecx, var_7C
  loc_0041F7B4: push eax
  loc_0041F7B5: lea edx, var_6C
  loc_0041F7B8: push ecx
  loc_0041F7B9: lea eax, var_5C
  loc_0041F7BC: push edx
  loc_0041F7BD: push eax
  loc_0041F7BE: push 00000004h
  loc_0041F7C0: call ebx
  loc_0041F7C2: mov eax, [00430150h]
  loc_0041F7C7: add esp, 00000048h
  loc_0041F7CA: test eax, eax
  loc_0041F7CC: jnz 0041F7DEh
  loc_0041F7CE: push 00430150h
  loc_0041F7D3: push 00403F48h
  loc_0041F7D8: call [004010D4h] ; __vbaNew2
  loc_0041F7DE: mov esi, [00430150h]
  loc_0041F7E4: push esi
  loc_0041F7E5: mov ecx, [esi]
  loc_0041F7E7: call [ecx+000002B4h]
  loc_0041F7ED: test eax, eax
  loc_0041F7EF: fnclex
  loc_0041F7F1: jge 0041F805h
  loc_0041F7F3: push 000002B4h
  loc_0041F7F8: push 0040FD4Ch
  loc_0041F7FD: push esi
  loc_0041F7FE: push eax
  loc_0041F7FF: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F805: xor esi, esi
  loc_0041F807: mov var_4, esi
  loc_0041F80A: fwait
  loc_0041F80B: push 0041F876h
  loc_0041F810: jmp 0041F875h
  loc_0041F812: lea edx, var_48
  loc_0041F815: lea eax, var_44
  loc_0041F818: push edx
  loc_0041F819: lea ecx, var_40
  loc_0041F81C: push eax
  loc_0041F81D: lea edx, var_3C
  loc_0041F820: push ecx
  loc_0041F821: lea eax, var_38
  loc_0041F824: push edx
  loc_0041F825: lea ecx, var_34
  loc_0041F828: push eax
  loc_0041F829: lea edx, var_30
  loc_0041F82C: push ecx
  loc_0041F82D: lea eax, var_2C
  loc_0041F830: push edx
  loc_0041F831: lea ecx, var_28
  loc_0041F834: push eax
  loc_0041F835: lea edx, var_24
  loc_0041F838: push ecx
  loc_0041F839: lea eax, var_20
  loc_0041F83C: push edx
  loc_0041F83D: lea ecx, var_1C
  loc_0041F840: push eax
  loc_0041F841: push ecx
  loc_0041F842: push 0000000Ch
  loc_0041F844: call [004010E4h] ; __vbaFreeStrList
  loc_0041F84A: add esp, 00000034h
  loc_0041F84D: lea ecx, var_4C
  loc_0041F850: call [00401114h] ; __vbaFreeObj
  loc_0041F856: lea edx, var_8C
  loc_0041F85C: lea eax, var_7C
  loc_0041F85F: push edx
  loc_0041F860: lea ecx, var_6C
  loc_0041F863: push eax
  loc_0041F864: lea edx, var_5C
  loc_0041F867: push ecx
  loc_0041F868: push edx
  loc_0041F869: push 00000004h
  loc_0041F86B: call [00401018h] ; __vbaFreeVarList
  loc_0041F871: add esp, 00000014h
  loc_0041F874: ret
  loc_0041F875: ret
  loc_0041F876: mov eax, Me
  loc_0041F879: push eax
  loc_0041F87A: mov ecx, [eax]
  loc_0041F87C: call [ecx+00000008h]
  loc_0041F87F: mov eax, var_4
  loc_0041F882: mov ecx, var_14
  loc_0041F885: pop edi
  loc_0041F886: pop esi
  loc_0041F887: mov fs:[00000000h], ecx
  loc_0041F88E: pop ebx
  loc_0041F88F: mov esp, ebp
  loc_0041F891: pop ebp
  loc_0041F892: retn 0004h
End Sub

Private Sub cmdRestore_Click() '41EDD0
  loc_0041EDD0: push ebp
  loc_0041EDD1: mov ebp, esp
  loc_0041EDD3: sub esp, 0000000Ch
  loc_0041EDD6: push 00401926h ; __vbaExceptHandler
  loc_0041EDDB: mov eax, fs:[00000000h]
  loc_0041EDE1: push eax
  loc_0041EDE2: mov fs:[00000000h], esp
  loc_0041EDE9: sub esp, 00000018h
  loc_0041EDEC: push ebx
  loc_0041EDED: push esi
  loc_0041EDEE: push edi
  loc_0041EDEF: mov var_C, esp
  loc_0041EDF2: mov var_8, 00401318h
  loc_0041EDF9: mov esi, Me
  loc_0041EDFC: mov eax, esi
  loc_0041EDFE: and eax, 00000001h
  loc_0041EE01: mov var_4, eax
  loc_0041EE04: and esi, FFFFFFFEh
  loc_0041EE07: push esi
  loc_0041EE08: mov Me, esi
  loc_0041EE0B: mov ecx, [esi]
  loc_0041EE0D: call [ecx+00000004h]
  loc_0041EE10: mov [00430060h], 41233333h
  loc_0041EE1A: mov [00430074h], 3F99999Ah
  loc_0041EE24: mov edx, [esi]
  loc_0041EE26: xor eax, eax
  loc_0041EE28: push esi
  loc_0041EE29: mov var_18, eax
  loc_0041EE2C: mov var_1C, eax
  loc_0041EE2F: call [edx+00000310h]
  loc_0041EE35: push eax
  loc_0041EE36: lea eax, var_1C
  loc_0041EE39: push eax
  loc_0041EE3A: call [0040103Ch] ; __vbaObjSet
  loc_0041EE40: mov ecx, [00430060h]
  loc_0041EE46: mov edi, eax
  loc_0041EE48: push ecx
  loc_0041EE49: mov ebx, [edi]
  loc_0041EE4B: call [0040107Ch] ; __vbaStrR4
  loc_0041EE51: mov edx, eax
  loc_0041EE53: lea ecx, var_18
  loc_0041EE56: call [004010FCh] ; __vbaStrMove
  loc_0041EE5C: push eax
  loc_0041EE5D: push edi
  loc_0041EE5E: call [ebx+000000A4h]
  loc_0041EE64: test eax, eax
  loc_0041EE66: fnclex
  loc_0041EE68: jge 0041EE7Ch
  loc_0041EE6A: push 000000A4h
  loc_0041EE6F: push 0040F02Ch
  loc_0041EE74: push edi
  loc_0041EE75: push eax
  loc_0041EE76: call [00401030h] ; __vbaHresultCheckObj
  loc_0041EE7C: mov edi, [00401110h] ; __vbaFreeStr
  loc_0041EE82: lea ecx, var_18
  loc_0041EE85: call edi
  loc_0041EE87: lea ecx, var_1C
  loc_0041EE8A: call [00401114h] ; __vbaFreeObj
  loc_0041EE90: mov edx, [esi]
  loc_0041EE92: push esi
  loc_0041EE93: call [edx+00000318h]
  loc_0041EE99: push eax
  loc_0041EE9A: lea eax, var_1C
  loc_0041EE9D: push eax
  loc_0041EE9E: call [0040103Ch] ; __vbaObjSet
  loc_0041EEA4: mov ecx, [00430074h]
  loc_0041EEAA: mov esi, eax
  loc_0041EEAC: push ecx
  loc_0041EEAD: mov ebx, [esi]
  loc_0041EEAF: call [0040107Ch] ; __vbaStrR4
  loc_0041EEB5: mov edx, eax
  loc_0041EEB7: lea ecx, var_18
  loc_0041EEBA: call [004010FCh] ; __vbaStrMove
  loc_0041EEC0: push eax
  loc_0041EEC1: push esi
  loc_0041EEC2: call [ebx+000000A4h]
  loc_0041EEC8: test eax, eax
  loc_0041EECA: fnclex
  loc_0041EECC: jge 0041EEE0h
  loc_0041EECE: push 000000A4h
  loc_0041EED3: push 0040F02Ch
  loc_0041EED8: push esi
  loc_0041EED9: push eax
  loc_0041EEDA: call [00401030h] ; __vbaHresultCheckObj
  loc_0041EEE0: lea ecx, var_18
  loc_0041EEE3: call edi
  loc_0041EEE5: lea ecx, var_1C
  loc_0041EEE8: call [00401114h] ; __vbaFreeObj
  loc_0041EEEE: mov var_4, 00000000h
  loc_0041EEF5: fwait
  loc_0041EEF6: push 0041EF11h
  loc_0041EEFB: jmp 0041EF10h
  loc_0041EEFD: lea ecx, var_18
  loc_0041EF00: call [00401110h] ; __vbaFreeStr
  loc_0041EF06: lea ecx, var_1C
  loc_0041EF09: call [00401114h] ; __vbaFreeObj
  loc_0041EF0F: ret
  loc_0041EF10: ret
  loc_0041EF11: mov eax, Me
  loc_0041EF14: push eax
  loc_0041EF15: mov edx, [eax]
  loc_0041EF17: call [edx+00000008h]
  loc_0041EF1A: mov eax, var_4
  loc_0041EF1D: mov ecx, var_14
  loc_0041EF20: pop edi
  loc_0041EF21: pop esi
  loc_0041EF22: mov fs:[00000000h], ecx
  loc_0041EF29: pop ebx
  loc_0041EF2A: mov esp, ebp
  loc_0041EF2C: pop ebp
  loc_0041EF2D: retn 0004h
End Sub

Private Sub Form_Activate() '41F8A0
  loc_0041F8A0: push ebp
  loc_0041F8A1: mov ebp, esp
  loc_0041F8A3: sub esp, 0000000Ch
  loc_0041F8A6: push 00401926h ; __vbaExceptHandler
  loc_0041F8AB: mov eax, fs:[00000000h]
  loc_0041F8B1: push eax
  loc_0041F8B2: mov fs:[00000000h], esp
  loc_0041F8B9: sub esp, 00000018h
  loc_0041F8BC: push ebx
  loc_0041F8BD: push esi
  loc_0041F8BE: push edi
  loc_0041F8BF: mov var_C, esp
  loc_0041F8C2: mov var_8, 00401338h
  loc_0041F8C9: mov esi, Me
  loc_0041F8CC: mov eax, esi
  loc_0041F8CE: and eax, 00000001h
  loc_0041F8D1: mov var_4, eax
  loc_0041F8D4: and esi, FFFFFFFEh
  loc_0041F8D7: push esi
  loc_0041F8D8: mov Me, esi
  loc_0041F8DB: mov ecx, [esi]
  loc_0041F8DD: call [ecx+00000004h]
  loc_0041F8E0: mov edx, [esi]
  loc_0041F8E2: xor eax, eax
  loc_0041F8E4: push esi
  loc_0041F8E5: mov var_18, eax
  loc_0041F8E8: mov var_1C, eax
  loc_0041F8EB: call [edx+00000310h]
  loc_0041F8F1: push eax
  loc_0041F8F2: lea eax, var_1C
  loc_0041F8F5: push eax
  loc_0041F8F6: call [0040103Ch] ; __vbaObjSet
  loc_0041F8FC: mov ecx, [00430060h]
  loc_0041F902: mov edi, eax
  loc_0041F904: push ecx
  loc_0041F905: mov ebx, [edi]
  loc_0041F907: call [0040107Ch] ; __vbaStrR4
  loc_0041F90D: mov edx, eax
  loc_0041F90F: lea ecx, var_18
  loc_0041F912: call [004010FCh] ; __vbaStrMove
  loc_0041F918: push eax
  loc_0041F919: push edi
  loc_0041F91A: call [ebx+000000A4h]
  loc_0041F920: test eax, eax
  loc_0041F922: fnclex
  loc_0041F924: jge 0041F938h
  loc_0041F926: push 000000A4h
  loc_0041F92B: push 0040F02Ch
  loc_0041F930: push edi
  loc_0041F931: push eax
  loc_0041F932: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F938: lea ecx, var_18
  loc_0041F93B: call [00401110h] ; __vbaFreeStr
  loc_0041F941: lea ecx, var_1C
  loc_0041F944: call [00401114h] ; __vbaFreeObj
  loc_0041F94A: mov eax, [00430150h]
  loc_0041F94F: test eax, eax
  loc_0041F951: jnz 0041F968h
  loc_0041F953: push 00430150h
  loc_0041F958: push 00403F48h
  loc_0041F95D: call [004010D4h] ; __vbaNew2
  loc_0041F963: mov eax, [00430150h]
  loc_0041F968: mov edx, [eax]
  loc_0041F96A: push eax
  loc_0041F96B: call [edx+00000318h]
  loc_0041F971: push eax
  loc_0041F972: lea eax, var_1C
  loc_0041F975: push eax
  loc_0041F976: call [0040103Ch] ; __vbaObjSet
  loc_0041F97C: mov ecx, [00430074h]
  loc_0041F982: mov edi, eax
  loc_0041F984: push ecx
  loc_0041F985: mov ebx, [edi]
  loc_0041F987: call [0040107Ch] ; __vbaStrR4
  loc_0041F98D: mov edx, eax
  loc_0041F98F: lea ecx, var_18
  loc_0041F992: call [004010FCh] ; __vbaStrMove
  loc_0041F998: push eax
  loc_0041F999: push edi
  loc_0041F99A: call [ebx+000000A4h]
  loc_0041F9A0: test eax, eax
  loc_0041F9A2: fnclex
  loc_0041F9A4: jge 0041F9B8h
  loc_0041F9A6: push 000000A4h
  loc_0041F9AB: push 0040F02Ch
  loc_0041F9B0: push edi
  loc_0041F9B1: push eax
  loc_0041F9B2: call [00401030h] ; __vbaHresultCheckObj
  loc_0041F9B8: lea ecx, var_18
  loc_0041F9BB: call [00401110h] ; __vbaFreeStr
  loc_0041F9C1: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041F9C7: lea ecx, var_1C
  loc_0041F9CA: call edi
  loc_0041F9CC: mov edx, [esi]
  loc_0041F9CE: push esi
  loc_0041F9CF: call [edx+00000310h]
  loc_0041F9D5: push eax
  loc_0041F9D6: lea eax, var_1C
  loc_0041F9D9: push eax
  loc_0041F9DA: call [0040103Ch] ; __vbaObjSet
  loc_0041F9E0: mov esi, eax
  loc_0041F9E2: push esi
  loc_0041F9E3: mov ecx, [esi]
  loc_0041F9E5: call [ecx+00000204h]
  loc_0041F9EB: test eax, eax
  loc_0041F9ED: fnclex
  loc_0041F9EF: jge 0041FA03h
  loc_0041F9F1: push 00000204h
  loc_0041F9F6: push 0040F02Ch
  loc_0041F9FB: push esi
  loc_0041F9FC: push eax
  loc_0041F9FD: call [00401030h] ; __vbaHresultCheckObj
  loc_0041FA03: lea ecx, var_1C
  loc_0041FA06: call edi
  loc_0041FA08: mov var_4, 00000000h
  loc_0041FA0F: fwait
  loc_0041FA10: push 0041FA2Bh
  loc_0041FA15: jmp 0041FA2Ah
  loc_0041FA17: lea ecx, var_18
  loc_0041FA1A: call [00401110h] ; __vbaFreeStr
  loc_0041FA20: lea ecx, var_1C
  loc_0041FA23: call [00401114h] ; __vbaFreeObj
  loc_0041FA29: ret
  loc_0041FA2A: ret
  loc_0041FA2B: mov eax, Me
  loc_0041FA2E: push eax
  loc_0041FA2F: mov edx, [eax]
  loc_0041FA31: call [edx+00000008h]
  loc_0041FA34: mov eax, var_4
  loc_0041FA37: mov ecx, var_14
  loc_0041FA3A: pop edi
  loc_0041FA3B: pop esi
  loc_0041FA3C: mov fs:[00000000h], ecx
  loc_0041FA43: pop ebx
  loc_0041FA44: mov esp, ebp
  loc_0041FA46: pop ebp
  loc_0041FA47: retn 0004h
End Sub

Private Sub txtbeta0Hat_KeyPress(KeyAscii As Integer) '41FBD0
  loc_0041FBD0: push ebp
  loc_0041FBD1: mov ebp, esp
  loc_0041FBD3: sub esp, 0000000Ch
  loc_0041FBD6: push 00401926h ; __vbaExceptHandler
  loc_0041FBDB: mov eax, fs:[00000000h]
  loc_0041FBE1: push eax
  loc_0041FBE2: mov fs:[00000000h], esp
  loc_0041FBE9: sub esp, 0000007Ch
  loc_0041FBEC: push ebx
  loc_0041FBED: push esi
  loc_0041FBEE: push edi
  loc_0041FBEF: mov var_C, esp
  loc_0041FBF2: mov var_8, 00401358h
  loc_0041FBF9: mov esi, Me
  loc_0041FBFC: mov eax, esi
  loc_0041FBFE: and eax, 00000001h
  loc_0041FC01: mov var_4, eax
  loc_0041FC04: and esi, FFFFFFFEh
  loc_0041FC07: push esi
  loc_0041FC08: mov Me, esi
  loc_0041FC0B: mov ecx, [esi]
  loc_0041FC0D: call [ecx+00000004h]
  loc_0041FC10: mov ebx, KeyAscii
  loc_0041FC13: lea edx, var_24
  loc_0041FC16: xor edi, edi
  loc_0041FC18: push ebx
  loc_0041FC19: push edx
  loc_0041FC1A: mov var_24, edi
  loc_0041FC1D: mov var_34, edi
  loc_0041FC20: mov var_44, edi
  loc_0041FC23: mov var_54, edi
  loc_0041FC26: mov var_64, edi
  loc_0041FC29: mov var_74, edi
  loc_0041FC2C: mov var_84, edi
  loc_0041FC32: call 0041A480h
  loc_0041FC37: lea eax, var_24
  loc_0041FC3A: push eax
  loc_0041FC3B: call [004010C4h] ; __vbaI2Var
  loc_0041FC41: lea ecx, var_24
  loc_0041FC44: mov [ebx], ax
  loc_0041FC47: call [00401010h] ; __vbaFreeVar
  loc_0041FC4D: push 0040DD3Ch ; "."
  loc_0041FC52: call [00401024h] ; rtcAnsiValueBstr
  loc_0041FC58: mov edx, [esi]
  loc_0041FC5A: xor ecx, ecx
  loc_0041FC5C: cmp [ebx], ax
  loc_0041FC5F: push esi
  loc_0041FC60: mov var_84, 0000000Bh
  loc_0041FC6A: setz cl
  loc_0041FC6D: neg ecx
  loc_0041FC6F: mov var_7C, cx
  loc_0041FC73: call [edx+00000318h]
  loc_0041FC79: mov var_1C, eax
  loc_0041FC7C: lea eax, var_84
  loc_0041FC82: push eax
  loc_0041FC83: lea ecx, var_24
  loc_0041FC86: push 00000001h
  loc_0041FC88: lea edx, var_64
  loc_0041FC8B: push ecx
  loc_0041FC8C: push edx
  loc_0041FC8D: lea eax, var_34
  loc_0041FC90: push edi
  loc_0041FC91: push eax
  loc_0041FC92: mov var_24, 00000009h
  loc_0041FC99: mov var_5C, 0040DD3Ch ; "."
  loc_0041FCA0: mov var_64, 00000008h
  loc_0041FCA7: mov var_6C, edi
  loc_0041FCAA: mov var_74, 00008002h
  loc_0041FCB1: call [004010B4h] ; __vbaInStrVar
  loc_0041FCB7: lea ecx, var_74
  loc_0041FCBA: push eax
  loc_0041FCBB: lea edx, var_44
  loc_0041FCBE: push ecx
  loc_0041FCBF: push edx
  loc_0041FCC0: call [00401060h] ; __vbaVarCmpGt
  loc_0041FCC6: push eax
  loc_0041FCC7: lea eax, var_54
  loc_0041FCCA: push eax
  loc_0041FCCB: call [00401094h] ; __vbaVarAnd
  loc_0041FCD1: push eax
  loc_0041FCD2: call [00401058h] ; __vbaBoolVarNull
  loc_0041FCD8: lea ecx, var_84
  loc_0041FCDE: mov esi, eax
  loc_0041FCE0: lea edx, var_34
  loc_0041FCE3: push ecx
  loc_0041FCE4: lea eax, var_24
  loc_0041FCE7: push edx
  loc_0041FCE8: push eax
  loc_0041FCE9: push 00000003h
  loc_0041FCEB: call [00401018h] ; __vbaFreeVarList
  loc_0041FCF1: add esp, 00000010h
  loc_0041FCF4: cmp si, di
  loc_0041FCF7: jz 0041FCFCh
  loc_0041FCF9: mov [ebx], di
  loc_0041FCFC: mov var_4, edi
  loc_0041FCFF: push 0041FD23h
  loc_0041FD04: jmp 0041FD22h
  loc_0041FD06: lea ecx, var_54
  loc_0041FD09: lea edx, var_44
  loc_0041FD0C: push ecx
  loc_0041FD0D: lea eax, var_34
  loc_0041FD10: push edx
  loc_0041FD11: lea ecx, var_24
  loc_0041FD14: push eax
  loc_0041FD15: push ecx
  loc_0041FD16: push 00000004h
  loc_0041FD18: call [00401018h] ; __vbaFreeVarList
  loc_0041FD1E: add esp, 00000014h
  loc_0041FD21: ret
  loc_0041FD22: ret
  loc_0041FD23: mov eax, Me
  loc_0041FD26: push eax
  loc_0041FD27: mov edx, [eax]
  loc_0041FD29: call [edx+00000008h]
  loc_0041FD2C: mov eax, var_4
  loc_0041FD2F: mov ecx, var_14
  loc_0041FD32: pop edi
  loc_0041FD33: pop esi
  loc_0041FD34: mov fs:[00000000h], ecx
  loc_0041FD3B: pop ebx
  loc_0041FD3C: mov esp, ebp
  loc_0041FD3E: pop ebp
  loc_0041FD3F: retn 0008h
End Sub
