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
  ClientWidth = 5475
  ClientHeight = 4875
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Distribution values"
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
      Width = 4545
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
        Left = 1290
        Top = 420
        Width = 1725
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
      Left = 480
      Top = 3750
      Width = 855
      Height = 795
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
      Top = 3750
      Width = 1095
      Height = 825
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
      Left = 3180
      Top = 3750
      Width = 1455
      Height = 825
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
      Left = 1350
      Top = 480
      Width = 2595
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
      Left = 270
      Top = 1590
      Width = 4605
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
        Left = 1410
        Top = 420
        Width = 1695
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


Private Sub txt1SidedT_KeyPress(KeyAscii As Integer) '413A60
  loc_00413A60: push ebp
  loc_00413A61: mov ebp, esp
  loc_00413A63: sub esp, 0000000Ch
  loc_00413A66: push 00401546h ; __vbaExceptHandler
  loc_00413A6B: mov eax, fs:[00000000h]
  loc_00413A71: push eax
  loc_00413A72: mov fs:[00000000h], esp
  loc_00413A79: sub esp, 0000008Ch
  loc_00413A7F: push ebx
  loc_00413A80: push esi
  loc_00413A81: push edi
  loc_00413A82: mov var_C, esp
  loc_00413A85: mov var_8, 00401530h
  loc_00413A8C: mov esi, Me
  loc_00413A8F: mov eax, esi
  loc_00413A91: and eax, 00000001h
  loc_00413A94: mov var_4, eax
  loc_00413A97: and esi, FFFFFFFEh
  loc_00413A9A: push esi
  loc_00413A9B: mov Me, esi
  loc_00413A9E: mov ecx, [esi]
  loc_00413AA0: call [ecx+00000004h]
  loc_00413AA3: mov edi, KeyAscii
  loc_00413AA6: lea edx, var_24
  loc_00413AA9: xor ebx, ebx
  loc_00413AAB: push edi
  loc_00413AAC: push edx
  loc_00413AAD: mov var_24, ebx
  loc_00413AB0: mov var_34, ebx
  loc_00413AB3: mov var_44, ebx
  loc_00413AB6: mov var_54, ebx
  loc_00413AB9: mov var_64, ebx
  loc_00413ABC: mov var_74, ebx
  loc_00413ABF: mov var_84, ebx
  loc_00413AC5: call 00408F40h
  loc_00413ACA: lea eax, var_24
  loc_00413ACD: push eax
  loc_00413ACE: call [00401104h] ; __vbaI2Var
  loc_00413AD4: lea ecx, var_24
  loc_00413AD7: mov [edi], ax
  loc_00413ADA: call [00401014h] ; __vbaFreeVar
  loc_00413AE0: push 004035B0h ; "."
  loc_00413AE5: call [00401030h] ; rtcAnsiValueBstr
  loc_00413AEB: mov edx, [esi]
  loc_00413AED: xor ecx, ecx
  loc_00413AEF: cmp [edi], ax
  loc_00413AF2: push esi
  loc_00413AF3: mov var_84, 0000000Bh
  loc_00413AFD: setz cl
  loc_00413B00: neg ecx
  loc_00413B02: mov var_7C, cx
  loc_00413B06: call [edx+00000304h]
  loc_00413B0C: mov var_1C, eax
  loc_00413B0F: lea eax, var_84
  loc_00413B15: push eax
  loc_00413B16: lea ecx, var_24
  loc_00413B19: push 00000001h
  loc_00413B1B: lea edx, var_64
  loc_00413B1E: push ecx
  loc_00413B1F: push edx
  loc_00413B20: lea eax, var_34
  loc_00413B23: push ebx
  loc_00413B24: push eax
  loc_00413B25: mov var_24, 00000009h
  loc_00413B2C: mov var_5C, 004035B0h ; "."
  loc_00413B33: mov var_64, 00000008h
  loc_00413B3A: mov var_6C, ebx
  loc_00413B3D: mov var_74, 00008002h
  loc_00413B44: call [004010F4h] ; __vbaInStrVar
  loc_00413B4A: lea ecx, var_74
  loc_00413B4D: push eax
  loc_00413B4E: lea edx, var_44
  loc_00413B51: push ecx
  loc_00413B52: push edx
  loc_00413B53: call [00401084h] ; __vbaVarCmpGt
  loc_00413B59: push eax
  loc_00413B5A: lea eax, var_54
  loc_00413B5D: push eax
  loc_00413B5E: call [004010C4h] ; __vbaVarAnd
  loc_00413B64: push eax
  loc_00413B65: call [0040107Ch] ; __vbaBoolVarNull
  loc_00413B6B: lea ecx, var_84
  loc_00413B71: mov esi, eax
  loc_00413B73: lea edx, var_34
  loc_00413B76: push ecx
  loc_00413B77: lea eax, var_24
  loc_00413B7A: push edx
  loc_00413B7B: push eax
  loc_00413B7C: push 00000003h
  loc_00413B7E: call [00401024h] ; __vbaFreeVarList
  loc_00413B84: add esp, 00000010h
  loc_00413B87: cmp si, bx
  loc_00413B8A: jz 00413B8Fh
  loc_00413B8C: mov [edi], bx
  loc_00413B8F: push 004035B8h ; "-"
  loc_00413B94: call [00401030h] ; rtcAnsiValueBstr
  loc_00413B9A: cmp [edi], ax
  loc_00413B9D: jnz 00413C0Ah
  loc_00413B9F: mov ecx, 80020004h
  loc_00413BA4: mov eax, 0000000Ah
  loc_00413BA9: mov var_4C, ecx
  loc_00413BAC: mov var_3C, ecx
  loc_00413BAF: mov var_2C, ecx
  loc_00413BB2: lea edx, var_64
  loc_00413BB5: lea ecx, var_24
  loc_00413BB8: mov var_54, eax
  loc_00413BBB: mov var_44, eax
  loc_00413BBE: mov var_34, eax
  loc_00413BC1: mov var_5C, 0040430Ch ; "Please enter the right-sided (or positive) t-table value."
  loc_00413BC8: mov var_64, 00000008h
  loc_00413BCF: call [00401150h] ; __vbaVarDup
  loc_00413BD5: lea ecx, var_54
  loc_00413BD8: lea edx, var_44
  loc_00413BDB: push ecx
  loc_00413BDC: lea eax, var_34
  loc_00413BDF: push edx
  loc_00413BE0: push eax
  loc_00413BE1: lea ecx, var_24
  loc_00413BE4: push ebx
  loc_00413BE5: push ecx
  loc_00413BE6: call [00401060h] ; rtcMsgBox
  loc_00413BEC: lea edx, var_54
  loc_00413BEF: lea eax, var_44
  loc_00413BF2: push edx
  loc_00413BF3: lea ecx, var_34
  loc_00413BF6: push eax
  loc_00413BF7: lea edx, var_24
  loc_00413BFA: push ecx
  loc_00413BFB: push edx
  loc_00413BFC: push 00000004h
  loc_00413BFE: call [00401024h] ; __vbaFreeVarList
  loc_00413C04: add esp, 00000014h
  loc_00413C07: mov [edi], bx
  loc_00413C0A: mov var_4, ebx
  loc_00413C0D: push 00413C31h
  loc_00413C12: jmp 00413C30h
  loc_00413C14: lea eax, var_54
  loc_00413C17: lea ecx, var_44
  loc_00413C1A: push eax
  loc_00413C1B: lea edx, var_34
  loc_00413C1E: push ecx
  loc_00413C1F: lea eax, var_24
  loc_00413C22: push edx
  loc_00413C23: push eax
  loc_00413C24: push 00000004h
  loc_00413C26: call [00401024h] ; __vbaFreeVarList
  loc_00413C2C: add esp, 00000014h
  loc_00413C2F: ret
  loc_00413C30: ret
  loc_00413C31: mov eax, Me
  loc_00413C34: push eax
  loc_00413C35: mov ecx, [eax]
  loc_00413C37: call [ecx+00000008h]
  loc_00413C3A: mov eax, var_4
  loc_00413C3D: mov ecx, var_14
  loc_00413C40: pop edi
  loc_00413C41: pop esi
  loc_00413C42: mov fs:[00000000h], ecx
  loc_00413C49: pop ebx
  loc_00413C4A: mov esp, ebp
  loc_00413C4C: pop ebp
  loc_00413C4D: retn 0008h
End Sub

Private Sub Form_Activate() '413500
  loc_00413500: push ebp
  loc_00413501: mov ebp, esp
  loc_00413503: sub esp, 0000000Ch
  loc_00413506: push 00401546h ; __vbaExceptHandler
  loc_0041350B: mov eax, fs:[00000000h]
  loc_00413511: push eax
  loc_00413512: mov fs:[00000000h], esp
  loc_00413519: sub esp, 00000018h
  loc_0041351C: push ebx
  loc_0041351D: push esi
  loc_0041351E: push edi
  loc_0041351F: mov var_C, esp
  loc_00413522: mov var_8, 00401500h
  loc_00413529: mov esi, Me
  loc_0041352C: mov eax, esi
  loc_0041352E: and eax, 00000001h
  loc_00413531: mov var_4, eax
  loc_00413534: and esi, FFFFFFFEh
  loc_00413537: push esi
  loc_00413538: mov Me, esi
  loc_0041353B: mov ecx, [esi]
  loc_0041353D: call [ecx+00000004h]
  loc_00413540: mov edx, [esi]
  loc_00413542: xor eax, eax
  loc_00413544: push esi
  loc_00413545: mov var_18, eax
  loc_00413548: mov var_1C, eax
  loc_0041354B: call [edx+00000318h]
  loc_00413551: push eax
  loc_00413552: lea eax, var_1C
  loc_00413555: push eax
  loc_00413556: call [0040105Ch] ; __vbaObjSet
  loc_0041355C: mov ecx, [00415020h]
  loc_00413562: mov edi, eax
  loc_00413564: push ecx
  loc_00413565: mov ebx, [edi]
  loc_00413567: call [004010A8h] ; __vbaStrR4
  loc_0041356D: mov edx, eax
  loc_0041356F: lea ecx, var_18
  loc_00413572: call [00401164h] ; __vbaStrMove
  loc_00413578: push eax
  loc_00413579: push edi
  loc_0041357A: call [ebx+000000A4h]
  loc_00413580: test eax, eax
  loc_00413582: fnclex
  loc_00413584: jge 00413598h
  loc_00413586: push 000000A4h
  loc_0041358B: push 00403848h
  loc_00413590: push edi
  loc_00413591: push eax
  loc_00413592: call [00401040h] ; __vbaHresultCheckObj
  loc_00413598: lea ecx, var_18
  loc_0041359B: call [00401184h] ; __vbaFreeStr
  loc_004135A1: lea ecx, var_1C
  loc_004135A4: call [00401180h] ; __vbaFreeObj
  loc_004135AA: mov edx, [esi]
  loc_004135AC: push esi
  loc_004135AD: call [edx+00000320h]
  loc_004135B3: push eax
  loc_004135B4: lea eax, var_1C
  loc_004135B7: push eax
  loc_004135B8: call [0040105Ch] ; __vbaObjSet
  loc_004135BE: mov ecx, [00415028h]
  loc_004135C4: mov edi, eax
  loc_004135C6: push ecx
  loc_004135C7: mov ebx, [edi]
  loc_004135C9: call [004010A8h] ; __vbaStrR4
  loc_004135CF: mov edx, eax
  loc_004135D1: lea ecx, var_18
  loc_004135D4: call [00401164h] ; __vbaStrMove
  loc_004135DA: push eax
  loc_004135DB: push edi
  loc_004135DC: call [ebx+000000A4h]
  loc_004135E2: test eax, eax
  loc_004135E4: fnclex
  loc_004135E6: jge 004135FAh
  loc_004135E8: push 000000A4h
  loc_004135ED: push 00403848h
  loc_004135F2: push edi
  loc_004135F3: push eax
  loc_004135F4: call [00401040h] ; __vbaHresultCheckObj
  loc_004135FA: lea ecx, var_18
  loc_004135FD: call [00401184h] ; __vbaFreeStr
  loc_00413603: lea ecx, var_1C
  loc_00413606: call [00401180h] ; __vbaFreeObj
  loc_0041360C: mov edx, [esi]
  loc_0041360E: push esi
  loc_0041360F: call [edx+00000304h]
  loc_00413615: push eax
  loc_00413616: lea eax, var_1C
  loc_00413619: push eax
  loc_0041361A: call [0040105Ch] ; __vbaObjSet
  loc_00413620: mov ecx, [00415024h]
  loc_00413626: mov edi, eax
  loc_00413628: push ecx
  loc_00413629: mov ebx, [edi]
  loc_0041362B: call [004010A8h] ; __vbaStrR4
  loc_00413631: mov edx, eax
  loc_00413633: lea ecx, var_18
  loc_00413636: call [00401164h] ; __vbaStrMove
  loc_0041363C: push eax
  loc_0041363D: push edi
  loc_0041363E: call [ebx+000000A4h]
  loc_00413644: test eax, eax
  loc_00413646: fnclex
  loc_00413648: jge 0041365Ch
  loc_0041364A: push 000000A4h
  loc_0041364F: push 00403848h
  loc_00413654: push edi
  loc_00413655: push eax
  loc_00413656: call [00401040h] ; __vbaHresultCheckObj
  loc_0041365C: lea ecx, var_18
  loc_0041365F: call [00401184h] ; __vbaFreeStr
  loc_00413665: mov edi, [00401180h] ; __vbaFreeObj
  loc_0041366B: lea ecx, var_1C
  loc_0041366E: call edi
  loc_00413670: mov edx, [esi]
  loc_00413672: push esi
  loc_00413673: call [edx+00000318h]
  loc_00413679: push eax
  loc_0041367A: lea eax, var_1C
  loc_0041367D: push eax
  loc_0041367E: call [0040105Ch] ; __vbaObjSet
  loc_00413684: mov esi, eax
  loc_00413686: push esi
  loc_00413687: mov ecx, [esi]
  loc_00413689: call [ecx+00000204h]
  loc_0041368F: test eax, eax
  loc_00413691: fnclex
  loc_00413693: jge 004136A7h
  loc_00413695: push 00000204h
  loc_0041369A: push 00403848h
  loc_0041369F: push esi
  loc_004136A0: push eax
  loc_004136A1: call [00401040h] ; __vbaHresultCheckObj
  loc_004136A7: lea ecx, var_1C
  loc_004136AA: call edi
  loc_004136AC: mov var_4, 00000000h
  loc_004136B3: fwait
  loc_004136B4: push 004136CFh
  loc_004136B9: jmp 004136CEh
  loc_004136BB: lea ecx, var_18
  loc_004136BE: call [00401184h] ; __vbaFreeStr
  loc_004136C4: lea ecx, var_1C
  loc_004136C7: call [00401180h] ; __vbaFreeObj
  loc_004136CD: ret
  loc_004136CE: ret
  loc_004136CF: mov eax, Me
  loc_004136D2: push eax
  loc_004136D3: mov edx, [eax]
  loc_004136D5: call [edx+00000008h]
  loc_004136D8: mov eax, var_4
  loc_004136DB: mov ecx, var_14
  loc_004136DE: pop edi
  loc_004136DF: pop esi
  loc_004136E0: mov fs:[00000000h], ecx
  loc_004136E7: pop ebx
  loc_004136E8: mov esp, ebp
  loc_004136EA: pop ebp
  loc_004136EB: retn 0004h
End Sub

Private Sub cmdClearChange_Click() '412A80
  loc_00412A80: push ebp
  loc_00412A81: mov ebp, esp
  loc_00412A83: sub esp, 0000000Ch
  loc_00412A86: push 00401546h ; __vbaExceptHandler
  loc_00412A8B: mov eax, fs:[00000000h]
  loc_00412A91: push eax
  loc_00412A92: mov fs:[00000000h], esp
  loc_00412A99: sub esp, 00000014h
  loc_00412A9C: push ebx
  loc_00412A9D: push esi
  loc_00412A9E: push edi
  loc_00412A9F: mov var_C, esp
  loc_00412AA2: mov var_8, 004014C8h
  loc_00412AA9: mov esi, Me
  loc_00412AAC: mov eax, esi
  loc_00412AAE: and eax, 00000001h
  loc_00412AB1: mov var_4, eax
  loc_00412AB4: and esi, FFFFFFFEh
  loc_00412AB7: push esi
  loc_00412AB8: mov Me, esi
  loc_00412ABB: mov ecx, [esi]
  loc_00412ABD: call [ecx+00000004h]
  loc_00412AC0: mov edx, [esi]
  loc_00412AC2: push esi
  loc_00412AC3: mov var_18, 00000000h
  loc_00412ACA: call [edx+00000318h]
  loc_00412AD0: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_00412AD6: push eax
  loc_00412AD7: lea eax, var_18
  loc_00412ADA: push eax
  loc_00412ADB: call ebx
  loc_00412ADD: mov edi, eax
  loc_00412ADF: push 00403D64h
  loc_00412AE4: push edi
  loc_00412AE5: mov ecx, [edi]
  loc_00412AE7: call [ecx+000000A4h]
  loc_00412AED: test eax, eax
  loc_00412AEF: fnclex
  loc_00412AF1: jge 00412B05h
  loc_00412AF3: push 000000A4h
  loc_00412AF8: push 00403848h
  loc_00412AFD: push edi
  loc_00412AFE: push eax
  loc_00412AFF: call [00401040h] ; __vbaHresultCheckObj
  loc_00412B05: lea ecx, var_18
  loc_00412B08: call [00401180h] ; __vbaFreeObj
  loc_00412B0E: mov edx, [esi]
  loc_00412B10: push esi
  loc_00412B11: call [edx+00000320h]
  loc_00412B17: push eax
  loc_00412B18: lea eax, var_18
  loc_00412B1B: push eax
  loc_00412B1C: call ebx
  loc_00412B1E: mov edi, eax
  loc_00412B20: push 00403D64h
  loc_00412B25: push edi
  loc_00412B26: mov ecx, [edi]
  loc_00412B28: call [ecx+000000A4h]
  loc_00412B2E: test eax, eax
  loc_00412B30: fnclex
  loc_00412B32: jge 00412B46h
  loc_00412B34: push 000000A4h
  loc_00412B39: push 00403848h
  loc_00412B3E: push edi
  loc_00412B3F: push eax
  loc_00412B40: call [00401040h] ; __vbaHresultCheckObj
  loc_00412B46: lea ecx, var_18
  loc_00412B49: call [00401180h] ; __vbaFreeObj
  loc_00412B4F: mov edx, [esi]
  loc_00412B51: push esi
  loc_00412B52: call [edx+00000304h]
  loc_00412B58: push eax
  loc_00412B59: lea eax, var_18
  loc_00412B5C: push eax
  loc_00412B5D: call ebx
  loc_00412B5F: mov edi, eax
  loc_00412B61: push 00403D64h
  loc_00412B66: push edi
  loc_00412B67: mov ecx, [edi]
  loc_00412B69: call [ecx+000000A4h]
  loc_00412B6F: test eax, eax
  loc_00412B71: fnclex
  loc_00412B73: jge 00412B87h
  loc_00412B75: push 000000A4h
  loc_00412B7A: push 00403848h
  loc_00412B7F: push edi
  loc_00412B80: push eax
  loc_00412B81: call [00401040h] ; __vbaHresultCheckObj
  loc_00412B87: mov edi, [00401180h] ; __vbaFreeObj
  loc_00412B8D: lea ecx, var_18
  loc_00412B90: call edi
  loc_00412B92: mov edx, [esi]
  loc_00412B94: push esi
  loc_00412B95: call [edx+00000318h]
  loc_00412B9B: push eax
  loc_00412B9C: lea eax, var_18
  loc_00412B9F: push eax
  loc_00412BA0: call ebx
  loc_00412BA2: mov esi, eax
  loc_00412BA4: push esi
  loc_00412BA5: mov ecx, [esi]
  loc_00412BA7: call [ecx+00000204h]
  loc_00412BAD: test eax, eax
  loc_00412BAF: fnclex
  loc_00412BB1: jge 00412BC5h
  loc_00412BB3: push 00000204h
  loc_00412BB8: push 00403848h
  loc_00412BBD: push esi
  loc_00412BBE: push eax
  loc_00412BBF: call [00401040h] ; __vbaHresultCheckObj
  loc_00412BC5: lea ecx, var_18
  loc_00412BC8: call edi
  loc_00412BCA: mov var_4, 00000000h
  loc_00412BD1: push 00412BE3h
  loc_00412BD6: jmp 00412BE2h
  loc_00412BD8: lea ecx, var_18
  loc_00412BDB: call [00401180h] ; __vbaFreeObj
  loc_00412BE1: ret
  loc_00412BE2: ret
  loc_00412BE3: mov eax, Me
  loc_00412BE6: push eax
  loc_00412BE7: mov edx, [eax]
  loc_00412BE9: call [edx+00000008h]
  loc_00412BEC: mov eax, var_4
  loc_00412BEF: mov ecx, var_14
  loc_00412BF2: pop edi
  loc_00412BF3: pop esi
  loc_00412BF4: mov fs:[00000000h], ecx
  loc_00412BFB: pop ebx
  loc_00412BFC: mov esp, ebp
  loc_00412BFE: pop ebp
  loc_00412BFF: retn 0004h
End Sub

Private Sub cmdChange_Click() '412DE0
  loc_00412DE0: push ebp
  loc_00412DE1: mov ebp, esp
  loc_00412DE3: sub esp, 0000000Ch
  loc_00412DE6: push 00401546h ; __vbaExceptHandler
  loc_00412DEB: mov eax, fs:[00000000h]
  loc_00412DF1: push eax
  loc_00412DF2: mov fs:[00000000h], esp
  loc_00412DF9: sub esp, 000000A8h
  loc_00412DFF: push ebx
  loc_00412E00: push esi
  loc_00412E01: push edi
  loc_00412E02: mov var_C, esp
  loc_00412E05: mov var_8, 004014F0h
  loc_00412E0C: mov esi, Me
  loc_00412E0F: mov eax, esi
  loc_00412E11: and eax, 00000001h
  loc_00412E14: mov var_4, eax
  loc_00412E17: and esi, FFFFFFFEh
  loc_00412E1A: push esi
  loc_00412E1B: mov Me, esi
  loc_00412E1E: mov ecx, [esi]
  loc_00412E20: call [ecx+00000004h]
  loc_00412E23: mov edx, [esi]
  loc_00412E25: xor eax, eax
  loc_00412E27: push esi
  loc_00412E28: mov var_18, eax
  loc_00412E2B: mov var_1C, eax
  loc_00412E2E: mov var_2C, eax
  loc_00412E31: mov var_3C, eax
  loc_00412E34: mov var_4C, eax
  loc_00412E37: mov var_5C, eax
  loc_00412E3A: mov var_6C, eax
  loc_00412E3D: mov var_7C, eax
  loc_00412E40: call [edx+00000318h]
  loc_00412E46: mov ebx, [00401078h] ; rtcTrimVar
  loc_00412E4C: mov var_24, eax
  loc_00412E4F: lea eax, var_2C
  loc_00412E52: lea ecx, var_3C
  loc_00412E55: push eax
  loc_00412E56: push ecx
  loc_00412E57: mov var_2C, 00000009h
  loc_00412E5E: call ebx
  loc_00412E60: lea edx, var_3C
  loc_00412E63: lea eax, var_6C
  loc_00412E66: push edx
  loc_00412E67: push eax
  loc_00412E68: mov var_64, 00403D64h
  loc_00412E6F: mov var_6C, 00008008h
  loc_00412E76: call [0040109Ch] ; __vbaVarTstEq
  loc_00412E7C: mov edi, [00401024h] ; __vbaFreeVarList
  loc_00412E82: lea ecx, var_3C
  loc_00412E85: lea edx, var_2C
  loc_00412E88: push ecx
  loc_00412E89: push edx
  loc_00412E8A: push 00000002h
  loc_00412E8C: mov var_A0, ax
  loc_00412E93: call edi
  loc_00412E95: add esp, 0000000Ch
  loc_00412E98: cmp var_A0, 0000h
  loc_00412EA0: jz 00412F21h
  loc_00412EA2: mov eax, [esi]
  loc_00412EA4: push esi
  loc_00412EA5: call [eax+00000318h]
  loc_00412EAB: lea ecx, var_1C
  loc_00412EAE: push eax
  loc_00412EAF: push ecx
  loc_00412EB0: call [0040105Ch] ; __vbaObjSet
  loc_00412EB6: mov esi, eax
  loc_00412EB8: push esi
  loc_00412EB9: mov edx, [esi]
  loc_00412EBB: call [edx+00000204h]
  loc_00412EC1: test eax, eax
  loc_00412EC3: fnclex
  loc_00412EC5: jge 00412ED9h
  loc_00412EC7: push 00000204h
  loc_00412ECC: push 00403848h
  loc_00412ED1: push esi
  loc_00412ED2: push eax
  loc_00412ED3: call [00401040h] ; __vbaHresultCheckObj
  loc_00412ED9: lea ecx, var_1C
  loc_00412EDC: call [00401180h] ; __vbaFreeObj
  loc_00412EE2: mov esi, [00401150h] ; __vbaVarDup
  loc_00412EE8: mov ecx, 80020004h
  loc_00412EED: mov var_54, ecx
  loc_00412EF0: mov eax, 0000000Ah
  loc_00412EF5: mov var_44, ecx
  loc_00412EF8: mov ebx, 00000008h
  loc_00412EFD: lea edx, var_7C
  loc_00412F00: lea ecx, var_3C
  loc_00412F03: mov var_5C, eax
  loc_00412F06: mov var_4C, eax
  loc_00412F09: mov var_74, 00404104h ; "Check alpha"
  loc_00412F10: mov var_7C, ebx
  loc_00412F13: call __vbaVarDup
  loc_00412F15: mov var_64, 004040C0h ; "Please enter a value for alpha."
  loc_00412F1C: jmp 00412FF8h
  loc_00412F21: mov edx, [esi]
  loc_00412F23: push esi
  loc_00412F24: call [edx+00000304h]
  loc_00412F2A: mov var_24, eax
  loc_00412F2D: lea eax, var_2C
  loc_00412F30: lea ecx, var_3C
  loc_00412F33: push eax
  loc_00412F34: push ecx
  loc_00412F35: mov var_2C, 00000009h
  loc_00412F3C: call ebx
  loc_00412F3E: lea edx, var_3C
  loc_00412F41: lea eax, var_6C
  loc_00412F44: push edx
  loc_00412F45: push eax
  loc_00412F46: mov var_64, 00403D64h
  loc_00412F4D: mov var_6C, 00008008h
  loc_00412F54: call [0040109Ch] ; __vbaVarTstEq
  loc_00412F5A: lea ecx, var_3C
  loc_00412F5D: lea edx, var_2C
  loc_00412F60: push ecx
  loc_00412F61: push edx
  loc_00412F62: push 00000002h
  loc_00412F64: mov var_A0, ax
  loc_00412F6B: call edi
  loc_00412F6D: add esp, 0000000Ch
  loc_00412F70: cmp var_A0, 0000h
  loc_00412F78: jz 00413037h
  loc_00412F7E: mov eax, [esi]
  loc_00412F80: push esi
  loc_00412F81: call [eax+00000304h]
  loc_00412F87: lea ecx, var_1C
  loc_00412F8A: push eax
  loc_00412F8B: push ecx
  loc_00412F8C: call [0040105Ch] ; __vbaObjSet
  loc_00412F92: mov esi, eax
  loc_00412F94: push esi
  loc_00412F95: mov edx, [esi]
  loc_00412F97: call [edx+00000204h]
  loc_00412F9D: test eax, eax
  loc_00412F9F: fnclex
  loc_00412FA1: jge 00412FB5h
  loc_00412FA3: push 00000204h
  loc_00412FA8: push 00403848h
  loc_00412FAD: push esi
  loc_00412FAE: push eax
  loc_00412FAF: call [00401040h] ; __vbaHresultCheckObj
  loc_00412FB5: lea ecx, var_1C
  loc_00412FB8: call [00401180h] ; __vbaFreeObj
  loc_00412FBE: mov esi, [00401150h] ; __vbaVarDup
  loc_00412FC4: mov ecx, 80020004h
  loc_00412FC9: mov var_54, ecx
  loc_00412FCC: mov eax, 0000000Ah
  loc_00412FD1: mov var_44, ecx
  loc_00412FD4: mov ebx, 00000008h
  loc_00412FD9: lea edx, var_7C
  loc_00412FDC: lea ecx, var_3C
  loc_00412FDF: mov var_5C, eax
  loc_00412FE2: mov var_4C, eax
  loc_00412FE5: mov var_74, 00403C5Ch ; "Check Name"
  loc_00412FEC: mov var_7C, ebx
  loc_00412FEF: call __vbaVarDup
  loc_00412FF1: mov var_64, 00404120h ; "Please enter a value from the t-table for a right-sided test."
  loc_00412FF8: lea edx, var_6C
  loc_00412FFB: lea ecx, var_2C
  loc_00412FFE: mov var_6C, ebx
  loc_00413001: call __vbaVarDup
  loc_00413003: lea eax, var_5C
  loc_00413006: lea ecx, var_4C
  loc_00413009: push eax
  loc_0041300A: lea edx, var_3C
  loc_0041300D: push ecx
  loc_0041300E: push edx
  loc_0041300F: lea eax, var_2C
  loc_00413012: push 00000000h
  loc_00413014: push eax
  loc_00413015: call [00401060h] ; rtcMsgBox
  loc_0041301B: lea ecx, var_5C
  loc_0041301E: lea edx, var_4C
  loc_00413021: push ecx
  loc_00413022: lea eax, var_3C
  loc_00413025: push edx
  loc_00413026: lea ecx, var_2C
  loc_00413029: push eax
  loc_0041302A: push ecx
  loc_0041302B: push 00000004h
  loc_0041302D: call edi
  loc_0041302F: add esp, 00000014h
  loc_00413032: jmp 0041349Ch
  loc_00413037: mov edx, [esi]
  loc_00413039: push esi
  loc_0041303A: call [edx+00000320h]
  loc_00413040: mov var_24, eax
  loc_00413043: lea eax, var_2C
  loc_00413046: lea ecx, var_3C
  loc_00413049: push eax
  loc_0041304A: push ecx
  loc_0041304B: mov var_2C, 00000009h
  loc_00413052: call ebx
  loc_00413054: lea edx, var_3C
  loc_00413057: lea eax, var_6C
  loc_0041305A: push edx
  loc_0041305B: push eax
  loc_0041305C: mov var_64, 00403D64h
  loc_00413063: mov var_6C, 00008008h
  loc_0041306A: call [0040109Ch] ; __vbaVarTstEq
  loc_00413070: lea ecx, var_3C
  loc_00413073: lea edx, var_2C
  loc_00413076: push ecx
  loc_00413077: push edx
  loc_00413078: push 00000002h
  loc_0041307A: mov var_A0, ax
  loc_00413081: call edi
  loc_00413083: add esp, 0000000Ch
  loc_00413086: cmp var_A0, 0000h
  loc_0041308E: jz 00413150h
  loc_00413094: mov ebx, [00401150h] ; __vbaVarDup
  loc_0041309A: mov ecx, 80020004h
  loc_0041309F: mov var_54, ecx
  loc_004130A2: mov eax, 0000000Ah
  loc_004130A7: mov var_44, ecx
  loc_004130AA: lea edx, var_7C
  loc_004130AD: lea ecx, var_3C
  loc_004130B0: mov var_5C, eax
  loc_004130B3: mov var_4C, eax
  loc_004130B6: mov var_74, 0040424Ch ; "Check ttable"
  loc_004130BD: mov var_7C, 00000008h
  loc_004130C4: call ebx
  loc_004130C6: lea edx, var_6C
  loc_004130C9: lea ecx, var_2C
  loc_004130CC: mov var_64, 004041A0h ; "Please enter a value from the t-table for a two-sided test or confidence interval."
  loc_004130D3: mov var_6C, 00000008h
  loc_004130DA: call ebx
  loc_004130DC: lea eax, var_5C
  loc_004130DF: lea ecx, var_4C
  loc_004130E2: push eax
  loc_004130E3: lea edx, var_3C
  loc_004130E6: push ecx
  loc_004130E7: push edx
  loc_004130E8: lea eax, var_2C
  loc_004130EB: push 00000000h
  loc_004130ED: push eax
  loc_004130EE: call [00401060h] ; rtcMsgBox
  loc_004130F4: lea ecx, var_5C
  loc_004130F7: lea edx, var_4C
  loc_004130FA: push ecx
  loc_004130FB: lea eax, var_3C
  loc_004130FE: push edx
  loc_004130FF: lea ecx, var_2C
  loc_00413102: push eax
  loc_00413103: push ecx
  loc_00413104: push 00000004h
  loc_00413106: call edi
  loc_00413108: mov edx, [esi]
  loc_0041310A: add esp, 00000014h
  loc_0041310D: push esi
  loc_0041310E: call [edx+00000320h]
  loc_00413114: push eax
  loc_00413115: lea eax, var_1C
  loc_00413118: push eax
  loc_00413119: call [0040105Ch] ; __vbaObjSet
  loc_0041311F: mov esi, eax
  loc_00413121: push esi
  loc_00413122: mov ecx, [esi]
  loc_00413124: call [ecx+00000204h]
  loc_0041312A: test eax, eax
  loc_0041312C: fnclex
  loc_0041312E: jge 00413142h
  loc_00413130: push 00000204h
  loc_00413135: push 00403848h
  loc_0041313A: push esi
  loc_0041313B: push eax
  loc_0041313C: call [00401040h] ; __vbaHresultCheckObj
  loc_00413142: lea ecx, var_1C
  loc_00413145: call [00401180h] ; __vbaFreeObj
  loc_0041314B: jmp 0041349Ch
  loc_00413150: mov edx, [esi]
  loc_00413152: push esi
  loc_00413153: call [edx+00000318h]
  loc_00413159: mov var_24, eax
  loc_0041315C: lea eax, var_2C
  loc_0041315F: lea ecx, var_3C
  loc_00413162: push eax
  loc_00413163: push ecx
  loc_00413164: mov var_2C, 00000009h
  loc_0041316B: call ebx
  loc_0041316D: lea edx, var_3C
  loc_00413170: push edx
  loc_00413171: call [004010DCh] ; __vbaR4ErrVar
  loc_00413177: fcomp real4 ptr [004014E8h]
  loc_0041317D: mov var_BC, 00000001h
  loc_00413187: fnstsw ax
  loc_00413189: test ah, 41h
  loc_0041318C: jz 00413198h
  loc_0041318E: mov var_BC, 00000000h
  loc_00413198: lea eax, var_3C
  loc_0041319B: lea ecx, var_3C
  loc_0041319E: push eax
  loc_0041319F: lea edx, var_2C
  loc_004131A2: push ecx
  loc_004131A3: push edx
  loc_004131A4: push 00000003h
  loc_004131A6: call edi
  loc_004131A8: mov eax, var_BC
  loc_004131AE: add esp, 00000010h
  loc_004131B1: neg eax
  loc_004131B3: test ax, ax
  loc_004131B6: jz 004132B9h
  loc_004131BC: mov ebx, [00401150h] ; __vbaVarDup
  loc_004131C2: mov ecx, 80020004h
  loc_004131C7: mov var_54, ecx
  loc_004131CA: mov eax, 0000000Ah
  loc_004131CF: mov var_44, ecx
  loc_004131D2: lea edx, var_7C
  loc_004131D5: lea ecx, var_3C
  loc_004131D8: mov var_5C, eax
  loc_004131DB: mov var_4C, eax
  loc_004131DE: mov var_74, 004042B8h ; "Check Alpha"
  loc_004131E5: mov var_7C, 00000008h
  loc_004131EC: call ebx
  loc_004131EE: lea edx, var_6C
  loc_004131F1: lea ecx, var_2C
  loc_004131F4: mov var_64, 0040426Ch ; "Alpha can not be greater than 1.00"
  loc_004131FB: mov var_6C, 00000008h
  loc_00413202: call ebx
  loc_00413204: lea eax, var_5C
  loc_00413207: lea ecx, var_4C
  loc_0041320A: push eax
  loc_0041320B: lea edx, var_3C
  loc_0041320E: push ecx
  loc_0041320F: push edx
  loc_00413210: lea eax, var_2C
  loc_00413213: push 00000000h
  loc_00413215: push eax
  loc_00413216: call [00401060h] ; rtcMsgBox
  loc_0041321C: lea ecx, var_5C
  loc_0041321F: lea edx, var_4C
  loc_00413222: push ecx
  loc_00413223: lea eax, var_3C
  loc_00413226: push edx
  loc_00413227: lea ecx, var_2C
  loc_0041322A: push eax
  loc_0041322B: push ecx
  loc_0041322C: push 00000004h
  loc_0041322E: call edi
  loc_00413230: mov edx, [esi]
  loc_00413232: add esp, 00000014h
  loc_00413235: push esi
  loc_00413236: call [edx+00000318h]
  loc_0041323C: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_00413242: push eax
  loc_00413243: lea eax, var_1C
  loc_00413246: push eax
  loc_00413247: call ebx
  loc_00413249: mov edi, eax
  loc_0041324B: push 00403D64h
  loc_00413250: push edi
  loc_00413251: mov ecx, [edi]
  loc_00413253: call [ecx+000000A4h]
  loc_00413259: test eax, eax
  loc_0041325B: fnclex
  loc_0041325D: jge 00413271h
  loc_0041325F: push 000000A4h
  loc_00413264: push 00403848h
  loc_00413269: push edi
  loc_0041326A: push eax
  loc_0041326B: call [00401040h] ; __vbaHresultCheckObj
  loc_00413271: mov edi, [00401180h] ; __vbaFreeObj
  loc_00413277: lea ecx, var_1C
  loc_0041327A: call edi
  loc_0041327C: mov edx, [esi]
  loc_0041327E: push esi
  loc_0041327F: call [edx+00000318h]
  loc_00413285: push eax
  loc_00413286: lea eax, var_1C
  loc_00413289: push eax
  loc_0041328A: call ebx
  loc_0041328C: mov esi, eax
  loc_0041328E: push esi
  loc_0041328F: mov ecx, [esi]
  loc_00413291: call [ecx+00000204h]
  loc_00413297: test eax, eax
  loc_00413299: fnclex
  loc_0041329B: jge 004132AFh
  loc_0041329D: push 00000204h
  loc_004132A2: push 00403848h
  loc_004132A7: push esi
  loc_004132A8: push eax
  loc_004132A9: call [00401040h] ; __vbaHresultCheckObj
  loc_004132AF: lea ecx, var_1C
  loc_004132B2: call edi
  loc_004132B4: jmp 0041349Ch
  loc_004132B9: mov edx, [esi]
  loc_004132BB: push esi
  loc_004132BC: call [edx+00000318h]
  loc_004132C2: push eax
  loc_004132C3: lea eax, var_1C
  loc_004132C6: push eax
  loc_004132C7: call [0040105Ch] ; __vbaObjSet
  loc_004132CD: mov ecx, [eax]
  loc_004132CF: lea edx, var_18
  loc_004132D2: push edx
  loc_004132D3: push eax
  loc_004132D4: mov var_A0, eax
  loc_004132DA: call [ecx+000000A0h]
  loc_004132E0: test eax, eax
  loc_004132E2: fnclex
  loc_004132E4: jge 004132FEh
  loc_004132E6: mov ecx, var_A0
  loc_004132EC: push 000000A0h
  loc_004132F1: push 00403848h
  loc_004132F6: push ecx
  loc_004132F7: push eax
  loc_004132F8: call [00401040h] ; __vbaHresultCheckObj
  loc_004132FE: mov eax, var_18
  loc_00413301: lea edx, var_2C
  loc_00413304: mov var_24, eax
  loc_00413307: lea eax, var_3C
  loc_0041330A: push edx
  loc_0041330B: push eax
  loc_0041330C: mov var_18, 00000000h
  loc_00413313: mov var_2C, 00000008h
  loc_0041331A: call ebx
  loc_0041331C: lea ecx, var_3C
  loc_0041331F: push ecx
  loc_00413320: call [004010DCh] ; __vbaR4ErrVar
  loc_00413326: fstp real4 ptr [00415020h]
  loc_0041332C: lea ecx, var_1C
  loc_0041332F: call [00401180h] ; __vbaFreeObj
  loc_00413335: lea edx, var_3C
  loc_00413338: lea eax, var_3C
  loc_0041333B: push edx
  loc_0041333C: lea ecx, var_2C
  loc_0041333F: push eax
  loc_00413340: push ecx
  loc_00413341: push 00000003h
  loc_00413343: call edi
  loc_00413345: mov edx, [esi]
  loc_00413347: add esp, 00000010h
  loc_0041334A: push esi
  loc_0041334B: call [edx+00000320h]
  loc_00413351: push eax
  loc_00413352: lea eax, var_1C
  loc_00413355: push eax
  loc_00413356: call [0040105Ch] ; __vbaObjSet
  loc_0041335C: mov ecx, [eax]
  loc_0041335E: lea edx, var_18
  loc_00413361: push edx
  loc_00413362: push eax
  loc_00413363: mov var_A0, eax
  loc_00413369: call [ecx+000000A0h]
  loc_0041336F: test eax, eax
  loc_00413371: fnclex
  loc_00413373: jge 0041338Dh
  loc_00413375: mov ecx, var_A0
  loc_0041337B: push 000000A0h
  loc_00413380: push 00403848h
  loc_00413385: push ecx
  loc_00413386: push eax
  loc_00413387: call [00401040h] ; __vbaHresultCheckObj
  loc_0041338D: mov eax, var_18
  loc_00413390: lea edx, var_2C
  loc_00413393: mov var_24, eax
  loc_00413396: lea eax, var_3C
  loc_00413399: push edx
  loc_0041339A: push eax
  loc_0041339B: mov var_18, 00000000h
  loc_004133A2: mov var_2C, 00000008h
  loc_004133A9: call ebx
  loc_004133AB: lea ecx, var_3C
  loc_004133AE: push ecx
  loc_004133AF: call [004010DCh] ; __vbaR4ErrVar
  loc_004133B5: fstp real4 ptr [00415028h]
  loc_004133BB: lea ecx, var_1C
  loc_004133BE: call [00401180h] ; __vbaFreeObj
  loc_004133C4: lea edx, var_3C
  loc_004133C7: lea eax, var_3C
  loc_004133CA: push edx
  loc_004133CB: lea ecx, var_2C
  loc_004133CE: push eax
  loc_004133CF: push ecx
  loc_004133D0: push 00000003h
  loc_004133D2: call edi
  loc_004133D4: mov edx, [esi]
  loc_004133D6: add esp, 00000010h
  loc_004133D9: push esi
  loc_004133DA: call [edx+00000304h]
  loc_004133E0: push eax
  loc_004133E1: lea eax, var_1C
  loc_004133E4: push eax
  loc_004133E5: call [0040105Ch] ; __vbaObjSet
  loc_004133EB: mov esi, eax
  loc_004133ED: lea edx, var_18
  loc_004133F0: push edx
  loc_004133F1: push esi
  loc_004133F2: mov ecx, [esi]
  loc_004133F4: call [ecx+000000A0h]
  loc_004133FA: test eax, eax
  loc_004133FC: fnclex
  loc_004133FE: jge 00413412h
  loc_00413400: push 000000A0h
  loc_00413405: push 00403848h
  loc_0041340A: push esi
  loc_0041340B: push eax
  loc_0041340C: call [00401040h] ; __vbaHresultCheckObj
  loc_00413412: mov eax, var_18
  loc_00413415: lea ecx, var_3C
  loc_00413418: mov var_24, eax
  loc_0041341B: lea eax, var_2C
  loc_0041341E: push eax
  loc_0041341F: push ecx
  loc_00413420: mov var_18, 00000000h
  loc_00413427: mov var_2C, 00000008h
  loc_0041342E: call ebx
  loc_00413430: lea edx, var_3C
  loc_00413433: push edx
  loc_00413434: call [004010DCh] ; __vbaR4ErrVar
  loc_0041343A: fstp real4 ptr [00415024h]
  loc_00413440: lea ecx, var_1C
  loc_00413443: call [00401180h] ; __vbaFreeObj
  loc_00413449: lea eax, var_3C
  loc_0041344C: lea ecx, var_3C
  loc_0041344F: push eax
  loc_00413450: lea edx, var_2C
  loc_00413453: push ecx
  loc_00413454: push edx
  loc_00413455: push 00000003h
  loc_00413457: call edi
  loc_00413459: mov eax, [0041520Ch]
  loc_0041345E: add esp, 00000010h
  loc_00413461: test eax, eax
  loc_00413463: jnz 00413475h
  loc_00413465: push 0041520Ch
  loc_0041346A: push 00402644h
  loc_0041346F: call [0040111Ch] ; __vbaNew2
  loc_00413475: mov esi, [0041520Ch]
  loc_0041347B: push esi
  loc_0041347C: mov eax, [esi]
  loc_0041347E: call [eax+000002B4h]
  loc_00413484: test eax, eax
  loc_00413486: fnclex
  loc_00413488: jge 0041349Ch
  loc_0041348A: push 000002B4h
  loc_0041348F: push 00403A74h
  loc_00413494: push esi
  loc_00413495: push eax
  loc_00413496: call [00401040h] ; __vbaHresultCheckObj
  loc_0041349C: mov var_4, 00000000h
  loc_004134A3: fwait
  loc_004134A4: push 004134DAh
  loc_004134A9: jmp 004134D9h
  loc_004134AB: lea ecx, var_18
  loc_004134AE: call [00401184h] ; __vbaFreeStr
  loc_004134B4: lea ecx, var_1C
  loc_004134B7: call [00401180h] ; __vbaFreeObj
  loc_004134BD: lea ecx, var_5C
  loc_004134C0: lea edx, var_4C
  loc_004134C3: push ecx
  loc_004134C4: lea eax, var_3C
  loc_004134C7: push edx
  loc_004134C8: lea ecx, var_2C
  loc_004134CB: push eax
  loc_004134CC: push ecx
  loc_004134CD: push 00000004h
  loc_004134CF: call [00401024h] ; __vbaFreeVarList
  loc_004134D5: add esp, 00000014h
  loc_004134D8: ret
  loc_004134D9: ret
  loc_004134DA: mov eax, Me
  loc_004134DD: push eax
  loc_004134DE: mov edx, [eax]
  loc_004134E0: call [edx+00000008h]
  loc_004134E3: mov eax, var_4
  loc_004134E6: mov ecx, var_14
  loc_004134E9: pop edi
  loc_004134EA: pop esi
  loc_004134EB: mov fs:[00000000h], ecx
  loc_004134F2: pop ebx
  loc_004134F3: mov esp, ebp
  loc_004134F5: pop ebp
  loc_004134F6: retn 0004h
End Sub

Private Sub cmdRestore_Click() '412C10
  loc_00412C10: push ebp
  loc_00412C11: mov ebp, esp
  loc_00412C13: sub esp, 0000000Ch
  loc_00412C16: push 00401546h ; __vbaExceptHandler
  loc_00412C1B: mov eax, fs:[00000000h]
  loc_00412C21: push eax
  loc_00412C22: mov fs:[00000000h], esp
  loc_00412C29: sub esp, 00000018h
  loc_00412C2C: push ebx
  loc_00412C2D: push esi
  loc_00412C2E: push edi
  loc_00412C2F: mov var_C, esp
  loc_00412C32: mov var_8, 004014D8h
  loc_00412C39: mov esi, Me
  loc_00412C3C: mov eax, esi
  loc_00412C3E: and eax, 00000001h
  loc_00412C41: mov var_4, eax
  loc_00412C44: and esi, FFFFFFFEh
  loc_00412C47: push esi
  loc_00412C48: mov Me, esi
  loc_00412C4B: mov ecx, [esi]
  loc_00412C4D: call [ecx+00000004h]
  loc_00412C50: mov [00415020h], 3D4CCCCDh
  loc_00412C5A: mov [00415028h], 3FFAE148h
  loc_00412C64: mov [00415024h], 3FD28F5Ch
  loc_00412C6E: mov edx, [esi]
  loc_00412C70: xor eax, eax
  loc_00412C72: push esi
  loc_00412C73: mov var_18, eax
  loc_00412C76: mov var_1C, eax
  loc_00412C79: call [edx+00000318h]
  loc_00412C7F: push eax
  loc_00412C80: lea eax, var_1C
  loc_00412C83: push eax
  loc_00412C84: call [0040105Ch] ; __vbaObjSet
  loc_00412C8A: mov ecx, [00415020h]
  loc_00412C90: mov edi, eax
  loc_00412C92: push ecx
  loc_00412C93: mov ebx, [edi]
  loc_00412C95: call [004010A8h] ; __vbaStrR4
  loc_00412C9B: mov edx, eax
  loc_00412C9D: lea ecx, var_18
  loc_00412CA0: call [00401164h] ; __vbaStrMove
  loc_00412CA6: push eax
  loc_00412CA7: push edi
  loc_00412CA8: call [ebx+000000A4h]
  loc_00412CAE: test eax, eax
  loc_00412CB0: fnclex
  loc_00412CB2: jge 00412CC6h
  loc_00412CB4: push 000000A4h
  loc_00412CB9: push 00403848h
  loc_00412CBE: push edi
  loc_00412CBF: push eax
  loc_00412CC0: call [00401040h] ; __vbaHresultCheckObj
  loc_00412CC6: lea ecx, var_18
  loc_00412CC9: call [00401184h] ; __vbaFreeStr
  loc_00412CCF: lea ecx, var_1C
  loc_00412CD2: call [00401180h] ; __vbaFreeObj
  loc_00412CD8: mov edx, [esi]
  loc_00412CDA: push esi
  loc_00412CDB: call [edx+00000320h]
  loc_00412CE1: push eax
  loc_00412CE2: lea eax, var_1C
  loc_00412CE5: push eax
  loc_00412CE6: call [0040105Ch] ; __vbaObjSet
  loc_00412CEC: mov ecx, [00415028h]
  loc_00412CF2: mov edi, eax
  loc_00412CF4: push ecx
  loc_00412CF5: mov ebx, [edi]
  loc_00412CF7: call [004010A8h] ; __vbaStrR4
  loc_00412CFD: mov edx, eax
  loc_00412CFF: lea ecx, var_18
  loc_00412D02: call [00401164h] ; __vbaStrMove
  loc_00412D08: push eax
  loc_00412D09: push edi
  loc_00412D0A: call [ebx+000000A4h]
  loc_00412D10: test eax, eax
  loc_00412D12: fnclex
  loc_00412D14: jge 00412D28h
  loc_00412D16: push 000000A4h
  loc_00412D1B: push 00403848h
  loc_00412D20: push edi
  loc_00412D21: push eax
  loc_00412D22: call [00401040h] ; __vbaHresultCheckObj
  loc_00412D28: mov edi, [00401184h] ; __vbaFreeStr
  loc_00412D2E: lea ecx, var_18
  loc_00412D31: call edi
  loc_00412D33: lea ecx, var_1C
  loc_00412D36: call [00401180h] ; __vbaFreeObj
  loc_00412D3C: mov edx, [esi]
  loc_00412D3E: push esi
  loc_00412D3F: call [edx+00000304h]
  loc_00412D45: push eax
  loc_00412D46: lea eax, var_1C
  loc_00412D49: push eax
  loc_00412D4A: call [0040105Ch] ; __vbaObjSet
  loc_00412D50: mov ecx, [00415024h]
  loc_00412D56: mov esi, eax
  loc_00412D58: push ecx
  loc_00412D59: mov ebx, [esi]
  loc_00412D5B: call [004010A8h] ; __vbaStrR4
  loc_00412D61: mov edx, eax
  loc_00412D63: lea ecx, var_18
  loc_00412D66: call [00401164h] ; __vbaStrMove
  loc_00412D6C: push eax
  loc_00412D6D: push esi
  loc_00412D6E: call [ebx+000000A4h]
  loc_00412D74: test eax, eax
  loc_00412D76: fnclex
  loc_00412D78: jge 00412D8Ch
  loc_00412D7A: push 000000A4h
  loc_00412D7F: push 00403848h
  loc_00412D84: push esi
  loc_00412D85: push eax
  loc_00412D86: call [00401040h] ; __vbaHresultCheckObj
  loc_00412D8C: lea ecx, var_18
  loc_00412D8F: call edi
  loc_00412D91: lea ecx, var_1C
  loc_00412D94: call [00401180h] ; __vbaFreeObj
  loc_00412D9A: mov var_4, 00000000h
  loc_00412DA1: fwait
  loc_00412DA2: push 00412DBDh
  loc_00412DA7: jmp 00412DBCh
  loc_00412DA9: lea ecx, var_18
  loc_00412DAC: call [00401184h] ; __vbaFreeStr
  loc_00412DB2: lea ecx, var_1C
  loc_00412DB5: call [00401180h] ; __vbaFreeObj
  loc_00412DBB: ret
  loc_00412DBC: ret
  loc_00412DBD: mov eax, Me
  loc_00412DC0: push eax
  loc_00412DC1: mov edx, [eax]
  loc_00412DC3: call [edx+00000008h]
  loc_00412DC6: mov eax, var_4
  loc_00412DC9: mov ecx, var_14
  loc_00412DCC: pop edi
  loc_00412DCD: pop esi
  loc_00412DCE: mov fs:[00000000h], ecx
  loc_00412DD5: pop ebx
  loc_00412DD6: mov esp, ebp
  loc_00412DD8: pop ebp
  loc_00412DD9: retn 0004h
End Sub

Private Sub txtTtable_KeyPress(KeyAscii As Integer) '4138E0
  loc_004138E0: push ebp
  loc_004138E1: mov ebp, esp
  loc_004138E3: sub esp, 0000000Ch
  loc_004138E6: push 00401546h ; __vbaExceptHandler
  loc_004138EB: mov eax, fs:[00000000h]
  loc_004138F1: push eax
  loc_004138F2: mov fs:[00000000h], esp
  loc_004138F9: sub esp, 0000007Ch
  loc_004138FC: push ebx
  loc_004138FD: push esi
  loc_004138FE: push edi
  loc_004138FF: mov var_C, esp
  loc_00413902: mov var_8, 00401520h
  loc_00413909: mov esi, Me
  loc_0041390C: mov eax, esi
  loc_0041390E: and eax, 00000001h
  loc_00413911: mov var_4, eax
  loc_00413914: and esi, FFFFFFFEh
  loc_00413917: push esi
  loc_00413918: mov Me, esi
  loc_0041391B: mov ecx, [esi]
  loc_0041391D: call [ecx+00000004h]
  loc_00413920: mov ebx, KeyAscii
  loc_00413923: lea edx, var_24
  loc_00413926: xor edi, edi
  loc_00413928: push ebx
  loc_00413929: push edx
  loc_0041392A: mov var_24, edi
  loc_0041392D: mov var_34, edi
  loc_00413930: mov var_44, edi
  loc_00413933: mov var_54, edi
  loc_00413936: mov var_64, edi
  loc_00413939: mov var_74, edi
  loc_0041393C: mov var_84, edi
  loc_00413942: call 00408F40h
  loc_00413947: lea eax, var_24
  loc_0041394A: push eax
  loc_0041394B: call [00401104h] ; __vbaI2Var
  loc_00413951: lea ecx, var_24
  loc_00413954: mov [ebx], ax
  loc_00413957: call [00401014h] ; __vbaFreeVar
  loc_0041395D: push 004035B0h ; "."
  loc_00413962: call [00401030h] ; rtcAnsiValueBstr
  loc_00413968: mov edx, [esi]
  loc_0041396A: xor ecx, ecx
  loc_0041396C: cmp [ebx], ax
  loc_0041396F: push esi
  loc_00413970: mov var_84, 0000000Bh
  loc_0041397A: setz cl
  loc_0041397D: neg ecx
  loc_0041397F: mov var_7C, cx
  loc_00413983: call [edx+00000320h]
  loc_00413989: mov var_1C, eax
  loc_0041398C: lea eax, var_84
  loc_00413992: push eax
  loc_00413993: lea ecx, var_24
  loc_00413996: push 00000001h
  loc_00413998: lea edx, var_64
  loc_0041399B: push ecx
  loc_0041399C: push edx
  loc_0041399D: lea eax, var_34
  loc_004139A0: push edi
  loc_004139A1: push eax
  loc_004139A2: mov var_24, 00000009h
  loc_004139A9: mov var_5C, 004035B0h ; "."
  loc_004139B0: mov var_64, 00000008h
  loc_004139B7: mov var_6C, edi
  loc_004139BA: mov var_74, 00008002h
  loc_004139C1: call [004010F4h] ; __vbaInStrVar
  loc_004139C7: lea ecx, var_74
  loc_004139CA: push eax
  loc_004139CB: lea edx, var_44
  loc_004139CE: push ecx
  loc_004139CF: push edx
  loc_004139D0: call [00401084h] ; __vbaVarCmpGt
  loc_004139D6: push eax
  loc_004139D7: lea eax, var_54
  loc_004139DA: push eax
  loc_004139DB: call [004010C4h] ; __vbaVarAnd
  loc_004139E1: push eax
  loc_004139E2: call [0040107Ch] ; __vbaBoolVarNull
  loc_004139E8: lea ecx, var_84
  loc_004139EE: mov esi, eax
  loc_004139F0: lea edx, var_34
  loc_004139F3: push ecx
  loc_004139F4: lea eax, var_24
  loc_004139F7: push edx
  loc_004139F8: push eax
  loc_004139F9: push 00000003h
  loc_004139FB: call [00401024h] ; __vbaFreeVarList
  loc_00413A01: add esp, 00000010h
  loc_00413A04: cmp si, di
  loc_00413A07: jz 00413A0Ch
  loc_00413A09: mov [ebx], di
  loc_00413A0C: mov var_4, edi
  loc_00413A0F: push 00413A33h
  loc_00413A14: jmp 00413A32h
  loc_00413A16: lea ecx, var_54
  loc_00413A19: lea edx, var_44
  loc_00413A1C: push ecx
  loc_00413A1D: lea eax, var_34
  loc_00413A20: push edx
  loc_00413A21: lea ecx, var_24
  loc_00413A24: push eax
  loc_00413A25: push ecx
  loc_00413A26: push 00000004h
  loc_00413A28: call [00401024h] ; __vbaFreeVarList
  loc_00413A2E: add esp, 00000014h
  loc_00413A31: ret
  loc_00413A32: ret
  loc_00413A33: mov eax, Me
  loc_00413A36: push eax
  loc_00413A37: mov edx, [eax]
  loc_00413A39: call [edx+00000008h]
  loc_00413A3C: mov eax, var_4
  loc_00413A3F: mov ecx, var_14
  loc_00413A42: pop edi
  loc_00413A43: pop esi
  loc_00413A44: mov fs:[00000000h], ecx
  loc_00413A4B: pop ebx
  loc_00413A4C: mov esp, ebp
  loc_00413A4E: pop ebp
  loc_00413A4F: retn 0008h
End Sub

Private Sub txtAlpha_KeyPress(KeyAscii As Integer) '4136F0
  loc_004136F0: push ebp
  loc_004136F1: mov ebp, esp
  loc_004136F3: sub esp, 0000000Ch
  loc_004136F6: push 00401546h ; __vbaExceptHandler
  loc_004136FB: mov eax, fs:[00000000h]
  loc_00413701: push eax
  loc_00413702: mov fs:[00000000h], esp
  loc_00413709: sub esp, 0000008Ch
  loc_0041370F: push ebx
  loc_00413710: push esi
  loc_00413711: push edi
  loc_00413712: mov var_C, esp
  loc_00413715: mov var_8, 00401510h
  loc_0041371C: mov esi, Me
  loc_0041371F: mov eax, esi
  loc_00413721: and eax, 00000001h
  loc_00413724: mov var_4, eax
  loc_00413727: and esi, FFFFFFFEh
  loc_0041372A: push esi
  loc_0041372B: mov Me, esi
  loc_0041372E: mov ecx, [esi]
  loc_00413730: call [ecx+00000004h]
  loc_00413733: mov edi, KeyAscii
  loc_00413736: lea edx, var_24
  loc_00413739: xor ebx, ebx
  loc_0041373B: push edi
  loc_0041373C: push edx
  loc_0041373D: mov var_24, ebx
  loc_00413740: mov var_34, ebx
  loc_00413743: mov var_44, ebx
  loc_00413746: mov var_54, ebx
  loc_00413749: mov var_64, ebx
  loc_0041374C: mov var_74, ebx
  loc_0041374F: mov var_84, ebx
  loc_00413755: call 00408F40h
  loc_0041375A: lea eax, var_24
  loc_0041375D: push eax
  loc_0041375E: call [00401104h] ; __vbaI2Var
  loc_00413764: lea ecx, var_24
  loc_00413767: mov [edi], ax
  loc_0041376A: call [00401014h] ; __vbaFreeVar
  loc_00413770: push 004035B0h ; "."
  loc_00413775: call [00401030h] ; rtcAnsiValueBstr
  loc_0041377B: mov edx, [esi]
  loc_0041377D: xor ecx, ecx
  loc_0041377F: cmp [edi], ax
  loc_00413782: push esi
  loc_00413783: mov var_84, 0000000Bh
  loc_0041378D: setz cl
  loc_00413790: neg ecx
  loc_00413792: mov var_7C, cx
  loc_00413796: call [edx+00000318h]
  loc_0041379C: mov var_1C, eax
  loc_0041379F: lea eax, var_84
  loc_004137A5: push eax
  loc_004137A6: lea ecx, var_24
  loc_004137A9: push 00000001h
  loc_004137AB: lea edx, var_64
  loc_004137AE: push ecx
  loc_004137AF: push edx
  loc_004137B0: lea eax, var_34
  loc_004137B3: push ebx
  loc_004137B4: push eax
  loc_004137B5: mov var_24, 00000009h
  loc_004137BC: mov var_5C, 004035B0h ; "."
  loc_004137C3: mov var_64, 00000008h
  loc_004137CA: mov var_6C, ebx
  loc_004137CD: mov var_74, 00008002h
  loc_004137D4: call [004010F4h] ; __vbaInStrVar
  loc_004137DA: lea ecx, var_74
  loc_004137DD: push eax
  loc_004137DE: lea edx, var_44
  loc_004137E1: push ecx
  loc_004137E2: push edx
  loc_004137E3: call [00401084h] ; __vbaVarCmpGt
  loc_004137E9: push eax
  loc_004137EA: lea eax, var_54
  loc_004137ED: push eax
  loc_004137EE: call [004010C4h] ; __vbaVarAnd
  loc_004137F4: push eax
  loc_004137F5: call [0040107Ch] ; __vbaBoolVarNull
  loc_004137FB: lea ecx, var_84
  loc_00413801: mov esi, eax
  loc_00413803: lea edx, var_34
  loc_00413806: push ecx
  loc_00413807: lea eax, var_24
  loc_0041380A: push edx
  loc_0041380B: push eax
  loc_0041380C: push 00000003h
  loc_0041380E: call [00401024h] ; __vbaFreeVarList
  loc_00413814: add esp, 00000010h
  loc_00413817: cmp si, bx
  loc_0041381A: jz 0041381Fh
  loc_0041381C: mov [edi], bx
  loc_0041381F: push 004035B8h ; "-"
  loc_00413824: call [00401030h] ; rtcAnsiValueBstr
  loc_0041382A: cmp [edi], ax
  loc_0041382D: jnz 0041389Ah
  loc_0041382F: mov ecx, 80020004h
  loc_00413834: mov eax, 0000000Ah
  loc_00413839: mov var_4C, ecx
  loc_0041383C: mov var_3C, ecx
  loc_0041383F: mov var_2C, ecx
  loc_00413842: lea edx, var_64
  loc_00413845: lea ecx, var_24
  loc_00413848: mov var_54, eax
  loc_0041384B: mov var_44, eax
  loc_0041384E: mov var_34, eax
  loc_00413851: mov var_5C, 004042D4h ; "Alpha is always positive."
  loc_00413858: mov var_64, 00000008h
  loc_0041385F: call [00401150h] ; __vbaVarDup
  loc_00413865: lea ecx, var_54
  loc_00413868: lea edx, var_44
  loc_0041386B: push ecx
  loc_0041386C: lea eax, var_34
  loc_0041386F: push edx
  loc_00413870: push eax
  loc_00413871: lea ecx, var_24
  loc_00413874: push ebx
  loc_00413875: push ecx
  loc_00413876: call [00401060h] ; rtcMsgBox
  loc_0041387C: lea edx, var_54
  loc_0041387F: lea eax, var_44
  loc_00413882: push edx
  loc_00413883: lea ecx, var_34
  loc_00413886: push eax
  loc_00413887: lea edx, var_24
  loc_0041388A: push ecx
  loc_0041388B: push edx
  loc_0041388C: push 00000004h
  loc_0041388E: call [00401024h] ; __vbaFreeVarList
  loc_00413894: add esp, 00000014h
  loc_00413897: mov [edi], bx
  loc_0041389A: mov var_4, ebx
  loc_0041389D: push 004138C1h
  loc_004138A2: jmp 004138C0h
  loc_004138A4: lea eax, var_54
  loc_004138A7: lea ecx, var_44
  loc_004138AA: push eax
  loc_004138AB: lea edx, var_34
  loc_004138AE: push ecx
  loc_004138AF: lea eax, var_24
  loc_004138B2: push edx
  loc_004138B3: push eax
  loc_004138B4: push 00000004h
  loc_004138B6: call [00401024h] ; __vbaFreeVarList
  loc_004138BC: add esp, 00000014h
  loc_004138BF: ret
  loc_004138C0: ret
  loc_004138C1: mov eax, Me
  loc_004138C4: push eax
  loc_004138C5: mov ecx, [eax]
  loc_004138C7: call [ecx+00000008h]
  loc_004138CA: mov eax, var_4
  loc_004138CD: mov ecx, var_14
  loc_004138D0: pop edi
  loc_004138D1: pop esi
  loc_004138D2: mov fs:[00000000h], ecx
  loc_004138D9: pop ebx
  loc_004138DA: mov esp, ebp
  loc_004138DC: pop ebp
  loc_004138DD: retn 0008h
End Sub
