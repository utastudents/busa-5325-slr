VERSION 5.00
Begin VB.Form frmChangeSlopeEstimates
  Caption = "Change Slope Estimates"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 7095
  ClientHeight = 4410
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Calculated by hand or found on a printout"
    Left = 0
    Top = 0
    Width = 6975
    Height = 4335
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
    Begin VB.Frame Frame2
      Caption = "Estimate of Standard Error of Slope Estimate"
      Left = 360
      Top = 2040
      Width = 5775
      Height = 1335
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
      Begin VB.TextBox txtSbetaHat
        Left = 240
        Top = 480
        Width = 4335
        Height = 615
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
    End
    Begin VB.Frame Frame1
      Caption = "Slope Estimate"
      Left = 360
      Top = 600
      Width = 5775
      Height = 1215
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
      Begin VB.TextBox txtBetaHat
        Left = 240
        Top = 480
        Width = 4335
        Height = 615
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
    End
    Begin VB.CommandButton cmdRestore
      Caption = "Reset Defaults"
      Left = 3480
      Top = 3600
      Width = 2655
      Height = 495
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
    Begin VB.CommandButton cmdClearChange
      Caption = "Clear"
      Left = 1920
      Top = 3600
      Width = 1335
      Height = 495
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
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 360
      Top = 3600
      Width = 1215
      Height = 495
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
End

Attribute VB_Name = "frmChangeSlopeEstimates"


Private Sub cmdClearChange_Click() '421770
  loc_00421770: push ebp
  loc_00421771: mov ebp, esp
  loc_00421773: sub esp, 0000000Ch
  loc_00421776: push 00401926h ; __vbaExceptHandler
  loc_0042177B: mov eax, fs:[00000000h]
  loc_00421781: push eax
  loc_00421782: mov fs:[00000000h], esp
  loc_00421789: sub esp, 00000014h
  loc_0042178C: push ebx
  loc_0042178D: push esi
  loc_0042178E: push edi
  loc_0042178F: mov var_C, esp
  loc_00421792: mov var_8, 00401420h
  loc_00421799: mov esi, Me
  loc_0042179C: mov eax, esi
  loc_0042179E: and eax, 00000001h
  loc_004217A1: mov var_4, eax
  loc_004217A4: and esi, FFFFFFFEh
  loc_004217A7: push esi
  loc_004217A8: mov Me, esi
  loc_004217AB: mov ecx, [esi]
  loc_004217AD: call [ecx+00000004h]
  loc_004217B0: mov edx, [esi]
  loc_004217B2: push esi
  loc_004217B3: mov var_18, 00000000h
  loc_004217BA: call [edx+0000030Ch]
  loc_004217C0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_004217C6: push eax
  loc_004217C7: lea eax, var_18
  loc_004217CA: push eax
  loc_004217CB: call ebx
  loc_004217CD: mov edi, eax
  loc_004217CF: push 0040F040h
  loc_004217D4: push edi
  loc_004217D5: mov ecx, [edi]
  loc_004217D7: call [ecx+000000A4h]
  loc_004217DD: test eax, eax
  loc_004217DF: fnclex
  loc_004217E1: jge 004217F5h
  loc_004217E3: push 000000A4h
  loc_004217E8: push 0040F02Ch
  loc_004217ED: push edi
  loc_004217EE: push eax
  loc_004217EF: call [00401030h] ; __vbaHresultCheckObj
  loc_004217F5: lea ecx, var_18
  loc_004217F8: call [00401114h] ; __vbaFreeObj
  loc_004217FE: mov edx, [esi]
  loc_00421800: push esi
  loc_00421801: call [edx+00000304h]
  loc_00421807: push eax
  loc_00421808: lea eax, var_18
  loc_0042180B: push eax
  loc_0042180C: call ebx
  loc_0042180E: mov edi, eax
  loc_00421810: push 0040F040h
  loc_00421815: push edi
  loc_00421816: mov ecx, [edi]
  loc_00421818: call [ecx+000000A4h]
  loc_0042181E: test eax, eax
  loc_00421820: fnclex
  loc_00421822: jge 00421836h
  loc_00421824: push 000000A4h
  loc_00421829: push 0040F02Ch
  loc_0042182E: push edi
  loc_0042182F: push eax
  loc_00421830: call [00401030h] ; __vbaHresultCheckObj
  loc_00421836: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042183C: lea ecx, var_18
  loc_0042183F: call edi
  loc_00421841: mov edx, [esi]
  loc_00421843: push esi
  loc_00421844: call [edx+0000030Ch]
  loc_0042184A: push eax
  loc_0042184B: lea eax, var_18
  loc_0042184E: push eax
  loc_0042184F: call ebx
  loc_00421851: mov esi, eax
  loc_00421853: push esi
  loc_00421854: mov ecx, [esi]
  loc_00421856: call [ecx+00000204h]
  loc_0042185C: test eax, eax
  loc_0042185E: fnclex
  loc_00421860: jge 00421874h
  loc_00421862: push 00000204h
  loc_00421867: push 0040F02Ch
  loc_0042186C: push esi
  loc_0042186D: push eax
  loc_0042186E: call [00401030h] ; __vbaHresultCheckObj
  loc_00421874: lea ecx, var_18
  loc_00421877: call edi
  loc_00421879: mov var_4, 00000000h
  loc_00421880: push 00421892h
  loc_00421885: jmp 00421891h
  loc_00421887: lea ecx, var_18
  loc_0042188A: call [00401114h] ; __vbaFreeObj
  loc_00421890: ret
  loc_00421891: ret
  loc_00421892: mov eax, Me
  loc_00421895: push eax
  loc_00421896: mov edx, [eax]
  loc_00421898: call [edx+00000008h]
  loc_0042189B: mov eax, var_4
  loc_0042189E: mov ecx, var_14
  loc_004218A1: pop edi
  loc_004218A2: pop esi
  loc_004218A3: mov fs:[00000000h], ecx
  loc_004218AA: pop ebx
  loc_004218AB: mov esp, ebp
  loc_004218AD: pop ebp
  loc_004218AE: retn 0004h
End Sub

Private Sub txtBetaHat_KeyPress(KeyAscii As Integer) '422560
  loc_00422560: push ebp
  loc_00422561: mov ebp, esp
  loc_00422563: sub esp, 0000000Ch
  loc_00422566: push 00401926h ; __vbaExceptHandler
  loc_0042256B: mov eax, fs:[00000000h]
  loc_00422571: push eax
  loc_00422572: mov fs:[00000000h], esp
  loc_00422579: sub esp, 0000007Ch
  loc_0042257C: push ebx
  loc_0042257D: push esi
  loc_0042257E: push edi
  loc_0042257F: mov var_C, esp
  loc_00422582: mov var_8, 00401460h
  loc_00422589: mov esi, Me
  loc_0042258C: mov eax, esi
  loc_0042258E: and eax, 00000001h
  loc_00422591: mov var_4, eax
  loc_00422594: and esi, FFFFFFFEh
  loc_00422597: push esi
  loc_00422598: mov Me, esi
  loc_0042259B: mov ecx, [esi]
  loc_0042259D: call [ecx+00000004h]
  loc_004225A0: mov ebx, KeyAscii
  loc_004225A3: lea edx, var_24
  loc_004225A6: xor edi, edi
  loc_004225A8: push ebx
  loc_004225A9: push edx
  loc_004225AA: mov var_24, edi
  loc_004225AD: mov var_34, edi
  loc_004225B0: mov var_44, edi
  loc_004225B3: mov var_54, edi
  loc_004225B6: mov var_64, edi
  loc_004225B9: mov var_74, edi
  loc_004225BC: mov var_84, edi
  loc_004225C2: call 0041A480h
  loc_004225C7: lea eax, var_24
  loc_004225CA: push eax
  loc_004225CB: call [004010C4h] ; __vbaI2Var
  loc_004225D1: lea ecx, var_24
  loc_004225D4: mov [ebx], ax
  loc_004225D7: call [00401010h] ; __vbaFreeVar
  loc_004225DD: push 0040DD3Ch ; "."
  loc_004225E2: call [00401024h] ; rtcAnsiValueBstr
  loc_004225E8: mov edx, [esi]
  loc_004225EA: xor ecx, ecx
  loc_004225EC: cmp [ebx], ax
  loc_004225EF: push esi
  loc_004225F0: mov var_84, 0000000Bh
  loc_004225FA: setz cl
  loc_004225FD: neg ecx
  loc_004225FF: mov var_7C, cx
  loc_00422603: call [edx+0000030Ch]
  loc_00422609: mov var_1C, eax
  loc_0042260C: lea eax, var_84
  loc_00422612: push eax
  loc_00422613: lea ecx, var_24
  loc_00422616: push 00000001h
  loc_00422618: lea edx, var_64
  loc_0042261B: push ecx
  loc_0042261C: push edx
  loc_0042261D: lea eax, var_34
  loc_00422620: push edi
  loc_00422621: push eax
  loc_00422622: mov var_24, 00000009h
  loc_00422629: mov var_5C, 0040DD3Ch ; "."
  loc_00422630: mov var_64, 00000008h
  loc_00422637: mov var_6C, edi
  loc_0042263A: mov var_74, 00008002h
  loc_00422641: call [004010B4h] ; __vbaInStrVar
  loc_00422647: lea ecx, var_74
  loc_0042264A: push eax
  loc_0042264B: lea edx, var_44
  loc_0042264E: push ecx
  loc_0042264F: push edx
  loc_00422650: call [00401060h] ; __vbaVarCmpGt
  loc_00422656: push eax
  loc_00422657: lea eax, var_54
  loc_0042265A: push eax
  loc_0042265B: call [00401094h] ; __vbaVarAnd
  loc_00422661: push eax
  loc_00422662: call [00401058h] ; __vbaBoolVarNull
  loc_00422668: lea ecx, var_84
  loc_0042266E: mov esi, eax
  loc_00422670: lea edx, var_34
  loc_00422673: push ecx
  loc_00422674: lea eax, var_24
  loc_00422677: push edx
  loc_00422678: push eax
  loc_00422679: push 00000003h
  loc_0042267B: call [00401018h] ; __vbaFreeVarList
  loc_00422681: add esp, 00000010h
  loc_00422684: cmp si, di
  loc_00422687: jz 0042268Ch
  loc_00422689: mov [ebx], di
  loc_0042268C: mov var_4, edi
  loc_0042268F: push 004226B3h
  loc_00422694: jmp 004226B2h
  loc_00422696: lea ecx, var_54
  loc_00422699: lea edx, var_44
  loc_0042269C: push ecx
  loc_0042269D: lea eax, var_34
  loc_004226A0: push edx
  loc_004226A1: lea ecx, var_24
  loc_004226A4: push eax
  loc_004226A5: push ecx
  loc_004226A6: push 00000004h
  loc_004226A8: call [00401018h] ; __vbaFreeVarList
  loc_004226AE: add esp, 00000014h
  loc_004226B1: ret
  loc_004226B2: ret
  loc_004226B3: mov eax, Me
  loc_004226B6: push eax
  loc_004226B7: mov edx, [eax]
  loc_004226B9: call [edx+00000008h]
  loc_004226BC: mov eax, var_4
  loc_004226BF: mov ecx, var_14
  loc_004226C2: pop edi
  loc_004226C3: pop esi
  loc_004226C4: mov fs:[00000000h], ecx
  loc_004226CB: pop ebx
  loc_004226CC: mov esp, ebp
  loc_004226CE: pop ebp
  loc_004226CF: retn 0008h
End Sub

Private Sub txtSbetaHat_KeyPress(KeyAscii As Integer) '4226E0
  loc_004226E0: push ebp
  loc_004226E1: mov ebp, esp
  loc_004226E3: sub esp, 0000000Ch
  loc_004226E6: push 00401926h ; __vbaExceptHandler
  loc_004226EB: mov eax, fs:[00000000h]
  loc_004226F1: push eax
  loc_004226F2: mov fs:[00000000h], esp
  loc_004226F9: sub esp, 0000008Ch
  loc_004226FF: push ebx
  loc_00422700: push esi
  loc_00422701: push edi
  loc_00422702: mov var_C, esp
  loc_00422705: mov var_8, 00401470h
  loc_0042270C: mov esi, Me
  loc_0042270F: mov eax, esi
  loc_00422711: and eax, 00000001h
  loc_00422714: mov var_4, eax
  loc_00422717: and esi, FFFFFFFEh
  loc_0042271A: push esi
  loc_0042271B: mov Me, esi
  loc_0042271E: mov ecx, [esi]
  loc_00422720: call [ecx+00000004h]
  loc_00422723: mov edi, KeyAscii
  loc_00422726: lea edx, var_24
  loc_00422729: xor ebx, ebx
  loc_0042272B: push edi
  loc_0042272C: push edx
  loc_0042272D: mov var_24, ebx
  loc_00422730: mov var_34, ebx
  loc_00422733: mov var_44, ebx
  loc_00422736: mov var_54, ebx
  loc_00422739: mov var_64, ebx
  loc_0042273C: mov var_74, ebx
  loc_0042273F: mov var_84, ebx
  loc_00422745: call 0041A480h
  loc_0042274A: lea eax, var_24
  loc_0042274D: push eax
  loc_0042274E: call [004010C4h] ; __vbaI2Var
  loc_00422754: lea ecx, var_24
  loc_00422757: mov [edi], ax
  loc_0042275A: call [00401010h] ; __vbaFreeVar
  loc_00422760: push 0040DD3Ch ; "."
  loc_00422765: call [00401024h] ; rtcAnsiValueBstr
  loc_0042276B: mov edx, [esi]
  loc_0042276D: xor ecx, ecx
  loc_0042276F: cmp [edi], ax
  loc_00422772: push esi
  loc_00422773: mov var_84, 0000000Bh
  loc_0042277D: setz cl
  loc_00422780: neg ecx
  loc_00422782: mov var_7C, cx
  loc_00422786: call [edx+00000304h]
  loc_0042278C: mov var_1C, eax
  loc_0042278F: lea eax, var_84
  loc_00422795: push eax
  loc_00422796: lea ecx, var_24
  loc_00422799: push 00000001h
  loc_0042279B: lea edx, var_64
  loc_0042279E: push ecx
  loc_0042279F: push edx
  loc_004227A0: lea eax, var_34
  loc_004227A3: push ebx
  loc_004227A4: push eax
  loc_004227A5: mov var_24, 00000009h
  loc_004227AC: mov var_5C, 0040DD3Ch ; "."
  loc_004227B3: mov var_64, 00000008h
  loc_004227BA: mov var_6C, ebx
  loc_004227BD: mov var_74, 00008002h
  loc_004227C4: call [004010B4h] ; __vbaInStrVar
  loc_004227CA: lea ecx, var_74
  loc_004227CD: push eax
  loc_004227CE: lea edx, var_44
  loc_004227D1: push ecx
  loc_004227D2: push edx
  loc_004227D3: call [00401060h] ; __vbaVarCmpGt
  loc_004227D9: push eax
  loc_004227DA: lea eax, var_54
  loc_004227DD: push eax
  loc_004227DE: call [00401094h] ; __vbaVarAnd
  loc_004227E4: push eax
  loc_004227E5: call [00401058h] ; __vbaBoolVarNull
  loc_004227EB: lea ecx, var_84
  loc_004227F1: mov esi, eax
  loc_004227F3: lea edx, var_34
  loc_004227F6: push ecx
  loc_004227F7: lea eax, var_24
  loc_004227FA: push edx
  loc_004227FB: push eax
  loc_004227FC: push 00000003h
  loc_004227FE: call [00401018h] ; __vbaFreeVarList
  loc_00422804: add esp, 00000010h
  loc_00422807: cmp si, bx
  loc_0042280A: jz 0042280Fh
  loc_0042280C: mov [edi], bx
  loc_0042280F: push 0040DD44h ; "-"
  loc_00422814: call [00401024h] ; rtcAnsiValueBstr
  loc_0042281A: cmp [edi], ax
  loc_0042281D: jnz 0042288Ah
  loc_0042281F: mov ecx, 80020004h
  loc_00422824: mov eax, 0000000Ah
  loc_00422829: mov var_4C, ecx
  loc_0042282C: mov var_3C, ecx
  loc_0042282F: mov var_2C, ecx
  loc_00422832: lea edx, var_64
  loc_00422835: lea ecx, var_24
  loc_00422838: mov var_54, eax
  loc_0042283B: mov var_44, eax
  loc_0042283E: mov var_34, eax
  loc_00422841: mov var_5C, 0040F98Ch ; "Standard errors are always positive."
  loc_00422848: mov var_64, 00000008h
  loc_0042284F: call [004010F4h] ; __vbaVarDup
  loc_00422855: lea ecx, var_54
  loc_00422858: lea edx, var_44
  loc_0042285B: push ecx
  loc_0042285C: lea eax, var_34
  loc_0042285F: push edx
  loc_00422860: push eax
  loc_00422861: lea ecx, var_24
  loc_00422864: push ebx
  loc_00422865: push ecx
  loc_00422866: call [00401038h] ; rtcMsgBox
  loc_0042286C: lea edx, var_54
  loc_0042286F: lea eax, var_44
  loc_00422872: push edx
  loc_00422873: lea ecx, var_34
  loc_00422876: push eax
  loc_00422877: lea edx, var_24
  loc_0042287A: push ecx
  loc_0042287B: push edx
  loc_0042287C: push 00000004h
  loc_0042287E: call [00401018h] ; __vbaFreeVarList
  loc_00422884: add esp, 00000014h
  loc_00422887: mov [edi], bx
  loc_0042288A: mov var_4, ebx
  loc_0042288D: push 004228B1h
  loc_00422892: jmp 004228B0h
  loc_00422894: lea eax, var_54
  loc_00422897: lea ecx, var_44
  loc_0042289A: push eax
  loc_0042289B: lea edx, var_34
  loc_0042289E: push ecx
  loc_0042289F: lea eax, var_24
  loc_004228A2: push edx
  loc_004228A3: push eax
  loc_004228A4: push 00000004h
  loc_004228A6: call [00401018h] ; __vbaFreeVarList
  loc_004228AC: add esp, 00000014h
  loc_004228AF: ret
  loc_004228B0: ret
  loc_004228B1: mov eax, Me
  loc_004228B4: push eax
  loc_004228B5: mov ecx, [eax]
  loc_004228B7: call [ecx+00000008h]
  loc_004228BA: mov eax, var_4
  loc_004228BD: mov ecx, var_14
  loc_004228C0: pop edi
  loc_004228C1: pop esi
  loc_004228C2: mov fs:[00000000h], ecx
  loc_004228C9: pop ebx
  loc_004228CA: mov esp, ebp
  loc_004228CC: pop ebp
  loc_004228CD: retn 0008h
End Sub

Private Sub cmdChange_Click() '421A20
  loc_00421A20: push ebp
  loc_00421A21: mov ebp, esp
  loc_00421A23: sub esp, 0000000Ch
  loc_00421A26: push 00401926h ; __vbaExceptHandler
  loc_00421A2B: mov eax, fs:[00000000h]
  loc_00421A31: push eax
  loc_00421A32: mov fs:[00000000h], esp
  loc_00421A39: sub esp, 000000DCh
  loc_00421A3F: push ebx
  loc_00421A40: push esi
  loc_00421A41: push edi
  loc_00421A42: mov var_C, esp
  loc_00421A45: mov var_8, 00401440h
  loc_00421A4C: mov ebx, Me
  loc_00421A4F: mov eax, ebx
  loc_00421A51: and eax, 00000001h
  loc_00421A54: mov var_4, eax
  loc_00421A57: and ebx, FFFFFFFEh
  loc_00421A5A: push ebx
  loc_00421A5B: mov Me, ebx
  loc_00421A5E: mov ecx, [ebx]
  loc_00421A60: call [ecx+00000004h]
  loc_00421A63: mov edx, [ebx]
  loc_00421A65: xor esi, esi
  loc_00421A67: push ebx
  loc_00421A68: mov var_20, esi
  loc_00421A6B: mov var_24, esi
  loc_00421A6E: mov var_28, esi
  loc_00421A71: mov var_2C, esi
  loc_00421A74: mov var_30, esi
  loc_00421A77: mov var_34, esi
  loc_00421A7A: mov var_38, esi
  loc_00421A7D: mov var_3C, esi
  loc_00421A80: mov var_40, esi
  loc_00421A83: mov var_44, esi
  loc_00421A86: mov var_48, esi
  loc_00421A89: mov var_4C, esi
  loc_00421A8C: mov var_50, esi
  loc_00421A8F: mov var_60, esi
  loc_00421A92: mov var_70, esi
  loc_00421A95: mov var_80, esi
  loc_00421A98: mov var_90, esi
  loc_00421A9E: mov var_A0, esi
  loc_00421AA4: mov var_B0, esi
  loc_00421AAA: call [edx+0000030Ch]
  loc_00421AB0: mov var_58, eax
  loc_00421AB3: lea eax, var_60
  loc_00421AB6: lea ecx, var_70
  loc_00421AB9: push eax
  loc_00421ABA: push ecx
  loc_00421ABB: mov var_60, 00000009h
  loc_00421AC2: call [00401050h] ; rtcTrimVar
  loc_00421AC8: lea edx, var_70
  loc_00421ACB: lea eax, var_A0
  loc_00421AD1: push edx
  loc_00421AD2: push eax
  loc_00421AD3: mov var_98, 0040F040h
  loc_00421ADD: mov var_A0, 00008008h
  loc_00421AE7: call [00401070h] ; __vbaVarTstEq
  loc_00421AED: mov edi, [00401018h] ; __vbaFreeVarList
  loc_00421AF3: lea ecx, var_70
  loc_00421AF6: lea edx, var_60
  loc_00421AF9: push ecx
  loc_00421AFA: push edx
  loc_00421AFB: push 00000002h
  loc_00421AFD: mov var_D4, ax
  loc_00421B04: call edi
  loc_00421B06: add esp, 0000000Ch
  loc_00421B09: cmp var_D4, si
  loc_00421B10: jz 00421BEFh
  loc_00421B16: mov eax, [ebx]
  loc_00421B18: push ebx
  loc_00421B19: call [eax+0000030Ch]
  loc_00421B1F: lea ecx, var_50
  loc_00421B22: push eax
  loc_00421B23: push ecx
  loc_00421B24: call [0040103Ch] ; __vbaObjSet
  loc_00421B2A: mov ebx, eax
  loc_00421B2C: push ebx
  loc_00421B2D: mov edx, [ebx]
  loc_00421B2F: call [edx+00000204h]
  loc_00421B35: cmp eax, esi
  loc_00421B37: fnclex
  loc_00421B39: jge 00421B4Dh
  loc_00421B3B: push 00000204h
  loc_00421B40: push 0040F02Ch
  loc_00421B45: push ebx
  loc_00421B46: push eax
  loc_00421B47: call [00401030h] ; __vbaHresultCheckObj
  loc_00421B4D: lea ecx, var_50
  loc_00421B50: call [00401114h] ; __vbaFreeObj
  loc_00421B56: mov ebx, [004010F4h] ; __vbaVarDup
  loc_00421B5C: mov ecx, 80020004h
  loc_00421B61: mov var_88, ecx
  loc_00421B67: mov eax, 0000000Ah
  loc_00421B6C: mov var_78, ecx
  loc_00421B6F: lea edx, var_B0
  loc_00421B75: lea ecx, var_70
  loc_00421B78: mov var_90, eax
  loc_00421B7E: mov var_80, eax
  loc_00421B81: mov var_A8, 0040F9DCh ; "Check Slope"
  loc_00421B8B: mov var_B0, 00000008h
  loc_00421B95: call ebx
  loc_00421B97: lea edx, var_A0
  loc_00421B9D: lea ecx, var_60
  loc_00421BA0: mov var_98, 0040F3A0h ; "Please enter a value for the estimate of the slope."
  loc_00421BAA: mov var_A0, 00000008h
  loc_00421BB4: call ebx
  loc_00421BB6: lea eax, var_90
  loc_00421BBC: lea ecx, var_80
  loc_00421BBF: push eax
  loc_00421BC0: lea edx, var_70
  loc_00421BC3: push ecx
  loc_00421BC4: push edx
  loc_00421BC5: lea eax, var_60
  loc_00421BC8: push esi
  loc_00421BC9: push eax
  loc_00421BCA: call [00401038h] ; rtcMsgBox
  loc_00421BD0: lea ecx, var_90
  loc_00421BD6: lea edx, var_80
  loc_00421BD9: push ecx
  loc_00421BDA: lea eax, var_70
  loc_00421BDD: push edx
  loc_00421BDE: lea ecx, var_60
  loc_00421BE1: push eax
  loc_00421BE2: push ecx
  loc_00421BE3: push 00000004h
  loc_00421BE5: call edi
  loc_00421BE7: add esp, 00000014h
  loc_00421BEA: jmp 00422337h
  loc_00421BEF: mov edx, [ebx]
  loc_00421BF1: push ebx
  loc_00421BF2: call [edx+00000304h]
  loc_00421BF8: mov var_58, eax
  loc_00421BFB: lea eax, var_60
  loc_00421BFE: lea ecx, var_70
  loc_00421C01: push eax
  loc_00421C02: push ecx
  loc_00421C03: mov var_60, 00000009h
  loc_00421C0A: call [00401050h] ; rtcTrimVar
  loc_00421C10: lea edx, var_70
  loc_00421C13: lea eax, var_A0
  loc_00421C19: push edx
  loc_00421C1A: push eax
  loc_00421C1B: mov var_98, 0040F040h
  loc_00421C25: mov var_A0, 00008008h
  loc_00421C2F: call [00401070h] ; __vbaVarTstEq
  loc_00421C35: lea ecx, var_70
  loc_00421C38: lea edx, var_60
  loc_00421C3B: push ecx
  loc_00421C3C: push edx
  loc_00421C3D: push 00000002h
  loc_00421C3F: mov var_D4, ax
  loc_00421C46: call edi
  loc_00421C48: add esp, 0000000Ch
  loc_00421C4B: cmp var_D4, si
  loc_00421C52: jz 00421D33h
  loc_00421C58: mov ecx, 80020004h
  loc_00421C5D: mov eax, 0000000Ah
  loc_00421C62: mov var_88, ecx
  loc_00421C68: mov var_78, ecx
  loc_00421C6B: lea edx, var_B0
  loc_00421C71: lea ecx, var_70
  loc_00421C74: mov var_90, eax
  loc_00421C7A: mov var_80, eax
  loc_00421C7D: mov var_A8, 0040FE20h ; "Check Intercept"
  loc_00421C87: mov var_B0, 00000008h
  loc_00421C91: call [004010F4h] ; __vbaVarDup
  loc_00421C97: lea edx, var_A0
  loc_00421C9D: lea ecx, var_60
  loc_00421CA0: mov var_98, 0040F40Ch ; "Please enter a value for the estimate of the standard error."
  loc_00421CAA: mov var_A0, 00000008h
  loc_00421CB4: call [004010F4h] ; __vbaVarDup
  loc_00421CBA: lea eax, var_90
  loc_00421CC0: lea ecx, var_80
  loc_00421CC3: push eax
  loc_00421CC4: lea edx, var_70
  loc_00421CC7: push ecx
  loc_00421CC8: push edx
  loc_00421CC9: lea eax, var_60
  loc_00421CCC: push esi
  loc_00421CCD: push eax
  loc_00421CCE: call [00401038h] ; rtcMsgBox
  loc_00421CD4: lea ecx, var_90
  loc_00421CDA: lea edx, var_80
  loc_00421CDD: push ecx
  loc_00421CDE: lea eax, var_70
  loc_00421CE1: push edx
  loc_00421CE2: lea ecx, var_60
  loc_00421CE5: push eax
  loc_00421CE6: push ecx
  loc_00421CE7: push 00000004h
  loc_00421CE9: call edi
  loc_00421CEB: mov edx, [ebx]
  loc_00421CED: add esp, 00000014h
  loc_00421CF0: push ebx
  loc_00421CF1: call [edx+00000304h]
  loc_00421CF7: push eax
  loc_00421CF8: lea eax, var_50
  loc_00421CFB: push eax
  loc_00421CFC: call [0040103Ch] ; __vbaObjSet
  loc_00421D02: mov edi, eax
  loc_00421D04: push edi
  loc_00421D05: mov ecx, [edi]
  loc_00421D07: call [ecx+00000204h]
  loc_00421D0D: cmp eax, esi
  loc_00421D0F: fnclex
  loc_00421D11: jge 00421D25h
  loc_00421D13: push 00000204h
  loc_00421D18: push 0040F02Ch
  loc_00421D1D: push edi
  loc_00421D1E: push eax
  loc_00421D1F: call [00401030h] ; __vbaHresultCheckObj
  loc_00421D25: lea ecx, var_50
  loc_00421D28: call [00401114h] ; __vbaFreeObj
  loc_00421D2E: jmp 00422337h
  loc_00421D33: mov edx, [ebx]
  loc_00421D35: push ebx
  loc_00421D36: call [edx+0000030Ch]
  loc_00421D3C: mov var_58, eax
  loc_00421D3F: lea eax, var_60
  loc_00421D42: lea ecx, var_70
  loc_00421D45: push eax
  loc_00421D46: push ecx
  loc_00421D47: mov var_60, 00000009h
  loc_00421D4E: call [00401050h] ; rtcTrimVar
  loc_00421D54: lea edx, var_70
  loc_00421D57: lea eax, var_20
  loc_00421D5A: push edx
  loc_00421D5B: push eax
  loc_00421D5C: call [004010B8h] ; __vbaStrVarVal
  loc_00421D62: push eax
  loc_00421D63: call [00401118h] ; rtcR8ValFromBstr
  loc_00421D69: call [00401054h] ; __vbaFpR8
  loc_00421D6F: fcomp real8 ptr [004011F0h]
  loc_00421D75: mov var_EC, 00000001h
  loc_00421D7F: fnstsw ax
  loc_00421D81: test ah, 41h
  loc_00421D84: jz 00421D8Ch
  loc_00421D86: mov var_EC, esi
  loc_00421D8C: fld real4 ptr [0043006Ch]
  loc_00421D92: fcomp real4 ptr [004011E8h]
  loc_00421D98: fnstsw ax
  loc_00421D9A: test ah, 01h
  loc_00421D9D: jz 00421DA6h
  loc_00421D9F: mov edi, 00000001h
  loc_00421DA4: jmp 00421DA8h
  loc_00421DA6: xor edi, edi
  loc_00421DA8: mov ecx, [ebx]
  loc_00421DAA: push ebx
  loc_00421DAB: call [ecx+0000030Ch]
  loc_00421DB1: mov var_78, eax
  loc_00421DB4: lea edx, var_80
  loc_00421DB7: lea eax, var_90
  loc_00421DBD: push edx
  loc_00421DBE: push eax
  loc_00421DBF: mov var_80, 00000009h
  loc_00421DC6: call [00401050h] ; rtcTrimVar
  loc_00421DCC: lea ecx, var_90
  loc_00421DD2: lea edx, var_24
  loc_00421DD5: push ecx
  loc_00421DD6: push edx
  loc_00421DD7: call [004010B8h] ; __vbaStrVarVal
  loc_00421DDD: push eax
  loc_00421DDE: call [00401118h] ; rtcR8ValFromBstr
  loc_00421DE4: call [00401054h] ; __vbaFpR8
  loc_00421DEA: fcomp real8 ptr [004011F0h]
  loc_00421DF0: mov var_F0, 00000001h
  loc_00421DFA: fnstsw ax
  loc_00421DFC: test ah, 01h
  loc_00421DFF: jnz 00421E07h
  loc_00421E01: mov var_F0, esi
  loc_00421E07: fld real4 ptr [0043006Ch]
  loc_00421E0D: fcomp real4 ptr [004011E8h]
  loc_00421E13: fnstsw ax
  loc_00421E15: test ah, 41h
  loc_00421E18: jnz 00421E1Fh
  loc_00421E1A: mov esi, 00000001h
  loc_00421E1F: lea eax, var_24
  loc_00421E22: lea ecx, var_20
  loc_00421E25: push eax
  loc_00421E26: push ecx
  loc_00421E27: push 00000002h
  loc_00421E29: call [004010E4h] ; __vbaFreeStrList
  loc_00421E2F: lea edx, var_90
  loc_00421E35: lea eax, var_80
  loc_00421E38: push edx
  loc_00421E39: lea ecx, var_70
  loc_00421E3C: push eax
  loc_00421E3D: lea edx, var_60
  loc_00421E40: push ecx
  loc_00421E41: push edx
  loc_00421E42: push 00000004h
  loc_00421E44: call [00401018h] ; __vbaFreeVarList
  loc_00421E4A: mov eax, var_F0
  loc_00421E50: add esp, 00000020h
  loc_00421E53: neg esi
  loc_00421E55: neg eax
  loc_00421E57: and esi, eax
  loc_00421E59: mov eax, var_EC
  loc_00421E5F: neg edi
  loc_00421E61: neg eax
  loc_00421E63: and edi, eax
  loc_00421E65: or esi, edi
  loc_00421E67: test si, si
  loc_00421E6A: jz 00421FDBh
  loc_00421E70: mov ecx, 80020004h
  loc_00421E75: mov eax, 0000000Ah
  loc_00421E7A: mov var_88, ecx
  loc_00421E80: mov var_78, ecx
  loc_00421E83: lea edx, var_A0
  loc_00421E89: lea ecx, var_70
  loc_00421E8C: mov var_90, eax
  loc_00421E92: mov var_80, eax
  loc_00421E95: mov var_98, 00410428h ; "Check Slope Estimate"
  loc_00421E9F: mov var_A0, 00000008h
  loc_00421EA9: call [004010F4h] ; __vbaVarDup
  loc_00421EAF: mov esi, [0040102Ch] ; __vbaStrCat
  loc_00421EB5: push 0040F674h ; "Warning, the correlation coefficient MUST have the same sign as the estimated slope"
  loc_00421EBA: push 0040F720h ; vbCrLf
  loc_00421EBF: call __vbaStrCat
  loc_00421EC1: mov edi, [004010FCh] ; __vbaStrMove
  loc_00421EC7: mov edx, eax
  loc_00421EC9: lea ecx, var_20
  loc_00421ECC: call edi
  loc_00421ECE: push eax
  loc_00421ECF: push 0040FE44h ; "Currently they do not. The correlation coefficient = "
  loc_00421ED4: call __vbaStrCat
  loc_00421ED6: mov edx, eax
  loc_00421ED8: lea ecx, var_24
  loc_00421EDB: call edi
  loc_00421EDD: push eax
  loc_00421EDE: mov eax, [0043006Ch]
  loc_00421EE3: push eax
  loc_00421EE4: call [0040107Ch] ; __vbaStrR4
  loc_00421EEA: mov edx, eax
  loc_00421EEC: lea ecx, var_28
  loc_00421EEF: call edi
  loc_00421EF1: push eax
  loc_00421EF2: call __vbaStrCat
  loc_00421EF4: mov edx, eax
  loc_00421EF6: lea ecx, var_2C
  loc_00421EF9: call edi
  loc_00421EFB: push eax
  loc_00421EFC: push 0040F720h ; vbCrLf
  loc_00421F01: call __vbaStrCat
  loc_00421F03: mov edx, eax
  loc_00421F05: lea ecx, var_30
  loc_00421F08: call edi
  loc_00421F0A: push eax
  loc_00421F0B: push 0040F720h ; vbCrLf
  loc_00421F10: call __vbaStrCat
  loc_00421F12: mov edx, eax
  loc_00421F14: lea ecx, var_34
  loc_00421F17: call edi
  loc_00421F19: push eax
  loc_00421F1A: push 0040F78Ch ; "Do you wish to continue with this value?"
  loc_00421F1F: call __vbaStrCat
  loc_00421F21: lea ecx, var_90
  loc_00421F27: mov var_58, eax
  loc_00421F2A: lea edx, var_80
  loc_00421F2D: push ecx
  loc_00421F2E: lea eax, var_70
  loc_00421F31: push edx
  loc_00421F32: push eax
  loc_00421F33: lea ecx, var_60
  loc_00421F36: push 00000004h
  loc_00421F38: push ecx
  loc_00421F39: mov var_60, 00000008h
  loc_00421F40: call [00401038h] ; rtcMsgBox
  loc_00421F46: mov ecx, eax
  loc_00421F48: call [00401078h] ; __vbaI2I4
  loc_00421F4E: mov var_18, eax
  loc_00421F51: lea edx, var_34
  loc_00421F54: lea eax, var_30
  loc_00421F57: push edx
  loc_00421F58: lea ecx, var_2C
  loc_00421F5B: push eax
  loc_00421F5C: lea edx, var_28
  loc_00421F5F: push ecx
  loc_00421F60: lea eax, var_24
  loc_00421F63: push edx
  loc_00421F64: lea ecx, var_20
  loc_00421F67: push eax
  loc_00421F68: push ecx
  loc_00421F69: push 00000006h
  loc_00421F6B: call [004010E4h] ; __vbaFreeStrList
  loc_00421F71: lea edx, var_90
  loc_00421F77: lea eax, var_80
  loc_00421F7A: push edx
  loc_00421F7B: lea ecx, var_70
  loc_00421F7E: push eax
  loc_00421F7F: lea edx, var_60
  loc_00421F82: push ecx
  loc_00421F83: push edx
  loc_00421F84: push 00000004h
  loc_00421F86: call [00401018h] ; __vbaFreeVarList
  loc_00421F8C: add esp, 00000030h
  loc_00421F8F: cmp var_18, 0007h
  loc_00421F94: jnz 00421FE7h
  loc_00421F96: mov eax, [ebx]
  loc_00421F98: push ebx
  loc_00421F99: call [eax+0000030Ch]
  loc_00421F9F: lea ecx, var_50
  loc_00421FA2: push eax
  loc_00421FA3: push ecx
  loc_00421FA4: call [0040103Ch] ; __vbaObjSet
  loc_00421FAA: mov esi, eax
  loc_00421FAC: push esi
  loc_00421FAD: mov edx, [esi]
  loc_00421FAF: call [edx+00000204h]
  loc_00421FB5: test eax, eax
  loc_00421FB7: fnclex
  loc_00421FB9: jge 00421FCDh
  loc_00421FBB: push 00000204h
  loc_00421FC0: push 0040F02Ch
  loc_00421FC5: push esi
  loc_00421FC6: push eax
  loc_00421FC7: call [00401030h] ; __vbaHresultCheckObj
  loc_00421FCD: lea ecx, var_50
  loc_00421FD0: call [00401114h] ; __vbaFreeObj
  loc_00421FD6: jmp 00422335h
  loc_00421FDB: mov esi, [0040102Ch] ; __vbaStrCat
  loc_00421FE1: mov edi, [004010FCh] ; __vbaStrMove
  loc_00421FE7: mov eax, [ebx]
  loc_00421FE9: push ebx
  loc_00421FEA: call [eax+0000030Ch]
  loc_00421FF0: lea ecx, var_50
  loc_00421FF3: push eax
  loc_00421FF4: push ecx
  loc_00421FF5: call [0040103Ch] ; __vbaObjSet
  loc_00421FFB: mov edx, [eax]
  loc_00421FFD: lea ecx, var_20
  loc_00422000: push ecx
  loc_00422001: push eax
  loc_00422002: mov var_D4, eax
  loc_00422008: call [edx+000000A0h]
  loc_0042200E: test eax, eax
  loc_00422010: fnclex
  loc_00422012: jge 0042202Ch
  loc_00422014: mov edx, var_D4
  loc_0042201A: push 000000A0h
  loc_0042201F: push 0040F02Ch
  loc_00422024: push edx
  loc_00422025: push eax
  loc_00422026: call [00401030h] ; __vbaHresultCheckObj
  loc_0042202C: mov eax, var_20
  loc_0042202F: lea ecx, var_70
  loc_00422032: mov var_58, eax
  loc_00422035: lea eax, var_60
  loc_00422038: push eax
  loc_00422039: push ecx
  loc_0042203A: mov var_20, 00000000h
  loc_00422041: mov var_60, 00000008h
  loc_00422048: call [00401050h] ; rtcTrimVar
  loc_0042204E: lea edx, var_70
  loc_00422051: push edx
  loc_00422052: call [004010A0h] ; __vbaR4ErrVar
  loc_00422058: fstp real4 ptr [00430060h]
  loc_0042205E: lea ecx, var_50
  loc_00422061: call [00401114h] ; __vbaFreeObj
  loc_00422067: lea eax, var_70
  loc_0042206A: lea ecx, var_70
  loc_0042206D: push eax
  loc_0042206E: lea edx, var_60
  loc_00422071: push ecx
  loc_00422072: push edx
  loc_00422073: push 00000003h
  loc_00422075: call [00401018h] ; __vbaFreeVarList
  loc_0042207B: mov eax, [ebx]
  loc_0042207D: add esp, 00000010h
  loc_00422080: push ebx
  loc_00422081: call [eax+00000304h]
  loc_00422087: lea ecx, var_50
  loc_0042208A: push eax
  loc_0042208B: push ecx
  loc_0042208C: call [0040103Ch] ; __vbaObjSet
  loc_00422092: mov ebx, eax
  loc_00422094: lea eax, var_20
  loc_00422097: push eax
  loc_00422098: push ebx
  loc_00422099: mov edx, [ebx]
  loc_0042209B: call [edx+000000A0h]
  loc_004220A1: test eax, eax
  loc_004220A3: fnclex
  loc_004220A5: jge 004220B9h
  loc_004220A7: push 000000A0h
  loc_004220AC: push 0040F02Ch
  loc_004220B1: push ebx
  loc_004220B2: push eax
  loc_004220B3: call [00401030h] ; __vbaHresultCheckObj
  loc_004220B9: mov eax, var_20
  loc_004220BC: lea ecx, var_60
  loc_004220BF: lea edx, var_70
  loc_004220C2: push ecx
  loc_004220C3: push edx
  loc_004220C4: mov var_20, 00000000h
  loc_004220CB: mov var_58, eax
  loc_004220CE: mov var_60, 00000008h
  loc_004220D5: call [00401050h] ; rtcTrimVar
  loc_004220DB: lea eax, var_70
  loc_004220DE: push eax
  loc_004220DF: call [004010A0h] ; __vbaR4ErrVar
  loc_004220E5: fstp real4 ptr [00430064h]
  loc_004220EB: lea ecx, var_50
  loc_004220EE: call [00401114h] ; __vbaFreeObj
  loc_004220F4: mov ebx, [00401018h] ; __vbaFreeVarList
  loc_004220FA: lea ecx, var_70
  loc_004220FD: lea edx, var_70
  loc_00422100: push ecx
  loc_00422101: lea eax, var_60
  loc_00422104: push edx
  loc_00422105: push eax
  loc_00422106: push 00000003h
  loc_00422108: call ebx
  loc_0042210A: mov ecx, [0043002Ch]
  loc_00422110: add esp, 00000010h
  loc_00422113: push ecx
  loc_00422114: call [004010D0h] ; __vbaR8Str
  loc_0042211A: fmul st0, real4 ptr [00430060h]
  loc_00422120: fadd st0, real4 ptr [00430074h]
  loc_00422126: mov ecx, 80020004h
  loc_0042212B: lea edx, var_A0
  loc_00422131: fstp real4 ptr [00430038h]
  loc_00422137: fnstsw ax
  loc_00422139: test al, 0Dh
  loc_0042213B: jnz 004223C5h
  loc_00422141: mov var_88, ecx
  loc_00422147: mov eax, 0000000Ah
  loc_0042214C: mov var_78, ecx
  loc_0042214F: lea ecx, var_70
  loc_00422152: mov var_90, eax
  loc_00422158: mov var_80, eax
  loc_0042215B: mov var_98, 0040FF38h ; "Change in Yhat"
  loc_00422165: mov var_A0, 00000008h
  loc_0042216F: call [004010F4h] ; __vbaVarDup
  loc_00422175: mov edx, [00430010h]
  loc_0042217B: push 0040FEB4h ; "Note the estimated mean "
  loc_00422180: push edx
  loc_00422181: call __vbaStrCat
  loc_00422183: mov edx, eax
  loc_00422185: lea ecx, var_20
  loc_00422188: call edi
  loc_0042218A: push eax
  loc_0042218B: push 0040FEECh ; " is now "
  loc_00422190: call __vbaStrCat
  loc_00422192: mov edx, eax
  loc_00422194: lea ecx, var_24
  loc_00422197: call edi
  loc_00422199: push eax
  loc_0042219A: mov eax, [00430038h]
  loc_0042219F: push eax
  loc_004221A0: call [0040107Ch] ; __vbaStrR4
  loc_004221A6: mov edx, eax
  loc_004221A8: lea ecx, var_28
  loc_004221AB: call edi
  loc_004221AD: push eax
  loc_004221AE: call __vbaStrCat
  loc_004221B0: mov edx, eax
  loc_004221B2: lea ecx, var_2C
  loc_004221B5: call edi
  loc_004221B7: push eax
  loc_004221B8: push 0040F028h
  loc_004221BD: call __vbaStrCat
  loc_004221BF: mov edx, eax
  loc_004221C1: lea ecx, var_30
  loc_004221C4: call edi
  loc_004221C6: mov ecx, [00430014h]
  loc_004221CC: push eax
  loc_004221CD: push ecx
  loc_004221CE: call __vbaStrCat
  loc_004221D0: mov edx, eax
  loc_004221D2: lea ecx, var_34
  loc_004221D5: call edi
  loc_004221D7: push eax
  loc_004221D8: push 0040FF04h ; " when "
  loc_004221DD: call __vbaStrCat
  loc_004221DF: mov edx, eax
  loc_004221E1: lea ecx, var_38
  loc_004221E4: call edi
  loc_004221E6: mov edx, [00430018h]
  loc_004221EC: push eax
  loc_004221ED: push edx
  loc_004221EE: call __vbaStrCat
  loc_004221F0: mov edx, eax
  loc_004221F2: lea ecx, var_3C
  loc_004221F5: call edi
  loc_004221F7: push eax
  loc_004221F8: push 0040FF18h ; " is equal to "
  loc_004221FD: call __vbaStrCat
  loc_004221FF: mov edx, eax
  loc_00422201: lea ecx, var_40
  loc_00422204: call edi
  loc_00422206: push eax
  loc_00422207: mov eax, [0043002Ch]
  loc_0042220C: push eax
  loc_0042220D: call __vbaStrCat
  loc_0042220F: mov edx, eax
  loc_00422211: lea ecx, var_44
  loc_00422214: call edi
  loc_00422216: push eax
  loc_00422217: push 0040F028h
  loc_0042221C: call __vbaStrCat
  loc_0042221E: mov edx, eax
  loc_00422220: lea ecx, var_48
  loc_00422223: call edi
  loc_00422225: mov ecx, [0043001Ch]
  loc_0042222B: push eax
  loc_0042222C: push ecx
  loc_0042222D: call __vbaStrCat
  loc_0042222F: mov edx, eax
  loc_00422231: lea ecx, var_4C
  loc_00422234: call edi
  loc_00422236: push eax
  loc_00422237: push 0040DD3Ch ; "."
  loc_0042223C: call __vbaStrCat
  loc_0042223E: mov var_58, eax
  loc_00422241: lea edx, var_90
  loc_00422247: lea eax, var_80
  loc_0042224A: push edx
  loc_0042224B: lea ecx, var_70
  loc_0042224E: push eax
  loc_0042224F: push ecx
  loc_00422250: lea edx, var_60
  loc_00422253: push 00000000h
  loc_00422255: push edx
  loc_00422256: mov var_60, 00000008h
  loc_0042225D: call [00401038h] ; rtcMsgBox
  loc_00422263: lea eax, var_4C
  loc_00422266: lea ecx, var_48
  loc_00422269: push eax
  loc_0042226A: lea edx, var_44
  loc_0042226D: push ecx
  loc_0042226E: lea eax, var_40
  loc_00422271: push edx
  loc_00422272: lea ecx, var_3C
  loc_00422275: push eax
  loc_00422276: lea edx, var_38
  loc_00422279: push ecx
  loc_0042227A: lea eax, var_34
  loc_0042227D: push edx
  loc_0042227E: lea ecx, var_30
  loc_00422281: push eax
  loc_00422282: lea edx, var_2C
  loc_00422285: push ecx
  loc_00422286: lea eax, var_28
  loc_00422289: push edx
  loc_0042228A: lea ecx, var_24
  loc_0042228D: push eax
  loc_0042228E: lea edx, var_20
  loc_00422291: push ecx
  loc_00422292: push edx
  loc_00422293: push 0000000Ch
  loc_00422295: call [004010E4h] ; __vbaFreeStrList
  loc_0042229B: lea eax, var_90
  loc_004222A1: lea ecx, var_80
  loc_004222A4: push eax
  loc_004222A5: lea edx, var_70
  loc_004222A8: push ecx
  loc_004222A9: lea eax, var_60
  loc_004222AC: push edx
  loc_004222AD: push eax
  loc_004222AE: push 00000004h
  loc_004222B0: call ebx
  loc_004222B2: mov eax, [00430150h]
  loc_004222B7: add esp, 00000048h
  loc_004222BA: test eax, eax
  loc_004222BC: jnz 004222CEh
  loc_004222BE: push 00430150h
  loc_004222C3: push 00403F48h
  loc_004222C8: call [004010D4h] ; __vbaNew2
  loc_004222CE: mov esi, [00430150h]
  loc_004222D4: push esi
  loc_004222D5: mov ecx, [esi]
  loc_004222D7: call [ecx+000002B4h]
  loc_004222DD: test eax, eax
  loc_004222DF: fnclex
  loc_004222E1: jge 004222F5h
  loc_004222E3: push 000002B4h
  loc_004222E8: push 0040FD4Ch
  loc_004222ED: push esi
  loc_004222EE: push eax
  loc_004222EF: call [00401030h] ; __vbaHresultCheckObj
  loc_004222F5: mov eax, [0043018Ch]
  loc_004222FA: test eax, eax
  loc_004222FC: jnz 0042230Eh
  loc_004222FE: push 0043018Ch
  loc_00422303: push 0040397Ch
  loc_00422308: call [004010D4h] ; __vbaNew2
  loc_0042230E: mov esi, [0043018Ch]
  loc_00422314: push esi
  loc_00422315: mov edx, [esi]
  loc_00422317: call [edx+000002B4h]
  loc_0042231D: test eax, eax
  loc_0042231F: fnclex
  loc_00422321: jge 00422335h
  loc_00422323: push 000002B4h
  loc_00422328: push 004103E4h
  loc_0042232D: push esi
  loc_0042232E: push eax
  loc_0042232F: call [00401030h] ; __vbaHresultCheckObj
  loc_00422335: xor esi, esi
  loc_00422337: mov var_4, esi
  loc_0042233A: fwait
  loc_0042233B: push 004223A6h
  loc_00422340: jmp 004223A5h
  loc_00422342: lea eax, var_4C
  loc_00422345: lea ecx, var_48
  loc_00422348: push eax
  loc_00422349: lea edx, var_44
  loc_0042234C: push ecx
  loc_0042234D: lea eax, var_40
  loc_00422350: push edx
  loc_00422351: lea ecx, var_3C
  loc_00422354: push eax
  loc_00422355: lea edx, var_38
  loc_00422358: push ecx
  loc_00422359: lea eax, var_34
  loc_0042235C: push edx
  loc_0042235D: lea ecx, var_30
  loc_00422360: push eax
  loc_00422361: lea edx, var_2C
  loc_00422364: push ecx
  loc_00422365: lea eax, var_28
  loc_00422368: push edx
  loc_00422369: lea ecx, var_24
  loc_0042236C: push eax
  loc_0042236D: lea edx, var_20
  loc_00422370: push ecx
  loc_00422371: push edx
  loc_00422372: push 0000000Ch
  loc_00422374: call [004010E4h] ; __vbaFreeStrList
  loc_0042237A: add esp, 00000034h
  loc_0042237D: lea ecx, var_50
  loc_00422380: call [00401114h] ; __vbaFreeObj
  loc_00422386: lea eax, var_90
  loc_0042238C: lea ecx, var_80
  loc_0042238F: push eax
  loc_00422390: lea edx, var_70
  loc_00422393: push ecx
  loc_00422394: lea eax, var_60
  loc_00422397: push edx
  loc_00422398: push eax
  loc_00422399: push 00000004h
  loc_0042239B: call [00401018h] ; __vbaFreeVarList
  loc_004223A1: add esp, 00000014h
  loc_004223A4: ret
  loc_004223A5: ret
  loc_004223A6: mov eax, Me
  loc_004223A9: push eax
  loc_004223AA: mov ecx, [eax]
  loc_004223AC: call [ecx+00000008h]
  loc_004223AF: mov eax, var_4
  loc_004223B2: mov ecx, var_14
  loc_004223B5: pop edi
  loc_004223B6: pop esi
  loc_004223B7: mov fs:[00000000h], ecx
  loc_004223BE: pop ebx
  loc_004223BF: mov esp, ebp
  loc_004223C1: pop ebp
  loc_004223C2: retn 0004h
End Sub

Private Sub cmdRestore_Click() '4218C0
  loc_004218C0: push ebp
  loc_004218C1: mov ebp, esp
  loc_004218C3: sub esp, 0000000Ch
  loc_004218C6: push 00401926h ; __vbaExceptHandler
  loc_004218CB: mov eax, fs:[00000000h]
  loc_004218D1: push eax
  loc_004218D2: mov fs:[00000000h], esp
  loc_004218D9: sub esp, 00000018h
  loc_004218DC: push ebx
  loc_004218DD: push esi
  loc_004218DE: push edi
  loc_004218DF: mov var_C, esp
  loc_004218E2: mov var_8, 00401430h
  loc_004218E9: mov esi, Me
  loc_004218EC: mov eax, esi
  loc_004218EE: and eax, 00000001h
  loc_004218F1: mov var_4, eax
  loc_004218F4: and esi, FFFFFFFEh
  loc_004218F7: push esi
  loc_004218F8: mov Me, esi
  loc_004218FB: mov ecx, [esi]
  loc_004218FD: call [ecx+00000004h]
  loc_00421900: mov [00430060h], 41233333h
  loc_0042190A: mov [00430064h], 3F99999Ah
  loc_00421914: mov edx, [esi]
  loc_00421916: xor eax, eax
  loc_00421918: push esi
  loc_00421919: mov var_18, eax
  loc_0042191C: mov var_1C, eax
  loc_0042191F: call [edx+0000030Ch]
  loc_00421925: push eax
  loc_00421926: lea eax, var_1C
  loc_00421929: push eax
  loc_0042192A: call [0040103Ch] ; __vbaObjSet
  loc_00421930: mov ecx, [00430060h]
  loc_00421936: mov edi, eax
  loc_00421938: push ecx
  loc_00421939: mov ebx, [edi]
  loc_0042193B: call [0040107Ch] ; __vbaStrR4
  loc_00421941: mov edx, eax
  loc_00421943: lea ecx, var_18
  loc_00421946: call [004010FCh] ; __vbaStrMove
  loc_0042194C: push eax
  loc_0042194D: push edi
  loc_0042194E: call [ebx+000000A4h]
  loc_00421954: test eax, eax
  loc_00421956: fnclex
  loc_00421958: jge 0042196Ch
  loc_0042195A: push 000000A4h
  loc_0042195F: push 0040F02Ch
  loc_00421964: push edi
  loc_00421965: push eax
  loc_00421966: call [00401030h] ; __vbaHresultCheckObj
  loc_0042196C: mov edi, [00401110h] ; __vbaFreeStr
  loc_00421972: lea ecx, var_18
  loc_00421975: call edi
  loc_00421977: lea ecx, var_1C
  loc_0042197A: call [00401114h] ; __vbaFreeObj
  loc_00421980: mov edx, [esi]
  loc_00421982: push esi
  loc_00421983: call [edx+00000304h]
  loc_00421989: push eax
  loc_0042198A: lea eax, var_1C
  loc_0042198D: push eax
  loc_0042198E: call [0040103Ch] ; __vbaObjSet
  loc_00421994: mov ecx, [00430064h]
  loc_0042199A: mov esi, eax
  loc_0042199C: push ecx
  loc_0042199D: mov ebx, [esi]
  loc_0042199F: call [0040107Ch] ; __vbaStrR4
  loc_004219A5: mov edx, eax
  loc_004219A7: lea ecx, var_18
  loc_004219AA: call [004010FCh] ; __vbaStrMove
  loc_004219B0: push eax
  loc_004219B1: push esi
  loc_004219B2: call [ebx+000000A4h]
  loc_004219B8: test eax, eax
  loc_004219BA: fnclex
  loc_004219BC: jge 004219D0h
  loc_004219BE: push 000000A4h
  loc_004219C3: push 0040F02Ch
  loc_004219C8: push esi
  loc_004219C9: push eax
  loc_004219CA: call [00401030h] ; __vbaHresultCheckObj
  loc_004219D0: lea ecx, var_18
  loc_004219D3: call edi
  loc_004219D5: lea ecx, var_1C
  loc_004219D8: call [00401114h] ; __vbaFreeObj
  loc_004219DE: mov var_4, 00000000h
  loc_004219E5: fwait
  loc_004219E6: push 00421A01h
  loc_004219EB: jmp 00421A00h
  loc_004219ED: lea ecx, var_18
  loc_004219F0: call [00401110h] ; __vbaFreeStr
  loc_004219F6: lea ecx, var_1C
  loc_004219F9: call [00401114h] ; __vbaFreeObj
  loc_004219FF: ret
  loc_00421A00: ret
  loc_00421A01: mov eax, Me
  loc_00421A04: push eax
  loc_00421A05: mov edx, [eax]
  loc_00421A07: call [edx+00000008h]
  loc_00421A0A: mov eax, var_4
  loc_00421A0D: mov ecx, var_14
  loc_00421A10: pop edi
  loc_00421A11: pop esi
  loc_00421A12: mov fs:[00000000h], ecx
  loc_00421A19: pop ebx
  loc_00421A1A: mov esp, ebp
  loc_00421A1C: pop ebp
  loc_00421A1D: retn 0004h
End Sub

Private Sub Form_Activate() '4223D0
  loc_004223D0: push ebp
  loc_004223D1: mov ebp, esp
  loc_004223D3: sub esp, 0000000Ch
  loc_004223D6: push 00401926h ; __vbaExceptHandler
  loc_004223DB: mov eax, fs:[00000000h]
  loc_004223E1: push eax
  loc_004223E2: mov fs:[00000000h], esp
  loc_004223E9: sub esp, 00000018h
  loc_004223EC: push ebx
  loc_004223ED: push esi
  loc_004223EE: push edi
  loc_004223EF: mov var_C, esp
  loc_004223F2: mov var_8, 00401450h
  loc_004223F9: mov esi, Me
  loc_004223FC: mov eax, esi
  loc_004223FE: and eax, 00000001h
  loc_00422401: mov var_4, eax
  loc_00422404: and esi, FFFFFFFEh
  loc_00422407: push esi
  loc_00422408: mov Me, esi
  loc_0042240B: mov ecx, [esi]
  loc_0042240D: call [ecx+00000004h]
  loc_00422410: mov edx, [esi]
  loc_00422412: xor eax, eax
  loc_00422414: push esi
  loc_00422415: mov var_18, eax
  loc_00422418: mov var_1C, eax
  loc_0042241B: call [edx+0000030Ch]
  loc_00422421: push eax
  loc_00422422: lea eax, var_1C
  loc_00422425: push eax
  loc_00422426: call [0040103Ch] ; __vbaObjSet
  loc_0042242C: mov ecx, [00430060h]
  loc_00422432: mov edi, eax
  loc_00422434: push ecx
  loc_00422435: mov ebx, [edi]
  loc_00422437: call [0040107Ch] ; __vbaStrR4
  loc_0042243D: mov edx, eax
  loc_0042243F: lea ecx, var_18
  loc_00422442: call [004010FCh] ; __vbaStrMove
  loc_00422448: push eax
  loc_00422449: push edi
  loc_0042244A: call [ebx+000000A4h]
  loc_00422450: test eax, eax
  loc_00422452: fnclex
  loc_00422454: jge 00422468h
  loc_00422456: push 000000A4h
  loc_0042245B: push 0040F02Ch
  loc_00422460: push edi
  loc_00422461: push eax
  loc_00422462: call [00401030h] ; __vbaHresultCheckObj
  loc_00422468: lea ecx, var_18
  loc_0042246B: call [00401110h] ; __vbaFreeStr
  loc_00422471: lea ecx, var_1C
  loc_00422474: call [00401114h] ; __vbaFreeObj
  loc_0042247A: mov edx, [esi]
  loc_0042247C: push esi
  loc_0042247D: call [edx+00000304h]
  loc_00422483: push eax
  loc_00422484: lea eax, var_1C
  loc_00422487: push eax
  loc_00422488: call [0040103Ch] ; __vbaObjSet
  loc_0042248E: mov ecx, [00430064h]
  loc_00422494: mov edi, eax
  loc_00422496: push ecx
  loc_00422497: mov ebx, [edi]
  loc_00422499: call [0040107Ch] ; __vbaStrR4
  loc_0042249F: mov edx, eax
  loc_004224A1: lea ecx, var_18
  loc_004224A4: call [004010FCh] ; __vbaStrMove
  loc_004224AA: push eax
  loc_004224AB: push edi
  loc_004224AC: call [ebx+000000A4h]
  loc_004224B2: test eax, eax
  loc_004224B4: fnclex
  loc_004224B6: jge 004224CAh
  loc_004224B8: push 000000A4h
  loc_004224BD: push 0040F02Ch
  loc_004224C2: push edi
  loc_004224C3: push eax
  loc_004224C4: call [00401030h] ; __vbaHresultCheckObj
  loc_004224CA: lea ecx, var_18
  loc_004224CD: call [00401110h] ; __vbaFreeStr
  loc_004224D3: mov edi, [00401114h] ; __vbaFreeObj
  loc_004224D9: lea ecx, var_1C
  loc_004224DC: call edi
  loc_004224DE: mov edx, [esi]
  loc_004224E0: push esi
  loc_004224E1: call [edx+0000030Ch]
  loc_004224E7: push eax
  loc_004224E8: lea eax, var_1C
  loc_004224EB: push eax
  loc_004224EC: call [0040103Ch] ; __vbaObjSet
  loc_004224F2: mov esi, eax
  loc_004224F4: push esi
  loc_004224F5: mov ecx, [esi]
  loc_004224F7: call [ecx+00000204h]
  loc_004224FD: test eax, eax
  loc_004224FF: fnclex
  loc_00422501: jge 00422515h
  loc_00422503: push 00000204h
  loc_00422508: push 0040F02Ch
  loc_0042250D: push esi
  loc_0042250E: push eax
  loc_0042250F: call [00401030h] ; __vbaHresultCheckObj
  loc_00422515: lea ecx, var_1C
  loc_00422518: call edi
  loc_0042251A: mov var_4, 00000000h
  loc_00422521: fwait
  loc_00422522: push 0042253Dh
  loc_00422527: jmp 0042253Ch
  loc_00422529: lea ecx, var_18
  loc_0042252C: call [00401110h] ; __vbaFreeStr
  loc_00422532: lea ecx, var_1C
  loc_00422535: call [00401114h] ; __vbaFreeObj
  loc_0042253B: ret
  loc_0042253C: ret
  loc_0042253D: mov eax, Me
  loc_00422540: push eax
  loc_00422541: mov edx, [eax]
  loc_00422543: call [edx+00000008h]
  loc_00422546: mov eax, var_4
  loc_00422549: mov ecx, var_14
  loc_0042254C: pop edi
  loc_0042254D: pop esi
  loc_0042254E: mov fs:[00000000h], ecx
  loc_00422555: pop ebx
  loc_00422556: mov esp, ebp
  loc_00422558: pop ebp
  loc_00422559: retn 0004h
End Sub
