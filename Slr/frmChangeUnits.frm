VERSION 5.00
Begin VB.Form frmChangeUnits
  Caption = "Change Units of Y and X"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 5970
  ClientHeight = 3825
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Units"
    Left = 0
    Top = 0
    Width = 5895
    Height = 3735
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
      Caption = "Independent Variable Units"
      Left = 240
      Top = 1800
      Width = 5415
      Height = 1215
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
      Begin VB.TextBox txtXUnits
        Left = 240
        Top = 480
        Width = 4935
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
      Caption = "Dependent Variable Units"
      Left = 240
      Top = 480
      Width = 5415
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
      Begin VB.TextBox txtYUnits
        Left = 240
        Top = 480
        Width = 4935
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
      Left = 2880
      Top = 3120
      Width = 2775
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
      Left = 1440
      Top = 3120
      Width = 1095
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
      Left = 240
      Top = 3120
      Width = 975
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

Attribute VB_Name = "frmChangeUnits"


Private Sub cmdClearChange_Click() '4228D0
  loc_004228D0: push ebp
  loc_004228D1: mov ebp, esp
  loc_004228D3: sub esp, 0000000Ch
  loc_004228D6: push 00401926h ; __vbaExceptHandler
  loc_004228DB: mov eax, fs:[00000000h]
  loc_004228E1: push eax
  loc_004228E2: mov fs:[00000000h], esp
  loc_004228E9: sub esp, 00000014h
  loc_004228EC: push ebx
  loc_004228ED: push esi
  loc_004228EE: push edi
  loc_004228EF: mov var_C, esp
  loc_004228F2: mov var_8, 00401480h
  loc_004228F9: mov esi, Me
  loc_004228FC: mov eax, esi
  loc_004228FE: and eax, 00000001h
  loc_00422901: mov var_4, eax
  loc_00422904: and esi, FFFFFFFEh
  loc_00422907: push esi
  loc_00422908: mov Me, esi
  loc_0042290B: mov ecx, [esi]
  loc_0042290D: call [ecx+00000004h]
  loc_00422910: mov edx, [esi]
  loc_00422912: push esi
  loc_00422913: mov var_18, 00000000h
  loc_0042291A: call [edx+0000030Ch]
  loc_00422920: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_00422926: push eax
  loc_00422927: lea eax, var_18
  loc_0042292A: push eax
  loc_0042292B: call ebx
  loc_0042292D: mov edi, eax
  loc_0042292F: push 0040F040h
  loc_00422934: push edi
  loc_00422935: mov ecx, [edi]
  loc_00422937: call [ecx+000000A4h]
  loc_0042293D: test eax, eax
  loc_0042293F: fnclex
  loc_00422941: jge 00422955h
  loc_00422943: push 000000A4h
  loc_00422948: push 0040F02Ch
  loc_0042294D: push edi
  loc_0042294E: push eax
  loc_0042294F: call [00401030h] ; __vbaHresultCheckObj
  loc_00422955: lea ecx, var_18
  loc_00422958: call [00401114h] ; __vbaFreeObj
  loc_0042295E: mov edx, [esi]
  loc_00422960: push esi
  loc_00422961: call [edx+00000304h]
  loc_00422967: push eax
  loc_00422968: lea eax, var_18
  loc_0042296B: push eax
  loc_0042296C: call ebx
  loc_0042296E: mov edi, eax
  loc_00422970: push 0040F040h
  loc_00422975: push edi
  loc_00422976: mov ecx, [edi]
  loc_00422978: call [ecx+000000A4h]
  loc_0042297E: test eax, eax
  loc_00422980: fnclex
  loc_00422982: jge 00422996h
  loc_00422984: push 000000A4h
  loc_00422989: push 0040F02Ch
  loc_0042298E: push edi
  loc_0042298F: push eax
  loc_00422990: call [00401030h] ; __vbaHresultCheckObj
  loc_00422996: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042299C: lea ecx, var_18
  loc_0042299F: call edi
  loc_004229A1: mov edx, [esi]
  loc_004229A3: push esi
  loc_004229A4: call [edx+0000030Ch]
  loc_004229AA: push eax
  loc_004229AB: lea eax, var_18
  loc_004229AE: push eax
  loc_004229AF: call ebx
  loc_004229B1: mov esi, eax
  loc_004229B3: push esi
  loc_004229B4: mov ecx, [esi]
  loc_004229B6: call [ecx+00000204h]
  loc_004229BC: test eax, eax
  loc_004229BE: fnclex
  loc_004229C0: jge 004229D4h
  loc_004229C2: push 00000204h
  loc_004229C7: push 0040F02Ch
  loc_004229CC: push esi
  loc_004229CD: push eax
  loc_004229CE: call [00401030h] ; __vbaHresultCheckObj
  loc_004229D4: lea ecx, var_18
  loc_004229D7: call edi
  loc_004229D9: mov var_4, 00000000h
  loc_004229E0: push 004229F2h
  loc_004229E5: jmp 004229F1h
  loc_004229E7: lea ecx, var_18
  loc_004229EA: call [00401114h] ; __vbaFreeObj
  loc_004229F0: ret
  loc_004229F1: ret
  loc_004229F2: mov eax, Me
  loc_004229F5: push eax
  loc_004229F6: mov edx, [eax]
  loc_004229F8: call [edx+00000008h]
  loc_004229FB: mov eax, var_4
  loc_004229FE: mov ecx, var_14
  loc_00422A01: pop edi
  loc_00422A02: pop esi
  loc_00422A03: mov fs:[00000000h], ecx
  loc_00422A0A: pop ebx
  loc_00422A0B: mov esp, ebp
  loc_00422A0D: pop ebp
  loc_00422A0E: retn 0004h
End Sub

Private Sub cmdChange_Click() '422C20
  loc_00422C20: push ebp
  loc_00422C21: mov ebp, esp
  loc_00422C23: sub esp, 0000000Ch
  loc_00422C26: push 00401926h ; __vbaExceptHandler
  loc_00422C2B: mov eax, fs:[00000000h]
  loc_00422C31: push eax
  loc_00422C32: mov fs:[00000000h], esp
  loc_00422C39: sub esp, 000000A0h
  loc_00422C3F: push ebx
  loc_00422C40: push esi
  loc_00422C41: push edi
  loc_00422C42: mov var_C, esp
  loc_00422C45: mov var_8, 004014A0h
  loc_00422C4C: mov esi, Me
  loc_00422C4F: mov eax, esi
  loc_00422C51: and eax, 00000001h
  loc_00422C54: mov var_4, eax
  loc_00422C57: and esi, FFFFFFFEh
  loc_00422C5A: push esi
  loc_00422C5B: mov Me, esi
  loc_00422C5E: mov ecx, [esi]
  loc_00422C60: call [ecx+00000004h]
  loc_00422C63: mov edx, [esi]
  loc_00422C65: xor ebx, ebx
  loc_00422C67: mov var_6C, ebx
  loc_00422C6A: push esi
  loc_00422C6B: mov var_18, ebx
  loc_00422C6E: mov var_1C, ebx
  loc_00422C71: mov var_2C, ebx
  loc_00422C74: mov var_3C, ebx
  loc_00422C77: mov var_4C, ebx
  loc_00422C7A: mov var_5C, ebx
  loc_00422C7D: mov var_7C, ebx
  loc_00422C80: mov var_64, 0040F028h
  loc_00422C87: mov var_6C, 00000008h
  loc_00422C8E: call [edx+0000030Ch]
  loc_00422C94: push eax
  loc_00422C95: lea eax, var_1C
  loc_00422C98: push eax
  loc_00422C99: call [0040103Ch] ; __vbaObjSet
  loc_00422C9F: mov edi, eax
  loc_00422CA1: lea edx, var_18
  loc_00422CA4: push edx
  loc_00422CA5: push edi
  loc_00422CA6: mov ecx, [edi]
  loc_00422CA8: call [ecx+000000A0h]
  loc_00422CAE: cmp eax, ebx
  loc_00422CB0: fnclex
  loc_00422CB2: jge 00422CC6h
  loc_00422CB4: push 000000A0h
  loc_00422CB9: push 0040F02Ch
  loc_00422CBE: push edi
  loc_00422CBF: push eax
  loc_00422CC0: call [00401030h] ; __vbaHresultCheckObj
  loc_00422CC6: mov eax, var_18
  loc_00422CC9: lea ecx, var_3C
  loc_00422CCC: mov var_24, eax
  loc_00422CCF: lea eax, var_2C
  loc_00422CD2: mov var_18, ebx
  loc_00422CD5: mov ebx, [00401050h] ; rtcTrimVar
  loc_00422CDB: push eax
  loc_00422CDC: push ecx
  loc_00422CDD: mov var_2C, 00000008h
  loc_00422CE4: call ebx
  loc_00422CE6: lea edx, var_6C
  loc_00422CE9: lea eax, var_3C
  loc_00422CEC: push edx
  loc_00422CED: lea ecx, var_4C
  loc_00422CF0: push eax
  loc_00422CF1: push ecx
  loc_00422CF2: call [004010C0h] ; __vbaVarCat
  loc_00422CF8: push eax
  loc_00422CF9: call [00401014h] ; __vbaStrVarMove
  loc_00422CFF: mov edx, eax
  loc_00422D01: mov ecx, 00430014h
  loc_00422D06: call [004010FCh] ; __vbaStrMove
  loc_00422D0C: lea ecx, var_1C
  loc_00422D0F: call [00401114h] ; __vbaFreeObj
  loc_00422D15: mov edi, [00401018h] ; __vbaFreeVarList
  loc_00422D1B: lea edx, var_4C
  loc_00422D1E: lea eax, var_3C
  loc_00422D21: push edx
  loc_00422D22: lea ecx, var_2C
  loc_00422D25: push eax
  loc_00422D26: push ecx
  loc_00422D27: push 00000003h
  loc_00422D29: call edi
  loc_00422D2B: mov edx, [esi]
  loc_00422D2D: add esp, 00000010h
  loc_00422D30: mov var_64, 0040F028h
  loc_00422D37: mov var_6C, 00000008h
  loc_00422D3E: push esi
  loc_00422D3F: call [edx+00000304h]
  loc_00422D45: push eax
  loc_00422D46: lea eax, var_1C
  loc_00422D49: push eax
  loc_00422D4A: call [0040103Ch] ; __vbaObjSet
  loc_00422D50: mov ecx, [eax]
  loc_00422D52: lea edx, var_18
  loc_00422D55: push edx
  loc_00422D56: push eax
  loc_00422D57: mov var_A0, eax
  loc_00422D5D: call [ecx+000000A0h]
  loc_00422D63: test eax, eax
  loc_00422D65: fnclex
  loc_00422D67: jge 00422D81h
  loc_00422D69: mov ecx, var_A0
  loc_00422D6F: push 000000A0h
  loc_00422D74: push 0040F02Ch
  loc_00422D79: push ecx
  loc_00422D7A: push eax
  loc_00422D7B: call [00401030h] ; __vbaHresultCheckObj
  loc_00422D81: mov eax, var_18
  loc_00422D84: lea edx, var_2C
  loc_00422D87: mov var_24, eax
  loc_00422D8A: lea eax, var_3C
  loc_00422D8D: push edx
  loc_00422D8E: push eax
  loc_00422D8F: mov var_18, 00000000h
  loc_00422D96: mov var_2C, 00000008h
  loc_00422D9D: call ebx
  loc_00422D9F: lea ecx, var_6C
  loc_00422DA2: lea edx, var_3C
  loc_00422DA5: push ecx
  loc_00422DA6: lea eax, var_4C
  loc_00422DA9: push edx
  loc_00422DAA: push eax
  loc_00422DAB: call [004010C0h] ; __vbaVarCat
  loc_00422DB1: push eax
  loc_00422DB2: call [00401014h] ; __vbaStrVarMove
  loc_00422DB8: mov edx, eax
  loc_00422DBA: mov ecx, 0043001Ch
  loc_00422DBF: call [004010FCh] ; __vbaStrMove
  loc_00422DC5: lea ecx, var_1C
  loc_00422DC8: call [00401114h] ; __vbaFreeObj
  loc_00422DCE: lea ecx, var_4C
  loc_00422DD1: lea edx, var_3C
  loc_00422DD4: push ecx
  loc_00422DD5: lea eax, var_2C
  loc_00422DD8: push edx
  loc_00422DD9: push eax
  loc_00422DDA: push 00000003h
  loc_00422DDC: call edi
  loc_00422DDE: mov ecx, [esi]
  loc_00422DE0: add esp, 00000010h
  loc_00422DE3: push esi
  loc_00422DE4: call [ecx+0000030Ch]
  loc_00422DEA: mov var_24, eax
  loc_00422DED: lea edx, var_2C
  loc_00422DF0: lea eax, var_3C
  loc_00422DF3: push edx
  loc_00422DF4: push eax
  loc_00422DF5: mov var_2C, 00000009h
  loc_00422DFC: call ebx
  loc_00422DFE: lea ecx, var_3C
  loc_00422E01: lea edx, var_6C
  loc_00422E04: push ecx
  loc_00422E05: push edx
  loc_00422E06: mov var_64, 0040F040h
  loc_00422E0D: mov var_6C, 00008008h
  loc_00422E14: call [00401070h] ; __vbaVarTstEq
  loc_00422E1A: mov var_A0, ax
  loc_00422E21: lea eax, var_3C
  loc_00422E24: lea ecx, var_2C
  loc_00422E27: push eax
  loc_00422E28: push ecx
  loc_00422E29: push 00000002h
  loc_00422E2B: call edi
  loc_00422E2D: add esp, 0000000Ch
  loc_00422E30: cmp var_A0, 0000h
  loc_00422E38: jz 00422EF7h
  loc_00422E3E: mov edx, [esi]
  loc_00422E40: push esi
  loc_00422E41: call [edx+0000030Ch]
  loc_00422E47: push eax
  loc_00422E48: lea eax, var_1C
  loc_00422E4B: push eax
  loc_00422E4C: call [0040103Ch] ; __vbaObjSet
  loc_00422E52: mov esi, eax
  loc_00422E54: push esi
  loc_00422E55: mov ecx, [esi]
  loc_00422E57: call [ecx+00000204h]
  loc_00422E5D: test eax, eax
  loc_00422E5F: fnclex
  loc_00422E61: jge 00422E75h
  loc_00422E63: push 00000204h
  loc_00422E68: push 0040F02Ch
  loc_00422E6D: push esi
  loc_00422E6E: push eax
  loc_00422E6F: call [00401030h] ; __vbaHresultCheckObj
  loc_00422E75: lea ecx, var_1C
  loc_00422E78: call [00401114h] ; __vbaFreeObj
  loc_00422E7E: mov esi, [004010F4h] ; __vbaVarDup
  loc_00422E84: mov ecx, 80020004h
  loc_00422E89: mov var_54, ecx
  loc_00422E8C: mov eax, 0000000Ah
  loc_00422E91: mov var_44, ecx
  loc_00422E94: mov ebx, 00000008h
  loc_00422E99: lea edx, var_7C
  loc_00422E9C: lea ecx, var_3C
  loc_00422E9F: mov var_5C, eax
  loc_00422EA2: mov var_4C, eax
  loc_00422EA5: mov var_74, 0040F1F0h ; "Check Units"
  loc_00422EAC: mov var_7C, ebx
  loc_00422EAF: call __vbaVarDup
  loc_00422EB1: lea edx, var_6C
  loc_00422EB4: lea ecx, var_2C
  loc_00422EB7: mov var_64, 0041049Ch ; "Please enter the Units for the dependent variable."
  loc_00422EBE: mov var_6C, ebx
  loc_00422EC1: call __vbaVarDup
  loc_00422EC3: lea edx, var_5C
  loc_00422EC6: lea eax, var_4C
  loc_00422EC9: push edx
  loc_00422ECA: lea ecx, var_3C
  loc_00422ECD: push eax
  loc_00422ECE: push ecx
  loc_00422ECF: lea edx, var_2C
  loc_00422ED2: push 00000000h
  loc_00422ED4: push edx
  loc_00422ED5: call [00401038h] ; rtcMsgBox
  loc_00422EDB: lea eax, var_5C
  loc_00422EDE: lea ecx, var_4C
  loc_00422EE1: push eax
  loc_00422EE2: lea edx, var_3C
  loc_00422EE5: push ecx
  loc_00422EE6: lea eax, var_2C
  loc_00422EE9: push edx
  loc_00422EEA: push eax
  loc_00422EEB: push 00000004h
  loc_00422EED: call edi
  loc_00422EEF: add esp, 00000014h
  loc_00422EF2: jmp 00423044h
  loc_00422EF7: mov ecx, [esi]
  loc_00422EF9: push esi
  loc_00422EFA: call [ecx+00000304h]
  loc_00422F00: mov var_24, eax
  loc_00422F03: lea edx, var_2C
  loc_00422F06: lea eax, var_3C
  loc_00422F09: push edx
  loc_00422F0A: push eax
  loc_00422F0B: mov var_2C, 00000009h
  loc_00422F12: call ebx
  loc_00422F14: lea ecx, var_3C
  loc_00422F17: lea edx, var_6C
  loc_00422F1A: push ecx
  loc_00422F1B: push edx
  loc_00422F1C: mov var_64, 0040F040h
  loc_00422F23: mov var_6C, 00008008h
  loc_00422F2A: call [00401070h] ; __vbaVarTstEq
  loc_00422F30: mov bx, ax
  loc_00422F33: lea eax, var_3C
  loc_00422F36: lea ecx, var_2C
  loc_00422F39: push eax
  loc_00422F3A: push ecx
  loc_00422F3B: push 00000002h
  loc_00422F3D: call edi
  loc_00422F3F: add esp, 0000000Ch
  loc_00422F42: test bx, bx
  loc_00422F45: jz 00423004h
  loc_00422F4B: mov ebx, [004010F4h] ; __vbaVarDup
  loc_00422F51: mov ecx, 80020004h
  loc_00422F56: mov var_54, ecx
  loc_00422F59: mov eax, 0000000Ah
  loc_00422F5E: mov var_44, ecx
  loc_00422F61: lea edx, var_7C
  loc_00422F64: lea ecx, var_3C
  loc_00422F67: mov var_5C, eax
  loc_00422F6A: mov var_4C, eax
  loc_00422F6D: mov var_74, 0040F1F0h ; "Check Units"
  loc_00422F74: mov var_7C, 00000008h
  loc_00422F7B: call ebx
  loc_00422F7D: lea edx, var_6C
  loc_00422F80: lea ecx, var_2C
  loc_00422F83: mov var_64, 00410508h ; "Please enter the Units for the independent variable."
  loc_00422F8A: mov var_6C, 00000008h
  loc_00422F91: call ebx
  loc_00422F93: lea edx, var_5C
  loc_00422F96: lea eax, var_4C
  loc_00422F99: push edx
  loc_00422F9A: lea ecx, var_3C
  loc_00422F9D: push eax
  loc_00422F9E: push ecx
  loc_00422F9F: lea edx, var_2C
  loc_00422FA2: push 00000000h
  loc_00422FA4: push edx
  loc_00422FA5: call [00401038h] ; rtcMsgBox
  loc_00422FAB: lea eax, var_5C
  loc_00422FAE: lea ecx, var_4C
  loc_00422FB1: push eax
  loc_00422FB2: lea edx, var_3C
  loc_00422FB5: push ecx
  loc_00422FB6: lea eax, var_2C
  loc_00422FB9: push edx
  loc_00422FBA: push eax
  loc_00422FBB: push 00000004h
  loc_00422FBD: call edi
  loc_00422FBF: mov ecx, [esi]
  loc_00422FC1: add esp, 00000014h
  loc_00422FC4: push esi
  loc_00422FC5: call [ecx+00000304h]
  loc_00422FCB: lea edx, var_1C
  loc_00422FCE: push eax
  loc_00422FCF: push edx
  loc_00422FD0: call [0040103Ch] ; __vbaObjSet
  loc_00422FD6: mov esi, eax
  loc_00422FD8: push esi
  loc_00422FD9: mov eax, [esi]
  loc_00422FDB: call [eax+00000204h]
  loc_00422FE1: test eax, eax
  loc_00422FE3: fnclex
  loc_00422FE5: jge 00422FF9h
  loc_00422FE7: push 00000204h
  loc_00422FEC: push 0040F02Ch
  loc_00422FF1: push esi
  loc_00422FF2: push eax
  loc_00422FF3: call [00401030h] ; __vbaHresultCheckObj
  loc_00422FF9: lea ecx, var_1C
  loc_00422FFC: call [00401114h] ; __vbaFreeObj
  loc_00423002: jmp 00423044h
  loc_00423004: mov eax, [004301A0h]
  loc_00423009: test eax, eax
  loc_0042300B: jnz 0042301Dh
  loc_0042300D: push 004301A0h
  loc_00423012: push 004033C0h
  loc_00423017: call [004010D4h] ; __vbaNew2
  loc_0042301D: mov esi, [004301A0h]
  loc_00423023: push esi
  loc_00423024: mov ecx, [esi]
  loc_00423026: call [ecx+000002B4h]
  loc_0042302C: test eax, eax
  loc_0042302E: fnclex
  loc_00423030: jge 00423044h
  loc_00423032: push 000002B4h
  loc_00423037: push 00410454h
  loc_0042303C: push esi
  loc_0042303D: push eax
  loc_0042303E: call [00401030h] ; __vbaHresultCheckObj
  loc_00423044: mov var_4, 00000000h
  loc_0042304B: push 00423081h
  loc_00423050: jmp 00423080h
  loc_00423052: lea ecx, var_18
  loc_00423055: call [00401110h] ; __vbaFreeStr
  loc_0042305B: lea ecx, var_1C
  loc_0042305E: call [00401114h] ; __vbaFreeObj
  loc_00423064: lea edx, var_5C
  loc_00423067: lea eax, var_4C
  loc_0042306A: push edx
  loc_0042306B: lea ecx, var_3C
  loc_0042306E: push eax
  loc_0042306F: lea edx, var_2C
  loc_00423072: push ecx
  loc_00423073: push edx
  loc_00423074: push 00000004h
  loc_00423076: call [00401018h] ; __vbaFreeVarList
  loc_0042307C: add esp, 00000014h
  loc_0042307F: ret
  loc_00423080: ret
  loc_00423081: mov eax, Me
  loc_00423084: push eax
  loc_00423085: mov ecx, [eax]
  loc_00423087: call [ecx+00000008h]
  loc_0042308A: mov eax, var_4
  loc_0042308D: mov ecx, var_14
  loc_00423090: pop edi
  loc_00423091: pop esi
  loc_00423092: mov fs:[00000000h], ecx
  loc_00423099: pop ebx
  loc_0042309A: mov esp, ebp
  loc_0042309C: pop ebp
  loc_0042309D: retn 0004h
End Sub

Private Sub cmdRestore_Click() '422A20
  loc_00422A20: push ebp
  loc_00422A21: mov ebp, esp
  loc_00422A23: sub esp, 0000000Ch
  loc_00422A26: push 00401926h ; __vbaExceptHandler
  loc_00422A2B: mov eax, fs:[00000000h]
  loc_00422A31: push eax
  loc_00422A32: mov fs:[00000000h], esp
  loc_00422A39: sub esp, 00000014h
  loc_00422A3C: push ebx
  loc_00422A3D: push esi
  loc_00422A3E: push edi
  loc_00422A3F: mov var_C, esp
  loc_00422A42: mov var_8, 00401490h
  loc_00422A49: mov esi, Me
  loc_00422A4C: mov eax, esi
  loc_00422A4E: and eax, 00000001h
  loc_00422A51: mov var_4, eax
  loc_00422A54: and esi, FFFFFFFEh
  loc_00422A57: push esi
  loc_00422A58: mov Me, esi
  loc_00422A5B: mov ecx, [esi]
  loc_00422A5D: call [ecx+00000004h]
  loc_00422A60: mov edi, [004010E0h] ; __vbaStrCopy
  loc_00422A66: mov edx, 004102A0h ; " millions of dollars"
  loc_00422A6B: mov ecx, 00430014h
  loc_00422A70: mov var_18, 00000000h
  loc_00422A77: call edi
  loc_00422A79: mov edx, 004102A0h ; " millions of dollars"
  loc_00422A7E: mov ecx, 0043001Ch
  loc_00422A83: call edi
  loc_00422A85: mov edx, [esi]
  loc_00422A87: push esi
  loc_00422A88: call [edx+0000030Ch]
  loc_00422A8E: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_00422A94: push eax
  loc_00422A95: lea eax, var_18
  loc_00422A98: push eax
  loc_00422A99: call ebx
  loc_00422A9B: mov edi, eax
  loc_00422A9D: push 004102A0h ; " millions of dollars"
  loc_00422AA2: push edi
  loc_00422AA3: mov ecx, [edi]
  loc_00422AA5: call [ecx+000000A4h]
  loc_00422AAB: test eax, eax
  loc_00422AAD: fnclex
  loc_00422AAF: jge 00422AC3h
  loc_00422AB1: push 000000A4h
  loc_00422AB6: push 0040F02Ch
  loc_00422ABB: push edi
  loc_00422ABC: push eax
  loc_00422ABD: call [00401030h] ; __vbaHresultCheckObj
  loc_00422AC3: mov edi, [00401114h] ; __vbaFreeObj
  loc_00422AC9: lea ecx, var_18
  loc_00422ACC: call edi
  loc_00422ACE: mov edx, [esi]
  loc_00422AD0: push esi
  loc_00422AD1: call [edx+00000304h]
  loc_00422AD7: push eax
  loc_00422AD8: lea eax, var_18
  loc_00422ADB: push eax
  loc_00422ADC: call ebx
  loc_00422ADE: mov esi, eax
  loc_00422AE0: push 004102A0h ; " millions of dollars"
  loc_00422AE5: push esi
  loc_00422AE6: mov ecx, [esi]
  loc_00422AE8: call [ecx+000000A4h]
  loc_00422AEE: test eax, eax
  loc_00422AF0: fnclex
  loc_00422AF2: jge 00422B06h
  loc_00422AF4: push 000000A4h
  loc_00422AF9: push 0040F02Ch
  loc_00422AFE: push esi
  loc_00422AFF: push eax
  loc_00422B00: call [00401030h] ; __vbaHresultCheckObj
  loc_00422B06: lea ecx, var_18
  loc_00422B09: call edi
  loc_00422B0B: mov var_4, 00000000h
  loc_00422B12: push 00422B24h
  loc_00422B17: jmp 00422B23h
  loc_00422B19: lea ecx, var_18
  loc_00422B1C: call [00401114h] ; __vbaFreeObj
  loc_00422B22: ret
  loc_00422B23: ret
  loc_00422B24: mov eax, Me
  loc_00422B27: push eax
  loc_00422B28: mov edx, [eax]
  loc_00422B2A: call [edx+00000008h]
  loc_00422B2D: mov eax, var_4
  loc_00422B30: mov ecx, var_14
  loc_00422B33: pop edi
  loc_00422B34: pop esi
  loc_00422B35: mov fs:[00000000h], ecx
  loc_00422B3C: pop ebx
  loc_00422B3D: mov esp, ebp
  loc_00422B3F: pop ebp
  loc_00422B40: retn 0004h
End Sub

Private Sub Form_Activate() '4230A0
  loc_004230A0: push ebp
  loc_004230A1: mov ebp, esp
  loc_004230A3: sub esp, 0000000Ch
  loc_004230A6: push 00401926h ; __vbaExceptHandler
  loc_004230AB: mov eax, fs:[00000000h]
  loc_004230B1: push eax
  loc_004230B2: mov fs:[00000000h], esp
  loc_004230B9: sub esp, 00000014h
  loc_004230BC: push ebx
  loc_004230BD: push esi
  loc_004230BE: push edi
  loc_004230BF: mov var_C, esp
  loc_004230C2: mov var_8, 004014B0h
  loc_004230C9: mov esi, Me
  loc_004230CC: mov eax, esi
  loc_004230CE: and eax, 00000001h
  loc_004230D1: mov var_4, eax
  loc_004230D4: and esi, FFFFFFFEh
  loc_004230D7: push esi
  loc_004230D8: mov Me, esi
  loc_004230DB: mov ecx, [esi]
  loc_004230DD: call [ecx+00000004h]
  loc_004230E0: mov edx, [esi]
  loc_004230E2: push esi
  loc_004230E3: mov var_18, 00000000h
  loc_004230EA: call [edx+0000030Ch]
  loc_004230F0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_004230F6: push eax
  loc_004230F7: lea eax, var_18
  loc_004230FA: push eax
  loc_004230FB: call ebx
  loc_004230FD: mov edx, [00430014h]
  loc_00423103: mov edi, eax
  loc_00423105: push edx
  loc_00423106: push edi
  loc_00423107: mov ecx, [edi]
  loc_00423109: call [ecx+000000A4h]
  loc_0042310F: test eax, eax
  loc_00423111: fnclex
  loc_00423113: jge 00423127h
  loc_00423115: push 000000A4h
  loc_0042311A: push 0040F02Ch
  loc_0042311F: push edi
  loc_00423120: push eax
  loc_00423121: call [00401030h] ; __vbaHresultCheckObj
  loc_00423127: lea ecx, var_18
  loc_0042312A: call [00401114h] ; __vbaFreeObj
  loc_00423130: mov eax, [esi]
  loc_00423132: push esi
  loc_00423133: call [eax+00000304h]
  loc_00423139: lea ecx, var_18
  loc_0042313C: push eax
  loc_0042313D: push ecx
  loc_0042313E: call ebx
  loc_00423140: mov edi, eax
  loc_00423142: mov eax, [0043001Ch]
  loc_00423147: push eax
  loc_00423148: push edi
  loc_00423149: mov edx, [edi]
  loc_0042314B: call [edx+000000A4h]
  loc_00423151: test eax, eax
  loc_00423153: fnclex
  loc_00423155: jge 00423169h
  loc_00423157: push 000000A4h
  loc_0042315C: push 0040F02Ch
  loc_00423161: push edi
  loc_00423162: push eax
  loc_00423163: call [00401030h] ; __vbaHresultCheckObj
  loc_00423169: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042316F: lea ecx, var_18
  loc_00423172: call edi
  loc_00423174: mov ecx, [esi]
  loc_00423176: push esi
  loc_00423177: call [ecx+0000030Ch]
  loc_0042317D: lea edx, var_18
  loc_00423180: push eax
  loc_00423181: push edx
  loc_00423182: call ebx
  loc_00423184: mov esi, eax
  loc_00423186: push esi
  loc_00423187: mov eax, [esi]
  loc_00423189: call [eax+00000204h]
  loc_0042318F: test eax, eax
  loc_00423191: fnclex
  loc_00423193: jge 004231A7h
  loc_00423195: push 00000204h
  loc_0042319A: push 0040F02Ch
  loc_0042319F: push esi
  loc_004231A0: push eax
  loc_004231A1: call [00401030h] ; __vbaHresultCheckObj
  loc_004231A7: lea ecx, var_18
  loc_004231AA: call edi
  loc_004231AC: mov var_4, 00000000h
  loc_004231B3: push 004231C5h
  loc_004231B8: jmp 004231C4h
  loc_004231BA: lea ecx, var_18
  loc_004231BD: call [00401114h] ; __vbaFreeObj
  loc_004231C3: ret
  loc_004231C4: ret
  loc_004231C5: mov eax, Me
  loc_004231C8: push eax
  loc_004231C9: mov ecx, [eax]
  loc_004231CB: call [ecx+00000008h]
  loc_004231CE: mov eax, var_4
  loc_004231D1: mov ecx, var_14
  loc_004231D4: pop edi
  loc_004231D5: pop esi
  loc_004231D6: mov fs:[00000000h], ecx
  loc_004231DD: pop ebx
  loc_004231DE: mov esp, ebp
  loc_004231E0: pop ebp
  loc_004231E1: retn 0004h
End Sub

Private Function Proc_10_4_422B50(arg_C) '422B50
  loc_00422B50: mov eax, [00430088h]
  loc_00422B55: sub esp, 00000010h
  loc_00422B58: test eax, eax
  loc_00422B5A: push ebx
  loc_00422B5B: push ebp
  loc_00422B5C: push esi
  loc_00422B5D: push edi
  loc_00422B5E: mov edi, [004010D4h] ; __vbaNew2
  loc_00422B64: jnz 00422B72h
  loc_00422B66: push 00430088h
  loc_00422B6B: push 004058C0h
  loc_00422B70: call edi
  loc_00422B72: mov esi, [00430088h]
  loc_00422B78: push esi
  loc_00422B79: mov eax, [esi]
  loc_00422B7B: call [eax+000002B4h]
  loc_00422B81: test eax, eax
  loc_00422B83: fnclex
  loc_00422B85: jge 00422B99h
  loc_00422B87: push 000002B4h
  loc_00422B8C: push 0040DB64h
  loc_00422B91: push esi
  loc_00422B92: push eax
  loc_00422B93: call [00401030h] ; __vbaHresultCheckObj
  loc_00422B99: mov eax, [00430164h]
  loc_00422B9E: test eax, eax
  loc_00422BA0: jnz 00422BAEh
  loc_00422BA2: push 00430164h
  loc_00422BA7: push 00408108h
  loc_00422BAC: call edi
  loc_00422BAE: sub esp, 00000010h
  loc_00422BB1: mov ecx, 0000000Ah
  loc_00422BB6: mov ebp, esp
  loc_00422BB8: mov edi, ecx
  loc_00422BBA: mov eax, 80020004h
  loc_00422BBF: sub esp, 00000010h
  loc_00422BC2: mov [ebp], ecx
  loc_00422BC5: mov ecx, arg_2C
  loc_00422BC9: mov esi, [00430164h]
  loc_00422BCF: mov edx, eax
  loc_00422BD1: mov Proc_10_4_422B50, ecx
  loc_00422BD4: mov ecx, esp
  loc_00422BD6: mov ebx, [esi]
  loc_00422BD8: push esi
  loc_00422BD9: mov Me, eax
  loc_00422BDC: mov eax, arg_38
  loc_00422BE0: mov arg_C, eax
  loc_00422BE3: mov eax, arg_30
  loc_00422BE7: mov [ecx], edi
  loc_00422BE9: mov [ecx+00000004h], eax
  loc_00422BEC: mov [ecx+00000008h], edx
  loc_00422BEF: mov edx, arg_38
  loc_00422BF3: mov [ecx+0000000Ch], edx
  loc_00422BF6: call [ebx+000002B0h]
  loc_00422BFC: test eax, eax
  loc_00422BFE: fnclex
  loc_00422C00: jge 00422C14h
  loc_00422C02: push 000002B0h
  loc_00422C07: push 0040FF58h
  loc_00422C0C: push esi
  loc_00422C0D: push eax
  loc_00422C0E: call [00401030h] ; __vbaHresultCheckObj
  loc_00422C14: pop edi
  loc_00422C15: pop esi
  loc_00422C16: pop ebp
  loc_00422C17: xor eax, eax
  loc_00422C19: pop ebx
  loc_00422C1A: add esp, 00000010h
  loc_00422C1D: retn 0004h
End Function
