VERSION 5.00
Begin VB.Form frmChangeNames
  Caption = "Change Variable Names"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 6300
  ClientHeight = 4200
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Names"
    Left = 0
    Top = 0
    Width = 6255
    Height = 4215
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
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 240
      Top = 3480
      Width = 1215
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
    Begin VB.CommandButton cmdClearChange
      Caption = "Clear"
      Left = 1680
      Top = 3480
      Width = 1335
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
    Begin VB.CommandButton cmdRestore
      Caption = "Reset Defaults"
      Left = 3240
      Top = 3480
      Width = 2775
      Height = 495
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
    Begin VB.Frame Frame1
      Caption = "Dependent Variable Name"
      Left = 240
      Top = 480
      Width = 5775
      Height = 1335
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
      Begin VB.TextBox txtYName
        Left = 240
        Top = 480
        Width = 4335
        Height = 735
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
      End
    End
    Begin VB.Frame Frame2
      Caption = "Independent Variable Name"
      Left = 240
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
      Begin VB.TextBox txtXName
        Left = 240
        Top = 480
        Width = 4335
        Height = 735
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
End

Attribute VB_Name = "frmChangeNames"


Private Sub cmdClearChange_Click() '41FD50
  loc_0041FD50: push ebp
  loc_0041FD51: mov ebp, esp
  loc_0041FD53: sub esp, 0000000Ch
  loc_0041FD56: push 00401926h ; __vbaExceptHandler
  loc_0041FD5B: mov eax, fs:[00000000h]
  loc_0041FD61: push eax
  loc_0041FD62: mov fs:[00000000h], esp
  loc_0041FD69: sub esp, 00000014h
  loc_0041FD6C: push ebx
  loc_0041FD6D: push esi
  loc_0041FD6E: push edi
  loc_0041FD6F: mov var_C, esp
  loc_0041FD72: mov var_8, 00401368h
  loc_0041FD79: mov esi, Me
  loc_0041FD7C: mov eax, esi
  loc_0041FD7E: and eax, 00000001h
  loc_0041FD81: mov var_4, eax
  loc_0041FD84: and esi, FFFFFFFEh
  loc_0041FD87: push esi
  loc_0041FD88: mov Me, esi
  loc_0041FD8B: mov ecx, [esi]
  loc_0041FD8D: call [ecx+00000004h]
  loc_0041FD90: mov edx, [esi]
  loc_0041FD92: push esi
  loc_0041FD93: mov var_18, 00000000h
  loc_0041FD9A: call [edx+00000310h]
  loc_0041FDA0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041FDA6: push eax
  loc_0041FDA7: lea eax, var_18
  loc_0041FDAA: push eax
  loc_0041FDAB: call ebx
  loc_0041FDAD: mov edi, eax
  loc_0041FDAF: push 0040F040h
  loc_0041FDB4: push edi
  loc_0041FDB5: mov ecx, [edi]
  loc_0041FDB7: call [ecx+000000A4h]
  loc_0041FDBD: test eax, eax
  loc_0041FDBF: fnclex
  loc_0041FDC1: jge 0041FDD5h
  loc_0041FDC3: push 000000A4h
  loc_0041FDC8: push 0040F02Ch
  loc_0041FDCD: push edi
  loc_0041FDCE: push eax
  loc_0041FDCF: call [00401030h] ; __vbaHresultCheckObj
  loc_0041FDD5: lea ecx, var_18
  loc_0041FDD8: call [00401114h] ; __vbaFreeObj
  loc_0041FDDE: mov edx, [esi]
  loc_0041FDE0: push esi
  loc_0041FDE1: call [edx+00000318h]
  loc_0041FDE7: push eax
  loc_0041FDE8: lea eax, var_18
  loc_0041FDEB: push eax
  loc_0041FDEC: call ebx
  loc_0041FDEE: mov edi, eax
  loc_0041FDF0: push 0040F040h
  loc_0041FDF5: push edi
  loc_0041FDF6: mov ecx, [edi]
  loc_0041FDF8: call [ecx+000000A4h]
  loc_0041FDFE: test eax, eax
  loc_0041FE00: fnclex
  loc_0041FE02: jge 0041FE16h
  loc_0041FE04: push 000000A4h
  loc_0041FE09: push 0040F02Ch
  loc_0041FE0E: push edi
  loc_0041FE0F: push eax
  loc_0041FE10: call [00401030h] ; __vbaHresultCheckObj
  loc_0041FE16: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041FE1C: lea ecx, var_18
  loc_0041FE1F: call edi
  loc_0041FE21: mov edx, [esi]
  loc_0041FE23: push esi
  loc_0041FE24: call [edx+00000310h]
  loc_0041FE2A: push eax
  loc_0041FE2B: lea eax, var_18
  loc_0041FE2E: push eax
  loc_0041FE2F: call ebx
  loc_0041FE31: mov esi, eax
  loc_0041FE33: push esi
  loc_0041FE34: mov ecx, [esi]
  loc_0041FE36: call [ecx+00000204h]
  loc_0041FE3C: test eax, eax
  loc_0041FE3E: fnclex
  loc_0041FE40: jge 0041FE54h
  loc_0041FE42: push 00000204h
  loc_0041FE47: push 0040F02Ch
  loc_0041FE4C: push esi
  loc_0041FE4D: push eax
  loc_0041FE4E: call [00401030h] ; __vbaHresultCheckObj
  loc_0041FE54: lea ecx, var_18
  loc_0041FE57: call edi
  loc_0041FE59: mov var_4, 00000000h
  loc_0041FE60: push 0041FE72h
  loc_0041FE65: jmp 0041FE71h
  loc_0041FE67: lea ecx, var_18
  loc_0041FE6A: call [00401114h] ; __vbaFreeObj
  loc_0041FE70: ret
  loc_0041FE71: ret
  loc_0041FE72: mov eax, Me
  loc_0041FE75: push eax
  loc_0041FE76: mov edx, [eax]
  loc_0041FE78: call [edx+00000008h]
  loc_0041FE7B: mov eax, var_4
  loc_0041FE7E: mov ecx, var_14
  loc_0041FE81: pop edi
  loc_0041FE82: pop esi
  loc_0041FE83: mov fs:[00000000h], ecx
  loc_0041FE8A: pop ebx
  loc_0041FE8B: mov esp, ebp
  loc_0041FE8D: pop ebp
  loc_0041FE8E: retn 0004h
End Sub

Private Sub cmdChange_Click() '41FFD0
  loc_0041FFD0: push ebp
  loc_0041FFD1: mov ebp, esp
  loc_0041FFD3: sub esp, 0000000Ch
  loc_0041FFD6: push 00401926h ; __vbaExceptHandler
  loc_0041FFDB: mov eax, fs:[00000000h]
  loc_0041FFE1: push eax
  loc_0041FFE2: mov fs:[00000000h], esp
  loc_0041FFE9: sub esp, 000000A0h
  loc_0041FFEF: push ebx
  loc_0041FFF0: push esi
  loc_0041FFF1: push edi
  loc_0041FFF2: mov var_C, esp
  loc_0041FFF5: mov var_8, 00401388h
  loc_0041FFFC: mov esi, Me
  loc_0041FFFF: mov eax, esi
  loc_00420001: and eax, 00000001h
  loc_00420004: mov var_4, eax
  loc_00420007: and esi, FFFFFFFEh
  loc_0042000A: push esi
  loc_0042000B: mov Me, esi
  loc_0042000E: mov ecx, [esi]
  loc_00420010: call [ecx+00000004h]
  loc_00420013: mov edx, [esi]
  loc_00420015: xor ebx, ebx
  loc_00420017: mov var_6C, ebx
  loc_0042001A: push esi
  loc_0042001B: mov var_18, ebx
  loc_0042001E: mov var_1C, ebx
  loc_00420021: mov var_2C, ebx
  loc_00420024: mov var_3C, ebx
  loc_00420027: mov var_4C, ebx
  loc_0042002A: mov var_5C, ebx
  loc_0042002D: mov var_7C, ebx
  loc_00420030: mov var_64, 0040F028h
  loc_00420037: mov var_6C, 00000008h
  loc_0042003E: call [edx+00000310h]
  loc_00420044: push eax
  loc_00420045: lea eax, var_1C
  loc_00420048: push eax
  loc_00420049: call [0040103Ch] ; __vbaObjSet
  loc_0042004F: mov edi, eax
  loc_00420051: lea edx, var_18
  loc_00420054: push edx
  loc_00420055: push edi
  loc_00420056: mov ecx, [edi]
  loc_00420058: call [ecx+000000A0h]
  loc_0042005E: cmp eax, ebx
  loc_00420060: fnclex
  loc_00420062: jge 00420076h
  loc_00420064: push 000000A0h
  loc_00420069: push 0040F02Ch
  loc_0042006E: push edi
  loc_0042006F: push eax
  loc_00420070: call [00401030h] ; __vbaHresultCheckObj
  loc_00420076: mov eax, var_18
  loc_00420079: lea ecx, var_3C
  loc_0042007C: mov var_24, eax
  loc_0042007F: lea eax, var_2C
  loc_00420082: mov var_18, ebx
  loc_00420085: mov ebx, [00401050h] ; rtcTrimVar
  loc_0042008B: push eax
  loc_0042008C: push ecx
  loc_0042008D: mov var_2C, 00000008h
  loc_00420094: call ebx
  loc_00420096: lea edx, var_6C
  loc_00420099: lea eax, var_3C
  loc_0042009C: push edx
  loc_0042009D: lea ecx, var_4C
  loc_004200A0: push eax
  loc_004200A1: push ecx
  loc_004200A2: call [004010C0h] ; __vbaVarCat
  loc_004200A8: push eax
  loc_004200A9: call [00401014h] ; __vbaStrVarMove
  loc_004200AF: mov edx, eax
  loc_004200B1: mov ecx, 00430010h
  loc_004200B6: call [004010FCh] ; __vbaStrMove
  loc_004200BC: lea ecx, var_1C
  loc_004200BF: call [00401114h] ; __vbaFreeObj
  loc_004200C5: mov edi, [00401018h] ; __vbaFreeVarList
  loc_004200CB: lea edx, var_4C
  loc_004200CE: lea eax, var_3C
  loc_004200D1: push edx
  loc_004200D2: lea ecx, var_2C
  loc_004200D5: push eax
  loc_004200D6: push ecx
  loc_004200D7: push 00000003h
  loc_004200D9: call edi
  loc_004200DB: mov edx, [esi]
  loc_004200DD: add esp, 00000010h
  loc_004200E0: mov var_64, 0040F028h
  loc_004200E7: mov var_6C, 00000008h
  loc_004200EE: push esi
  loc_004200EF: call [edx+00000318h]
  loc_004200F5: push eax
  loc_004200F6: lea eax, var_1C
  loc_004200F9: push eax
  loc_004200FA: call [0040103Ch] ; __vbaObjSet
  loc_00420100: mov ecx, [eax]
  loc_00420102: lea edx, var_18
  loc_00420105: push edx
  loc_00420106: push eax
  loc_00420107: mov var_A0, eax
  loc_0042010D: call [ecx+000000A0h]
  loc_00420113: test eax, eax
  loc_00420115: fnclex
  loc_00420117: jge 00420131h
  loc_00420119: mov ecx, var_A0
  loc_0042011F: push 000000A0h
  loc_00420124: push 0040F02Ch
  loc_00420129: push ecx
  loc_0042012A: push eax
  loc_0042012B: call [00401030h] ; __vbaHresultCheckObj
  loc_00420131: mov eax, var_18
  loc_00420134: lea edx, var_2C
  loc_00420137: mov var_24, eax
  loc_0042013A: lea eax, var_3C
  loc_0042013D: push edx
  loc_0042013E: push eax
  loc_0042013F: mov var_18, 00000000h
  loc_00420146: mov var_2C, 00000008h
  loc_0042014D: call ebx
  loc_0042014F: lea ecx, var_6C
  loc_00420152: lea edx, var_3C
  loc_00420155: push ecx
  loc_00420156: lea eax, var_4C
  loc_00420159: push edx
  loc_0042015A: push eax
  loc_0042015B: call [004010C0h] ; __vbaVarCat
  loc_00420161: push eax
  loc_00420162: call [00401014h] ; __vbaStrVarMove
  loc_00420168: mov edx, eax
  loc_0042016A: mov ecx, 00430018h
  loc_0042016F: call [004010FCh] ; __vbaStrMove
  loc_00420175: lea ecx, var_1C
  loc_00420178: call [00401114h] ; __vbaFreeObj
  loc_0042017E: lea ecx, var_4C
  loc_00420181: lea edx, var_3C
  loc_00420184: push ecx
  loc_00420185: lea eax, var_2C
  loc_00420188: push edx
  loc_00420189: push eax
  loc_0042018A: push 00000003h
  loc_0042018C: call edi
  loc_0042018E: mov ecx, [esi]
  loc_00420190: add esp, 00000010h
  loc_00420193: push esi
  loc_00420194: call [ecx+00000310h]
  loc_0042019A: mov var_24, eax
  loc_0042019D: lea edx, var_2C
  loc_004201A0: lea eax, var_3C
  loc_004201A3: push edx
  loc_004201A4: push eax
  loc_004201A5: mov var_2C, 00000009h
  loc_004201AC: call ebx
  loc_004201AE: lea ecx, var_3C
  loc_004201B1: lea edx, var_6C
  loc_004201B4: push ecx
  loc_004201B5: push edx
  loc_004201B6: mov var_64, 0040F040h
  loc_004201BD: mov var_6C, 00008008h
  loc_004201C4: call [00401070h] ; __vbaVarTstEq
  loc_004201CA: mov var_A0, ax
  loc_004201D1: lea eax, var_3C
  loc_004201D4: lea ecx, var_2C
  loc_004201D7: push eax
  loc_004201D8: push ecx
  loc_004201D9: push 00000002h
  loc_004201DB: call edi
  loc_004201DD: add esp, 0000000Ch
  loc_004201E0: cmp var_A0, 0000h
  loc_004201E8: jz 004202A7h
  loc_004201EE: mov edx, [esi]
  loc_004201F0: push esi
  loc_004201F1: call [edx+00000310h]
  loc_004201F7: push eax
  loc_004201F8: lea eax, var_1C
  loc_004201FB: push eax
  loc_004201FC: call [0040103Ch] ; __vbaObjSet
  loc_00420202: mov esi, eax
  loc_00420204: push esi
  loc_00420205: mov ecx, [esi]
  loc_00420207: call [ecx+00000204h]
  loc_0042020D: test eax, eax
  loc_0042020F: fnclex
  loc_00420211: jge 00420225h
  loc_00420213: push 00000204h
  loc_00420218: push 0040F02Ch
  loc_0042021D: push esi
  loc_0042021E: push eax
  loc_0042021F: call [00401030h] ; __vbaHresultCheckObj
  loc_00420225: lea ecx, var_1C
  loc_00420228: call [00401114h] ; __vbaFreeObj
  loc_0042022E: mov esi, [004010F4h] ; __vbaVarDup
  loc_00420234: mov ecx, 80020004h
  loc_00420239: mov var_54, ecx
  loc_0042023C: mov eax, 0000000Ah
  loc_00420241: mov var_44, ecx
  loc_00420244: mov ebx, 00000008h
  loc_00420249: lea edx, var_7C
  loc_0042024C: lea ecx, var_3C
  loc_0042024F: mov var_5C, eax
  loc_00420252: mov var_4C, eax
  loc_00420255: mov var_74, 0040F100h ; "Check Name"
  loc_0042025C: mov var_7C, ebx
  loc_0042025F: call __vbaVarDup
  loc_00420261: lea edx, var_6C
  loc_00420264: lea ecx, var_2C
  loc_00420267: mov var_64, 0040F09Ch ; "Please enter a name for the dependent variable."
  loc_0042026E: mov var_6C, ebx
  loc_00420271: call __vbaVarDup
  loc_00420273: lea edx, var_5C
  loc_00420276: lea eax, var_4C
  loc_00420279: push edx
  loc_0042027A: lea ecx, var_3C
  loc_0042027D: push eax
  loc_0042027E: push ecx
  loc_0042027F: lea edx, var_2C
  loc_00420282: push 00000000h
  loc_00420284: push edx
  loc_00420285: call [00401038h] ; rtcMsgBox
  loc_0042028B: lea eax, var_5C
  loc_0042028E: lea ecx, var_4C
  loc_00420291: push eax
  loc_00420292: lea edx, var_3C
  loc_00420295: push ecx
  loc_00420296: lea eax, var_2C
  loc_00420299: push edx
  loc_0042029A: push eax
  loc_0042029B: push 00000004h
  loc_0042029D: call edi
  loc_0042029F: add esp, 00000014h
  loc_004202A2: jmp 004203F4h
  loc_004202A7: mov ecx, [esi]
  loc_004202A9: push esi
  loc_004202AA: call [ecx+00000318h]
  loc_004202B0: mov var_24, eax
  loc_004202B3: lea edx, var_2C
  loc_004202B6: lea eax, var_3C
  loc_004202B9: push edx
  loc_004202BA: push eax
  loc_004202BB: mov var_2C, 00000009h
  loc_004202C2: call ebx
  loc_004202C4: lea ecx, var_3C
  loc_004202C7: lea edx, var_6C
  loc_004202CA: push ecx
  loc_004202CB: push edx
  loc_004202CC: mov var_64, 0040F040h
  loc_004202D3: mov var_6C, 00008008h
  loc_004202DA: call [00401070h] ; __vbaVarTstEq
  loc_004202E0: mov bx, ax
  loc_004202E3: lea eax, var_3C
  loc_004202E6: lea ecx, var_2C
  loc_004202E9: push eax
  loc_004202EA: push ecx
  loc_004202EB: push 00000002h
  loc_004202ED: call edi
  loc_004202EF: add esp, 0000000Ch
  loc_004202F2: test bx, bx
  loc_004202F5: jz 004203B4h
  loc_004202FB: mov ebx, [004010F4h] ; __vbaVarDup
  loc_00420301: mov ecx, 80020004h
  loc_00420306: mov var_54, ecx
  loc_00420309: mov eax, 0000000Ah
  loc_0042030E: mov var_44, ecx
  loc_00420311: lea edx, var_7C
  loc_00420314: lea ecx, var_3C
  loc_00420317: mov var_5C, eax
  loc_0042031A: mov var_4C, eax
  loc_0042031D: mov var_74, 0040F100h ; "Check Name"
  loc_00420324: mov var_7C, 00000008h
  loc_0042032B: call ebx
  loc_0042032D: lea edx, var_6C
  loc_00420330: lea ecx, var_2C
  loc_00420333: mov var_64, 0040F11Ch ; "Please enter a name for the independent variable."
  loc_0042033A: mov var_6C, 00000008h
  loc_00420341: call ebx
  loc_00420343: lea edx, var_5C
  loc_00420346: lea eax, var_4C
  loc_00420349: push edx
  loc_0042034A: lea ecx, var_3C
  loc_0042034D: push eax
  loc_0042034E: push ecx
  loc_0042034F: lea edx, var_2C
  loc_00420352: push 00000000h
  loc_00420354: push edx
  loc_00420355: call [00401038h] ; rtcMsgBox
  loc_0042035B: lea eax, var_5C
  loc_0042035E: lea ecx, var_4C
  loc_00420361: push eax
  loc_00420362: lea edx, var_3C
  loc_00420365: push ecx
  loc_00420366: lea eax, var_2C
  loc_00420369: push edx
  loc_0042036A: push eax
  loc_0042036B: push 00000004h
  loc_0042036D: call edi
  loc_0042036F: mov ecx, [esi]
  loc_00420371: add esp, 00000014h
  loc_00420374: push esi
  loc_00420375: call [ecx+00000318h]
  loc_0042037B: lea edx, var_1C
  loc_0042037E: push eax
  loc_0042037F: push edx
  loc_00420380: call [0040103Ch] ; __vbaObjSet
  loc_00420386: mov esi, eax
  loc_00420388: push esi
  loc_00420389: mov eax, [esi]
  loc_0042038B: call [eax+00000204h]
  loc_00420391: test eax, eax
  loc_00420393: fnclex
  loc_00420395: jge 004203A9h
  loc_00420397: push 00000204h
  loc_0042039C: push 0040F02Ch
  loc_004203A1: push esi
  loc_004203A2: push eax
  loc_004203A3: call [00401030h] ; __vbaHresultCheckObj
  loc_004203A9: lea ecx, var_1C
  loc_004203AC: call [00401114h] ; __vbaFreeObj
  loc_004203B2: jmp 004203F4h
  loc_004203B4: mov eax, [004300D8h]
  loc_004203B9: test eax, eax
  loc_004203BB: jnz 004203CDh
  loc_004203BD: push 004300D8h
  loc_004203C2: push 00402E04h
  loc_004203C7: call [004010D4h] ; __vbaNew2
  loc_004203CD: mov esi, [004300D8h]
  loc_004203D3: push esi
  loc_004203D4: mov ecx, [esi]
  loc_004203D6: call [ecx+000002B4h]
  loc_004203DC: test eax, eax
  loc_004203DE: fnclex
  loc_004203E0: jge 004203F4h
  loc_004203E2: push 000002B4h
  loc_004203E7: push 0040E260h
  loc_004203EC: push esi
  loc_004203ED: push eax
  loc_004203EE: call [00401030h] ; __vbaHresultCheckObj
  loc_004203F4: mov var_4, 00000000h
  loc_004203FB: push 00420431h
  loc_00420400: jmp 00420430h
  loc_00420402: lea ecx, var_18
  loc_00420405: call [00401110h] ; __vbaFreeStr
  loc_0042040B: lea ecx, var_1C
  loc_0042040E: call [00401114h] ; __vbaFreeObj
  loc_00420414: lea edx, var_5C
  loc_00420417: lea eax, var_4C
  loc_0042041A: push edx
  loc_0042041B: lea ecx, var_3C
  loc_0042041E: push eax
  loc_0042041F: lea edx, var_2C
  loc_00420422: push ecx
  loc_00420423: push edx
  loc_00420424: push 00000004h
  loc_00420426: call [00401018h] ; __vbaFreeVarList
  loc_0042042C: add esp, 00000014h
  loc_0042042F: ret
  loc_00420430: ret
  loc_00420431: mov eax, Me
  loc_00420434: push eax
  loc_00420435: mov ecx, [eax]
  loc_00420437: call [ecx+00000008h]
  loc_0042043A: mov eax, var_4
  loc_0042043D: mov ecx, var_14
  loc_00420440: pop edi
  loc_00420441: pop esi
  loc_00420442: mov fs:[00000000h], ecx
  loc_00420449: pop ebx
  loc_0042044A: mov esp, ebp
  loc_0042044C: pop ebp
  loc_0042044D: retn 0004h
End Sub

Private Sub cmdRestore_Click() '41FEA0
  loc_0041FEA0: push ebp
  loc_0041FEA1: mov ebp, esp
  loc_0041FEA3: sub esp, 0000000Ch
  loc_0041FEA6: push 00401926h ; __vbaExceptHandler
  loc_0041FEAB: mov eax, fs:[00000000h]
  loc_0041FEB1: push eax
  loc_0041FEB2: mov fs:[00000000h], esp
  loc_0041FEB9: sub esp, 00000014h
  loc_0041FEBC: push ebx
  loc_0041FEBD: push esi
  loc_0041FEBE: push edi
  loc_0041FEBF: mov var_C, esp
  loc_0041FEC2: mov var_8, 00401378h
  loc_0041FEC9: mov esi, Me
  loc_0041FECC: mov eax, esi
  loc_0041FECE: and eax, 00000001h
  loc_0041FED1: mov var_4, eax
  loc_0041FED4: and esi, FFFFFFFEh
  loc_0041FED7: push esi
  loc_0041FED8: mov Me, esi
  loc_0041FEDB: mov ecx, [esi]
  loc_0041FEDD: call [ecx+00000004h]
  loc_0041FEE0: mov edi, [004010E0h] ; __vbaStrCopy
  loc_0041FEE6: mov edx, 0040DB10h ; " Net Income"
  loc_0041FEEB: mov ecx, 00430010h
  loc_0041FEF0: mov var_18, 00000000h
  loc_0041FEF7: call edi
  loc_0041FEF9: mov edx, 0040DB54h ; " Sales"
  loc_0041FEFE: mov ecx, 00430018h
  loc_0041FF03: call edi
  loc_0041FF05: mov edx, [esi]
  loc_0041FF07: push esi
  loc_0041FF08: call [edx+00000310h]
  loc_0041FF0E: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041FF14: push eax
  loc_0041FF15: lea eax, var_18
  loc_0041FF18: push eax
  loc_0041FF19: call ebx
  loc_0041FF1B: mov edi, eax
  loc_0041FF1D: push 0040DB10h ; " Net Income"
  loc_0041FF22: push edi
  loc_0041FF23: mov ecx, [edi]
  loc_0041FF25: call [ecx+000000A4h]
  loc_0041FF2B: test eax, eax
  loc_0041FF2D: fnclex
  loc_0041FF2F: jge 0041FF43h
  loc_0041FF31: push 000000A4h
  loc_0041FF36: push 0040F02Ch
  loc_0041FF3B: push edi
  loc_0041FF3C: push eax
  loc_0041FF3D: call [00401030h] ; __vbaHresultCheckObj
  loc_0041FF43: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041FF49: lea ecx, var_18
  loc_0041FF4C: call edi
  loc_0041FF4E: mov edx, [esi]
  loc_0041FF50: push esi
  loc_0041FF51: call [edx+00000318h]
  loc_0041FF57: push eax
  loc_0041FF58: lea eax, var_18
  loc_0041FF5B: push eax
  loc_0041FF5C: call ebx
  loc_0041FF5E: mov esi, eax
  loc_0041FF60: push 0040DB54h ; " Sales"
  loc_0041FF65: push esi
  loc_0041FF66: mov ecx, [esi]
  loc_0041FF68: call [ecx+000000A4h]
  loc_0041FF6E: test eax, eax
  loc_0041FF70: fnclex
  loc_0041FF72: jge 0041FF86h
  loc_0041FF74: push 000000A4h
  loc_0041FF79: push 0040F02Ch
  loc_0041FF7E: push esi
  loc_0041FF7F: push eax
  loc_0041FF80: call [00401030h] ; __vbaHresultCheckObj
  loc_0041FF86: lea ecx, var_18
  loc_0041FF89: call edi
  loc_0041FF8B: mov var_4, 00000000h
  loc_0041FF92: push 0041FFA4h
  loc_0041FF97: jmp 0041FFA3h
  loc_0041FF99: lea ecx, var_18
  loc_0041FF9C: call [00401114h] ; __vbaFreeObj
  loc_0041FFA2: ret
  loc_0041FFA3: ret
  loc_0041FFA4: mov eax, Me
  loc_0041FFA7: push eax
  loc_0041FFA8: mov edx, [eax]
  loc_0041FFAA: call [edx+00000008h]
  loc_0041FFAD: mov eax, var_4
  loc_0041FFB0: mov ecx, var_14
  loc_0041FFB3: pop edi
  loc_0041FFB4: pop esi
  loc_0041FFB5: mov fs:[00000000h], ecx
  loc_0041FFBC: pop ebx
  loc_0041FFBD: mov esp, ebp
  loc_0041FFBF: pop ebp
  loc_0041FFC0: retn 0004h
End Sub

Private Sub Form_Activate() '420450
  loc_00420450: push ebp
  loc_00420451: mov ebp, esp
  loc_00420453: sub esp, 0000000Ch
  loc_00420456: push 00401926h ; __vbaExceptHandler
  loc_0042045B: mov eax, fs:[00000000h]
  loc_00420461: push eax
  loc_00420462: mov fs:[00000000h], esp
  loc_00420469: sub esp, 00000014h
  loc_0042046C: push ebx
  loc_0042046D: push esi
  loc_0042046E: push edi
  loc_0042046F: mov var_C, esp
  loc_00420472: mov var_8, 00401398h
  loc_00420479: mov esi, Me
  loc_0042047C: mov eax, esi
  loc_0042047E: and eax, 00000001h
  loc_00420481: mov var_4, eax
  loc_00420484: and esi, FFFFFFFEh
  loc_00420487: push esi
  loc_00420488: mov Me, esi
  loc_0042048B: mov ecx, [esi]
  loc_0042048D: call [ecx+00000004h]
  loc_00420490: mov edx, [esi]
  loc_00420492: push esi
  loc_00420493: mov var_18, 00000000h
  loc_0042049A: call [edx+00000310h]
  loc_004204A0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_004204A6: push eax
  loc_004204A7: lea eax, var_18
  loc_004204AA: push eax
  loc_004204AB: call ebx
  loc_004204AD: mov edx, [00430010h]
  loc_004204B3: mov edi, eax
  loc_004204B5: push edx
  loc_004204B6: push edi
  loc_004204B7: mov ecx, [edi]
  loc_004204B9: call [ecx+000000A4h]
  loc_004204BF: test eax, eax
  loc_004204C1: fnclex
  loc_004204C3: jge 004204D7h
  loc_004204C5: push 000000A4h
  loc_004204CA: push 0040F02Ch
  loc_004204CF: push edi
  loc_004204D0: push eax
  loc_004204D1: call [00401030h] ; __vbaHresultCheckObj
  loc_004204D7: lea ecx, var_18
  loc_004204DA: call [00401114h] ; __vbaFreeObj
  loc_004204E0: mov eax, [esi]
  loc_004204E2: push esi
  loc_004204E3: call [eax+00000318h]
  loc_004204E9: lea ecx, var_18
  loc_004204EC: push eax
  loc_004204ED: push ecx
  loc_004204EE: call ebx
  loc_004204F0: mov edi, eax
  loc_004204F2: mov eax, [00430018h]
  loc_004204F7: push eax
  loc_004204F8: push edi
  loc_004204F9: mov edx, [edi]
  loc_004204FB: call [edx+000000A4h]
  loc_00420501: test eax, eax
  loc_00420503: fnclex
  loc_00420505: jge 00420519h
  loc_00420507: push 000000A4h
  loc_0042050C: push 0040F02Ch
  loc_00420511: push edi
  loc_00420512: push eax
  loc_00420513: call [00401030h] ; __vbaHresultCheckObj
  loc_00420519: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042051F: lea ecx, var_18
  loc_00420522: call edi
  loc_00420524: mov ecx, [esi]
  loc_00420526: push esi
  loc_00420527: call [ecx+00000310h]
  loc_0042052D: lea edx, var_18
  loc_00420530: push eax
  loc_00420531: push edx
  loc_00420532: call ebx
  loc_00420534: mov esi, eax
  loc_00420536: push esi
  loc_00420537: mov eax, [esi]
  loc_00420539: call [eax+00000204h]
  loc_0042053F: test eax, eax
  loc_00420541: fnclex
  loc_00420543: jge 00420557h
  loc_00420545: push 00000204h
  loc_0042054A: push 0040F02Ch
  loc_0042054F: push esi
  loc_00420550: push eax
  loc_00420551: call [00401030h] ; __vbaHresultCheckObj
  loc_00420557: lea ecx, var_18
  loc_0042055A: call edi
  loc_0042055C: mov var_4, 00000000h
  loc_00420563: push 00420575h
  loc_00420568: jmp 00420574h
  loc_0042056A: lea ecx, var_18
  loc_0042056D: call [00401114h] ; __vbaFreeObj
  loc_00420573: ret
  loc_00420574: ret
  loc_00420575: mov eax, Me
  loc_00420578: push eax
  loc_00420579: mov ecx, [eax]
  loc_0042057B: call [ecx+00000008h]
  loc_0042057E: mov eax, var_4
  loc_00420581: mov ecx, var_14
  loc_00420584: pop edi
  loc_00420585: pop esi
  loc_00420586: mov fs:[00000000h], ecx
  loc_0042058D: pop ebx
  loc_0042058E: mov esp, ebp
  loc_00420590: pop ebp
  loc_00420591: retn 0004h
End Sub
