VERSION 5.00
Begin VB.Form frmChangeXg
  Caption = "Change Xg"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  MaxButton = 0   'False
  MinButton = 0   'False
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 3645
  ClientHeight = 2235
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Given Value of X:"
    Left = 0
    Top = 0
    Width = 3615
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
    Begin VB.TextBox txtXg
      Left = 120
      Top = 600
      Width = 3375
      Height = 615
      Text = "100"
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
    Begin VB.CommandButton cmdRestore
      Caption = "Reset "
      Left = 2160
      Top = 1440
      Width = 1335
      Height = 615
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
      Left = 960
      Top = 1440
      Width = 1095
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
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 120
      Top = 1440
      Width = 735
      Height = 615
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

Attribute VB_Name = "frmChangeXg"


Private Sub cmdRestore_Click() '4232F0
  loc_004232F0: push ebp
  loc_004232F1: mov ebp, esp
  loc_004232F3: sub esp, 0000000Ch
  loc_004232F6: push 00401926h ; __vbaExceptHandler
  loc_004232FB: mov eax, fs:[00000000h]
  loc_00423301: push eax
  loc_00423302: mov fs:[00000000h], esp
  loc_00423309: sub esp, 00000014h
  loc_0042330C: push ebx
  loc_0042330D: push esi
  loc_0042330E: push edi
  loc_0042330F: mov var_C, esp
  loc_00423312: mov var_8, 004014D0h
  loc_00423319: mov esi, Me
  loc_0042331C: mov eax, esi
  loc_0042331E: and eax, 00000001h
  loc_00423321: mov var_4, eax
  loc_00423324: and esi, FFFFFFFEh
  loc_00423327: push esi
  loc_00423328: mov Me, esi
  loc_0042332B: mov ecx, [esi]
  loc_0042332D: call [ecx+00000004h]
  loc_00423330: xor edi, edi
  loc_00423332: push 00000064h
  loc_00423334: mov var_18, edi
  loc_00423337: call [00401000h] ; __vbaStrI2
  loc_0042333D: mov edx, eax
  loc_0042333F: mov ecx, 0043002Ch
  loc_00423344: call [004010FCh] ; __vbaStrMove
  loc_0042334A: mov edx, [esi]
  loc_0042334C: push esi
  loc_0042334D: call [edx+00000300h]
  loc_00423353: push eax
  loc_00423354: lea eax, var_18
  loc_00423357: push eax
  loc_00423358: call [0040103Ch] ; __vbaObjSet
  loc_0042335E: mov esi, eax
  loc_00423360: push 0040FCA0h ; "100"
  loc_00423365: push esi
  loc_00423366: mov ecx, [esi]
  loc_00423368: call [ecx+000000A4h]
  loc_0042336E: cmp eax, edi
  loc_00423370: fnclex
  loc_00423372: jge 00423386h
  loc_00423374: push 000000A4h
  loc_00423379: push 0040F02Ch
  loc_0042337E: push esi
  loc_0042337F: push eax
  loc_00423380: call [00401030h] ; __vbaHresultCheckObj
  loc_00423386: lea ecx, var_18
  loc_00423389: call [00401114h] ; __vbaFreeObj
  loc_0042338F: mov var_4, edi
  loc_00423392: push 004233A4h
  loc_00423397: jmp 004233A3h
  loc_00423399: lea ecx, var_18
  loc_0042339C: call [00401114h] ; __vbaFreeObj
  loc_004233A2: ret
  loc_004233A3: ret
  loc_004233A4: mov eax, Me
  loc_004233A7: push eax
  loc_004233A8: mov edx, [eax]
  loc_004233AA: call [edx+00000008h]
  loc_004233AD: mov eax, var_4
  loc_004233B0: mov ecx, var_14
  loc_004233B3: pop edi
  loc_004233B4: pop esi
  loc_004233B5: mov fs:[00000000h], ecx
  loc_004233BC: pop ebx
  loc_004233BD: mov esp, ebp
  loc_004233BF: pop ebp
  loc_004233C0: retn 0004h
End Sub

Private Sub cmdClearChange_Click() '4231F0
  loc_004231F0: push ebp
  loc_004231F1: mov ebp, esp
  loc_004231F3: sub esp, 0000000Ch
  loc_004231F6: push 00401926h ; __vbaExceptHandler
  loc_004231FB: mov eax, fs:[00000000h]
  loc_00423201: push eax
  loc_00423202: mov fs:[00000000h], esp
  loc_00423209: sub esp, 00000014h
  loc_0042320C: push ebx
  loc_0042320D: push esi
  loc_0042320E: push edi
  loc_0042320F: mov var_C, esp
  loc_00423212: mov var_8, 004014C0h
  loc_00423219: mov esi, Me
  loc_0042321C: mov eax, esi
  loc_0042321E: and eax, 00000001h
  loc_00423221: mov var_4, eax
  loc_00423224: and esi, FFFFFFFEh
  loc_00423227: push esi
  loc_00423228: mov Me, esi
  loc_0042322B: mov ecx, [esi]
  loc_0042322D: call [ecx+00000004h]
  loc_00423230: mov edx, [esi]
  loc_00423232: push esi
  loc_00423233: mov var_18, 00000000h
  loc_0042323A: call [edx+00000300h]
  loc_00423240: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_00423246: push eax
  loc_00423247: lea eax, var_18
  loc_0042324A: push eax
  loc_0042324B: call ebx
  loc_0042324D: mov edi, eax
  loc_0042324F: push 0040F040h
  loc_00423254: push edi
  loc_00423255: mov ecx, [edi]
  loc_00423257: call [ecx+000000A4h]
  loc_0042325D: test eax, eax
  loc_0042325F: fnclex
  loc_00423261: jge 00423275h
  loc_00423263: push 000000A4h
  loc_00423268: push 0040F02Ch
  loc_0042326D: push edi
  loc_0042326E: push eax
  loc_0042326F: call [00401030h] ; __vbaHresultCheckObj
  loc_00423275: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042327B: lea ecx, var_18
  loc_0042327E: call edi
  loc_00423280: mov edx, [esi]
  loc_00423282: push esi
  loc_00423283: call [edx+00000300h]
  loc_00423289: push eax
  loc_0042328A: lea eax, var_18
  loc_0042328D: push eax
  loc_0042328E: call ebx
  loc_00423290: mov esi, eax
  loc_00423292: push esi
  loc_00423293: mov ecx, [esi]
  loc_00423295: call [ecx+00000204h]
  loc_0042329B: test eax, eax
  loc_0042329D: fnclex
  loc_0042329F: jge 004232B3h
  loc_004232A1: push 00000204h
  loc_004232A6: push 0040F02Ch
  loc_004232AB: push esi
  loc_004232AC: push eax
  loc_004232AD: call [00401030h] ; __vbaHresultCheckObj
  loc_004232B3: lea ecx, var_18
  loc_004232B6: call edi
  loc_004232B8: mov var_4, 00000000h
  loc_004232BF: push 004232D1h
  loc_004232C4: jmp 004232D0h
  loc_004232C6: lea ecx, var_18
  loc_004232C9: call [00401114h] ; __vbaFreeObj
  loc_004232CF: ret
  loc_004232D0: ret
  loc_004232D1: mov eax, Me
  loc_004232D4: push eax
  loc_004232D5: mov edx, [eax]
  loc_004232D7: call [edx+00000008h]
  loc_004232DA: mov eax, var_4
  loc_004232DD: mov ecx, var_14
  loc_004232E0: pop edi
  loc_004232E1: pop esi
  loc_004232E2: mov fs:[00000000h], ecx
  loc_004232E9: pop ebx
  loc_004232EA: mov esp, ebp
  loc_004232EC: pop ebp
  loc_004232ED: retn 0004h
End Sub

Private Sub Form_Activate() '423830
  loc_00423830: push ebp
  loc_00423831: mov ebp, esp
  loc_00423833: sub esp, 0000000Ch
  loc_00423836: push 00401926h ; __vbaExceptHandler
  loc_0042383B: mov eax, fs:[00000000h]
  loc_00423841: push eax
  loc_00423842: mov fs:[00000000h], esp
  loc_00423849: sub esp, 00000014h
  loc_0042384C: push ebx
  loc_0042384D: push esi
  loc_0042384E: push edi
  loc_0042384F: mov var_C, esp
  loc_00423852: mov var_8, 004014F0h
  loc_00423859: mov esi, Me
  loc_0042385C: mov eax, esi
  loc_0042385E: and eax, 00000001h
  loc_00423861: mov var_4, eax
  loc_00423864: and esi, FFFFFFFEh
  loc_00423867: push esi
  loc_00423868: mov Me, esi
  loc_0042386B: mov ecx, [esi]
  loc_0042386D: call [ecx+00000004h]
  loc_00423870: mov edx, [esi]
  loc_00423872: push esi
  loc_00423873: mov var_18, 00000000h
  loc_0042387A: call [edx+00000300h]
  loc_00423880: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_00423886: push eax
  loc_00423887: lea eax, var_18
  loc_0042388A: push eax
  loc_0042388B: call ebx
  loc_0042388D: mov edx, [0043002Ch]
  loc_00423893: mov edi, eax
  loc_00423895: push edx
  loc_00423896: push edi
  loc_00423897: mov ecx, [edi]
  loc_00423899: call [ecx+000000A4h]
  loc_0042389F: test eax, eax
  loc_004238A1: fnclex
  loc_004238A3: jge 004238B7h
  loc_004238A5: push 000000A4h
  loc_004238AA: push 0040F02Ch
  loc_004238AF: push edi
  loc_004238B0: push eax
  loc_004238B1: call [00401030h] ; __vbaHresultCheckObj
  loc_004238B7: mov edi, [00401114h] ; __vbaFreeObj
  loc_004238BD: lea ecx, var_18
  loc_004238C0: call edi
  loc_004238C2: mov eax, [esi]
  loc_004238C4: push esi
  loc_004238C5: call [eax+00000300h]
  loc_004238CB: lea ecx, var_18
  loc_004238CE: push eax
  loc_004238CF: push ecx
  loc_004238D0: call ebx
  loc_004238D2: mov esi, eax
  loc_004238D4: push esi
  loc_004238D5: mov edx, [esi]
  loc_004238D7: call [edx+00000204h]
  loc_004238DD: test eax, eax
  loc_004238DF: fnclex
  loc_004238E1: jge 004238F5h
  loc_004238E3: push 00000204h
  loc_004238E8: push 0040F02Ch
  loc_004238ED: push esi
  loc_004238EE: push eax
  loc_004238EF: call [00401030h] ; __vbaHresultCheckObj
  loc_004238F5: lea ecx, var_18
  loc_004238F8: call edi
  loc_004238FA: mov var_4, 00000000h
  loc_00423901: push 00423913h
  loc_00423906: jmp 00423912h
  loc_00423908: lea ecx, var_18
  loc_0042390B: call [00401114h] ; __vbaFreeObj
  loc_00423911: ret
  loc_00423912: ret
  loc_00423913: mov eax, Me
  loc_00423916: push eax
  loc_00423917: mov ecx, [eax]
  loc_00423919: call [ecx+00000008h]
  loc_0042391C: mov eax, var_4
  loc_0042391F: mov ecx, var_14
  loc_00423922: pop edi
  loc_00423923: pop esi
  loc_00423924: mov fs:[00000000h], ecx
  loc_0042392B: pop ebx
  loc_0042392C: mov esp, ebp
  loc_0042392E: pop ebp
  loc_0042392F: retn 0004h
End Sub

Private Sub txtXg_KeyPress(KeyAscii As Integer) '423940
  loc_00423940: push ebp
  loc_00423941: mov ebp, esp
  loc_00423943: sub esp, 0000000Ch
  loc_00423946: push 00401926h ; __vbaExceptHandler
  loc_0042394B: mov eax, fs:[00000000h]
  loc_00423951: push eax
  loc_00423952: mov fs:[00000000h], esp
  loc_00423959: sub esp, 0000007Ch
  loc_0042395C: push ebx
  loc_0042395D: push esi
  loc_0042395E: push edi
  loc_0042395F: mov var_C, esp
  loc_00423962: mov var_8, 00401500h
  loc_00423969: mov esi, Me
  loc_0042396C: mov eax, esi
  loc_0042396E: and eax, 00000001h
  loc_00423971: mov var_4, eax
  loc_00423974: and esi, FFFFFFFEh
  loc_00423977: push esi
  loc_00423978: mov Me, esi
  loc_0042397B: mov ecx, [esi]
  loc_0042397D: call [ecx+00000004h]
  loc_00423980: mov ebx, KeyAscii
  loc_00423983: lea edx, var_24
  loc_00423986: xor edi, edi
  loc_00423988: push ebx
  loc_00423989: push edx
  loc_0042398A: mov var_24, edi
  loc_0042398D: mov var_34, edi
  loc_00423990: mov var_44, edi
  loc_00423993: mov var_54, edi
  loc_00423996: mov var_64, edi
  loc_00423999: mov var_74, edi
  loc_0042399C: mov var_84, edi
  loc_004239A2: call 0041A480h
  loc_004239A7: lea eax, var_24
  loc_004239AA: push eax
  loc_004239AB: call [004010C4h] ; __vbaI2Var
  loc_004239B1: lea ecx, var_24
  loc_004239B4: mov [ebx], ax
  loc_004239B7: call [00401010h] ; __vbaFreeVar
  loc_004239BD: push 0040DD3Ch ; "."
  loc_004239C2: call [00401024h] ; rtcAnsiValueBstr
  loc_004239C8: mov edx, [esi]
  loc_004239CA: xor ecx, ecx
  loc_004239CC: cmp [ebx], ax
  loc_004239CF: push esi
  loc_004239D0: mov var_84, 0000000Bh
  loc_004239DA: setz cl
  loc_004239DD: neg ecx
  loc_004239DF: mov var_7C, cx
  loc_004239E3: call [edx+00000300h]
  loc_004239E9: mov var_1C, eax
  loc_004239EC: lea eax, var_84
  loc_004239F2: push eax
  loc_004239F3: lea ecx, var_24
  loc_004239F6: push 00000001h
  loc_004239F8: lea edx, var_64
  loc_004239FB: push ecx
  loc_004239FC: push edx
  loc_004239FD: lea eax, var_34
  loc_00423A00: push edi
  loc_00423A01: push eax
  loc_00423A02: mov var_24, 00000009h
  loc_00423A09: mov var_5C, 0040DD3Ch ; "."
  loc_00423A10: mov var_64, 00000008h
  loc_00423A17: mov var_6C, edi
  loc_00423A1A: mov var_74, 00008002h
  loc_00423A21: call [004010B4h] ; __vbaInStrVar
  loc_00423A27: lea ecx, var_74
  loc_00423A2A: push eax
  loc_00423A2B: lea edx, var_44
  loc_00423A2E: push ecx
  loc_00423A2F: push edx
  loc_00423A30: call [00401060h] ; __vbaVarCmpGt
  loc_00423A36: push eax
  loc_00423A37: lea eax, var_54
  loc_00423A3A: push eax
  loc_00423A3B: call [00401094h] ; __vbaVarAnd
  loc_00423A41: push eax
  loc_00423A42: call [00401058h] ; __vbaBoolVarNull
  loc_00423A48: lea ecx, var_84
  loc_00423A4E: mov esi, eax
  loc_00423A50: lea edx, var_34
  loc_00423A53: push ecx
  loc_00423A54: lea eax, var_24
  loc_00423A57: push edx
  loc_00423A58: push eax
  loc_00423A59: push 00000003h
  loc_00423A5B: call [00401018h] ; __vbaFreeVarList
  loc_00423A61: add esp, 00000010h
  loc_00423A64: cmp si, di
  loc_00423A67: jz 00423A6Ch
  loc_00423A69: mov [ebx], di
  loc_00423A6C: mov var_4, edi
  loc_00423A6F: push 00423A93h
  loc_00423A74: jmp 00423A92h
  loc_00423A76: lea ecx, var_54
  loc_00423A79: lea edx, var_44
  loc_00423A7C: push ecx
  loc_00423A7D: lea eax, var_34
  loc_00423A80: push edx
  loc_00423A81: lea ecx, var_24
  loc_00423A84: push eax
  loc_00423A85: push ecx
  loc_00423A86: push 00000004h
  loc_00423A88: call [00401018h] ; __vbaFreeVarList
  loc_00423A8E: add esp, 00000014h
  loc_00423A91: ret
  loc_00423A92: ret
  loc_00423A93: mov eax, Me
  loc_00423A96: push eax
  loc_00423A97: mov edx, [eax]
  loc_00423A99: call [edx+00000008h]
  loc_00423A9C: mov eax, var_4
  loc_00423A9F: mov ecx, var_14
  loc_00423AA2: pop edi
  loc_00423AA3: pop esi
  loc_00423AA4: mov fs:[00000000h], ecx
  loc_00423AAB: pop ebx
  loc_00423AAC: mov esp, ebp
  loc_00423AAE: pop ebp
  loc_00423AAF: retn 0008h
End Sub

Private Sub cmdChange_Click() '4233D0
  loc_004233D0: push ebp
  loc_004233D1: mov ebp, esp
  loc_004233D3: sub esp, 0000000Ch
  loc_004233D6: push 00401926h ; __vbaExceptHandler
  loc_004233DB: mov eax, fs:[00000000h]
  loc_004233E1: push eax
  loc_004233E2: mov fs:[00000000h], esp
  loc_004233E9: sub esp, 000000BCh
  loc_004233EF: push ebx
  loc_004233F0: push esi
  loc_004233F1: push edi
  loc_004233F2: mov var_C, esp
  loc_004233F5: mov var_8, 004014E0h
  loc_004233FC: mov edi, Me
  loc_004233FF: mov eax, edi
  loc_00423401: and eax, 00000001h
  loc_00423404: mov var_4, eax
  loc_00423407: and edi, FFFFFFFEh
  loc_0042340A: push edi
  loc_0042340B: mov Me, edi
  loc_0042340E: mov ecx, [edi]
  loc_00423410: call [ecx+00000004h]
  loc_00423413: mov edx, [edi]
  loc_00423415: xor ebx, ebx
  loc_00423417: push edi
  loc_00423418: mov var_18, ebx
  loc_0042341B: mov var_1C, ebx
  loc_0042341E: mov var_20, ebx
  loc_00423421: mov var_24, ebx
  loc_00423424: mov var_28, ebx
  loc_00423427: mov var_2C, ebx
  loc_0042342A: mov var_30, ebx
  loc_0042342D: mov var_34, ebx
  loc_00423430: mov var_38, ebx
  loc_00423433: mov var_3C, ebx
  loc_00423436: mov var_40, ebx
  loc_00423439: mov var_50, ebx
  loc_0042343C: mov var_60, ebx
  loc_0042343F: mov var_70, ebx
  loc_00423442: mov var_80, ebx
  loc_00423445: mov var_90, ebx
  loc_0042344B: mov var_A0, ebx
  loc_00423451: call [edx+00000300h]
  loc_00423457: mov var_48, eax
  loc_0042345A: lea eax, var_50
  loc_0042345D: lea ecx, var_60
  loc_00423460: push eax
  loc_00423461: push ecx
  loc_00423462: mov var_50, 00000009h
  loc_00423469: call [00401050h] ; rtcTrimVar
  loc_0042346F: lea edx, var_60
  loc_00423472: push edx
  loc_00423473: call [004010A0h] ; __vbaR4ErrVar
  loc_00423479: push ecx
  loc_0042347A: fstp real4 ptr [esp]
  loc_0042347D: call [0040107Ch] ; __vbaStrR4
  loc_00423483: mov esi, [004010FCh] ; __vbaStrMove
  loc_00423489: mov edx, eax
  loc_0042348B: mov ecx, 0043002Ch
  loc_00423490: call __vbaStrMove
  loc_00423492: lea eax, var_60
  loc_00423495: lea ecx, var_60
  loc_00423498: push eax
  loc_00423499: lea edx, var_50
  loc_0042349C: push ecx
  loc_0042349D: push edx
  loc_0042349E: push 00000003h
  loc_004234A0: call [00401018h] ; __vbaFreeVarList
  loc_004234A6: mov eax, [edi]
  loc_004234A8: add esp, 00000010h
  loc_004234AB: push edi
  loc_004234AC: call [eax+00000300h]
  loc_004234B2: lea ecx, var_50
  loc_004234B5: lea edx, var_60
  loc_004234B8: push ecx
  loc_004234B9: push edx
  loc_004234BA: mov var_48, eax
  loc_004234BD: mov var_50, 00000009h
  loc_004234C4: call [00401050h] ; rtcTrimVar
  loc_004234CA: lea eax, var_60
  loc_004234CD: lea ecx, var_90
  loc_004234D3: push eax
  loc_004234D4: push ecx
  loc_004234D5: mov var_88, 0040F040h
  loc_004234DF: mov var_90, 00008008h
  loc_004234E9: call [00401070h] ; __vbaVarTstEq
  loc_004234EF: lea edx, var_60
  loc_004234F2: mov var_C4, ax
  loc_004234F9: push edx
  loc_004234FA: lea eax, var_50
  loc_004234FD: push eax
  loc_004234FE: push 00000002h
  loc_00423500: call [00401018h] ; __vbaFreeVarList
  loc_00423506: add esp, 0000000Ch
  loc_00423509: cmp var_C4, bx
  loc_00423510: jz 004235E4h
  loc_00423516: mov ecx, [edi]
  loc_00423518: push edi
  loc_00423519: call [ecx+00000300h]
  loc_0042351F: lea edx, var_40
  loc_00423522: push eax
  loc_00423523: push edx
  loc_00423524: call [0040103Ch] ; __vbaObjSet
  loc_0042352A: mov esi, eax
  loc_0042352C: push esi
  loc_0042352D: mov eax, [esi]
  loc_0042352F: call [eax+00000204h]
  loc_00423535: cmp eax, ebx
  loc_00423537: fnclex
  loc_00423539: jge 0042354Dh
  loc_0042353B: push 00000204h
  loc_00423540: push 0040F02Ch
  loc_00423545: push esi
  loc_00423546: push eax
  loc_00423547: call [00401030h] ; __vbaHresultCheckObj
  loc_0042354D: lea ecx, var_40
  loc_00423550: call [00401114h] ; __vbaFreeObj
  loc_00423556: mov esi, [004010F4h] ; __vbaVarDup
  loc_0042355C: mov ecx, 0000000Ah
  loc_00423561: mov eax, 80020004h
  loc_00423566: mov var_80, ecx
  loc_00423569: mov var_70, ecx
  loc_0042356C: mov edi, 00000008h
  loc_00423571: lea edx, var_A0
  loc_00423577: lea ecx, var_60
  loc_0042357A: mov var_78, eax
  loc_0042357D: mov var_68, eax
  loc_00423580: mov var_98, 0040FD34h ; "Check Value"
  loc_0042358A: mov var_A0, edi
  loc_00423590: call __vbaVarDup
  loc_00423592: lea edx, var_90
  loc_00423598: lea ecx, var_50
  loc_0042359B: mov var_88, 004105B8h ; "Please enter a value for the given value of X: Xg"
  loc_004235A5: mov var_90, edi
  loc_004235AB: call __vbaVarDup
  loc_004235AD: lea ecx, var_80
  loc_004235B0: lea edx, var_70
  loc_004235B3: push ecx
  loc_004235B4: lea eax, var_60
  loc_004235B7: push edx
  loc_004235B8: push eax
  loc_004235B9: lea ecx, var_50
  loc_004235BC: push ebx
  loc_004235BD: push ecx
  loc_004235BE: call [00401038h] ; rtcMsgBox
  loc_004235C4: lea edx, var_80
  loc_004235C7: lea eax, var_70
  loc_004235CA: push edx
  loc_004235CB: lea ecx, var_60
  loc_004235CE: push eax
  loc_004235CF: lea edx, var_50
  loc_004235D2: push ecx
  loc_004235D3: push edx
  loc_004235D4: push 00000004h
  loc_004235D6: call [00401018h] ; __vbaFreeVarList
  loc_004235DC: add esp, 00000014h
  loc_004235DF: jmp 004237A1h
  loc_004235E4: mov eax, [0043002Ch]
  loc_004235E9: push eax
  loc_004235EA: call [004010D0h] ; __vbaR8Str
  loc_004235F0: fmul st0, real4 ptr [00430060h]
  loc_004235F6: fadd st0, real4 ptr [00430074h]
  loc_004235FC: mov ecx, 0000000Ah
  loc_00423601: lea edx, var_90
  loc_00423607: fstp real4 ptr [00430038h]
  loc_0042360D: fnstsw ax
  loc_0042360F: test al, 0Dh
  loc_00423611: jnz 00423824h
  loc_00423617: mov eax, 80020004h
  loc_0042361C: mov var_80, ecx
  loc_0042361F: mov var_70, ecx
  loc_00423622: lea ecx, var_60
  loc_00423625: mov var_78, eax
  loc_00423628: mov var_68, eax
  loc_0042362B: mov var_88, 0040FF38h ; "Change in Yhat"
  loc_00423635: mov var_90, 00000008h
  loc_0042363F: call [004010F4h] ; __vbaVarDup
  loc_00423645: mov ecx, [0043002Ch]
  loc_0042364B: mov edi, [0040102Ch] ; __vbaStrCat
  loc_00423651: push 00410620h ; "With this value of "
  loc_00423656: push ecx
  loc_00423657: call edi
  loc_00423659: mov edx, eax
  loc_0042365B: lea ecx, var_18
  loc_0042365E: call __vbaVarDup
  loc_00423660: push eax
  loc_00423661: push 0040F028h
  loc_00423666: call edi
  loc_00423668: mov edx, eax
  loc_0042366A: lea ecx, var_1C
  loc_0042366D: call __vbaVarDup
  loc_0042366F: mov edx, [0043001Ch]
  loc_00423675: push eax
  loc_00423676: push edx
  loc_00423677: call edi
  loc_00423679: mov edx, eax
  loc_0042367B: lea ecx, var_20
  loc_0042367E: call __vbaVarDup
  loc_00423680: push eax
  loc_00423681: push 0041064Ch ; ", the estimated mean of "
  loc_00423686: call edi
  loc_00423688: mov edx, eax
  loc_0042368A: lea ecx, var_24
  loc_0042368D: call __vbaVarDup
  loc_0042368F: push eax
  loc_00423690: mov eax, [00430010h]
  loc_00423695: push eax
  loc_00423696: call edi
  loc_00423698: mov edx, eax
  loc_0042369A: lea ecx, var_28
  loc_0042369D: call __vbaVarDup
  loc_0042369F: push eax
  loc_004236A0: push 00410684h ; " is "
  loc_004236A5: call edi
  loc_004236A7: mov edx, eax
  loc_004236A9: lea ecx, var_2C
  loc_004236AC: call __vbaVarDup
  loc_004236AE: mov ecx, [00430038h]
  loc_004236B4: push eax
  loc_004236B5: push ecx
  loc_004236B6: call [0040107Ch] ; __vbaStrR4
  loc_004236BC: mov edx, eax
  loc_004236BE: lea ecx, var_30
  loc_004236C1: call __vbaVarDup
  loc_004236C3: push eax
  loc_004236C4: call edi
  loc_004236C6: mov edx, eax
  loc_004236C8: lea ecx, var_34
  loc_004236CB: call __vbaVarDup
  loc_004236CD: push eax
  loc_004236CE: push 0040F028h
  loc_004236D3: call edi
  loc_004236D5: mov edx, eax
  loc_004236D7: lea ecx, var_38
  loc_004236DA: call __vbaVarDup
  loc_004236DC: mov edx, [00430014h]
  loc_004236E2: push eax
  loc_004236E3: push edx
  loc_004236E4: call edi
  loc_004236E6: mov edx, eax
  loc_004236E8: lea ecx, var_3C
  loc_004236EB: call __vbaVarDup
  loc_004236ED: push eax
  loc_004236EE: push 0040DD3Ch ; "."
  loc_004236F3: call edi
  loc_004236F5: mov var_48, eax
  loc_004236F8: lea eax, var_80
  loc_004236FB: lea ecx, var_70
  loc_004236FE: push eax
  loc_004236FF: lea edx, var_60
  loc_00423702: push ecx
  loc_00423703: push edx
  loc_00423704: lea eax, var_50
  loc_00423707: push ebx
  loc_00423708: push eax
  loc_00423709: mov var_50, 00000008h
  loc_00423710: call [00401038h] ; rtcMsgBox
  loc_00423716: lea ecx, var_3C
  loc_00423719: lea edx, var_38
  loc_0042371C: push ecx
  loc_0042371D: lea eax, var_34
  loc_00423720: push edx
  loc_00423721: lea ecx, var_30
  loc_00423724: push eax
  loc_00423725: lea edx, var_2C
  loc_00423728: push ecx
  loc_00423729: lea eax, var_28
  loc_0042372C: push edx
  loc_0042372D: lea ecx, var_24
  loc_00423730: push eax
  loc_00423731: lea edx, var_20
  loc_00423734: push ecx
  loc_00423735: lea eax, var_1C
  loc_00423738: push edx
  loc_00423739: lea ecx, var_18
  loc_0042373C: push eax
  loc_0042373D: push ecx
  loc_0042373E: push 0000000Ah
  loc_00423740: call [004010E4h] ; __vbaFreeStrList
  loc_00423746: lea edx, var_80
  loc_00423749: lea eax, var_70
  loc_0042374C: push edx
  loc_0042374D: lea ecx, var_60
  loc_00423750: push eax
  loc_00423751: lea edx, var_50
  loc_00423754: push ecx
  loc_00423755: push edx
  loc_00423756: push 00000004h
  loc_00423758: call [00401018h] ; __vbaFreeVarList
  loc_0042375E: mov eax, [004301B4h]
  loc_00423763: add esp, 00000040h
  loc_00423766: cmp eax, ebx
  loc_00423768: jnz 0042377Ah
  loc_0042376A: push 004301B4h
  loc_0042376F: push 00402480h
  loc_00423774: call [004010D4h] ; __vbaNew2
  loc_0042377A: mov esi, [004301B4h]
  loc_00423780: push esi
  loc_00423781: mov eax, [esi]
  loc_00423783: call [eax+000002B4h]
  loc_00423789: cmp eax, ebx
  loc_0042378B: fnclex
  loc_0042378D: jge 004237A1h
  loc_0042378F: push 000002B4h
  loc_00423794: push 00410574h
  loc_00423799: push esi
  loc_0042379A: push eax
  loc_0042379B: call [00401030h] ; __vbaHresultCheckObj
  loc_004237A1: mov var_4, ebx
  loc_004237A4: fwait
  loc_004237A5: push 00423805h
  loc_004237AA: jmp 00423804h
  loc_004237AC: lea ecx, var_3C
  loc_004237AF: lea edx, var_38
  loc_004237B2: push ecx
  loc_004237B3: lea eax, var_34
  loc_004237B6: push edx
  loc_004237B7: lea ecx, var_30
  loc_004237BA: push eax
  loc_004237BB: lea edx, var_2C
  loc_004237BE: push ecx
  loc_004237BF: lea eax, var_28
  loc_004237C2: push edx
  loc_004237C3: lea ecx, var_24
  loc_004237C6: push eax
  loc_004237C7: lea edx, var_20
  loc_004237CA: push ecx
  loc_004237CB: lea eax, var_1C
  loc_004237CE: push edx
  loc_004237CF: lea ecx, var_18
  loc_004237D2: push eax
  loc_004237D3: push ecx
  loc_004237D4: push 0000000Ah
  loc_004237D6: call [004010E4h] ; __vbaFreeStrList
  loc_004237DC: add esp, 0000002Ch
  loc_004237DF: lea ecx, var_40
  loc_004237E2: call [00401114h] ; __vbaFreeObj
  loc_004237E8: lea edx, var_80
  loc_004237EB: lea eax, var_70
  loc_004237EE: push edx
  loc_004237EF: lea ecx, var_60
  loc_004237F2: push eax
  loc_004237F3: lea edx, var_50
  loc_004237F6: push ecx
  loc_004237F7: push edx
  loc_004237F8: push 00000004h
  loc_004237FA: call [00401018h] ; __vbaFreeVarList
  loc_00423800: add esp, 00000014h
  loc_00423803: ret
  loc_00423804: ret
  loc_00423805: mov eax, Me
  loc_00423808: push eax
  loc_00423809: mov ecx, [eax]
  loc_0042380B: call [ecx+00000008h]
  loc_0042380E: mov eax, var_4
  loc_00423811: mov ecx, var_14
  loc_00423814: pop edi
  loc_00423815: pop esi
  loc_00423816: mov fs:[00000000h], ecx
  loc_0042381D: pop ebx
  loc_0042381E: mov esp, ebp
  loc_00423820: pop ebp
  loc_00423821: retn 0004h
End Sub
