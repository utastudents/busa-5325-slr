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


Private Sub cmdRestore_Click() '4124B0
  loc_004124B0: push ebp
  loc_004124B1: mov ebp, esp
  loc_004124B3: sub esp, 0000000Ch
  loc_004124B6: push 00401546h ; __vbaExceptHandler
  loc_004124BB: mov eax, fs:[00000000h]
  loc_004124C1: push eax
  loc_004124C2: mov fs:[00000000h], esp
  loc_004124C9: sub esp, 00000014h
  loc_004124CC: push ebx
  loc_004124CD: push esi
  loc_004124CE: push edi
  loc_004124CF: mov var_C, esp
  loc_004124D2: mov var_8, 00401488h
  loc_004124D9: mov esi, Me
  loc_004124DC: mov eax, esi
  loc_004124DE: and eax, 00000001h
  loc_004124E1: mov var_4, eax
  loc_004124E4: and esi, FFFFFFFEh
  loc_004124E7: push esi
  loc_004124E8: mov Me, esi
  loc_004124EB: mov ecx, [esi]
  loc_004124ED: call [ecx+00000004h]
  loc_004124F0: mov [0041505Ch], 42C80000h
  loc_004124FA: mov edx, [esi]
  loc_004124FC: xor edi, edi
  loc_004124FE: push esi
  loc_004124FF: mov var_18, edi
  loc_00412502: call [edx+0000030Ch]
  loc_00412508: push eax
  loc_00412509: lea eax, var_18
  loc_0041250C: push eax
  loc_0041250D: call [0040105Ch] ; __vbaObjSet
  loc_00412513: mov esi, eax
  loc_00412515: push 00403FB0h ; "100"
  loc_0041251A: push esi
  loc_0041251B: mov ecx, [esi]
  loc_0041251D: call [ecx+000000A4h]
  loc_00412523: cmp eax, edi
  loc_00412525: fnclex
  loc_00412527: jge 0041253Bh
  loc_00412529: push 000000A4h
  loc_0041252E: push 00403848h
  loc_00412533: push esi
  loc_00412534: push eax
  loc_00412535: call [00401040h] ; __vbaHresultCheckObj
  loc_0041253B: lea ecx, var_18
  loc_0041253E: call [00401180h] ; __vbaFreeObj
  loc_00412544: mov var_4, edi
  loc_00412547: fwait
  loc_00412548: push 0041255Ah
  loc_0041254D: jmp 00412559h
  loc_0041254F: lea ecx, var_18
  loc_00412552: call [00401180h] ; __vbaFreeObj
  loc_00412558: ret
  loc_00412559: ret
  loc_0041255A: mov eax, Me
  loc_0041255D: push eax
  loc_0041255E: mov edx, [eax]
  loc_00412560: call [edx+00000008h]
  loc_00412563: mov eax, var_4
  loc_00412566: mov ecx, var_14
  loc_00412569: pop edi
  loc_0041256A: pop esi
  loc_0041256B: mov fs:[00000000h], ecx
  loc_00412572: pop ebx
  loc_00412573: mov esp, ebp
  loc_00412575: pop ebp
  loc_00412576: retn 0004h
End Sub

Private Sub txtHypothesizedValue_KeyPress(KeyAscii As Integer) '412900
  loc_00412900: push ebp
  loc_00412901: mov ebp, esp
  loc_00412903: sub esp, 0000000Ch
  loc_00412906: push 00401546h ; __vbaExceptHandler
  loc_0041290B: mov eax, fs:[00000000h]
  loc_00412911: push eax
  loc_00412912: mov fs:[00000000h], esp
  loc_00412919: sub esp, 0000007Ch
  loc_0041291C: push ebx
  loc_0041291D: push esi
  loc_0041291E: push edi
  loc_0041291F: mov var_C, esp
  loc_00412922: mov var_8, 004014B8h
  loc_00412929: mov esi, Me
  loc_0041292C: mov eax, esi
  loc_0041292E: and eax, 00000001h
  loc_00412931: mov var_4, eax
  loc_00412934: and esi, FFFFFFFEh
  loc_00412937: push esi
  loc_00412938: mov Me, esi
  loc_0041293B: mov ecx, [esi]
  loc_0041293D: call [ecx+00000004h]
  loc_00412940: mov ebx, KeyAscii
  loc_00412943: lea edx, var_24
  loc_00412946: xor edi, edi
  loc_00412948: push ebx
  loc_00412949: push edx
  loc_0041294A: mov var_24, edi
  loc_0041294D: mov var_34, edi
  loc_00412950: mov var_44, edi
  loc_00412953: mov var_54, edi
  loc_00412956: mov var_64, edi
  loc_00412959: mov var_74, edi
  loc_0041295C: mov var_84, edi
  loc_00412962: call 00408F40h
  loc_00412967: lea eax, var_24
  loc_0041296A: push eax
  loc_0041296B: call [00401104h] ; __vbaI2Var
  loc_00412971: lea ecx, var_24
  loc_00412974: mov [ebx], ax
  loc_00412977: call [00401014h] ; __vbaFreeVar
  loc_0041297D: push 004035B0h ; "."
  loc_00412982: call [00401030h] ; rtcAnsiValueBstr
  loc_00412988: mov edx, [esi]
  loc_0041298A: xor ecx, ecx
  loc_0041298C: cmp [ebx], ax
  loc_0041298F: push esi
  loc_00412990: mov var_84, 0000000Bh
  loc_0041299A: setz cl
  loc_0041299D: neg ecx
  loc_0041299F: mov var_7C, cx
  loc_004129A3: call [edx+0000030Ch]
  loc_004129A9: mov var_1C, eax
  loc_004129AC: lea eax, var_84
  loc_004129B2: push eax
  loc_004129B3: lea ecx, var_24
  loc_004129B6: push 00000001h
  loc_004129B8: lea edx, var_64
  loc_004129BB: push ecx
  loc_004129BC: push edx
  loc_004129BD: lea eax, var_34
  loc_004129C0: push edi
  loc_004129C1: push eax
  loc_004129C2: mov var_24, 00000009h
  loc_004129C9: mov var_5C, 004035B0h ; "."
  loc_004129D0: mov var_64, 00000008h
  loc_004129D7: mov var_6C, edi
  loc_004129DA: mov var_74, 00008002h
  loc_004129E1: call [004010F4h] ; __vbaInStrVar
  loc_004129E7: lea ecx, var_74
  loc_004129EA: push eax
  loc_004129EB: lea edx, var_44
  loc_004129EE: push ecx
  loc_004129EF: push edx
  loc_004129F0: call [00401084h] ; __vbaVarCmpGt
  loc_004129F6: push eax
  loc_004129F7: lea eax, var_54
  loc_004129FA: push eax
  loc_004129FB: call [004010C4h] ; __vbaVarAnd
  loc_00412A01: push eax
  loc_00412A02: call [0040107Ch] ; __vbaBoolVarNull
  loc_00412A08: lea ecx, var_84
  loc_00412A0E: mov esi, eax
  loc_00412A10: lea edx, var_34
  loc_00412A13: push ecx
  loc_00412A14: lea eax, var_24
  loc_00412A17: push edx
  loc_00412A18: push eax
  loc_00412A19: push 00000003h
  loc_00412A1B: call [00401024h] ; __vbaFreeVarList
  loc_00412A21: add esp, 00000010h
  loc_00412A24: cmp si, di
  loc_00412A27: jz 00412A2Ch
  loc_00412A29: mov [ebx], di
  loc_00412A2C: mov var_4, edi
  loc_00412A2F: push 00412A53h
  loc_00412A34: jmp 00412A52h
  loc_00412A36: lea ecx, var_54
  loc_00412A39: lea edx, var_44
  loc_00412A3C: push ecx
  loc_00412A3D: lea eax, var_34
  loc_00412A40: push edx
  loc_00412A41: lea ecx, var_24
  loc_00412A44: push eax
  loc_00412A45: push ecx
  loc_00412A46: push 00000004h
  loc_00412A48: call [00401024h] ; __vbaFreeVarList
  loc_00412A4E: add esp, 00000014h
  loc_00412A51: ret
  loc_00412A52: ret
  loc_00412A53: mov eax, Me
  loc_00412A56: push eax
  loc_00412A57: mov edx, [eax]
  loc_00412A59: call [edx+00000008h]
  loc_00412A5C: mov eax, var_4
  loc_00412A5F: mov ecx, var_14
  loc_00412A62: pop edi
  loc_00412A63: pop esi
  loc_00412A64: mov fs:[00000000h], ecx
  loc_00412A6B: pop ebx
  loc_00412A6C: mov esp, ebp
  loc_00412A6E: pop ebp
  loc_00412A6F: retn 0008h
End Sub

Private Sub Form_Activate() '4127D0
  loc_004127D0: push ebp
  loc_004127D1: mov ebp, esp
  loc_004127D3: sub esp, 0000000Ch
  loc_004127D6: push 00401546h ; __vbaExceptHandler
  loc_004127DB: mov eax, fs:[00000000h]
  loc_004127E1: push eax
  loc_004127E2: mov fs:[00000000h], esp
  loc_004127E9: sub esp, 00000018h
  loc_004127EC: push ebx
  loc_004127ED: push esi
  loc_004127EE: push edi
  loc_004127EF: mov var_C, esp
  loc_004127F2: mov var_8, 004014A8h
  loc_004127F9: mov esi, Me
  loc_004127FC: mov eax, esi
  loc_004127FE: and eax, 00000001h
  loc_00412801: mov var_4, eax
  loc_00412804: and esi, FFFFFFFEh
  loc_00412807: push esi
  loc_00412808: mov Me, esi
  loc_0041280B: mov ecx, [esi]
  loc_0041280D: call [ecx+00000004h]
  loc_00412810: mov edx, [esi]
  loc_00412812: xor eax, eax
  loc_00412814: push esi
  loc_00412815: mov var_18, eax
  loc_00412818: mov var_1C, eax
  loc_0041281B: call [edx+0000030Ch]
  loc_00412821: push eax
  loc_00412822: lea eax, var_1C
  loc_00412825: push eax
  loc_00412826: call [0040105Ch] ; __vbaObjSet
  loc_0041282C: mov ecx, [0041505Ch]
  loc_00412832: mov edi, eax
  loc_00412834: push ecx
  loc_00412835: mov ebx, [edi]
  loc_00412837: call [004010A8h] ; __vbaStrR4
  loc_0041283D: mov edx, eax
  loc_0041283F: lea ecx, var_18
  loc_00412842: call [00401164h] ; __vbaStrMove
  loc_00412848: push eax
  loc_00412849: push edi
  loc_0041284A: call [ebx+000000A4h]
  loc_00412850: test eax, eax
  loc_00412852: fnclex
  loc_00412854: jge 00412868h
  loc_00412856: push 000000A4h
  loc_0041285B: push 00403848h
  loc_00412860: push edi
  loc_00412861: push eax
  loc_00412862: call [00401040h] ; __vbaHresultCheckObj
  loc_00412868: lea ecx, var_18
  loc_0041286B: call [00401184h] ; __vbaFreeStr
  loc_00412871: mov edi, [00401180h] ; __vbaFreeObj
  loc_00412877: lea ecx, var_1C
  loc_0041287A: call edi
  loc_0041287C: mov edx, [esi]
  loc_0041287E: push esi
  loc_0041287F: call [edx+0000030Ch]
  loc_00412885: push eax
  loc_00412886: lea eax, var_1C
  loc_00412889: push eax
  loc_0041288A: call [0040105Ch] ; __vbaObjSet
  loc_00412890: mov esi, eax
  loc_00412892: push esi
  loc_00412893: mov ecx, [esi]
  loc_00412895: call [ecx+00000204h]
  loc_0041289B: test eax, eax
  loc_0041289D: fnclex
  loc_0041289F: jge 004128B3h
  loc_004128A1: push 00000204h
  loc_004128A6: push 00403848h
  loc_004128AB: push esi
  loc_004128AC: push eax
  loc_004128AD: call [00401040h] ; __vbaHresultCheckObj
  loc_004128B3: lea ecx, var_1C
  loc_004128B6: call edi
  loc_004128B8: mov var_4, 00000000h
  loc_004128BF: fwait
  loc_004128C0: push 004128DBh
  loc_004128C5: jmp 004128DAh
  loc_004128C7: lea ecx, var_18
  loc_004128CA: call [00401184h] ; __vbaFreeStr
  loc_004128D0: lea ecx, var_1C
  loc_004128D3: call [00401180h] ; __vbaFreeObj
  loc_004128D9: ret
  loc_004128DA: ret
  loc_004128DB: mov eax, Me
  loc_004128DE: push eax
  loc_004128DF: mov edx, [eax]
  loc_004128E1: call [edx+00000008h]
  loc_004128E4: mov eax, var_4
  loc_004128E7: mov ecx, var_14
  loc_004128EA: pop edi
  loc_004128EB: pop esi
  loc_004128EC: mov fs:[00000000h], ecx
  loc_004128F3: pop ebx
  loc_004128F4: mov esp, ebp
  loc_004128F6: pop ebp
  loc_004128F7: retn 0004h
End Sub

Private Sub cmdClearChange_Click() '4123B0
  loc_004123B0: push ebp
  loc_004123B1: mov ebp, esp
  loc_004123B3: sub esp, 0000000Ch
  loc_004123B6: push 00401546h ; __vbaExceptHandler
  loc_004123BB: mov eax, fs:[00000000h]
  loc_004123C1: push eax
  loc_004123C2: mov fs:[00000000h], esp
  loc_004123C9: sub esp, 00000014h
  loc_004123CC: push ebx
  loc_004123CD: push esi
  loc_004123CE: push edi
  loc_004123CF: mov var_C, esp
  loc_004123D2: mov var_8, 00401478h
  loc_004123D9: mov esi, Me
  loc_004123DC: mov eax, esi
  loc_004123DE: and eax, 00000001h
  loc_004123E1: mov var_4, eax
  loc_004123E4: and esi, FFFFFFFEh
  loc_004123E7: push esi
  loc_004123E8: mov Me, esi
  loc_004123EB: mov ecx, [esi]
  loc_004123ED: call [ecx+00000004h]
  loc_004123F0: mov edx, [esi]
  loc_004123F2: push esi
  loc_004123F3: mov var_18, 00000000h
  loc_004123FA: call [edx+0000030Ch]
  loc_00412400: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_00412406: push eax
  loc_00412407: lea eax, var_18
  loc_0041240A: push eax
  loc_0041240B: call ebx
  loc_0041240D: mov edi, eax
  loc_0041240F: push 00403D64h
  loc_00412414: push edi
  loc_00412415: mov ecx, [edi]
  loc_00412417: call [ecx+000000A4h]
  loc_0041241D: test eax, eax
  loc_0041241F: fnclex
  loc_00412421: jge 00412435h
  loc_00412423: push 000000A4h
  loc_00412428: push 00403848h
  loc_0041242D: push edi
  loc_0041242E: push eax
  loc_0041242F: call [00401040h] ; __vbaHresultCheckObj
  loc_00412435: mov edi, [00401180h] ; __vbaFreeObj
  loc_0041243B: lea ecx, var_18
  loc_0041243E: call edi
  loc_00412440: mov edx, [esi]
  loc_00412442: push esi
  loc_00412443: call [edx+0000030Ch]
  loc_00412449: push eax
  loc_0041244A: lea eax, var_18
  loc_0041244D: push eax
  loc_0041244E: call ebx
  loc_00412450: mov esi, eax
  loc_00412452: push esi
  loc_00412453: mov ecx, [esi]
  loc_00412455: call [ecx+00000204h]
  loc_0041245B: test eax, eax
  loc_0041245D: fnclex
  loc_0041245F: jge 00412473h
  loc_00412461: push 00000204h
  loc_00412466: push 00403848h
  loc_0041246B: push esi
  loc_0041246C: push eax
  loc_0041246D: call [00401040h] ; __vbaHresultCheckObj
  loc_00412473: lea ecx, var_18
  loc_00412476: call edi
  loc_00412478: mov var_4, 00000000h
  loc_0041247F: push 00412491h
  loc_00412484: jmp 00412490h
  loc_00412486: lea ecx, var_18
  loc_00412489: call [00401180h] ; __vbaFreeObj
  loc_0041248F: ret
  loc_00412490: ret
  loc_00412491: mov eax, Me
  loc_00412494: push eax
  loc_00412495: mov edx, [eax]
  loc_00412497: call [edx+00000008h]
  loc_0041249A: mov eax, var_4
  loc_0041249D: mov ecx, var_14
  loc_004124A0: pop edi
  loc_004124A1: pop esi
  loc_004124A2: mov fs:[00000000h], ecx
  loc_004124A9: pop ebx
  loc_004124AA: mov esp, ebp
  loc_004124AC: pop ebp
  loc_004124AD: retn 0004h
End Sub

Private Sub cmdChange_Click() '412580
  loc_00412580: push ebp
  loc_00412581: mov ebp, esp
  loc_00412583: sub esp, 0000000Ch
  loc_00412586: push 00401546h ; __vbaExceptHandler
  loc_0041258B: mov eax, fs:[00000000h]
  loc_00412591: push eax
  loc_00412592: mov fs:[00000000h], esp
  loc_00412599: sub esp, 00000094h
  loc_0041259F: push ebx
  loc_004125A0: push esi
  loc_004125A1: push edi
  loc_004125A2: mov var_C, esp
  loc_004125A5: mov var_8, 00401498h
  loc_004125AC: mov esi, Me
  loc_004125AF: mov eax, esi
  loc_004125B1: and eax, 00000001h
  loc_004125B4: mov var_4, eax
  loc_004125B7: and esi, FFFFFFFEh
  loc_004125BA: push esi
  loc_004125BB: mov Me, esi
  loc_004125BE: mov ecx, [esi]
  loc_004125C0: call [ecx+00000004h]
  loc_004125C3: mov edx, [esi]
  loc_004125C5: xor edi, edi
  loc_004125C7: push esi
  loc_004125C8: mov var_18, edi
  loc_004125CB: mov var_28, edi
  loc_004125CE: mov var_38, edi
  loc_004125D1: mov var_48, edi
  loc_004125D4: mov var_58, edi
  loc_004125D7: mov var_68, edi
  loc_004125DA: mov var_78, edi
  loc_004125DD: call [edx+0000030Ch]
  loc_004125E3: mov var_20, eax
  loc_004125E6: lea eax, var_28
  loc_004125E9: lea ecx, var_38
  loc_004125EC: push eax
  loc_004125ED: push ecx
  loc_004125EE: mov var_28, 00000009h
  loc_004125F5: call [00401078h] ; rtcTrimVar
  loc_004125FB: lea edx, var_38
  loc_004125FE: push edx
  loc_004125FF: call [004010DCh] ; __vbaR4ErrVar
  loc_00412605: mov ebx, [00401024h] ; __vbaFreeVarList
  loc_0041260B: lea eax, var_38
  loc_0041260E: fstp real4 ptr [0041505Ch]
  loc_00412614: lea ecx, var_38
  loc_00412617: push eax
  loc_00412618: lea edx, var_28
  loc_0041261B: push ecx
  loc_0041261C: push edx
  loc_0041261D: push 00000003h
  loc_0041261F: call ebx
  loc_00412621: mov eax, [esi]
  loc_00412623: add esp, 00000010h
  loc_00412626: push esi
  loc_00412627: call [eax+0000030Ch]
  loc_0041262D: lea ecx, var_28
  loc_00412630: lea edx, var_38
  loc_00412633: push ecx
  loc_00412634: push edx
  loc_00412635: mov var_20, eax
  loc_00412638: mov var_28, 00000009h
  loc_0041263F: call [00401078h] ; rtcTrimVar
  loc_00412645: lea eax, var_38
  loc_00412648: lea ecx, var_68
  loc_0041264B: push eax
  loc_0041264C: push ecx
  loc_0041264D: mov var_60, 00403D64h
  loc_00412654: mov var_68, 00008008h
  loc_0041265B: call [0040109Ch] ; __vbaVarTstEq
  loc_00412661: mov var_9C, ax
  loc_00412668: lea edx, var_38
  loc_0041266B: lea eax, var_28
  loc_0041266E: push edx
  loc_0041266F: push eax
  loc_00412670: push 00000002h
  loc_00412672: call ebx
  loc_00412674: add esp, 0000000Ch
  loc_00412677: cmp var_9C, di
  loc_0041267E: jz 0041273Ch
  loc_00412684: mov ecx, [esi]
  loc_00412686: push esi
  loc_00412687: call [ecx+0000030Ch]
  loc_0041268D: lea edx, var_18
  loc_00412690: push eax
  loc_00412691: push edx
  loc_00412692: call [0040105Ch] ; __vbaObjSet
  loc_00412698: mov esi, eax
  loc_0041269A: push esi
  loc_0041269B: mov eax, [esi]
  loc_0041269D: call [eax+00000204h]
  loc_004126A3: cmp eax, edi
  loc_004126A5: fnclex
  loc_004126A7: jge 004126BBh
  loc_004126A9: push 00000204h
  loc_004126AE: push 00403848h
  loc_004126B3: push esi
  loc_004126B4: push eax
  loc_004126B5: call [00401040h] ; __vbaHresultCheckObj
  loc_004126BB: lea ecx, var_18
  loc_004126BE: call [00401180h] ; __vbaFreeObj
  loc_004126C4: mov esi, [00401150h] ; __vbaVarDup
  loc_004126CA: mov ecx, 80020004h
  loc_004126CF: mov var_50, ecx
  loc_004126D2: mov eax, 0000000Ah
  loc_004126D7: mov var_40, ecx
  loc_004126DA: lea edx, var_78
  loc_004126DD: lea ecx, var_38
  loc_004126E0: mov var_58, eax
  loc_004126E3: mov var_48, eax
  loc_004126E6: mov var_70, 00404044h ; "Check Value"
  loc_004126ED: mov var_78, 00000008h
  loc_004126F4: call __vbaVarDup
  loc_004126F6: lea edx, var_68
  loc_004126F9: lea ecx, var_28
  loc_004126FC: mov var_60, 00403FBCh ; "Please enter a value for the given value of X: HypothesizedValue"
  loc_00412703: mov var_68, 00000008h
  loc_0041270A: call __vbaVarDup
  loc_0041270C: lea ecx, var_58
  loc_0041270F: lea edx, var_48
  loc_00412712: push ecx
  loc_00412713: lea eax, var_38
  loc_00412716: push edx
  loc_00412717: push eax
  loc_00412718: lea ecx, var_28
  loc_0041271B: push edi
  loc_0041271C: push ecx
  loc_0041271D: call [00401060h] ; rtcMsgBox
  loc_00412723: lea edx, var_58
  loc_00412726: lea eax, var_48
  loc_00412729: push edx
  loc_0041272A: lea ecx, var_38
  loc_0041272D: push eax
  loc_0041272E: lea edx, var_28
  loc_00412731: push ecx
  loc_00412732: push edx
  loc_00412733: push 00000004h
  loc_00412735: call ebx
  loc_00412737: add esp, 00000014h
  loc_0041273A: jmp 0041277Bh
  loc_0041273C: cmp [004151F8h], edi
  loc_00412742: jnz 00412754h
  loc_00412744: push 004151F8h
  loc_00412749: push 00401C68h
  loc_0041274E: call [0040111Ch] ; __vbaNew2
  loc_00412754: mov esi, [004151F8h]
  loc_0041275A: push esi
  loc_0041275B: mov eax, [esi]
  loc_0041275D: call [eax+000002B4h]
  loc_00412763: cmp eax, edi
  loc_00412765: fnclex
  loc_00412767: jge 0041277Bh
  loc_00412769: push 000002B4h
  loc_0041276E: push 004039BCh
  loc_00412773: push esi
  loc_00412774: push eax
  loc_00412775: call [00401040h] ; __vbaHresultCheckObj
  loc_0041277B: mov var_4, edi
  loc_0041277E: fwait
  loc_0041277F: push 004127ACh
  loc_00412784: jmp 004127ABh
  loc_00412786: lea ecx, var_18
  loc_00412789: call [00401180h] ; __vbaFreeObj
  loc_0041278F: lea ecx, var_58
  loc_00412792: lea edx, var_48
  loc_00412795: push ecx
  loc_00412796: lea eax, var_38
  loc_00412799: push edx
  loc_0041279A: lea ecx, var_28
  loc_0041279D: push eax
  loc_0041279E: push ecx
  loc_0041279F: push 00000004h
  loc_004127A1: call [00401024h] ; __vbaFreeVarList
  loc_004127A7: add esp, 00000014h
  loc_004127AA: ret
  loc_004127AB: ret
  loc_004127AC: mov eax, Me
  loc_004127AF: push eax
  loc_004127B0: mov edx, [eax]
  loc_004127B2: call [edx+00000008h]
  loc_004127B5: mov eax, var_4
  loc_004127B8: mov ecx, var_14
  loc_004127BB: pop edi
  loc_004127BC: pop esi
  loc_004127BD: mov fs:[00000000h], ecx
  loc_004127C4: pop ebx
  loc_004127C5: mov esp, ebp
  loc_004127C7: pop ebp
  loc_004127C8: retn 0004h
End Sub
