VERSION 5.00
Begin VB.Form Form1
  Caption = "Simple Linear Regression"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 735
  ClientWidth = 8400
  ClientHeight = 4650
  StartUpPosition = 3 'Windows Default
  Begin VB.Frame fraExample
    Caption = "Example"
    Left = 840
    Top = 1800
    Width = 6615
    Height = 1095
    TabIndex = 16
    Begin VB.Label lblExample
      Left = 240
      Top = 240
      Width = 6255
      Height = 735
      TabIndex = 17
    End
  End
  Begin VB.Frame fraSituation
    Caption = "Situation"
    Left = 840
    Top = 600
    Width = 6615
    Height = 975
    TabIndex = 14
    Begin VB.Label lblSituation
      Left = 240
      Top = 240
      Width = 6255
      Height = 495
      TabIndex = 15
    End
  End
  Begin VB.Frame fraChange
    Caption = "Change Variable Descriptions"
    Left = 6240
    Top = 3720
    Width = 375
    Height = 135
    Visible = 0   'False
    TabIndex = 0
    Begin VB.CommandButton cmdRestore
      Caption = "Reset Defaults"
      Left = 3480
      Top = 3000
      Width = 1455
      Height = 255
      TabIndex = 13
    End
    Begin VB.CommandButton cmdClearChange
      Caption = "Clear"
      Left = 1920
      Top = 3000
      Width = 1335
      Height = 255
      TabIndex = 12
    End
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 360
      Top = 3000
      Width = 1215
      Height = 255
      TabIndex = 11
    End
    Begin VB.TextBox txtXUnits
      Left = 1680
      Top = 2640
      Width = 3375
      Height = 285
      TabIndex = 10
    End
    Begin VB.TextBox txtXName
      Left = 1680
      Top = 2160
      Width = 3375
      Height = 285
      TabIndex = 8
    End
    Begin VB.TextBox txtYUnits
      Left = 1680
      Top = 1320
      Width = 3375
      Height = 285
      TabIndex = 6
    End
    Begin VB.TextBox txtYName
      Left = 1680
      Top = 810
      Width = 3375
      Height = 285
      TabIndex = 4
    End
    Begin VB.Label Label6
      Caption = "Units:"
      Left = 840
      Top = 2640
      Width = 735
      Height = 255
      TabIndex = 9
    End
    Begin VB.Label Label5
      Caption = "Name:"
      Left = 840
      Top = 2160
      Width = 735
      Height = 255
      TabIndex = 7
    End
    Begin VB.Label Label4
      Caption = "Units:"
      Left = 840
      Top = 1320
      Width = 735
      Height = 255
      TabIndex = 5
    End
    Begin VB.Label Label3
      Caption = "Name:"
      Left = 840
      Top = 840
      Width = 615
      Height = 255
      TabIndex = 3
    End
    Begin VB.Label Label2
      Caption = "Independent"
      Left = 240
      Top = 1800
      Width = 1095
      Height = 255
      TabIndex = 2
    End
    Begin VB.Label Label1
      Caption = "Dependent"
      Left = 240
      Top = 480
      Width = 1095
      Height = 255
      TabIndex = 1
    End
  End
  Begin VB.Menu mnuVariables
    Caption = "&Variables"
  End
End

Attribute VB_Name = "Form1"


Private Sub mnuVariables_Click() '42E820
  loc_0042E820: push ebp
  loc_0042E821: mov ebp, esp
  loc_0042E823: sub esp, 0000000Ch
  loc_0042E826: push 00401926h ; __vbaExceptHandler
  loc_0042E82B: mov eax, fs:[00000000h]
  loc_0042E831: push eax
  loc_0042E832: mov fs:[00000000h], esp
  loc_0042E839: sub esp, 00000018h
  loc_0042E83C: push ebx
  loc_0042E83D: push esi
  loc_0042E83E: push edi
  loc_0042E83F: mov var_C, esp
  loc_0042E842: mov var_8, 00401868h
  loc_0042E849: mov esi, Me
  loc_0042E84C: mov eax, esi
  loc_0042E84E: and eax, 00000001h
  loc_0042E851: mov var_4, eax
  loc_0042E854: and esi, FFFFFFFEh
  loc_0042E857: push esi
  loc_0042E858: mov Me, esi
  loc_0042E85B: mov ecx, [esi]
  loc_0042E85D: call [ecx+00000004h]
  loc_0042E860: mov edx, [esi]
  loc_0042E862: xor edi, edi
  loc_0042E864: push esi
  loc_0042E865: mov var_18, edi
  loc_0042E868: mov var_24, edi
  loc_0042E86B: call [edx+0000030Ch]
  loc_0042E871: push eax
  loc_0042E872: lea eax, var_24
  loc_0042E875: push eax
  loc_0042E876: call [0040103Ch] ; __vbaObjSet
  loc_0042E87C: mov eax, var_24
  loc_0042E87F: push FFFFFFFFh
  loc_0042E881: push eax
  loc_0042E882: mov ecx, [eax]
  loc_0042E884: call [ecx+0000009Ch]
  loc_0042E88A: cmp eax, edi
  loc_0042E88C: fnclex
  loc_0042E88E: jge 0042E8A9h
  loc_0042E890: mov edx, var_24
  loc_0042E893: mov ebx, [00401030h] ; __vbaHresultCheckObj
  loc_0042E899: push 0000009Ch
  loc_0042E89E: push 0040E728h
  loc_0042E8A3: push edx
  loc_0042E8A4: push eax
  loc_0042E8A5: call ebx
  loc_0042E8A7: jmp 0042E8AFh
  loc_0042E8A9: mov ebx, [00401030h] ; __vbaHresultCheckObj
  loc_0042E8AF: mov eax, var_24
  loc_0042E8B2: push 43F00000h
  loc_0042E8B7: push eax
  loc_0042E8B8: mov ecx, [eax]
  loc_0042E8BA: call [ecx+0000007Ch]
  loc_0042E8BD: cmp eax, edi
  loc_0042E8BF: fnclex
  loc_0042E8C1: jge 0042E8D1h
  loc_0042E8C3: mov edx, var_24
  loc_0042E8C6: push 0000007Ch
  loc_0042E8C8: push 0040E728h
  loc_0042E8CD: push edx
  loc_0042E8CE: push eax
  loc_0042E8CF: call ebx
  loc_0042E8D1: mov eax, var_24
  loc_0042E8D4: push 4552F000h
  loc_0042E8D9: push eax
  loc_0042E8DA: mov ecx, [eax]
  loc_0042E8DC: call [ecx+0000008Ch]
  loc_0042E8E2: cmp eax, edi
  loc_0042E8E4: fnclex
  loc_0042E8E6: jge 0042E8F9h
  loc_0042E8E8: mov edx, var_24
  loc_0042E8EB: push 0000008Ch
  loc_0042E8F0: push 0040E728h
  loc_0042E8F5: push edx
  loc_0042E8F6: push eax
  loc_0042E8F7: call ebx
  loc_0042E8F9: mov eax, var_24
  loc_0042E8FC: push 44A50000h
  loc_0042E901: push eax
  loc_0042E902: mov ecx, [eax]
  loc_0042E904: call [ecx+00000074h]
  loc_0042E907: cmp eax, edi
  loc_0042E909: fnclex
  loc_0042E90B: jge 0042E91Bh
  loc_0042E90D: mov edx, var_24
  loc_0042E910: push 00000074h
  loc_0042E912: push 0040E728h
  loc_0042E917: push edx
  loc_0042E918: push eax
  loc_0042E919: call ebx
  loc_0042E91B: mov eax, var_24
  loc_0042E91E: push 45A57800h
  loc_0042E923: push eax
  loc_0042E924: mov ecx, [eax]
  loc_0042E926: call [ecx+00000084h]
  loc_0042E92C: cmp eax, edi
  loc_0042E92E: fnclex
  loc_0042E930: jge 0042E943h
  loc_0042E932: mov edx, var_24
  loc_0042E935: push 00000084h
  loc_0042E93A: push 0040E728h
  loc_0042E93F: push edx
  loc_0042E940: push eax
  loc_0042E941: call ebx
  loc_0042E943: lea eax, var_24
  loc_0042E946: push edi
  loc_0042E947: push eax
  loc_0042E948: call [00401044h] ; __vbaObjSetAddref
  loc_0042E94E: mov ecx, [esi]
  loc_0042E950: push esi
  loc_0042E951: call [ecx+00000304h]
  loc_0042E957: lea edx, var_18
  loc_0042E95A: push eax
  loc_0042E95B: push edx
  loc_0042E95C: call [0040103Ch] ; __vbaObjSet
  loc_0042E962: mov edi, eax
  loc_0042E964: push 00000000h
  loc_0042E966: push edi
  loc_0042E967: mov eax, [edi]
  loc_0042E969: call [eax+0000009Ch]
  loc_0042E96F: test eax, eax
  loc_0042E971: fnclex
  loc_0042E973: jge 0042E983h
  loc_0042E975: push 0000009Ch
  loc_0042E97A: push 0040E728h
  loc_0042E97F: push edi
  loc_0042E980: push eax
  loc_0042E981: call ebx
  loc_0042E983: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042E989: lea ecx, var_18
  loc_0042E98C: call edi
  loc_0042E98E: mov ecx, [esi]
  loc_0042E990: push esi
  loc_0042E991: call [ecx+000002FCh]
  loc_0042E997: lea edx, var_18
  loc_0042E99A: push eax
  loc_0042E99B: push edx
  loc_0042E99C: call [0040103Ch] ; __vbaObjSet
  loc_0042E9A2: mov esi, eax
  loc_0042E9A4: push 00000000h
  loc_0042E9A6: push esi
  loc_0042E9A7: mov eax, [esi]
  loc_0042E9A9: call [eax+0000009Ch]
  loc_0042E9AF: test eax, eax
  loc_0042E9B1: fnclex
  loc_0042E9B3: jge 0042E9C3h
  loc_0042E9B5: push 0000009Ch
  loc_0042E9BA: push 0040E728h
  loc_0042E9BF: push esi
  loc_0042E9C0: push eax
  loc_0042E9C1: call ebx
  loc_0042E9C3: lea ecx, var_18
  loc_0042E9C6: call edi
  loc_0042E9C8: mov var_4, 00000000h
  loc_0042E9CF: fwait
  loc_0042E9D0: push 0042E9EBh
  loc_0042E9D5: jmp 0042E9E1h
  loc_0042E9D7: lea ecx, var_18
  loc_0042E9DA: call [00401114h] ; __vbaFreeObj
  loc_0042E9E0: ret
  loc_0042E9E1: lea ecx, var_24
  loc_0042E9E4: call [00401114h] ; __vbaFreeObj
  loc_0042E9EA: ret
  loc_0042E9EB: mov eax, Me
  loc_0042E9EE: push eax
  loc_0042E9EF: mov ecx, [eax]
  loc_0042E9F1: call [ecx+00000008h]
  loc_0042E9F4: mov eax, var_4
  loc_0042E9F7: mov ecx, var_14
  loc_0042E9FA: pop edi
  loc_0042E9FB: pop esi
  loc_0042E9FC: mov fs:[00000000h], ecx
  loc_0042EA03: pop ebx
  loc_0042EA04: mov esp, ebp
  loc_0042EA06: pop ebp
  loc_0042EA07: retn 0004h
End Sub

Private Sub cmdClearChange_Click() '42E4D0
  loc_0042E4D0: push ebp
  loc_0042E4D1: mov ebp, esp
  loc_0042E4D3: sub esp, 0000000Ch
  loc_0042E4D6: push 00401926h ; __vbaExceptHandler
  loc_0042E4DB: mov eax, fs:[00000000h]
  loc_0042E4E1: push eax
  loc_0042E4E2: mov fs:[00000000h], esp
  loc_0042E4E9: sub esp, 00000014h
  loc_0042E4EC: push ebx
  loc_0042E4ED: push esi
  loc_0042E4EE: push edi
  loc_0042E4EF: mov var_C, esp
  loc_0042E4F2: mov var_8, 00401848h
  loc_0042E4F9: mov esi, Me
  loc_0042E4FC: mov eax, esi
  loc_0042E4FE: and eax, 00000001h
  loc_0042E501: mov var_4, eax
  loc_0042E504: and esi, FFFFFFFEh
  loc_0042E507: push esi
  loc_0042E508: mov Me, esi
  loc_0042E50B: mov ecx, [esi]
  loc_0042E50D: call [ecx+00000004h]
  loc_0042E510: mov edx, [esi]
  loc_0042E512: push esi
  loc_0042E513: mov var_18, 00000000h
  loc_0042E51A: call [edx+00000328h]
  loc_0042E520: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0042E526: push eax
  loc_0042E527: lea eax, var_18
  loc_0042E52A: push eax
  loc_0042E52B: call ebx
  loc_0042E52D: mov edi, eax
  loc_0042E52F: push 0040F040h
  loc_0042E534: push edi
  loc_0042E535: mov ecx, [edi]
  loc_0042E537: call [ecx+000000A4h]
  loc_0042E53D: test eax, eax
  loc_0042E53F: fnclex
  loc_0042E541: jge 0042E555h
  loc_0042E543: push 000000A4h
  loc_0042E548: push 0040F02Ch
  loc_0042E54D: push edi
  loc_0042E54E: push eax
  loc_0042E54F: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E555: lea ecx, var_18
  loc_0042E558: call [00401114h] ; __vbaFreeObj
  loc_0042E55E: mov edx, [esi]
  loc_0042E560: push esi
  loc_0042E561: call [edx+00000324h]
  loc_0042E567: push eax
  loc_0042E568: lea eax, var_18
  loc_0042E56B: push eax
  loc_0042E56C: call ebx
  loc_0042E56E: mov edi, eax
  loc_0042E570: push 0040F040h
  loc_0042E575: push edi
  loc_0042E576: mov ecx, [edi]
  loc_0042E578: call [ecx+000000A4h]
  loc_0042E57E: test eax, eax
  loc_0042E580: fnclex
  loc_0042E582: jge 0042E596h
  loc_0042E584: push 000000A4h
  loc_0042E589: push 0040F02Ch
  loc_0042E58E: push edi
  loc_0042E58F: push eax
  loc_0042E590: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E596: lea ecx, var_18
  loc_0042E599: call [00401114h] ; __vbaFreeObj
  loc_0042E59F: mov edx, [esi]
  loc_0042E5A1: push esi
  loc_0042E5A2: call [edx+00000320h]
  loc_0042E5A8: push eax
  loc_0042E5A9: lea eax, var_18
  loc_0042E5AC: push eax
  loc_0042E5AD: call ebx
  loc_0042E5AF: mov edi, eax
  loc_0042E5B1: push 0040F040h
  loc_0042E5B6: push edi
  loc_0042E5B7: mov ecx, [edi]
  loc_0042E5B9: call [ecx+000000A4h]
  loc_0042E5BF: test eax, eax
  loc_0042E5C1: fnclex
  loc_0042E5C3: jge 0042E5D7h
  loc_0042E5C5: push 000000A4h
  loc_0042E5CA: push 0040F02Ch
  loc_0042E5CF: push edi
  loc_0042E5D0: push eax
  loc_0042E5D1: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E5D7: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042E5DD: lea ecx, var_18
  loc_0042E5E0: call edi
  loc_0042E5E2: mov edx, [esi]
  loc_0042E5E4: push esi
  loc_0042E5E5: call [edx+0000031Ch]
  loc_0042E5EB: push eax
  loc_0042E5EC: lea eax, var_18
  loc_0042E5EF: push eax
  loc_0042E5F0: call ebx
  loc_0042E5F2: mov esi, eax
  loc_0042E5F4: push 0040F040h
  loc_0042E5F9: push esi
  loc_0042E5FA: mov ecx, [esi]
  loc_0042E5FC: call [ecx+000000A4h]
  loc_0042E602: test eax, eax
  loc_0042E604: fnclex
  loc_0042E606: jge 0042E61Ah
  loc_0042E608: push 000000A4h
  loc_0042E60D: push 0040F02Ch
  loc_0042E612: push esi
  loc_0042E613: push eax
  loc_0042E614: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E61A: lea ecx, var_18
  loc_0042E61D: call edi
  loc_0042E61F: mov var_4, 00000000h
  loc_0042E626: push 0042E638h
  loc_0042E62B: jmp 0042E637h
  loc_0042E62D: lea ecx, var_18
  loc_0042E630: call [00401114h] ; __vbaFreeObj
  loc_0042E636: ret
  loc_0042E637: ret
  loc_0042E638: mov eax, Me
  loc_0042E63B: push eax
  loc_0042E63C: mov edx, [eax]
  loc_0042E63E: call [edx+00000008h]
  loc_0042E641: mov eax, var_4
  loc_0042E644: mov ecx, var_14
  loc_0042E647: pop edi
  loc_0042E648: pop esi
  loc_0042E649: mov fs:[00000000h], ecx
  loc_0042E650: pop ebx
  loc_0042E651: mov esp, ebp
  loc_0042E653: pop ebp
  loc_0042E654: retn 0004h
End Sub

Private Sub Form_Load() '42D9E0
  loc_0042D9E0: push ebp
  loc_0042D9E1: mov ebp, esp
  loc_0042D9E3: sub esp, 0000000Ch
  loc_0042D9E6: push 00401926h ; __vbaExceptHandler
  loc_0042D9EB: mov eax, fs:[00000000h]
  loc_0042D9F1: push eax
  loc_0042D9F2: mov fs:[00000000h], esp
  loc_0042D9F9: sub esp, 00000014h
  loc_0042D9FC: push ebx
  loc_0042D9FD: push esi
  loc_0042D9FE: push edi
  loc_0042D9FF: mov var_C, esp
  loc_0042DA02: mov var_8, 00401818h
  loc_0042DA09: mov esi, Me
  loc_0042DA0C: mov eax, esi
  loc_0042DA0E: and eax, 00000001h
  loc_0042DA11: mov var_4, eax
  loc_0042DA14: and esi, FFFFFFFEh
  loc_0042DA17: push esi
  loc_0042DA18: mov Me, esi
  loc_0042DA1B: mov ecx, [esi]
  loc_0042DA1D: call [ecx+00000004h]
  loc_0042DA20: mov edi, [004010E0h] ; __vbaStrCopy
  loc_0042DA26: xor ebx, ebx
  loc_0042DA28: mov edx, 0040DB10h ; " Net Income"
  loc_0042DA2D: lea ecx, [esi+00000034h]
  loc_0042DA30: mov var_18, ebx
  loc_0042DA33: call edi
  loc_0042DA35: mov edx, 004102A0h ; " millions of dollars"
  loc_0042DA3A: lea ecx, [esi+00000038h]
  loc_0042DA3D: call edi
  loc_0042DA3F: mov edx, 0040DB54h ; " Sales"
  loc_0042DA44: lea ecx, [esi+0000003Ch]
  loc_0042DA47: call edi
  loc_0042DA49: mov edx, 004102A0h ; " millions of dollars"
  loc_0042DA4E: lea ecx, [esi+00000040h]
  loc_0042DA51: call edi
  loc_0042DA53: mov edx, [esi]
  loc_0042DA55: push esi
  loc_0042DA56: call [edx+00000308h]
  loc_0042DA5C: push eax
  loc_0042DA5D: lea eax, var_18
  loc_0042DA60: push eax
  loc_0042DA61: call [0040103Ch] ; __vbaObjSet
  loc_0042DA67: mov edi, eax
  loc_0042DA69: push 00412D14h ; "You wish to determine if X has an effect of Y and then use this relationship for inferences about Y."
  loc_0042DA6E: push edi
  loc_0042DA6F: mov ecx, [edi]
  loc_0042DA71: call [ecx+00000054h]
  loc_0042DA74: cmp eax, ebx
  loc_0042DA76: fnclex
  loc_0042DA78: jge 0042DA89h
  loc_0042DA7A: push 00000054h
  loc_0042DA7C: push 0040E390h
  loc_0042DA81: push edi
  loc_0042DA82: push eax
  loc_0042DA83: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DA89: lea ecx, var_18
  loc_0042DA8C: call [00401114h] ; __vbaFreeObj
  loc_0042DA92: mov edx, [esi]
  loc_0042DA94: push esi
  loc_0042DA95: call [edx+00000710h]
  loc_0042DA9B: mov eax, [esi]
  loc_0042DA9D: push esi
  loc_0042DA9E: call [eax+00000708h]
  loc_0042DAA4: cmp eax, ebx
  loc_0042DAA6: jge 0042DABAh
  loc_0042DAA8: push 00000708h
  loc_0042DAAD: push 00412CE0h
  loc_0042DAB2: push esi
  loc_0042DAB3: push eax
  loc_0042DAB4: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DABA: mov var_4, ebx
  loc_0042DABD: push 0042DACFh
  loc_0042DAC2: jmp 0042DACEh
  loc_0042DAC4: lea ecx, var_18
  loc_0042DAC7: call [00401114h] ; __vbaFreeObj
  loc_0042DACD: ret
  loc_0042DACE: ret
  loc_0042DACF: mov eax, Me
  loc_0042DAD2: push eax
  loc_0042DAD3: mov ecx, [eax]
  loc_0042DAD5: call [ecx+00000008h]
  loc_0042DAD8: mov eax, var_4
  loc_0042DADB: mov ecx, var_14
  loc_0042DADE: pop edi
  loc_0042DADF: pop esi
  loc_0042DAE0: mov fs:[00000000h], ecx
  loc_0042DAE7: pop ebx
  loc_0042DAE8: mov esp, ebp
  loc_0042DAEA: pop ebp
  loc_0042DAEB: retn 0004h
End Sub

Private Sub cmdChange_Click() '42DB90
  loc_0042DB90: push ebp
  loc_0042DB91: mov ebp, esp
  loc_0042DB93: sub esp, 0000000Ch
  loc_0042DB96: push 00401926h ; __vbaExceptHandler
  loc_0042DB9B: mov eax, fs:[00000000h]
  loc_0042DBA1: push eax
  loc_0042DBA2: mov fs:[00000000h], esp
  loc_0042DBA9: sub esp, 000000B0h
  loc_0042DBAF: push ebx
  loc_0042DBB0: push esi
  loc_0042DBB1: push edi
  loc_0042DBB2: mov var_C, esp
  loc_0042DBB5: mov var_8, 00401838h
  loc_0042DBBC: mov esi, Me
  loc_0042DBBF: mov eax, esi
  loc_0042DBC1: and eax, 00000001h
  loc_0042DBC4: mov var_4, eax
  loc_0042DBC7: and esi, FFFFFFFEh
  loc_0042DBCA: push esi
  loc_0042DBCB: mov Me, esi
  loc_0042DBCE: mov ecx, [esi]
  loc_0042DBD0: call [ecx+00000004h]
  loc_0042DBD3: mov edx, [esi]
  loc_0042DBD5: xor ebx, ebx
  loc_0042DBD7: mov var_70, ebx
  loc_0042DBDA: push esi
  loc_0042DBDB: mov var_18, ebx
  loc_0042DBDE: mov var_1C, ebx
  loc_0042DBE1: mov var_20, ebx
  loc_0042DBE4: mov var_30, ebx
  loc_0042DBE7: mov var_40, ebx
  loc_0042DBEA: mov var_50, ebx
  loc_0042DBED: mov var_60, ebx
  loc_0042DBF0: mov var_68, 0040F028h
  loc_0042DBF7: mov var_70, 00000008h
  loc_0042DBFE: call [edx+00000328h]
  loc_0042DC04: mov edi, [0040103Ch] ; __vbaObjSet
  loc_0042DC0A: push eax
  loc_0042DC0B: lea eax, var_20
  loc_0042DC0E: push eax
  loc_0042DC0F: call edi
  loc_0042DC11: mov ecx, [eax]
  loc_0042DC13: lea edx, var_18
  loc_0042DC16: push edx
  loc_0042DC17: push eax
  loc_0042DC18: mov var_A4, eax
  loc_0042DC1E: call [ecx+000000A0h]
  loc_0042DC24: cmp eax, ebx
  loc_0042DC26: fnclex
  loc_0042DC28: jge 0042DC42h
  loc_0042DC2A: mov ecx, var_A4
  loc_0042DC30: push 000000A0h
  loc_0042DC35: push 0040F02Ch
  loc_0042DC3A: push ecx
  loc_0042DC3B: push eax
  loc_0042DC3C: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DC42: mov eax, var_18
  loc_0042DC45: lea edx, var_30
  loc_0042DC48: mov var_28, eax
  loc_0042DC4B: lea eax, var_40
  loc_0042DC4E: push edx
  loc_0042DC4F: push eax
  loc_0042DC50: mov var_18, ebx
  loc_0042DC53: mov var_30, 00000008h
  loc_0042DC5A: call [00401050h] ; rtcTrimVar
  loc_0042DC60: lea ecx, var_70
  loc_0042DC63: lea edx, var_40
  loc_0042DC66: push ecx
  loc_0042DC67: lea eax, var_50
  loc_0042DC6A: push edx
  loc_0042DC6B: push eax
  loc_0042DC6C: call [004010C0h] ; __vbaVarCat
  loc_0042DC72: push eax
  loc_0042DC73: call [00401014h] ; __vbaStrVarMove
  loc_0042DC79: mov edx, eax
  loc_0042DC7B: lea ecx, var_1C
  loc_0042DC7E: call [004010FCh] ; __vbaStrMove
  loc_0042DC84: mov edx, eax
  loc_0042DC86: lea ecx, [esi+00000034h]
  loc_0042DC89: call [004010E0h] ; __vbaStrCopy
  loc_0042DC8F: mov ebx, [00401110h] ; __vbaFreeStr
  loc_0042DC95: lea ecx, var_1C
  loc_0042DC98: call ebx
  loc_0042DC9A: lea ecx, var_20
  loc_0042DC9D: call [00401114h] ; __vbaFreeObj
  loc_0042DCA3: lea ecx, var_50
  loc_0042DCA6: lea edx, var_40
  loc_0042DCA9: push ecx
  loc_0042DCAA: lea eax, var_30
  loc_0042DCAD: push edx
  loc_0042DCAE: push eax
  loc_0042DCAF: push 00000003h
  loc_0042DCB1: call [00401018h] ; __vbaFreeVarList
  loc_0042DCB7: mov ecx, [esi]
  loc_0042DCB9: add esp, 00000010h
  loc_0042DCBC: mov var_68, 0040F028h
  loc_0042DCC3: mov var_70, 00000008h
  loc_0042DCCA: push esi
  loc_0042DCCB: call [ecx+00000324h]
  loc_0042DCD1: lea edx, var_20
  loc_0042DCD4: push eax
  loc_0042DCD5: push edx
  loc_0042DCD6: call edi
  loc_0042DCD8: mov ecx, [eax]
  loc_0042DCDA: lea edx, var_18
  loc_0042DCDD: push edx
  loc_0042DCDE: push eax
  loc_0042DCDF: mov var_A4, eax
  loc_0042DCE5: call [ecx+000000A0h]
  loc_0042DCEB: test eax, eax
  loc_0042DCED: fnclex
  loc_0042DCEF: jge 0042DD09h
  loc_0042DCF1: mov ecx, var_A4
  loc_0042DCF7: push 000000A0h
  loc_0042DCFC: push 0040F02Ch
  loc_0042DD01: push ecx
  loc_0042DD02: push eax
  loc_0042DD03: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DD09: mov eax, var_18
  loc_0042DD0C: lea edx, var_30
  loc_0042DD0F: mov var_28, eax
  loc_0042DD12: lea eax, var_40
  loc_0042DD15: push edx
  loc_0042DD16: push eax
  loc_0042DD17: mov var_18, 00000000h
  loc_0042DD1E: mov var_30, 00000008h
  loc_0042DD25: call [00401050h] ; rtcTrimVar
  loc_0042DD2B: lea ecx, var_70
  loc_0042DD2E: lea edx, var_40
  loc_0042DD31: push ecx
  loc_0042DD32: lea eax, var_50
  loc_0042DD35: push edx
  loc_0042DD36: push eax
  loc_0042DD37: call [004010C0h] ; __vbaVarCat
  loc_0042DD3D: push eax
  loc_0042DD3E: call [00401014h] ; __vbaStrVarMove
  loc_0042DD44: mov edx, eax
  loc_0042DD46: lea ecx, var_1C
  loc_0042DD49: call [004010FCh] ; __vbaStrMove
  loc_0042DD4F: mov edx, eax
  loc_0042DD51: lea ecx, [esi+00000038h]
  loc_0042DD54: call [004010E0h] ; __vbaStrCopy
  loc_0042DD5A: lea ecx, var_1C
  loc_0042DD5D: call ebx
  loc_0042DD5F: lea ecx, var_20
  loc_0042DD62: call [00401114h] ; __vbaFreeObj
  loc_0042DD68: lea ecx, var_50
  loc_0042DD6B: lea edx, var_40
  loc_0042DD6E: push ecx
  loc_0042DD6F: lea eax, var_30
  loc_0042DD72: push edx
  loc_0042DD73: push eax
  loc_0042DD74: push 00000003h
  loc_0042DD76: call [00401018h] ; __vbaFreeVarList
  loc_0042DD7C: mov ecx, [esi]
  loc_0042DD7E: add esp, 00000010h
  loc_0042DD81: mov var_68, 0040F028h
  loc_0042DD88: mov var_70, 00000008h
  loc_0042DD8F: push esi
  loc_0042DD90: call [ecx+00000320h]
  loc_0042DD96: lea edx, var_20
  loc_0042DD99: push eax
  loc_0042DD9A: push edx
  loc_0042DD9B: call edi
  loc_0042DD9D: mov ecx, [eax]
  loc_0042DD9F: lea edx, var_18
  loc_0042DDA2: push edx
  loc_0042DDA3: push eax
  loc_0042DDA4: mov var_A4, eax
  loc_0042DDAA: call [ecx+000000A0h]
  loc_0042DDB0: test eax, eax
  loc_0042DDB2: fnclex
  loc_0042DDB4: jge 0042DDCEh
  loc_0042DDB6: mov ecx, var_A4
  loc_0042DDBC: push 000000A0h
  loc_0042DDC1: push 0040F02Ch
  loc_0042DDC6: push ecx
  loc_0042DDC7: push eax
  loc_0042DDC8: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DDCE: mov eax, var_18
  loc_0042DDD1: lea edx, var_30
  loc_0042DDD4: mov var_28, eax
  loc_0042DDD7: lea eax, var_40
  loc_0042DDDA: push edx
  loc_0042DDDB: push eax
  loc_0042DDDC: mov var_18, 00000000h
  loc_0042DDE3: mov var_30, 00000008h
  loc_0042DDEA: call [00401050h] ; rtcTrimVar
  loc_0042DDF0: lea ecx, var_70
  loc_0042DDF3: lea edx, var_40
  loc_0042DDF6: push ecx
  loc_0042DDF7: lea eax, var_50
  loc_0042DDFA: push edx
  loc_0042DDFB: push eax
  loc_0042DDFC: call [004010C0h] ; __vbaVarCat
  loc_0042DE02: push eax
  loc_0042DE03: call [00401014h] ; __vbaStrVarMove
  loc_0042DE09: mov edx, eax
  loc_0042DE0B: lea ecx, var_1C
  loc_0042DE0E: call [004010FCh] ; __vbaStrMove
  loc_0042DE14: mov edx, eax
  loc_0042DE16: lea ecx, [esi+0000003Ch]
  loc_0042DE19: call [004010E0h] ; __vbaStrCopy
  loc_0042DE1F: lea ecx, var_1C
  loc_0042DE22: call ebx
  loc_0042DE24: lea ecx, var_20
  loc_0042DE27: call [00401114h] ; __vbaFreeObj
  loc_0042DE2D: lea ecx, var_50
  loc_0042DE30: lea edx, var_40
  loc_0042DE33: push ecx
  loc_0042DE34: lea eax, var_30
  loc_0042DE37: push edx
  loc_0042DE38: push eax
  loc_0042DE39: push 00000003h
  loc_0042DE3B: call [00401018h] ; __vbaFreeVarList
  loc_0042DE41: mov ecx, [esi]
  loc_0042DE43: add esp, 00000010h
  loc_0042DE46: mov var_68, 0040F028h
  loc_0042DE4D: mov var_70, 00000008h
  loc_0042DE54: push esi
  loc_0042DE55: call [ecx+0000031Ch]
  loc_0042DE5B: lea edx, var_20
  loc_0042DE5E: push eax
  loc_0042DE5F: push edx
  loc_0042DE60: call edi
  loc_0042DE62: mov ecx, [eax]
  loc_0042DE64: lea edx, var_18
  loc_0042DE67: push edx
  loc_0042DE68: push eax
  loc_0042DE69: mov var_A4, eax
  loc_0042DE6F: call [ecx+000000A0h]
  loc_0042DE75: test eax, eax
  loc_0042DE77: fnclex
  loc_0042DE79: jge 0042DE93h
  loc_0042DE7B: mov ecx, var_A4
  loc_0042DE81: push 000000A0h
  loc_0042DE86: push 0040F02Ch
  loc_0042DE8B: push ecx
  loc_0042DE8C: push eax
  loc_0042DE8D: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DE93: mov eax, var_18
  loc_0042DE96: lea edx, var_30
  loc_0042DE99: mov var_28, eax
  loc_0042DE9C: lea eax, var_40
  loc_0042DE9F: push edx
  loc_0042DEA0: push eax
  loc_0042DEA1: mov var_18, 00000000h
  loc_0042DEA8: mov var_30, 00000008h
  loc_0042DEAF: call [00401050h] ; rtcTrimVar
  loc_0042DEB5: lea ecx, var_70
  loc_0042DEB8: lea edx, var_40
  loc_0042DEBB: push ecx
  loc_0042DEBC: lea eax, var_50
  loc_0042DEBF: push edx
  loc_0042DEC0: push eax
  loc_0042DEC1: call [004010C0h] ; __vbaVarCat
  loc_0042DEC7: push eax
  loc_0042DEC8: call [00401014h] ; __vbaStrVarMove
  loc_0042DECE: mov edx, eax
  loc_0042DED0: lea ecx, var_1C
  loc_0042DED3: call [004010FCh] ; __vbaStrMove
  loc_0042DED9: mov edx, eax
  loc_0042DEDB: lea ecx, [esi+00000040h]
  loc_0042DEDE: call [004010E0h] ; __vbaStrCopy
  loc_0042DEE4: lea ecx, var_1C
  loc_0042DEE7: call ebx
  loc_0042DEE9: lea ecx, var_20
  loc_0042DEEC: call [00401114h] ; __vbaFreeObj
  loc_0042DEF2: lea ecx, var_50
  loc_0042DEF5: lea edx, var_40
  loc_0042DEF8: push ecx
  loc_0042DEF9: lea eax, var_30
  loc_0042DEFC: push edx
  loc_0042DEFD: push eax
  loc_0042DEFE: push 00000003h
  loc_0042DF00: call [00401018h] ; __vbaFreeVarList
  loc_0042DF06: mov ecx, [esi]
  loc_0042DF08: add esp, 00000010h
  loc_0042DF0B: push esi
  loc_0042DF0C: call [ecx+00000328h]
  loc_0042DF12: lea edx, var_20
  loc_0042DF15: push eax
  loc_0042DF16: push edx
  loc_0042DF17: call edi
  loc_0042DF19: mov ecx, [eax]
  loc_0042DF1B: lea edx, var_18
  loc_0042DF1E: push edx
  loc_0042DF1F: push eax
  loc_0042DF20: mov var_A4, eax
  loc_0042DF26: call [ecx+000000A0h]
  loc_0042DF2C: test eax, eax
  loc_0042DF2E: fnclex
  loc_0042DF30: jge 0042DF4Ah
  loc_0042DF32: mov ecx, var_A4
  loc_0042DF38: push 000000A0h
  loc_0042DF3D: push 0040F02Ch
  loc_0042DF42: push ecx
  loc_0042DF43: push eax
  loc_0042DF44: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DF4A: mov edx, var_18
  loc_0042DF4D: push edx
  loc_0042DF4E: push 0040F040h
  loc_0042DF53: call [0040106Ch] ; __vbaStrCmp
  loc_0042DF59: neg eax
  loc_0042DF5B: sbb eax, eax
  loc_0042DF5D: lea ecx, var_18
  loc_0042DF60: inc eax
  loc_0042DF61: neg eax
  loc_0042DF63: mov var_AC, eax
  loc_0042DF69: call ebx
  loc_0042DF6B: lea ecx, var_20
  loc_0042DF6E: call [00401114h] ; __vbaFreeObj
  loc_0042DF74: cmp var_AC, 0000h
  loc_0042DF7C: jz 0042E02Ch
  loc_0042DF82: mov eax, [esi]
  loc_0042DF84: push esi
  loc_0042DF85: call [eax+00000328h]
  loc_0042DF8B: lea ecx, var_20
  loc_0042DF8E: push eax
  loc_0042DF8F: push ecx
  loc_0042DF90: call edi
  loc_0042DF92: mov esi, eax
  loc_0042DF94: push esi
  loc_0042DF95: mov edx, [esi]
  loc_0042DF97: call [edx+00000204h]
  loc_0042DF9D: test eax, eax
  loc_0042DF9F: fnclex
  loc_0042DFA1: jge 0042DFB5h
  loc_0042DFA3: push 00000204h
  loc_0042DFA8: push 0040F02Ch
  loc_0042DFAD: push esi
  loc_0042DFAE: push eax
  loc_0042DFAF: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DFB5: lea ecx, var_20
  loc_0042DFB8: call [00401114h] ; __vbaFreeObj
  loc_0042DFBE: mov ecx, 80020004h
  loc_0042DFC3: mov eax, 0000000Ah
  loc_0042DFC8: mov var_58, ecx
  loc_0042DFCB: mov var_48, ecx
  loc_0042DFCE: mov var_38, ecx
  loc_0042DFD1: lea edx, var_70
  loc_0042DFD4: lea ecx, var_30
  loc_0042DFD7: mov var_60, eax
  loc_0042DFDA: mov var_50, eax
  loc_0042DFDD: mov var_40, eax
  loc_0042DFE0: mov var_68, 0040F09Ch ; "Please enter a name for the dependent variable."
  loc_0042DFE7: mov var_70, 00000008h
  loc_0042DFEE: call [004010F4h] ; __vbaVarDup
  loc_0042DFF4: lea eax, var_60
  loc_0042DFF7: lea ecx, var_50
  loc_0042DFFA: push eax
  loc_0042DFFB: lea edx, var_40
  loc_0042DFFE: push ecx
  loc_0042DFFF: push edx
  loc_0042E000: lea eax, var_30
  loc_0042E003: push 00000000h
  loc_0042E005: push eax
  loc_0042E006: call [00401038h] ; rtcMsgBox
  loc_0042E00C: lea ecx, var_60
  loc_0042E00F: lea edx, var_50
  loc_0042E012: push ecx
  loc_0042E013: lea eax, var_40
  loc_0042E016: push edx
  loc_0042E017: lea ecx, var_30
  loc_0042E01A: push eax
  loc_0042E01B: push ecx
  loc_0042E01C: push 00000004h
  loc_0042E01E: call [00401018h] ; __vbaFreeVarList
  loc_0042E024: add esp, 00000014h
  loc_0042E027: jmp 0042E460h
  loc_0042E02C: mov edx, [esi]
  loc_0042E02E: push esi
  loc_0042E02F: call [edx+00000320h]
  loc_0042E035: push eax
  loc_0042E036: lea eax, var_20
  loc_0042E039: push eax
  loc_0042E03A: call edi
  loc_0042E03C: mov ecx, [eax]
  loc_0042E03E: lea edx, var_18
  loc_0042E041: push edx
  loc_0042E042: push eax
  loc_0042E043: mov var_A4, eax
  loc_0042E049: call [ecx+000000A0h]
  loc_0042E04F: test eax, eax
  loc_0042E051: fnclex
  loc_0042E053: jge 0042E06Dh
  loc_0042E055: mov ecx, var_A4
  loc_0042E05B: push 000000A0h
  loc_0042E060: push 0040F02Ch
  loc_0042E065: push ecx
  loc_0042E066: push eax
  loc_0042E067: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E06D: mov edx, var_18
  loc_0042E070: push edx
  loc_0042E071: push 0040F040h
  loc_0042E076: call [0040106Ch] ; __vbaStrCmp
  loc_0042E07C: neg eax
  loc_0042E07E: sbb eax, eax
  loc_0042E080: lea ecx, var_18
  loc_0042E083: inc eax
  loc_0042E084: neg eax
  loc_0042E086: mov var_AC, eax
  loc_0042E08C: call ebx
  loc_0042E08E: lea ecx, var_20
  loc_0042E091: call [00401114h] ; __vbaFreeObj
  loc_0042E097: cmp var_AC, 0000h
  loc_0042E09F: jz 0042E138h
  loc_0042E0A5: mov ecx, 80020004h
  loc_0042E0AA: mov eax, 0000000Ah
  loc_0042E0AF: mov var_58, ecx
  loc_0042E0B2: mov var_48, ecx
  loc_0042E0B5: mov var_38, ecx
  loc_0042E0B8: lea edx, var_70
  loc_0042E0BB: lea ecx, var_30
  loc_0042E0BE: mov var_60, eax
  loc_0042E0C1: mov var_50, eax
  loc_0042E0C4: mov var_40, eax
  loc_0042E0C7: mov var_68, 0040F11Ch ; "Please enter a name for the independent variable."
  loc_0042E0CE: mov var_70, 00000008h
  loc_0042E0D5: call [004010F4h] ; __vbaVarDup
  loc_0042E0DB: lea eax, var_60
  loc_0042E0DE: lea ecx, var_50
  loc_0042E0E1: push eax
  loc_0042E0E2: lea edx, var_40
  loc_0042E0E5: push ecx
  loc_0042E0E6: push edx
  loc_0042E0E7: lea eax, var_30
  loc_0042E0EA: push 00000000h
  loc_0042E0EC: push eax
  loc_0042E0ED: call [00401038h] ; rtcMsgBox
  loc_0042E0F3: lea ecx, var_60
  loc_0042E0F6: lea edx, var_50
  loc_0042E0F9: push ecx
  loc_0042E0FA: lea eax, var_40
  loc_0042E0FD: push edx
  loc_0042E0FE: lea ecx, var_30
  loc_0042E101: push eax
  loc_0042E102: push ecx
  loc_0042E103: push 00000004h
  loc_0042E105: call [00401018h] ; __vbaFreeVarList
  loc_0042E10B: mov edx, [esi]
  loc_0042E10D: add esp, 00000014h
  loc_0042E110: push esi
  loc_0042E111: call [edx+00000320h]
  loc_0042E117: push eax
  loc_0042E118: lea eax, var_20
  loc_0042E11B: push eax
  loc_0042E11C: call edi
  loc_0042E11E: mov esi, eax
  loc_0042E120: push esi
  loc_0042E121: mov ecx, [esi]
  loc_0042E123: call [ecx+00000204h]
  loc_0042E129: test eax, eax
  loc_0042E12B: fnclex
  loc_0042E12D: jge 0042E391h
  loc_0042E133: jmp 0042E37Fh
  loc_0042E138: mov edx, [esi]
  loc_0042E13A: push esi
  loc_0042E13B: call [edx+00000324h]
  loc_0042E141: push eax
  loc_0042E142: lea eax, var_20
  loc_0042E145: push eax
  loc_0042E146: call edi
  loc_0042E148: mov ecx, [eax]
  loc_0042E14A: lea edx, var_18
  loc_0042E14D: push edx
  loc_0042E14E: push eax
  loc_0042E14F: mov var_A4, eax
  loc_0042E155: call [ecx+000000A0h]
  loc_0042E15B: test eax, eax
  loc_0042E15D: fnclex
  loc_0042E15F: jge 0042E179h
  loc_0042E161: mov ecx, var_A4
  loc_0042E167: push 000000A0h
  loc_0042E16C: push 0040F02Ch
  loc_0042E171: push ecx
  loc_0042E172: push eax
  loc_0042E173: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E179: mov edx, var_18
  loc_0042E17C: push edx
  loc_0042E17D: push 0040F040h
  loc_0042E182: call [0040106Ch] ; __vbaStrCmp
  loc_0042E188: neg eax
  loc_0042E18A: sbb eax, eax
  loc_0042E18C: lea ecx, var_18
  loc_0042E18F: inc eax
  loc_0042E190: neg eax
  loc_0042E192: mov var_AC, eax
  loc_0042E198: call ebx
  loc_0042E19A: lea ecx, var_20
  loc_0042E19D: call [00401114h] ; __vbaFreeObj
  loc_0042E1A3: cmp var_AC, 0000h
  loc_0042E1AB: jz 0042E260h
  loc_0042E1B1: mov ecx, 80020004h
  loc_0042E1B6: mov eax, 0000000Ah
  loc_0042E1BB: push 0040F184h ; "Please enter the units for the dependent variable."
  loc_0042E1C0: push 0040F720h ; vbCrLf
  loc_0042E1C5: mov var_58, ecx
  loc_0042E1C8: mov var_60, eax
  loc_0042E1CB: mov var_48, ecx
  loc_0042E1CE: mov var_50, eax
  loc_0042E1D1: mov var_38, ecx
  loc_0042E1D4: mov var_40, eax
  loc_0042E1D7: call [0040102Ch] ; __vbaStrCat
  loc_0042E1DD: mov edx, eax
  loc_0042E1DF: lea ecx, var_18
  loc_0042E1E2: call [004010FCh] ; __vbaStrMove
  loc_0042E1E8: push eax
  loc_0042E1E9: push 00412DE4h ; "This will be used later."
  loc_0042E1EE: call [0040102Ch] ; __vbaStrCat
  loc_0042E1F4: mov var_28, eax
  loc_0042E1F7: lea eax, var_60
  loc_0042E1FA: lea ecx, var_50
  loc_0042E1FD: push eax
  loc_0042E1FE: lea edx, var_40
  loc_0042E201: push ecx
  loc_0042E202: push edx
  loc_0042E203: lea eax, var_30
  loc_0042E206: push 00000000h
  loc_0042E208: push eax
  loc_0042E209: mov var_30, 00000008h
  loc_0042E210: call [00401038h] ; rtcMsgBox
  loc_0042E216: lea ecx, var_18
  loc_0042E219: call ebx
  loc_0042E21B: lea ecx, var_60
  loc_0042E21E: lea edx, var_50
  loc_0042E221: push ecx
  loc_0042E222: lea eax, var_40
  loc_0042E225: push edx
  loc_0042E226: lea ecx, var_30
  loc_0042E229: push eax
  loc_0042E22A: push ecx
  loc_0042E22B: push 00000004h
  loc_0042E22D: call [00401018h] ; __vbaFreeVarList
  loc_0042E233: mov edx, [esi]
  loc_0042E235: add esp, 00000014h
  loc_0042E238: push esi
  loc_0042E239: call [edx+00000324h]
  loc_0042E23F: push eax
  loc_0042E240: lea eax, var_20
  loc_0042E243: push eax
  loc_0042E244: call edi
  loc_0042E246: mov esi, eax
  loc_0042E248: push esi
  loc_0042E249: mov ecx, [esi]
  loc_0042E24B: call [ecx+00000204h]
  loc_0042E251: test eax, eax
  loc_0042E253: fnclex
  loc_0042E255: jge 0042E391h
  loc_0042E25B: jmp 0042E37Fh
  loc_0042E260: mov edx, [esi]
  loc_0042E262: push esi
  loc_0042E263: call [edx+0000031Ch]
  loc_0042E269: push eax
  loc_0042E26A: lea eax, var_20
  loc_0042E26D: push eax
  loc_0042E26E: call edi
  loc_0042E270: mov ecx, [eax]
  loc_0042E272: lea edx, var_18
  loc_0042E275: push edx
  loc_0042E276: push eax
  loc_0042E277: mov var_A4, eax
  loc_0042E27D: call [ecx+000000A0h]
  loc_0042E283: test eax, eax
  loc_0042E285: fnclex
  loc_0042E287: jge 0042E2A1h
  loc_0042E289: mov ecx, var_A4
  loc_0042E28F: push 000000A0h
  loc_0042E294: push 0040F02Ch
  loc_0042E299: push ecx
  loc_0042E29A: push eax
  loc_0042E29B: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E2A1: mov edx, var_18
  loc_0042E2A4: push edx
  loc_0042E2A5: push 0040F040h
  loc_0042E2AA: call [0040106Ch] ; __vbaStrCmp
  loc_0042E2B0: neg eax
  loc_0042E2B2: sbb eax, eax
  loc_0042E2B4: lea ecx, var_18
  loc_0042E2B7: inc eax
  loc_0042E2B8: neg eax
  loc_0042E2BA: mov var_AC, eax
  loc_0042E2C0: call ebx
  loc_0042E2C2: lea ecx, var_20
  loc_0042E2C5: call [00401114h] ; __vbaFreeObj
  loc_0042E2CB: cmp var_AC, 0000h
  loc_0042E2D3: jz 0042E39Fh
  loc_0042E2D9: mov ecx, 80020004h
  loc_0042E2DE: mov eax, 0000000Ah
  loc_0042E2E3: push 0040F20Ch ; "Please enter the units for the independent variable."
  loc_0042E2E8: push 0040F720h ; vbCrLf
  loc_0042E2ED: mov var_58, ecx
  loc_0042E2F0: mov var_60, eax
  loc_0042E2F3: mov var_48, ecx
  loc_0042E2F6: mov var_50, eax
  loc_0042E2F9: mov var_38, ecx
  loc_0042E2FC: mov var_40, eax
  loc_0042E2FF: call [0040102Ch] ; __vbaStrCat
  loc_0042E305: mov edx, eax
  loc_0042E307: lea ecx, var_18
  loc_0042E30A: call [004010FCh] ; __vbaStrMove
  loc_0042E310: push eax
  loc_0042E311: push 00412DE4h ; "This will be used later."
  loc_0042E316: call [0040102Ch] ; __vbaStrCat
  loc_0042E31C: mov var_28, eax
  loc_0042E31F: lea eax, var_60
  loc_0042E322: lea ecx, var_50
  loc_0042E325: push eax
  loc_0042E326: lea edx, var_40
  loc_0042E329: push ecx
  loc_0042E32A: push edx
  loc_0042E32B: lea eax, var_30
  loc_0042E32E: push 00000000h
  loc_0042E330: push eax
  loc_0042E331: mov var_30, 00000008h
  loc_0042E338: call [00401038h] ; rtcMsgBox
  loc_0042E33E: lea ecx, var_18
  loc_0042E341: call ebx
  loc_0042E343: lea ecx, var_60
  loc_0042E346: lea edx, var_50
  loc_0042E349: push ecx
  loc_0042E34A: lea eax, var_40
  loc_0042E34D: push edx
  loc_0042E34E: lea ecx, var_30
  loc_0042E351: push eax
  loc_0042E352: push ecx
  loc_0042E353: push 00000004h
  loc_0042E355: call [00401018h] ; __vbaFreeVarList
  loc_0042E35B: mov edx, [esi]
  loc_0042E35D: add esp, 00000014h
  loc_0042E360: push esi
  loc_0042E361: call [edx+00000324h]
  loc_0042E367: push eax
  loc_0042E368: lea eax, var_20
  loc_0042E36B: push eax
  loc_0042E36C: call edi
  loc_0042E36E: mov esi, eax
  loc_0042E370: push esi
  loc_0042E371: mov ecx, [esi]
  loc_0042E373: call [ecx+00000204h]
  loc_0042E379: test eax, eax
  loc_0042E37B: fnclex
  loc_0042E37D: jge 0042E391h
  loc_0042E37F: push 00000204h
  loc_0042E384: push 0040F02Ch
  loc_0042E389: push esi
  loc_0042E38A: push eax
  loc_0042E38B: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E391: lea ecx, var_20
  loc_0042E394: call [00401114h] ; __vbaFreeObj
  loc_0042E39A: jmp 0042E460h
  loc_0042E39F: mov edx, [esi]
  loc_0042E3A1: push esi
  loc_0042E3A2: call [edx+0000030Ch]
  loc_0042E3A8: push eax
  loc_0042E3A9: lea eax, var_20
  loc_0042E3AC: push eax
  loc_0042E3AD: call edi
  loc_0042E3AF: mov ebx, eax
  loc_0042E3B1: push 00000000h
  loc_0042E3B3: push ebx
  loc_0042E3B4: mov ecx, [ebx]
  loc_0042E3B6: call [ecx+0000009Ch]
  loc_0042E3BC: test eax, eax
  loc_0042E3BE: fnclex
  loc_0042E3C0: jge 0042E3D4h
  loc_0042E3C2: push 0000009Ch
  loc_0042E3C7: push 0040E728h
  loc_0042E3CC: push ebx
  loc_0042E3CD: push eax
  loc_0042E3CE: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E3D4: lea ecx, var_20
  loc_0042E3D7: call [00401114h] ; __vbaFreeObj
  loc_0042E3DD: mov edx, [esi]
  loc_0042E3DF: push esi
  loc_0042E3E0: call [edx+00000304h]
  loc_0042E3E6: push eax
  loc_0042E3E7: lea eax, var_20
  loc_0042E3EA: push eax
  loc_0042E3EB: call edi
  loc_0042E3ED: mov ebx, eax
  loc_0042E3EF: push FFFFFFFFh
  loc_0042E3F1: push ebx
  loc_0042E3F2: mov ecx, [ebx]
  loc_0042E3F4: call [ecx+0000009Ch]
  loc_0042E3FA: test eax, eax
  loc_0042E3FC: fnclex
  loc_0042E3FE: jge 0042E412h
  loc_0042E400: push 0000009Ch
  loc_0042E405: push 0040E728h
  loc_0042E40A: push ebx
  loc_0042E40B: push eax
  loc_0042E40C: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E412: mov ebx, [00401114h] ; __vbaFreeObj
  loc_0042E418: lea ecx, var_20
  loc_0042E41B: call ebx
  loc_0042E41D: mov edx, [esi]
  loc_0042E41F: push esi
  loc_0042E420: call [edx+000002FCh]
  loc_0042E426: push eax
  loc_0042E427: lea eax, var_20
  loc_0042E42A: push eax
  loc_0042E42B: call edi
  loc_0042E42D: mov edi, eax
  loc_0042E42F: push FFFFFFFFh
  loc_0042E431: push edi
  loc_0042E432: mov ecx, [edi]
  loc_0042E434: call [ecx+0000009Ch]
  loc_0042E43A: test eax, eax
  loc_0042E43C: fnclex
  loc_0042E43E: jge 0042E452h
  loc_0042E440: push 0000009Ch
  loc_0042E445: push 0040E728h
  loc_0042E44A: push edi
  loc_0042E44B: push eax
  loc_0042E44C: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E452: lea ecx, var_20
  loc_0042E455: call ebx
  loc_0042E457: mov edx, [esi]
  loc_0042E459: push esi
  loc_0042E45A: call [edx+00000710h]
  loc_0042E460: mov var_4, 00000000h
  loc_0042E467: push 0042E4A7h
  loc_0042E46C: jmp 0042E4A6h
  loc_0042E46E: lea eax, var_1C
  loc_0042E471: lea ecx, var_18
  loc_0042E474: push eax
  loc_0042E475: push ecx
  loc_0042E476: push 00000002h
  loc_0042E478: call [004010E4h] ; __vbaFreeStrList
  loc_0042E47E: add esp, 0000000Ch
  loc_0042E481: lea ecx, var_20
  loc_0042E484: call [00401114h] ; __vbaFreeObj
  loc_0042E48A: lea edx, var_60
  loc_0042E48D: lea eax, var_50
  loc_0042E490: push edx
  loc_0042E491: lea ecx, var_40
  loc_0042E494: push eax
  loc_0042E495: lea edx, var_30
  loc_0042E498: push ecx
  loc_0042E499: push edx
  loc_0042E49A: push 00000004h
  loc_0042E49C: call [00401018h] ; __vbaFreeVarList
  loc_0042E4A2: add esp, 00000014h
  loc_0042E4A5: ret
  loc_0042E4A6: ret
  loc_0042E4A7: mov eax, Me
  loc_0042E4AA: push eax
  loc_0042E4AB: mov ecx, [eax]
  loc_0042E4AD: call [ecx+00000008h]
  loc_0042E4B0: mov eax, var_4
  loc_0042E4B3: mov ecx, var_14
  loc_0042E4B6: pop edi
  loc_0042E4B7: pop esi
  loc_0042E4B8: mov fs:[00000000h], ecx
  loc_0042E4BF: pop ebx
  loc_0042E4C0: mov esp, ebp
  loc_0042E4C2: pop ebp
  loc_0042E4C3: retn 0004h
End Sub

Private Sub cmdRestore_Click() '42E660
  loc_0042E660: push ebp
  loc_0042E661: mov ebp, esp
  loc_0042E663: sub esp, 0000000Ch
  loc_0042E666: push 00401926h ; __vbaExceptHandler
  loc_0042E66B: mov eax, fs:[00000000h]
  loc_0042E671: push eax
  loc_0042E672: mov fs:[00000000h], esp
  loc_0042E679: sub esp, 00000014h
  loc_0042E67C: push ebx
  loc_0042E67D: push esi
  loc_0042E67E: push edi
  loc_0042E67F: mov var_C, esp
  loc_0042E682: mov var_8, 00401858h
  loc_0042E689: mov esi, Me
  loc_0042E68C: mov eax, esi
  loc_0042E68E: and eax, 00000001h
  loc_0042E691: mov var_4, eax
  loc_0042E694: and esi, FFFFFFFEh
  loc_0042E697: push esi
  loc_0042E698: mov Me, esi
  loc_0042E69B: mov ecx, [esi]
  loc_0042E69D: call [ecx+00000004h]
  loc_0042E6A0: mov edi, [004010E0h] ; __vbaStrCopy
  loc_0042E6A6: mov edx, 0040DB10h ; " Net Income"
  loc_0042E6AB: lea ecx, [esi+00000034h]
  loc_0042E6AE: mov var_18, 00000000h
  loc_0042E6B5: call edi
  loc_0042E6B7: mov edx, 004102A0h ; " millions of dollars"
  loc_0042E6BC: lea ecx, [esi+00000038h]
  loc_0042E6BF: call edi
  loc_0042E6C1: mov edx, 0040DB54h ; " Sales"
  loc_0042E6C6: lea ecx, [esi+0000003Ch]
  loc_0042E6C9: call edi
  loc_0042E6CB: mov edx, 004102A0h ; " millions of dollars"
  loc_0042E6D0: lea ecx, [esi+00000040h]
  loc_0042E6D3: call edi
  loc_0042E6D5: mov edx, [esi]
  loc_0042E6D7: push esi
  loc_0042E6D8: call [edx+00000328h]
  loc_0042E6DE: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0042E6E4: push eax
  loc_0042E6E5: lea eax, var_18
  loc_0042E6E8: push eax
  loc_0042E6E9: call ebx
  loc_0042E6EB: mov edi, eax
  loc_0042E6ED: push 0040DB10h ; " Net Income"
  loc_0042E6F2: push edi
  loc_0042E6F3: mov ecx, [edi]
  loc_0042E6F5: call [ecx+000000A4h]
  loc_0042E6FB: test eax, eax
  loc_0042E6FD: fnclex
  loc_0042E6FF: jge 0042E713h
  loc_0042E701: push 000000A4h
  loc_0042E706: push 0040F02Ch
  loc_0042E70B: push edi
  loc_0042E70C: push eax
  loc_0042E70D: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E713: lea ecx, var_18
  loc_0042E716: call [00401114h] ; __vbaFreeObj
  loc_0042E71C: mov edx, [esi]
  loc_0042E71E: push esi
  loc_0042E71F: call [edx+00000324h]
  loc_0042E725: push eax
  loc_0042E726: lea eax, var_18
  loc_0042E729: push eax
  loc_0042E72A: call ebx
  loc_0042E72C: mov edi, eax
  loc_0042E72E: push 004102A0h ; " millions of dollars"
  loc_0042E733: push edi
  loc_0042E734: mov ecx, [edi]
  loc_0042E736: call [ecx+000000A4h]
  loc_0042E73C: test eax, eax
  loc_0042E73E: fnclex
  loc_0042E740: jge 0042E754h
  loc_0042E742: push 000000A4h
  loc_0042E747: push 0040F02Ch
  loc_0042E74C: push edi
  loc_0042E74D: push eax
  loc_0042E74E: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E754: lea ecx, var_18
  loc_0042E757: call [00401114h] ; __vbaFreeObj
  loc_0042E75D: mov edx, [esi]
  loc_0042E75F: push esi
  loc_0042E760: call [edx+00000320h]
  loc_0042E766: push eax
  loc_0042E767: lea eax, var_18
  loc_0042E76A: push eax
  loc_0042E76B: call ebx
  loc_0042E76D: mov edi, eax
  loc_0042E76F: push 0040DB54h ; " Sales"
  loc_0042E774: push edi
  loc_0042E775: mov ecx, [edi]
  loc_0042E777: call [ecx+000000A4h]
  loc_0042E77D: test eax, eax
  loc_0042E77F: fnclex
  loc_0042E781: jge 0042E795h
  loc_0042E783: push 000000A4h
  loc_0042E788: push 0040F02Ch
  loc_0042E78D: push edi
  loc_0042E78E: push eax
  loc_0042E78F: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E795: mov edi, [00401114h] ; __vbaFreeObj
  loc_0042E79B: lea ecx, var_18
  loc_0042E79E: call edi
  loc_0042E7A0: mov edx, [esi]
  loc_0042E7A2: push esi
  loc_0042E7A3: call [edx+0000031Ch]
  loc_0042E7A9: push eax
  loc_0042E7AA: lea eax, var_18
  loc_0042E7AD: push eax
  loc_0042E7AE: call ebx
  loc_0042E7B0: mov esi, eax
  loc_0042E7B2: push 004102A0h ; " millions of dollars"
  loc_0042E7B7: push esi
  loc_0042E7B8: mov ecx, [esi]
  loc_0042E7BA: call [ecx+000000A4h]
  loc_0042E7C0: test eax, eax
  loc_0042E7C2: fnclex
  loc_0042E7C4: jge 0042E7D8h
  loc_0042E7C6: push 000000A4h
  loc_0042E7CB: push 0040F02Ch
  loc_0042E7D0: push esi
  loc_0042E7D1: push eax
  loc_0042E7D2: call [00401030h] ; __vbaHresultCheckObj
  loc_0042E7D8: lea ecx, var_18
  loc_0042E7DB: call edi
  loc_0042E7DD: mov var_4, 00000000h
  loc_0042E7E4: push 0042E7F6h
  loc_0042E7E9: jmp 0042E7F5h
  loc_0042E7EB: lea ecx, var_18
  loc_0042E7EE: call [00401114h] ; __vbaFreeObj
  loc_0042E7F4: ret
  loc_0042E7F5: ret
  loc_0042E7F6: mov eax, Me
  loc_0042E7F9: push eax
  loc_0042E7FA: mov edx, [eax]
  loc_0042E7FC: call [edx+00000008h]
  loc_0042E7FF: mov eax, var_4
  loc_0042E802: mov ecx, var_14
  loc_0042E805: pop edi
  loc_0042E806: pop esi
  loc_0042E807: mov fs:[00000000h], ecx
  loc_0042E80E: pop ebx
  loc_0042E80F: mov esp, ebp
  loc_0042E811: pop ebp
  loc_0042E812: retn 0004h
End Sub

Private Sub Proc_16_5_42DAF0
  loc_0042DAF0: push ebp
  loc_0042DAF1: mov ebp, esp
  loc_0042DAF3: sub esp, 00000008h
  loc_0042DAF6: push 00401926h ; __vbaExceptHandler
  loc_0042DAFB: mov eax, fs:[00000000h]
  loc_0042DB01: push eax
  loc_0042DB02: mov fs:[00000000h], esp
  loc_0042DB09: sub esp, 00000010h
  loc_0042DB0C: push ebx
  loc_0042DB0D: push esi
  loc_0042DB0E: push edi
  loc_0042DB0F: mov var_8, esp
  loc_0042DB12: mov var_4, 00401828h
  loc_0042DB19: mov eax, Me
  loc_0042DB1C: mov var_14, 00000000h
  loc_0042DB23: push eax
  loc_0042DB24: mov ecx, [eax]
  loc_0042DB26: call [ecx+0000030Ch]
  loc_0042DB2C: lea edx, var_14
  loc_0042DB2F: push eax
  loc_0042DB30: push edx
  loc_0042DB31: call [0040103Ch] ; __vbaObjSet
  loc_0042DB37: mov esi, eax
  loc_0042DB39: push 00000000h
  loc_0042DB3B: push esi
  loc_0042DB3C: mov eax, [esi]
  loc_0042DB3E: call [eax+0000009Ch]
  loc_0042DB44: test eax, eax
  loc_0042DB46: fnclex
  loc_0042DB48: jge 0042DB5Ch
  loc_0042DB4A: push 0000009Ch
  loc_0042DB4F: push 0040E728h
  loc_0042DB54: push esi
  loc_0042DB55: push eax
  loc_0042DB56: call [00401030h] ; __vbaHresultCheckObj
  loc_0042DB5C: lea ecx, var_14
  loc_0042DB5F: call [00401114h] ; __vbaFreeObj
  loc_0042DB65: push 0042DB77h
  loc_0042DB6A: jmp 0042DB76h
  loc_0042DB6C: lea ecx, var_14
  loc_0042DB6F: call [00401114h] ; __vbaFreeObj
  loc_0042DB75: ret
  loc_0042DB76: ret
  loc_0042DB77: mov ecx, var_10
  loc_0042DB7A: pop edi
  loc_0042DB7B: pop esi
  loc_0042DB7C: xor eax, eax
  loc_0042DB7E: mov fs:[00000000h], ecx
  loc_0042DB85: pop ebx
  loc_0042DB86: mov esp, ebp
  loc_0042DB88: pop ebp
  loc_0042DB89: retn 0004h
End Sub

Private Sub Proc_16_6_42EA10
  loc_0042EA10: push ebp
  loc_0042EA11: mov ebp, esp
  loc_0042EA13: sub esp, 00000008h
  loc_0042EA16: push 00401926h ; __vbaExceptHandler
  loc_0042EA1B: mov eax, fs:[00000000h]
  loc_0042EA21: push eax
  loc_0042EA22: mov fs:[00000000h], esp
  loc_0042EA29: sub esp, 00000028h
  loc_0042EA2C: push ebx
  loc_0042EA2D: push esi
  loc_0042EA2E: push edi
  loc_0042EA2F: mov var_8, esp
  loc_0042EA32: mov var_4, 00401878h
  loc_0042EA39: mov esi, Me
  loc_0042EA3C: xor eax, eax
  loc_0042EA3E: mov var_14, eax
  loc_0042EA41: mov var_18, eax
  loc_0042EA44: mov var_1C, eax
  loc_0042EA47: mov var_20, eax
  loc_0042EA4A: mov var_24, eax
  loc_0042EA4D: mov var_28, eax
  loc_0042EA50: mov var_2C, eax
  loc_0042EA53: mov eax, [esi]
  loc_0042EA55: push esi
  loc_0042EA56: call [eax+00000300h]
  loc_0042EA5C: lea ecx, var_2C
  loc_0042EA5F: push eax
  loc_0042EA60: push ecx
  loc_0042EA61: call [0040103Ch] ; __vbaObjSet
  loc_0042EA67: mov edx, [esi+0000003Ch]
  loc_0042EA6A: mov esi, [0040102Ch] ; __vbaStrCat
  loc_0042EA70: mov ebx, [eax]
  loc_0042EA72: push 00412E1Ch ; "You wish to investigate the effect of"
  loc_0042EA77: push edx
  loc_0042EA78: mov var_30, eax
  loc_0042EA7B: call __vbaStrCat
  loc_0042EA7D: mov edi, [004010FCh] ; __vbaStrMove
  loc_0042EA83: mov edx, eax
  loc_0042EA85: lea ecx, var_14
  loc_0042EA88: call edi
  loc_0042EA8A: push eax
  loc_0042EA8B: push 00412E6Ch ; " on"
  loc_0042EA90: call __vbaStrCat
  loc_0042EA92: mov edx, eax
  loc_0042EA94: lea ecx, var_18
  loc_0042EA97: call edi
  loc_0042EA99: push eax
  loc_0042EA9A: mov eax, Me
  loc_0042EA9D: mov ecx, [eax+00000034h]
  loc_0042EAA0: push ecx
  loc_0042EAA1: call __vbaStrCat
  loc_0042EAA3: mov edx, eax
  loc_0042EAA5: lea ecx, var_1C
  loc_0042EAA8: call edi
  loc_0042EAAA: push eax
  loc_0042EAAB: push 00412E78h ; " and use this relationship to infer about"
  loc_0042EAB0: call __vbaStrCat
  loc_0042EAB2: mov edx, eax
  loc_0042EAB4: lea ecx, var_20
  loc_0042EAB7: call edi
  loc_0042EAB9: mov edx, Me
  loc_0042EABC: push eax
  loc_0042EABD: mov eax, [edx+00000034h]
  loc_0042EAC0: push eax
  loc_0042EAC1: call __vbaStrCat
  loc_0042EAC3: mov edx, eax
  loc_0042EAC5: lea ecx, var_24
  loc_0042EAC8: call edi
  loc_0042EACA: push eax
  loc_0042EACB: push 0040DD3Ch ; "."
  loc_0042EAD0: call __vbaStrCat
  loc_0042EAD2: mov edx, eax
  loc_0042EAD4: lea ecx, var_28
  loc_0042EAD7: call edi
  loc_0042EAD9: mov esi, var_30
  loc_0042EADC: push eax
  loc_0042EADD: push esi
  loc_0042EADE: call [ebx+00000054h]
  loc_0042EAE1: test eax, eax
  loc_0042EAE3: fnclex
  loc_0042EAE5: jge 0042EAF6h
  loc_0042EAE7: push 00000054h
  loc_0042EAE9: push 0040E390h
  loc_0042EAEE: push esi
  loc_0042EAEF: push eax
  loc_0042EAF0: call [00401030h] ; __vbaHresultCheckObj
  loc_0042EAF6: lea ecx, var_28
  loc_0042EAF9: lea edx, var_24
  loc_0042EAFC: push ecx
  loc_0042EAFD: lea eax, var_20
  loc_0042EB00: push edx
  loc_0042EB01: lea ecx, var_1C
  loc_0042EB04: push eax
  loc_0042EB05: lea edx, var_18
  loc_0042EB08: push ecx
  loc_0042EB09: lea eax, var_14
  loc_0042EB0C: push edx
  loc_0042EB0D: push eax
  loc_0042EB0E: push 00000006h
  loc_0042EB10: call [004010E4h] ; __vbaFreeStrList
  loc_0042EB16: add esp, 0000001Ch
  loc_0042EB19: lea ecx, var_2C
  loc_0042EB1C: call [00401114h] ; __vbaFreeObj
  loc_0042EB22: push 0042EB57h
  loc_0042EB27: jmp 0042EB56h
  loc_0042EB29: lea ecx, var_28
  loc_0042EB2C: lea edx, var_24
  loc_0042EB2F: push ecx
  loc_0042EB30: lea eax, var_20
  loc_0042EB33: push edx
  loc_0042EB34: lea ecx, var_1C
  loc_0042EB37: push eax
  loc_0042EB38: lea edx, var_18
  loc_0042EB3B: push ecx
  loc_0042EB3C: lea eax, var_14
  loc_0042EB3F: push edx
  loc_0042EB40: push eax
  loc_0042EB41: push 00000006h
  loc_0042EB43: call [004010E4h] ; __vbaFreeStrList
  loc_0042EB49: add esp, 0000001Ch
  loc_0042EB4C: lea ecx, var_2C
  loc_0042EB4F: call [00401114h] ; __vbaFreeObj
  loc_0042EB55: ret
  loc_0042EB56: ret
  loc_0042EB57: mov ecx, var_10
  loc_0042EB5A: pop edi
  loc_0042EB5B: pop esi
  loc_0042EB5C: xor eax, eax
  loc_0042EB5E: mov fs:[00000000h], ecx
  loc_0042EB65: pop ebx
  loc_0042EB66: mov esp, ebp
  loc_0042EB68: pop ebp
  loc_0042EB69: retn 0004h
End Sub
