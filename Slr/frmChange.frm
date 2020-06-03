VERSION 5.00
Begin VB.Form frmChange
  Caption = "Change Variable Characteristics"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 735
  ClientWidth = 8730
  ClientHeight = 6180
  StartUpPosition = 3 'Windows Default
  Begin VB.CommandButton cmdChange
    Caption = "Change"
    Left = 1080
    Top = 2040
    Width = 1695
    Height = 495
    TabIndex = 40
  End
  Begin VB.Frame fraAlpha
    Caption = "Value of Alpha"
    BackColor = &HC0C0C0&
    Left = 1200
    Top = 10000
    Width = 2295
    Height = 855
    TabIndex = 35
    Begin VB.TextBox txtAlpha
      Left = 240
      Top = 240
      Width = 1815
      Height = 405
      Text = "0.05"
      TabIndex = 36
    End
  End
  Begin VB.Frame fraTTable
    Caption = "T-Table Value with n-2 d.f."
    BackColor = &HC0C0C0&
    Left = 120
    Top = 10000
    Width = 2655
    Height = 735
    TabIndex = 33
    Begin VB.TextBox txtTTable
      Left = 240
      Top = 240
      Width = 2295
      Height = 405
      TabIndex = 34
    End
  End
  Begin VB.Frame fraEstimation
    Caption = "Values Needed for Prediction Interval"
    BackColor = &HC0C0C0&
    Left = 120
    Top = 10000
    Width = 8415
    Height = 1095
    TabIndex = 26
    Begin VB.Frame fraYhatg
      Caption = "Estimate of the mean Y at Xg"
      Left = 360
      Top = 240
      Width = 2295
      Height = 735
      TabIndex = 31
      Begin VB.TextBox txt
        Left = 240
        Top = 240
        Width = 1935
        Height = 375
        TabIndex = 32
      End
    End
    Begin VB.Frame Frame7
      Caption = "Xg, the Given Value of X"
      Left = 5640
      Top = 240
      Width = 2655
      Height = 735
      TabIndex = 29
      Begin VB.TextBox Text3
        Left = 240
        Top = 240
        Width = 2055
        Height = 375
        TabIndex = 30
      End
    End
    Begin VB.Frame Frame6
      Caption = "Estimate of Standard Error"
      Left = 2880
      Top = 240
      Width = 2655
      Height = 735
      TabIndex = 27
      Begin VB.TextBox Text2
        Left = 240
        Top = 240
        Width = 2055
        Height = 375
        TabIndex = 28
      End
    End
  End
  Begin VB.Frame fraPrediction
    Caption = "Values Needed for Prediction Interval"
    BackColor = &HC0C0C0&
    Left = 120
    Top = 10000
    Width = 8415
    Height = 1215
    TabIndex = 17
    Begin VB.Frame fraSyhatnew
      Caption = "Estimate of Standard Error"
      Left = 2880
      Top = 240
      Width = 2655
      Height = 735
      TabIndex = 24
      Begin VB.TextBox txtSyhatNew
        Left = 240
        Top = 240
        Width = 2055
        Height = 375
        TabIndex = 25
      End
    End
    Begin VB.Frame fraXg
      Caption = "Xg, the Given Value of X"
      Left = 5640
      Top = 240
      Width = 2655
      Height = 735
      TabIndex = 22
      Begin VB.TextBox txtXg
        Left = 240
        Top = 240
        Width = 2055
        Height = 375
        TabIndex = 23
      End
    End
    Begin VB.Frame fraYhatNewg
      Caption = "Estimate of Y at Xg"
      Left = 360
      Top = 240
      Width = 2295
      Height = 735
      TabIndex = 20
      Begin VB.TextBox Text1
        Left = 240
        Top = 240
        Width = 1935
        Height = 375
        TabIndex = 21
      End
    End
  End
  Begin VB.Frame fraSlope
    Caption = "Values Needed for Inferences on the Slope"
    BackColor = &HC0C0C0&
    Left = 120
    Top = 10000
    Width = 8415
    Height = 1095
    TabIndex = 12
    Begin VB.Frame fraHo
      Caption = "Hypothesized Value (usually zero)"
      Left = 5640
      Top = 240
      Width = 2655
      Height = 735
      TabIndex = 18
      Begin VB.TextBox txtHo
        Left = 240
        Top = 240
        Width = 2055
        Height = 375
        Text = "0"
        TabIndex = 19
      End
    End
    Begin VB.Frame fraStanErr
      Caption = "Estimate of Standard Error"
      Left = 2880
      Top = 240
      Width = 2655
      Height = 735
      TabIndex = 15
      Begin VB.TextBox txtSBetaHat
        Left = 240
        Top = 240
        Width = 2175
        Height = 405
        TabIndex = 16
      End
    End
    Begin VB.Frame fraSlopeEstimate
      Caption = "Estimate of Slope"
      Left = 360
      Top = 240
      Width = 2295
      Height = 735
      TabIndex = 13
      Begin VB.TextBox txtBetaHat
        Left = 240
        Top = 240
        Width = 1935
        Height = 375
        TabIndex = 14
      End
    End
  End
  Begin VB.Frame fraUnits
    Caption = "Units"
    BackColor = &HC0C0C0&
    Left = 5280
    Top = 10000
    Width = 3375
    Height = 1815
    Visible = 0   'False
    TabIndex = 7
    Begin VB.Frame Frame5
      Caption = "Independent Variable Units"
      Left = 120
      Top = 960
      Width = 3015
      Height = 735
      TabIndex = 10
      Begin VB.TextBox txtXUnits
        Left = 240
        Top = 240
        Width = 2655
        Height = 375
        TabIndex = 11
      End
    End
    Begin VB.Frame Frame4
      Caption = "Dependent Variable Units"
      Left = 120
      Top = 240
      Width = 3015
      Height = 735
      TabIndex = 8
      Begin VB.TextBox txtYUnits
        Left = 240
        Top = 240
        Width = 2655
        Height = 375
        TabIndex = 9
      End
    End
  End
  Begin VB.CommandButton cmdRestore
    Caption = "Reset Defaults"
    Left = 4680
    Top = 2040
    Width = 1455
    Height = 495
    TabIndex = 6
  End
  Begin VB.CommandButton cmdClearChange
    Caption = "Clear"
    Left = 3000
    Top = 2040
    Width = 1335
    Height = 495
    TabIndex = 5
  End
  Begin VB.Frame fraNames
    Caption = "Names"
    BackColor = &HC0C0C0&
    Left = 120
    Top = 10000
    Width = 5055
    Height = 2535
    Visible = 0   'False
    TabIndex = 0
    Begin VB.CommandButton cmdNamesReset
      Caption = "Reset Defaults"
      Left = 3240
      Top = 1920
      Width = 1575
      Height = 375
      TabIndex = 39
    End
    Begin VB.CommandButton cmdNamesClear
      Caption = "Clear"
      Left = 1800
      Top = 1920
      Width = 1215
      Height = 375
      TabIndex = 38
    End
    Begin VB.CommandButton cmdNamesChange
      Caption = "Change"
      Left = 240
      Top = 1920
      Width = 1335
      Height = 375
      TabIndex = 37
    End
    Begin VB.Frame Frame1
      Caption = "Dependent Variable Name"
      Left = 120
      Top = 240
      Width = 4815
      Height = 735
      TabIndex = 3
      Begin VB.TextBox txtYName
        Left = 240
        Top = 240
        Width = 4335
        Height = 375
        TabIndex = 4
      End
    End
    Begin VB.Frame Frame2
      Caption = "Independent Variable Name"
      Left = 120
      Top = 960
      Width = 4815
      Height = 735
      TabIndex = 1
      Begin VB.TextBox txtXName
        Left = 240
        Top = 240
        Width = 4335
        Height = 375
        TabIndex = 2
      End
    End
  End
  Begin VB.Menu mnuChange
    Caption = "Characteristic to Change"
    Begin VB.Menu mnuNames
      Caption = "Variable Names"
    End
    Begin VB.Menu mnuUnits
      Caption = "Variable Units"
    End
    Begin VB.Menu mnuSlope
      Caption = "Values for Slope Inferences"
    End
    Begin VB.Menu mnuPrediction
      Caption = "Values for Prediction Intervals"
    End
    Begin VB.Menu mnuEstimation
      Caption = "Values for Estimation Intervals"
    End
    Begin VB.Menu mnuAlpha
      Caption = "Alpha"
    End
    Begin VB.Menu mnuTtable
      Caption = "T-Table Value"
    End
  End
End

Attribute VB_Name = "frmChange"


Private Sub mnuNames_Click() '41B770
  loc_0041B770: push ebp
  loc_0041B771: mov ebp, esp
  loc_0041B773: sub esp, 0000000Ch
  loc_0041B776: push 00401926h ; __vbaExceptHandler
  loc_0041B77B: mov eax, fs:[00000000h]
  loc_0041B781: push eax
  loc_0041B782: mov fs:[00000000h], esp
  loc_0041B789: sub esp, 00000014h
  loc_0041B78C: push ebx
  loc_0041B78D: push esi
  loc_0041B78E: push edi
  loc_0041B78F: mov var_C, esp
  loc_0041B792: mov var_8, 004011A8h
  loc_0041B799: mov edi, Me
  loc_0041B79C: mov eax, edi
  loc_0041B79E: and eax, 00000001h
  loc_0041B7A1: mov var_4, eax
  loc_0041B7A4: and edi, FFFFFFFEh
  loc_0041B7A7: push edi
  loc_0041B7A8: mov Me, edi
  loc_0041B7AB: mov ecx, [edi]
  loc_0041B7AD: call [ecx+00000004h]
  loc_0041B7B0: mov eax, [00430100h]
  loc_0041B7B5: xor ebx, ebx
  loc_0041B7B7: cmp eax, ebx
  loc_0041B7B9: mov var_18, ebx
  loc_0041B7BC: jnz 0041B7D3h
  loc_0041B7BE: push 00430100h
  loc_0041B7C3: push 0040BA38h
  loc_0041B7C8: call [004010D4h] ; __vbaNew2
  loc_0041B7CE: mov eax, [00430100h]
  loc_0041B7D3: mov edx, [eax]
  loc_0041B7D5: push eax
  loc_0041B7D6: call [edx+00000380h]
  loc_0041B7DC: push eax
  loc_0041B7DD: lea eax, var_18
  loc_0041B7E0: push eax
  loc_0041B7E1: call [0040103Ch] ; __vbaObjSet
  loc_0041B7E7: mov esi, eax
  loc_0041B7E9: push 44160000h
  loc_0041B7EE: push esi
  loc_0041B7EF: mov ecx, [esi]
  loc_0041B7F1: call [ecx+0000007Ch]
  loc_0041B7F4: cmp eax, ebx
  loc_0041B7F6: fnclex
  loc_0041B7F8: jge 0041B809h
  loc_0041B7FA: push 0000007Ch
  loc_0041B7FC: push 0040E728h
  loc_0041B801: push esi
  loc_0041B802: push eax
  loc_0041B803: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B809: lea ecx, var_18
  loc_0041B80C: call [00401114h] ; __vbaFreeObj
  loc_0041B812: mov [edi+00000034h], 0001h
  loc_0041B818: mov var_4, ebx
  loc_0041B81B: fwait
  loc_0041B81C: push 0041B82Eh
  loc_0041B821: jmp 0041B82Dh
  loc_0041B823: lea ecx, var_18
  loc_0041B826: call [00401114h] ; __vbaFreeObj
  loc_0041B82C: ret
  loc_0041B82D: ret
  loc_0041B82E: mov eax, Me
  loc_0041B831: push eax
  loc_0041B832: mov edx, [eax]
  loc_0041B834: call [edx+00000008h]
  loc_0041B837: mov eax, var_4
  loc_0041B83A: mov ecx, var_14
  loc_0041B83D: pop edi
  loc_0041B83E: pop esi
  loc_0041B83F: mov fs:[00000000h], ecx
  loc_0041B846: pop ebx
  loc_0041B847: mov esp, ebp
  loc_0041B849: pop ebp
  loc_0041B84A: retn 0004h
End Sub

Private Sub cmdChange_Click() '41B850
  loc_0041B850: push ebp
  loc_0041B851: mov ebp, esp
  loc_0041B853: sub esp, 0000000Ch
  loc_0041B856: push 00401926h ; __vbaExceptHandler
  loc_0041B85B: mov eax, fs:[00000000h]
  loc_0041B861: push eax
  loc_0041B862: mov fs:[00000000h], esp
  loc_0041B869: sub esp, 000000ACh
  loc_0041B86F: push ebx
  loc_0041B870: push esi
  loc_0041B871: push edi
  loc_0041B872: mov var_C, esp
  loc_0041B875: mov var_8, 004011B8h
  loc_0041B87C: mov esi, Me
  loc_0041B87F: mov eax, esi
  loc_0041B881: and eax, 00000001h
  loc_0041B884: mov var_4, eax
  loc_0041B887: and esi, FFFFFFFEh
  loc_0041B88A: push esi
  loc_0041B88B: mov Me, esi
  loc_0041B88E: mov ecx, [esi]
  loc_0041B890: call [ecx+00000004h]
  loc_0041B893: xor eax, eax
  loc_0041B895: mov var_18, eax
  loc_0041B898: mov var_1C, eax
  loc_0041B89B: mov var_2C, eax
  loc_0041B89E: mov var_3C, eax
  loc_0041B8A1: mov var_4C, eax
  loc_0041B8A4: mov var_5C, eax
  loc_0041B8A7: mov var_6C, eax
  loc_0041B8AA: mov var_7C, eax
  loc_0041B8AD: movsx eax, [esi+00000034h]
  loc_0041B8B1: dec eax
  loc_0041B8B2: jz 0041BF20h
  loc_0041B8B8: dec eax
  loc_0041B8B9: jz 0041BBDEh
  loc_0041B8BF: dec eax
  loc_0041B8C0: jnz 0041C2D6h
  loc_0041B8C6: mov edx, [esi]
  loc_0041B8C8: push esi
  loc_0041B8C9: call [edx+00000360h]
  loc_0041B8CF: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041B8D5: push eax
  loc_0041B8D6: lea eax, var_1C
  loc_0041B8D9: push eax
  loc_0041B8DA: call ebx
  loc_0041B8DC: mov edi, eax
  loc_0041B8DE: lea edx, var_18
  loc_0041B8E1: push edx
  loc_0041B8E2: push edi
  loc_0041B8E3: mov ecx, [edi]
  loc_0041B8E5: call [ecx+000000A0h]
  loc_0041B8EB: test eax, eax
  loc_0041B8ED: fnclex
  loc_0041B8EF: jge 0041B903h
  loc_0041B8F1: push 000000A0h
  loc_0041B8F6: push 0040F02Ch
  loc_0041B8FB: push edi
  loc_0041B8FC: push eax
  loc_0041B8FD: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B903: mov eax, var_18
  loc_0041B906: push eax
  loc_0041B907: call [00401074h] ; __vbaR4Str
  loc_0041B90D: mov edi, [00401110h] ; __vbaFreeStr
  loc_0041B913: lea ecx, var_18
  loc_0041B916: fstp real4 ptr [00430060h]
  loc_0041B91C: call edi
  loc_0041B91E: lea ecx, var_1C
  loc_0041B921: call [00401114h] ; __vbaFreeObj
  loc_0041B927: mov ecx, [esi]
  loc_0041B929: push esi
  loc_0041B92A: call [ecx+00000358h]
  loc_0041B930: lea edx, var_1C
  loc_0041B933: push eax
  loc_0041B934: push edx
  loc_0041B935: call ebx
  loc_0041B937: mov ebx, eax
  loc_0041B939: lea ecx, var_18
  loc_0041B93C: push ecx
  loc_0041B93D: push ebx
  loc_0041B93E: mov eax, [ebx]
  loc_0041B940: call [eax+000000A0h]
  loc_0041B946: test eax, eax
  loc_0041B948: fnclex
  loc_0041B94A: jge 0041B95Eh
  loc_0041B94C: push 000000A0h
  loc_0041B951: push 0040F02Ch
  loc_0041B956: push ebx
  loc_0041B957: push eax
  loc_0041B958: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B95E: mov edx, var_18
  loc_0041B961: push edx
  loc_0041B962: call [00401074h] ; __vbaR4Str
  loc_0041B968: fstp real4 ptr [00430064h]
  loc_0041B96E: lea ecx, var_18
  loc_0041B971: call edi
  loc_0041B973: lea ecx, var_1C
  loc_0041B976: call [00401114h] ; __vbaFreeObj
  loc_0041B97C: mov eax, [esi]
  loc_0041B97E: push esi
  loc_0041B97F: call [eax+00000350h]
  loc_0041B985: lea ecx, var_1C
  loc_0041B988: push eax
  loc_0041B989: push ecx
  loc_0041B98A: call [0040103Ch] ; __vbaObjSet
  loc_0041B990: mov ebx, eax
  loc_0041B992: lea eax, var_18
  loc_0041B995: push eax
  loc_0041B996: push ebx
  loc_0041B997: mov edx, [ebx]
  loc_0041B999: call [edx+000000A0h]
  loc_0041B99F: test eax, eax
  loc_0041B9A1: fnclex
  loc_0041B9A3: jge 0041B9B7h
  loc_0041B9A5: push 000000A0h
  loc_0041B9AA: push 0040F02Ch
  loc_0041B9AF: push ebx
  loc_0041B9B0: push eax
  loc_0041B9B1: call [00401030h] ; __vbaHresultCheckObj
  loc_0041B9B7: mov ecx, var_18
  loc_0041B9BA: push ecx
  loc_0041B9BB: call [00401074h] ; __vbaR4Str
  loc_0041B9C1: fstp real4 ptr [0043005Ch]
  loc_0041B9C7: lea ecx, var_18
  loc_0041B9CA: call edi
  loc_0041B9CC: lea ecx, var_1C
  loc_0041B9CF: call [00401114h] ; __vbaFreeObj
  loc_0041B9D5: mov edx, [esi]
  loc_0041B9D7: push esi
  loc_0041B9D8: call [edx+00000360h]
  loc_0041B9DE: mov var_24, eax
  loc_0041B9E1: lea eax, var_2C
  loc_0041B9E4: lea ecx, var_3C
  loc_0041B9E7: push eax
  loc_0041B9E8: push ecx
  loc_0041B9E9: mov var_2C, 00000009h
  loc_0041B9F0: call [00401050h] ; rtcTrimVar
  loc_0041B9F6: lea edx, var_3C
  loc_0041B9F9: lea eax, var_6C
  loc_0041B9FC: push edx
  loc_0041B9FD: push eax
  loc_0041B9FE: mov var_64, 0040F040h
  loc_0041BA05: mov var_6C, 00008008h
  loc_0041BA0C: call [00401070h] ; __vbaVarTstEq
  loc_0041BA12: mov ebx, [00401018h] ; __vbaFreeVarList
  loc_0041BA18: lea ecx, var_3C
  loc_0041BA1B: lea edx, var_2C
  loc_0041BA1E: push ecx
  loc_0041BA1F: push edx
  loc_0041BA20: push 00000002h
  loc_0041BA22: mov di, ax
  loc_0041BA25: call ebx
  loc_0041BA27: add esp, 0000000Ch
  loc_0041BA2A: test di, di
  loc_0041BA2D: jz 0041BAECh
  loc_0041BA33: mov eax, [esi]
  loc_0041BA35: push esi
  loc_0041BA36: call [eax+00000360h]
  loc_0041BA3C: lea ecx, var_1C
  loc_0041BA3F: push eax
  loc_0041BA40: push ecx
  loc_0041BA41: call [0040103Ch] ; __vbaObjSet
  loc_0041BA47: mov esi, eax
  loc_0041BA49: push esi
  loc_0041BA4A: mov edx, [esi]
  loc_0041BA4C: call [edx+00000204h]
  loc_0041BA52: test eax, eax
  loc_0041BA54: fnclex
  loc_0041BA56: jge 0041BA6Ah
  loc_0041BA58: push 00000204h
  loc_0041BA5D: push 0040F02Ch
  loc_0041BA62: push esi
  loc_0041BA63: push eax
  loc_0041BA64: call [00401030h] ; __vbaHresultCheckObj
  loc_0041BA6A: lea ecx, var_1C
  loc_0041BA6D: call [00401114h] ; __vbaFreeObj
  loc_0041BA73: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041BA79: mov ecx, 80020004h
  loc_0041BA7E: mov var_54, ecx
  loc_0041BA81: mov eax, 0000000Ah
  loc_0041BA86: mov var_44, ecx
  loc_0041BA89: mov edi, 00000008h
  loc_0041BA8E: lea edx, var_7C
  loc_0041BA91: lea ecx, var_3C
  loc_0041BA94: mov var_5C, eax
  loc_0041BA97: mov var_4C, eax
  loc_0041BA9A: mov var_74, 0040F27Ch ; "Check Number"
  loc_0041BAA1: mov var_7C, edi
  loc_0041BAA4: call __vbaVarDup
  loc_0041BAA6: lea edx, var_6C
  loc_0041BAA9: lea ecx, var_2C
  loc_0041BAAC: mov var_64, 0040F048h ; "Please enter the estimate of the slope."
  loc_0041BAB3: mov var_6C, edi
  loc_0041BAB6: call __vbaVarDup
  loc_0041BAB8: lea eax, var_5C
  loc_0041BABB: lea ecx, var_4C
  loc_0041BABE: push eax
  loc_0041BABF: lea edx, var_3C
  loc_0041BAC2: push ecx
  loc_0041BAC3: push edx
  loc_0041BAC4: lea eax, var_2C
  loc_0041BAC7: push 00000000h
  loc_0041BAC9: push eax
  loc_0041BACA: call [00401038h] ; rtcMsgBox
  loc_0041BAD0: lea ecx, var_5C
  loc_0041BAD3: lea edx, var_4C
  loc_0041BAD6: push ecx
  loc_0041BAD7: lea eax, var_3C
  loc_0041BADA: push edx
  loc_0041BADB: lea ecx, var_2C
  loc_0041BADE: push eax
  loc_0041BADF: push ecx
  loc_0041BAE0: push 00000004h
  loc_0041BAE2: call ebx
  loc_0041BAE4: add esp, 00000014h
  loc_0041BAE7: jmp 0041C2D6h
  loc_0041BAEC: mov edx, [esi]
  loc_0041BAEE: push esi
  loc_0041BAEF: call [edx+00000358h]
  loc_0041BAF5: mov var_24, eax
  loc_0041BAF8: lea eax, var_2C
  loc_0041BAFB: lea ecx, var_3C
  loc_0041BAFE: push eax
  loc_0041BAFF: push ecx
  loc_0041BB00: mov var_2C, 00000009h
  loc_0041BB07: call [00401050h] ; rtcTrimVar
  loc_0041BB0D: lea edx, var_3C
  loc_0041BB10: lea eax, var_6C
  loc_0041BB13: push edx
  loc_0041BB14: push eax
  loc_0041BB15: mov var_64, 0040F040h
  loc_0041BB1C: mov var_6C, 00008008h
  loc_0041BB23: call [00401070h] ; __vbaVarTstEq
  loc_0041BB29: lea ecx, var_3C
  loc_0041BB2C: lea edx, var_2C
  loc_0041BB2F: push ecx
  loc_0041BB30: push edx
  loc_0041BB31: push 00000002h
  loc_0041BB33: mov di, ax
  loc_0041BB36: call ebx
  loc_0041BB38: add esp, 0000000Ch
  loc_0041BB3B: test di, di
  loc_0041BB3E: jz 0041C2D6h
  loc_0041BB44: mov ecx, 80020004h
  loc_0041BB49: mov eax, 0000000Ah
  loc_0041BB4E: mov var_54, ecx
  loc_0041BB51: mov var_44, ecx
  loc_0041BB54: mov edi, 00000008h
  loc_0041BB59: lea edx, var_7C
  loc_0041BB5C: lea ecx, var_3C
  loc_0041BB5F: mov var_5C, eax
  loc_0041BB62: mov var_4C, eax
  loc_0041BB65: mov var_74, 0040F1F0h ; "Check Units"
  loc_0041BB6C: mov var_7C, edi
  loc_0041BB6F: call [004010F4h] ; __vbaVarDup
  loc_0041BB75: lea edx, var_6C
  loc_0041BB78: lea ecx, var_2C
  loc_0041BB7B: mov var_64, 0040F29Ch ; "Please enter the estimate of the standard error."
  loc_0041BB82: mov var_6C, edi
  loc_0041BB85: call [004010F4h] ; __vbaVarDup
  loc_0041BB8B: lea eax, var_5C
  loc_0041BB8E: lea ecx, var_4C
  loc_0041BB91: push eax
  loc_0041BB92: lea edx, var_3C
  loc_0041BB95: push ecx
  loc_0041BB96: push edx
  loc_0041BB97: lea eax, var_2C
  loc_0041BB9A: push 00000000h
  loc_0041BB9C: push eax
  loc_0041BB9D: call [00401038h] ; rtcMsgBox
  loc_0041BBA3: lea ecx, var_5C
  loc_0041BBA6: lea edx, var_4C
  loc_0041BBA9: push ecx
  loc_0041BBAA: lea eax, var_3C
  loc_0041BBAD: push edx
  loc_0041BBAE: lea ecx, var_2C
  loc_0041BBB1: push eax
  loc_0041BBB2: push ecx
  loc_0041BBB3: push 00000004h
  loc_0041BBB5: call ebx
  loc_0041BBB7: mov edx, [esi]
  loc_0041BBB9: add esp, 00000014h
  loc_0041BBBC: push esi
  loc_0041BBBD: call [edx+00000358h]
  loc_0041BBC3: push eax
  loc_0041BBC4: lea eax, var_1C
  loc_0041BBC7: push eax
  loc_0041BBC8: call [0040103Ch] ; __vbaObjSet
  loc_0041BBCE: mov esi, eax
  loc_0041BBD0: push esi
  loc_0041BBD1: mov ecx, [esi]
  loc_0041BBD3: call [ecx+00000204h]
  loc_0041BBD9: jmp 0041C285h
  loc_0041BBDE: mov edx, [esi]
  loc_0041BBE0: mov edi, 00000008h
  loc_0041BBE5: push esi
  loc_0041BBE6: mov var_64, 0040F028h
  loc_0041BBED: mov var_6C, edi
  loc_0041BBF0: call [edx+00000374h]
  loc_0041BBF6: push eax
  loc_0041BBF7: lea eax, var_1C
  loc_0041BBFA: push eax
  loc_0041BBFB: call [0040103Ch] ; __vbaObjSet
  loc_0041BC01: mov ebx, eax
  loc_0041BC03: lea edx, var_18
  loc_0041BC06: push edx
  loc_0041BC07: push ebx
  loc_0041BC08: mov ecx, [ebx]
  loc_0041BC0A: call [ecx+000000A0h]
  loc_0041BC10: test eax, eax
  loc_0041BC12: fnclex
  loc_0041BC14: jge 0041BC28h
  loc_0041BC16: push 000000A0h
  loc_0041BC1B: push 0040F02Ch
  loc_0041BC20: push ebx
  loc_0041BC21: push eax
  loc_0041BC22: call [00401030h] ; __vbaHresultCheckObj
  loc_0041BC28: mov eax, var_18
  loc_0041BC2B: lea ecx, var_3C
  loc_0041BC2E: mov var_24, eax
  loc_0041BC31: lea eax, var_2C
  loc_0041BC34: push eax
  loc_0041BC35: push ecx
  loc_0041BC36: mov var_18, 00000000h
  loc_0041BC3D: mov var_2C, edi
  loc_0041BC40: call [00401050h] ; rtcTrimVar
  loc_0041BC46: lea edx, var_6C
  loc_0041BC49: lea eax, var_3C
  loc_0041BC4C: push edx
  loc_0041BC4D: lea ecx, var_4C
  loc_0041BC50: push eax
  loc_0041BC51: push ecx
  loc_0041BC52: call [004010C0h] ; __vbaVarCat
  loc_0041BC58: push eax
  loc_0041BC59: call [00401014h] ; __vbaStrVarMove
  loc_0041BC5F: mov edx, eax
  loc_0041BC61: mov ecx, 00430014h
  loc_0041BC66: call [004010FCh] ; __vbaStrMove
  loc_0041BC6C: lea ecx, var_1C
  loc_0041BC6F: call [00401114h] ; __vbaFreeObj
  loc_0041BC75: mov ebx, [00401018h] ; __vbaFreeVarList
  loc_0041BC7B: lea edx, var_4C
  loc_0041BC7E: lea eax, var_3C
  loc_0041BC81: push edx
  loc_0041BC82: lea ecx, var_2C
  loc_0041BC85: push eax
  loc_0041BC86: push ecx
  loc_0041BC87: push 00000003h
  loc_0041BC89: call ebx
  loc_0041BC8B: mov edx, [esi]
  loc_0041BC8D: add esp, 00000010h
  loc_0041BC90: mov var_64, 0040F028h
  loc_0041BC97: mov var_6C, edi
  loc_0041BC9A: push esi
  loc_0041BC9B: call [edx+0000036Ch]
  loc_0041BCA1: push eax
  loc_0041BCA2: lea eax, var_1C
  loc_0041BCA5: push eax
  loc_0041BCA6: call [0040103Ch] ; __vbaObjSet
  loc_0041BCAC: mov ecx, [eax]
  loc_0041BCAE: lea edx, var_18
  loc_0041BCB1: push edx
  loc_0041BCB2: push eax
  loc_0041BCB3: mov var_A0, eax
  loc_0041BCB9: call [ecx+000000A0h]
  loc_0041BCBF: test eax, eax
  loc_0041BCC1: fnclex
  loc_0041BCC3: jge 0041BCDDh
  loc_0041BCC5: mov ecx, var_A0
  loc_0041BCCB: push 000000A0h
  loc_0041BCD0: push 0040F02Ch
  loc_0041BCD5: push ecx
  loc_0041BCD6: push eax
  loc_0041BCD7: call [00401030h] ; __vbaHresultCheckObj
  loc_0041BCDD: mov eax, var_18
  loc_0041BCE0: lea edx, var_2C
  loc_0041BCE3: mov var_24, eax
  loc_0041BCE6: lea eax, var_3C
  loc_0041BCE9: push edx
  loc_0041BCEA: push eax
  loc_0041BCEB: mov var_18, 00000000h
  loc_0041BCF2: mov var_2C, edi
  loc_0041BCF5: call [00401050h] ; rtcTrimVar
  loc_0041BCFB: lea ecx, var_6C
  loc_0041BCFE: lea edx, var_3C
  loc_0041BD01: push ecx
  loc_0041BD02: lea eax, var_4C
  loc_0041BD05: push edx
  loc_0041BD06: push eax
  loc_0041BD07: call [004010C0h] ; __vbaVarCat
  loc_0041BD0D: push eax
  loc_0041BD0E: call [00401014h] ; __vbaStrVarMove
  loc_0041BD14: mov edx, eax
  loc_0041BD16: mov ecx, 0043001Ch
  loc_0041BD1B: call [004010FCh] ; __vbaStrMove
  loc_0041BD21: lea ecx, var_1C
  loc_0041BD24: call [00401114h] ; __vbaFreeObj
  loc_0041BD2A: lea ecx, var_4C
  loc_0041BD2D: lea edx, var_3C
  loc_0041BD30: push ecx
  loc_0041BD31: lea eax, var_2C
  loc_0041BD34: push edx
  loc_0041BD35: push eax
  loc_0041BD36: push 00000003h
  loc_0041BD38: call ebx
  loc_0041BD3A: mov ecx, [esi]
  loc_0041BD3C: add esp, 00000010h
  loc_0041BD3F: push esi
  loc_0041BD40: call [ecx+00000374h]
  loc_0041BD46: mov var_24, eax
  loc_0041BD49: lea edx, var_2C
  loc_0041BD4C: lea eax, var_3C
  loc_0041BD4F: push edx
  loc_0041BD50: push eax
  loc_0041BD51: mov var_2C, 00000009h
  loc_0041BD58: call [00401050h] ; rtcTrimVar
  loc_0041BD5E: lea ecx, var_3C
  loc_0041BD61: lea edx, var_6C
  loc_0041BD64: push ecx
  loc_0041BD65: push edx
  loc_0041BD66: mov var_64, 0040F040h
  loc_0041BD6D: mov var_6C, 00008008h
  loc_0041BD74: call [00401070h] ; __vbaVarTstEq
  loc_0041BD7A: mov var_A0, ax
  loc_0041BD81: lea eax, var_3C
  loc_0041BD84: lea ecx, var_2C
  loc_0041BD87: push eax
  loc_0041BD88: push ecx
  loc_0041BD89: push 00000002h
  loc_0041BD8B: call ebx
  loc_0041BD8D: add esp, 0000000Ch
  loc_0041BD90: cmp var_A0, 0000h
  loc_0041BD98: jz 0041BE14h
  loc_0041BD9A: mov edx, [esi]
  loc_0041BD9C: push esi
  loc_0041BD9D: call [edx+00000374h]
  loc_0041BDA3: push eax
  loc_0041BDA4: lea eax, var_1C
  loc_0041BDA7: push eax
  loc_0041BDA8: call [0040103Ch] ; __vbaObjSet
  loc_0041BDAE: mov esi, eax
  loc_0041BDB0: push esi
  loc_0041BDB1: mov ecx, [esi]
  loc_0041BDB3: call [ecx+00000204h]
  loc_0041BDB9: test eax, eax
  loc_0041BDBB: fnclex
  loc_0041BDBD: jge 0041BDD1h
  loc_0041BDBF: push 00000204h
  loc_0041BDC4: push 0040F02Ch
  loc_0041BDC9: push esi
  loc_0041BDCA: push eax
  loc_0041BDCB: call [00401030h] ; __vbaHresultCheckObj
  loc_0041BDD1: lea ecx, var_1C
  loc_0041BDD4: call [00401114h] ; __vbaFreeObj
  loc_0041BDDA: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041BDE0: mov ecx, 80020004h
  loc_0041BDE5: mov var_54, ecx
  loc_0041BDE8: mov eax, 0000000Ah
  loc_0041BDED: mov var_44, ecx
  loc_0041BDF0: lea edx, var_7C
  loc_0041BDF3: lea ecx, var_3C
  loc_0041BDF6: mov var_5C, eax
  loc_0041BDF9: mov var_4C, eax
  loc_0041BDFC: mov var_74, 0040F1F0h ; "Check Units"
  loc_0041BE03: mov var_7C, edi
  loc_0041BE06: call __vbaVarDup
  loc_0041BE08: mov var_64, 0040F184h ; "Please enter the units for the dependent variable."
  loc_0041BE0F: jmp 0041C155h
  loc_0041BE14: mov ecx, [esi]
  loc_0041BE16: push esi
  loc_0041BE17: call [ecx+0000036Ch]
  loc_0041BE1D: mov var_24, eax
  loc_0041BE20: lea edx, var_2C
  loc_0041BE23: lea eax, var_3C
  loc_0041BE26: push edx
  loc_0041BE27: push eax
  loc_0041BE28: mov var_2C, 00000009h
  loc_0041BE2F: call [00401050h] ; rtcTrimVar
  loc_0041BE35: lea ecx, var_3C
  loc_0041BE38: lea edx, var_6C
  loc_0041BE3B: push ecx
  loc_0041BE3C: push edx
  loc_0041BE3D: mov var_64, 0040F040h
  loc_0041BE44: mov var_6C, 00008008h
  loc_0041BE4B: call [00401070h] ; __vbaVarTstEq
  loc_0041BE51: mov var_A0, ax
  loc_0041BE58: lea eax, var_3C
  loc_0041BE5B: lea ecx, var_2C
  loc_0041BE5E: push eax
  loc_0041BE5F: push ecx
  loc_0041BE60: push 00000002h
  loc_0041BE62: call ebx
  loc_0041BE64: add esp, 0000000Ch
  loc_0041BE67: cmp var_A0, 0000h
  loc_0041BE6F: jz 0041BEF0h
  loc_0041BE71: mov ecx, 80020004h
  loc_0041BE76: mov eax, 0000000Ah
  loc_0041BE7B: mov var_54, ecx
  loc_0041BE7E: mov var_44, ecx
  loc_0041BE81: lea edx, var_7C
  loc_0041BE84: lea ecx, var_3C
  loc_0041BE87: mov var_5C, eax
  loc_0041BE8A: mov var_4C, eax
  loc_0041BE8D: mov var_74, 0040F1F0h ; "Check Units"
  loc_0041BE94: mov var_7C, edi
  loc_0041BE97: call [004010F4h] ; __vbaVarDup
  loc_0041BE9D: lea edx, var_6C
  loc_0041BEA0: lea ecx, var_2C
  loc_0041BEA3: mov var_64, 0040F20Ch ; "Please enter the units for the independent variable."
  loc_0041BEAA: mov var_6C, edi
  loc_0041BEAD: call [004010F4h] ; __vbaVarDup
  loc_0041BEB3: lea edx, var_5C
  loc_0041BEB6: lea eax, var_4C
  loc_0041BEB9: push edx
  loc_0041BEBA: lea ecx, var_3C
  loc_0041BEBD: push eax
  loc_0041BEBE: push ecx
  loc_0041BEBF: lea edx, var_2C
  loc_0041BEC2: push 00000000h
  loc_0041BEC4: push edx
  loc_0041BEC5: call [00401038h] ; rtcMsgBox
  loc_0041BECB: lea eax, var_5C
  loc_0041BECE: lea ecx, var_4C
  loc_0041BED1: push eax
  loc_0041BED2: lea edx, var_3C
  loc_0041BED5: push ecx
  loc_0041BED6: lea eax, var_2C
  loc_0041BED9: push edx
  loc_0041BEDA: push eax
  loc_0041BEDB: push 00000004h
  loc_0041BEDD: call ebx
  loc_0041BEDF: mov ecx, [esi]
  loc_0041BEE1: add esp, 00000014h
  loc_0041BEE4: push esi
  loc_0041BEE5: call [ecx+0000036Ch]
  loc_0041BEEB: jmp 0041C26Fh
  loc_0041BEF0: mov ecx, [esi]
  loc_0041BEF2: push esi
  loc_0041BEF3: call [ecx+00000364h]
  loc_0041BEF9: lea edx, var_1C
  loc_0041BEFC: push eax
  loc_0041BEFD: push edx
  loc_0041BEFE: call [0040103Ch] ; __vbaObjSet
  loc_0041BF04: mov esi, eax
  loc_0041BF06: push 461C4000h
  loc_0041BF0B: push esi
  loc_0041BF0C: mov eax, [esi]
  loc_0041BF0E: call [eax+0000007Ch]
  loc_0041BF11: test eax, eax
  loc_0041BF13: fnclex
  loc_0041BF15: jge 0041C2CDh
  loc_0041BF1B: jmp 0041C2BEh
  loc_0041BF20: mov ecx, [esi]
  loc_0041BF22: mov edi, 00000008h
  loc_0041BF27: push esi
  loc_0041BF28: mov var_64, 0040F028h
  loc_0041BF2F: mov var_6C, edi
  loc_0041BF32: call [ecx+00000394h]
  loc_0041BF38: lea edx, var_1C
  loc_0041BF3B: push eax
  loc_0041BF3C: push edx
  loc_0041BF3D: call [0040103Ch] ; __vbaObjSet
  loc_0041BF43: mov ebx, eax
  loc_0041BF45: lea ecx, var_18
  loc_0041BF48: push ecx
  loc_0041BF49: push ebx
  loc_0041BF4A: mov eax, [ebx]
  loc_0041BF4C: call [eax+000000A0h]
  loc_0041BF52: test eax, eax
  loc_0041BF54: fnclex
  loc_0041BF56: jge 0041BF6Ah
  loc_0041BF58: push 000000A0h
  loc_0041BF5D: push 0040F02Ch
  loc_0041BF62: push ebx
  loc_0041BF63: push eax
  loc_0041BF64: call [00401030h] ; __vbaHresultCheckObj
  loc_0041BF6A: mov eax, var_18
  loc_0041BF6D: lea edx, var_2C
  loc_0041BF70: mov var_24, eax
  loc_0041BF73: lea eax, var_3C
  loc_0041BF76: push edx
  loc_0041BF77: push eax
  loc_0041BF78: mov var_18, 00000000h
  loc_0041BF7F: mov var_2C, edi
  loc_0041BF82: call [00401050h] ; rtcTrimVar
  loc_0041BF88: lea ecx, var_6C
  loc_0041BF8B: lea edx, var_3C
  loc_0041BF8E: push ecx
  loc_0041BF8F: lea eax, var_4C
  loc_0041BF92: push edx
  loc_0041BF93: push eax
  loc_0041BF94: call [004010C0h] ; __vbaVarCat
  loc_0041BF9A: push eax
  loc_0041BF9B: call [00401014h] ; __vbaStrVarMove
  loc_0041BFA1: mov edx, eax
  loc_0041BFA3: mov ecx, 00430010h
  loc_0041BFA8: call [004010FCh] ; __vbaStrMove
  loc_0041BFAE: lea ecx, var_1C
  loc_0041BFB1: call [00401114h] ; __vbaFreeObj
  loc_0041BFB7: mov ebx, [00401018h] ; __vbaFreeVarList
  loc_0041BFBD: lea ecx, var_4C
  loc_0041BFC0: lea edx, var_3C
  loc_0041BFC3: push ecx
  loc_0041BFC4: lea eax, var_2C
  loc_0041BFC7: push edx
  loc_0041BFC8: push eax
  loc_0041BFC9: push 00000003h
  loc_0041BFCB: call ebx
  loc_0041BFCD: mov ecx, [esi]
  loc_0041BFCF: add esp, 00000010h
  loc_0041BFD2: mov var_64, 0040F028h
  loc_0041BFD9: mov var_6C, edi
  loc_0041BFDC: push esi
  loc_0041BFDD: call [ecx+0000039Ch]
  loc_0041BFE3: lea edx, var_1C
  loc_0041BFE6: push eax
  loc_0041BFE7: push edx
  loc_0041BFE8: call [0040103Ch] ; __vbaObjSet
  loc_0041BFEE: mov ecx, [eax]
  loc_0041BFF0: lea edx, var_18
  loc_0041BFF3: push edx
  loc_0041BFF4: push eax
  loc_0041BFF5: mov var_A0, eax
  loc_0041BFFB: call [ecx+000000A0h]
  loc_0041C001: test eax, eax
  loc_0041C003: fnclex
  loc_0041C005: jge 0041C01Fh
  loc_0041C007: mov ecx, var_A0
  loc_0041C00D: push 000000A0h
  loc_0041C012: push 0040F02Ch
  loc_0041C017: push ecx
  loc_0041C018: push eax
  loc_0041C019: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C01F: mov eax, var_18
  loc_0041C022: lea edx, var_2C
  loc_0041C025: mov var_24, eax
  loc_0041C028: lea eax, var_3C
  loc_0041C02B: push edx
  loc_0041C02C: push eax
  loc_0041C02D: mov var_18, 00000000h
  loc_0041C034: mov var_2C, edi
  loc_0041C037: call [00401050h] ; rtcTrimVar
  loc_0041C03D: lea ecx, var_6C
  loc_0041C040: lea edx, var_3C
  loc_0041C043: push ecx
  loc_0041C044: lea eax, var_4C
  loc_0041C047: push edx
  loc_0041C048: push eax
  loc_0041C049: call [004010C0h] ; __vbaVarCat
  loc_0041C04F: push eax
  loc_0041C050: call [00401014h] ; __vbaStrVarMove
  loc_0041C056: mov edx, eax
  loc_0041C058: mov ecx, 00430018h
  loc_0041C05D: call [004010FCh] ; __vbaStrMove
  loc_0041C063: lea ecx, var_1C
  loc_0041C066: call [00401114h] ; __vbaFreeObj
  loc_0041C06C: lea ecx, var_4C
  loc_0041C06F: lea edx, var_3C
  loc_0041C072: push ecx
  loc_0041C073: lea eax, var_2C
  loc_0041C076: push edx
  loc_0041C077: push eax
  loc_0041C078: push 00000003h
  loc_0041C07A: call ebx
  loc_0041C07C: mov ecx, [esi]
  loc_0041C07E: add esp, 00000010h
  loc_0041C081: push esi
  loc_0041C082: call [ecx+00000394h]
  loc_0041C088: mov var_24, eax
  loc_0041C08B: lea edx, var_2C
  loc_0041C08E: lea eax, var_3C
  loc_0041C091: push edx
  loc_0041C092: push eax
  loc_0041C093: mov var_2C, 00000009h
  loc_0041C09A: call [00401050h] ; rtcTrimVar
  loc_0041C0A0: lea ecx, var_3C
  loc_0041C0A3: lea edx, var_6C
  loc_0041C0A6: push ecx
  loc_0041C0A7: push edx
  loc_0041C0A8: mov var_64, 0040F040h
  loc_0041C0AF: mov var_6C, 00008008h
  loc_0041C0B6: call [00401070h] ; __vbaVarTstEq
  loc_0041C0BC: mov var_A0, ax
  loc_0041C0C3: lea eax, var_3C
  loc_0041C0C6: lea ecx, var_2C
  loc_0041C0C9: push eax
  loc_0041C0CA: push ecx
  loc_0041C0CB: push 00000002h
  loc_0041C0CD: call ebx
  loc_0041C0CF: add esp, 0000000Ch
  loc_0041C0D2: cmp var_A0, 0000h
  loc_0041C0DA: jz 0041C194h
  loc_0041C0E0: mov edx, [esi]
  loc_0041C0E2: push esi
  loc_0041C0E3: call [edx+00000394h]
  loc_0041C0E9: push eax
  loc_0041C0EA: lea eax, var_1C
  loc_0041C0ED: push eax
  loc_0041C0EE: call [0040103Ch] ; __vbaObjSet
  loc_0041C0F4: mov esi, eax
  loc_0041C0F6: push esi
  loc_0041C0F7: mov ecx, [esi]
  loc_0041C0F9: call [ecx+00000204h]
  loc_0041C0FF: test eax, eax
  loc_0041C101: fnclex
  loc_0041C103: jge 0041C117h
  loc_0041C105: push 00000204h
  loc_0041C10A: push 0040F02Ch
  loc_0041C10F: push esi
  loc_0041C110: push eax
  loc_0041C111: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C117: lea ecx, var_1C
  loc_0041C11A: call [00401114h] ; __vbaFreeObj
  loc_0041C120: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041C126: mov ecx, 80020004h
  loc_0041C12B: mov var_54, ecx
  loc_0041C12E: mov eax, 0000000Ah
  loc_0041C133: mov var_44, ecx
  loc_0041C136: lea edx, var_7C
  loc_0041C139: lea ecx, var_3C
  loc_0041C13C: mov var_5C, eax
  loc_0041C13F: mov var_4C, eax
  loc_0041C142: mov var_74, 0040F100h ; "Check Name"
  loc_0041C149: mov var_7C, edi
  loc_0041C14C: call __vbaVarDup
  loc_0041C14E: mov var_64, 0040F09Ch ; "Please enter a name for the dependent variable."
  loc_0041C155: lea edx, var_6C
  loc_0041C158: lea ecx, var_2C
  loc_0041C15B: mov var_6C, edi
  loc_0041C15E: call __vbaVarDup
  loc_0041C160: lea edx, var_5C
  loc_0041C163: lea eax, var_4C
  loc_0041C166: push edx
  loc_0041C167: lea ecx, var_3C
  loc_0041C16A: push eax
  loc_0041C16B: push ecx
  loc_0041C16C: lea edx, var_2C
  loc_0041C16F: push 00000000h
  loc_0041C171: push edx
  loc_0041C172: call [00401038h] ; rtcMsgBox
  loc_0041C178: lea eax, var_5C
  loc_0041C17B: lea ecx, var_4C
  loc_0041C17E: push eax
  loc_0041C17F: lea edx, var_3C
  loc_0041C182: push ecx
  loc_0041C183: lea eax, var_2C
  loc_0041C186: push edx
  loc_0041C187: push eax
  loc_0041C188: push 00000004h
  loc_0041C18A: call ebx
  loc_0041C18C: add esp, 00000014h
  loc_0041C18F: jmp 0041C2D6h
  loc_0041C194: mov ecx, [esi]
  loc_0041C196: push esi
  loc_0041C197: call [ecx+0000039Ch]
  loc_0041C19D: mov var_24, eax
  loc_0041C1A0: lea edx, var_2C
  loc_0041C1A3: lea eax, var_3C
  loc_0041C1A6: push edx
  loc_0041C1A7: push eax
  loc_0041C1A8: mov var_2C, 00000009h
  loc_0041C1AF: call [00401050h] ; rtcTrimVar
  loc_0041C1B5: lea ecx, var_3C
  loc_0041C1B8: lea edx, var_6C
  loc_0041C1BB: push ecx
  loc_0041C1BC: push edx
  loc_0041C1BD: mov var_64, 0040F040h
  loc_0041C1C4: mov var_6C, 00008008h
  loc_0041C1CB: call [00401070h] ; __vbaVarTstEq
  loc_0041C1D1: mov var_A0, ax
  loc_0041C1D8: lea eax, var_3C
  loc_0041C1DB: lea ecx, var_2C
  loc_0041C1DE: push eax
  loc_0041C1DF: push ecx
  loc_0041C1E0: push 00000002h
  loc_0041C1E2: call ebx
  loc_0041C1E4: add esp, 0000000Ch
  loc_0041C1E7: cmp var_A0, 0000h
  loc_0041C1EF: jz 0041C297h
  loc_0041C1F5: mov ecx, 80020004h
  loc_0041C1FA: mov eax, 0000000Ah
  loc_0041C1FF: mov var_54, ecx
  loc_0041C202: mov var_44, ecx
  loc_0041C205: lea edx, var_7C
  loc_0041C208: lea ecx, var_3C
  loc_0041C20B: mov var_5C, eax
  loc_0041C20E: mov var_4C, eax
  loc_0041C211: mov var_74, 0040F100h ; "Check Name"
  loc_0041C218: mov var_7C, edi
  loc_0041C21B: call [004010F4h] ; __vbaVarDup
  loc_0041C221: lea edx, var_6C
  loc_0041C224: lea ecx, var_2C
  loc_0041C227: mov var_64, 0040F11Ch ; "Please enter a name for the independent variable."
  loc_0041C22E: mov var_6C, edi
  loc_0041C231: call [004010F4h] ; __vbaVarDup
  loc_0041C237: lea edx, var_5C
  loc_0041C23A: lea eax, var_4C
  loc_0041C23D: push edx
  loc_0041C23E: lea ecx, var_3C
  loc_0041C241: push eax
  loc_0041C242: push ecx
  loc_0041C243: lea edx, var_2C
  loc_0041C246: push 00000000h
  loc_0041C248: push edx
  loc_0041C249: call [00401038h] ; rtcMsgBox
  loc_0041C24F: lea eax, var_5C
  loc_0041C252: lea ecx, var_4C
  loc_0041C255: push eax
  loc_0041C256: lea edx, var_3C
  loc_0041C259: push ecx
  loc_0041C25A: lea eax, var_2C
  loc_0041C25D: push edx
  loc_0041C25E: push eax
  loc_0041C25F: push 00000004h
  loc_0041C261: call ebx
  loc_0041C263: mov ecx, [esi]
  loc_0041C265: add esp, 00000014h
  loc_0041C268: push esi
  loc_0041C269: call [ecx+0000039Ch]
  loc_0041C26F: lea edx, var_1C
  loc_0041C272: push eax
  loc_0041C273: push edx
  loc_0041C274: call [0040103Ch] ; __vbaObjSet
  loc_0041C27A: mov esi, eax
  loc_0041C27C: push esi
  loc_0041C27D: mov eax, [esi]
  loc_0041C27F: call [eax+00000204h]
  loc_0041C285: test eax, eax
  loc_0041C287: fnclex
  loc_0041C289: jge 0041C2CDh
  loc_0041C28B: push 00000204h
  loc_0041C290: push 0040F02Ch
  loc_0041C295: jmp 0041C2C5h
  loc_0041C297: mov ecx, [esi]
  loc_0041C299: push esi
  loc_0041C29A: call [ecx+00000380h]
  loc_0041C2A0: lea edx, var_1C
  loc_0041C2A3: push eax
  loc_0041C2A4: push edx
  loc_0041C2A5: call [0040103Ch] ; __vbaObjSet
  loc_0041C2AB: mov esi, eax
  loc_0041C2AD: push 461C4000h
  loc_0041C2B2: push esi
  loc_0041C2B3: mov eax, [esi]
  loc_0041C2B5: call [eax+0000007Ch]
  loc_0041C2B8: test eax, eax
  loc_0041C2BA: fnclex
  loc_0041C2BC: jge 0041C2CDh
  loc_0041C2BE: push 0000007Ch
  loc_0041C2C0: push 0040E728h
  loc_0041C2C5: push esi
  loc_0041C2C6: push eax
  loc_0041C2C7: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C2CD: lea ecx, var_1C
  loc_0041C2D0: call [00401114h] ; __vbaFreeObj
  loc_0041C2D6: mov var_4, 00000000h
  loc_0041C2DD: fwait
  loc_0041C2DE: push 0041C314h
  loc_0041C2E3: jmp 0041C313h
  loc_0041C2E5: lea ecx, var_18
  loc_0041C2E8: call [00401110h] ; __vbaFreeStr
  loc_0041C2EE: lea ecx, var_1C
  loc_0041C2F1: call [00401114h] ; __vbaFreeObj
  loc_0041C2F7: lea ecx, var_5C
  loc_0041C2FA: lea edx, var_4C
  loc_0041C2FD: push ecx
  loc_0041C2FE: lea eax, var_3C
  loc_0041C301: push edx
  loc_0041C302: lea ecx, var_2C
  loc_0041C305: push eax
  loc_0041C306: push ecx
  loc_0041C307: push 00000004h
  loc_0041C309: call [00401018h] ; __vbaFreeVarList
  loc_0041C30F: add esp, 00000014h
  loc_0041C312: ret
  loc_0041C313: ret
  loc_0041C314: mov eax, Me
  loc_0041C317: push eax
  loc_0041C318: mov edx, [eax]
  loc_0041C31A: call [edx+00000008h]
  loc_0041C31D: mov eax, var_4
  loc_0041C320: mov ecx, var_14
  loc_0041C323: pop edi
  loc_0041C324: pop esi
  loc_0041C325: mov fs:[00000000h], ecx
  loc_0041C32C: pop ebx
  loc_0041C32D: mov esp, ebp
  loc_0041C32F: pop ebp
  loc_0041C330: retn 0004h
End Sub
