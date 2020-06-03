VERSION 5.00
Begin VB.Form frmChangeEstPred
  Caption = "Values for Prediction and Estimation"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 8040
  ClientHeight = 4950
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Prediction of Y | X and Estimation of Mean Y | X"
    Left = 0
    Top = 0
    Width = 7935
    Height = 4935
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
      Caption = "Estimated Standard Error For Estimation"
      Left = 480
      Top = 1800
      Width = 7095
      Height = 1095
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
      Begin VB.TextBox txtSyhatg
        Left = 120
        Top = 360
        Width = 2295
        Height = 615
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
    Begin VB.Frame Frame1
      Caption = "Value of Yhat when X=Xg"
      Left = 480
      Top = 600
      Width = 7095
      Height = 1095
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
      Begin VB.TextBox txtYhatg
        Left = 120
        Top = 360
        Width = 2295
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
    Begin VB.CommandButton cmdRestore
      Caption = "Reset Defaults"
      Left = 3120
      Top = 4200
      Width = 3255
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
    Begin VB.CommandButton cmdClearChange
      Caption = "Clear"
      Left = 1680
      Top = 4200
      Width = 1095
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
    Begin VB.CommandButton cmdChange
      Caption = "Ok"
      Left = 480
      Top = 4200
      Width = 855
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
    Begin VB.Frame Frame3
      Caption = "Estimated Standard Error for Prediction"
      Left = 480
      Top = 2880
      Width = 7095
      Height = 1095
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
      Begin VB.TextBox txtSyhatnewg
        Left = 120
        Top = 360
        Width = 2295
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

Attribute VB_Name = "frmChangeEstPred"


Private Sub cmdClearChange_Click() '41D550
  loc_0041D550: push ebp
  loc_0041D551: mov ebp, esp
  loc_0041D553: sub esp, 0000000Ch
  loc_0041D556: push 00401926h ; __vbaExceptHandler
  loc_0041D55B: mov eax, fs:[00000000h]
  loc_0041D561: push eax
  loc_0041D562: mov fs:[00000000h], esp
  loc_0041D569: sub esp, 00000014h
  loc_0041D56C: push ebx
  loc_0041D56D: push esi
  loc_0041D56E: push edi
  loc_0041D56F: mov var_C, esp
  loc_0041D572: mov var_8, 00401248h
  loc_0041D579: mov esi, Me
  loc_0041D57C: mov eax, esi
  loc_0041D57E: and eax, 00000001h
  loc_0041D581: mov var_4, eax
  loc_0041D584: and esi, FFFFFFFEh
  loc_0041D587: push esi
  loc_0041D588: mov Me, esi
  loc_0041D58B: mov ecx, [esi]
  loc_0041D58D: call [ecx+00000004h]
  loc_0041D590: mov edx, [esi]
  loc_0041D592: push esi
  loc_0041D593: mov var_18, 00000000h
  loc_0041D59A: call [edx+0000030Ch]
  loc_0041D5A0: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041D5A6: push eax
  loc_0041D5A7: lea eax, var_18
  loc_0041D5AA: push eax
  loc_0041D5AB: call ebx
  loc_0041D5AD: mov edi, eax
  loc_0041D5AF: push 0040F040h
  loc_0041D5B4: push edi
  loc_0041D5B5: mov ecx, [edi]
  loc_0041D5B7: call [ecx+000000A4h]
  loc_0041D5BD: test eax, eax
  loc_0041D5BF: fnclex
  loc_0041D5C1: jge 0041D5D5h
  loc_0041D5C3: push 000000A4h
  loc_0041D5C8: push 0040F02Ch
  loc_0041D5CD: push edi
  loc_0041D5CE: push eax
  loc_0041D5CF: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D5D5: lea ecx, var_18
  loc_0041D5D8: call [00401114h] ; __vbaFreeObj
  loc_0041D5DE: mov edx, [esi]
  loc_0041D5E0: push esi
  loc_0041D5E1: call [edx+00000320h]
  loc_0041D5E7: push eax
  loc_0041D5E8: lea eax, var_18
  loc_0041D5EB: push eax
  loc_0041D5EC: call ebx
  loc_0041D5EE: mov edi, eax
  loc_0041D5F0: push 0040F040h
  loc_0041D5F5: push edi
  loc_0041D5F6: mov ecx, [edi]
  loc_0041D5F8: call [ecx+000000A4h]
  loc_0041D5FE: test eax, eax
  loc_0041D600: fnclex
  loc_0041D602: jge 0041D616h
  loc_0041D604: push 000000A4h
  loc_0041D609: push 0040F02Ch
  loc_0041D60E: push edi
  loc_0041D60F: push eax
  loc_0041D610: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D616: lea ecx, var_18
  loc_0041D619: call [00401114h] ; __vbaFreeObj
  loc_0041D61F: mov edx, [esi]
  loc_0041D621: push esi
  loc_0041D622: call [edx+00000304h]
  loc_0041D628: push eax
  loc_0041D629: lea eax, var_18
  loc_0041D62C: push eax
  loc_0041D62D: call ebx
  loc_0041D62F: mov edi, eax
  loc_0041D631: push 0040F040h
  loc_0041D636: push edi
  loc_0041D637: mov ecx, [edi]
  loc_0041D639: call [ecx+000000A4h]
  loc_0041D63F: test eax, eax
  loc_0041D641: fnclex
  loc_0041D643: jge 0041D657h
  loc_0041D645: push 000000A4h
  loc_0041D64A: push 0040F02Ch
  loc_0041D64F: push edi
  loc_0041D650: push eax
  loc_0041D651: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D657: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041D65D: lea ecx, var_18
  loc_0041D660: call edi
  loc_0041D662: mov edx, [esi]
  loc_0041D664: push esi
  loc_0041D665: call [edx+0000030Ch]
  loc_0041D66B: push eax
  loc_0041D66C: lea eax, var_18
  loc_0041D66F: push eax
  loc_0041D670: call ebx
  loc_0041D672: mov esi, eax
  loc_0041D674: push esi
  loc_0041D675: mov ecx, [esi]
  loc_0041D677: call [ecx+00000204h]
  loc_0041D67D: test eax, eax
  loc_0041D67F: fnclex
  loc_0041D681: jge 0041D695h
  loc_0041D683: push 00000204h
  loc_0041D688: push 0040F02Ch
  loc_0041D68D: push esi
  loc_0041D68E: push eax
  loc_0041D68F: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D695: lea ecx, var_18
  loc_0041D698: call edi
  loc_0041D69A: mov var_4, 00000000h
  loc_0041D6A1: push 0041D6B3h
  loc_0041D6A6: jmp 0041D6B2h
  loc_0041D6A8: lea ecx, var_18
  loc_0041D6AB: call [00401114h] ; __vbaFreeObj
  loc_0041D6B1: ret
  loc_0041D6B2: ret
  loc_0041D6B3: mov eax, Me
  loc_0041D6B6: push eax
  loc_0041D6B7: mov edx, [eax]
  loc_0041D6B9: call [edx+00000008h]
  loc_0041D6BC: mov eax, var_4
  loc_0041D6BF: mov ecx, var_14
  loc_0041D6C2: pop edi
  loc_0041D6C3: pop esi
  loc_0041D6C4: mov fs:[00000000h], ecx
  loc_0041D6CB: pop ebx
  loc_0041D6CC: mov esp, ebp
  loc_0041D6CE: pop ebp
  loc_0041D6CF: retn 0004h
End Sub

Private Sub txtYhatg_KeyPress(KeyAscii As Integer) '41E240
  loc_0041E240: push ebp
  loc_0041E241: mov ebp, esp
  loc_0041E243: sub esp, 0000000Ch
  loc_0041E246: push 00401926h ; __vbaExceptHandler
  loc_0041E24B: mov eax, fs:[00000000h]
  loc_0041E251: push eax
  loc_0041E252: mov fs:[00000000h], esp
  loc_0041E259: sub esp, 0000007Ch
  loc_0041E25C: push ebx
  loc_0041E25D: push esi
  loc_0041E25E: push edi
  loc_0041E25F: mov var_C, esp
  loc_0041E262: mov var_8, 00401298h
  loc_0041E269: mov esi, Me
  loc_0041E26C: mov eax, esi
  loc_0041E26E: and eax, 00000001h
  loc_0041E271: mov var_4, eax
  loc_0041E274: and esi, FFFFFFFEh
  loc_0041E277: push esi
  loc_0041E278: mov Me, esi
  loc_0041E27B: mov ecx, [esi]
  loc_0041E27D: call [ecx+00000004h]
  loc_0041E280: mov ebx, KeyAscii
  loc_0041E283: lea edx, var_24
  loc_0041E286: xor edi, edi
  loc_0041E288: push ebx
  loc_0041E289: push edx
  loc_0041E28A: mov var_24, edi
  loc_0041E28D: mov var_34, edi
  loc_0041E290: mov var_44, edi
  loc_0041E293: mov var_54, edi
  loc_0041E296: mov var_64, edi
  loc_0041E299: mov var_74, edi
  loc_0041E29C: mov var_84, edi
  loc_0041E2A2: call 0041A480h
  loc_0041E2A7: lea eax, var_24
  loc_0041E2AA: push eax
  loc_0041E2AB: call [004010C4h] ; __vbaI2Var
  loc_0041E2B1: lea ecx, var_24
  loc_0041E2B4: mov [ebx], ax
  loc_0041E2B7: call [00401010h] ; __vbaFreeVar
  loc_0041E2BD: push 0040DD3Ch ; "."
  loc_0041E2C2: call [00401024h] ; rtcAnsiValueBstr
  loc_0041E2C8: mov edx, [esi]
  loc_0041E2CA: xor ecx, ecx
  loc_0041E2CC: cmp [ebx], ax
  loc_0041E2CF: push esi
  loc_0041E2D0: mov var_84, 0000000Bh
  loc_0041E2DA: setz cl
  loc_0041E2DD: neg ecx
  loc_0041E2DF: mov var_7C, cx
  loc_0041E2E3: call [edx+0000030Ch]
  loc_0041E2E9: mov var_1C, eax
  loc_0041E2EC: lea eax, var_84
  loc_0041E2F2: push eax
  loc_0041E2F3: lea ecx, var_24
  loc_0041E2F6: push 00000001h
  loc_0041E2F8: lea edx, var_64
  loc_0041E2FB: push ecx
  loc_0041E2FC: push edx
  loc_0041E2FD: lea eax, var_34
  loc_0041E300: push edi
  loc_0041E301: push eax
  loc_0041E302: mov var_24, 00000009h
  loc_0041E309: mov var_5C, 0040DD3Ch ; "."
  loc_0041E310: mov var_64, 00000008h
  loc_0041E317: mov var_6C, edi
  loc_0041E31A: mov var_74, 00008002h
  loc_0041E321: call [004010B4h] ; __vbaInStrVar
  loc_0041E327: lea ecx, var_74
  loc_0041E32A: push eax
  loc_0041E32B: lea edx, var_44
  loc_0041E32E: push ecx
  loc_0041E32F: push edx
  loc_0041E330: call [00401060h] ; __vbaVarCmpGt
  loc_0041E336: push eax
  loc_0041E337: lea eax, var_54
  loc_0041E33A: push eax
  loc_0041E33B: call [00401094h] ; __vbaVarAnd
  loc_0041E341: push eax
  loc_0041E342: call [00401058h] ; __vbaBoolVarNull
  loc_0041E348: lea ecx, var_84
  loc_0041E34E: mov esi, eax
  loc_0041E350: lea edx, var_34
  loc_0041E353: push ecx
  loc_0041E354: lea eax, var_24
  loc_0041E357: push edx
  loc_0041E358: push eax
  loc_0041E359: push 00000003h
  loc_0041E35B: call [00401018h] ; __vbaFreeVarList
  loc_0041E361: add esp, 00000010h
  loc_0041E364: cmp si, di
  loc_0041E367: jz 0041E36Ch
  loc_0041E369: mov [ebx], di
  loc_0041E36C: mov var_4, edi
  loc_0041E36F: push 0041E393h
  loc_0041E374: jmp 0041E392h
  loc_0041E376: lea ecx, var_54
  loc_0041E379: lea edx, var_44
  loc_0041E37C: push ecx
  loc_0041E37D: lea eax, var_34
  loc_0041E380: push edx
  loc_0041E381: lea ecx, var_24
  loc_0041E384: push eax
  loc_0041E385: push ecx
  loc_0041E386: push 00000004h
  loc_0041E388: call [00401018h] ; __vbaFreeVarList
  loc_0041E38E: add esp, 00000014h
  loc_0041E391: ret
  loc_0041E392: ret
  loc_0041E393: mov eax, Me
  loc_0041E396: push eax
  loc_0041E397: mov edx, [eax]
  loc_0041E399: call [edx+00000008h]
  loc_0041E39C: mov eax, var_4
  loc_0041E39F: mov ecx, var_14
  loc_0041E3A2: pop edi
  loc_0041E3A3: pop esi
  loc_0041E3A4: mov fs:[00000000h], ecx
  loc_0041E3AB: pop ebx
  loc_0041E3AC: mov esp, ebp
  loc_0041E3AE: pop ebp
  loc_0041E3AF: retn 0008h
End Sub

Private Sub txtSyhatnewg_KeyPress(KeyAscii As Integer) '41E3C0
  loc_0041E3C0: push ebp
  loc_0041E3C1: mov ebp, esp
  loc_0041E3C3: sub esp, 0000000Ch
  loc_0041E3C6: push 00401926h ; __vbaExceptHandler
  loc_0041E3CB: mov eax, fs:[00000000h]
  loc_0041E3D1: push eax
  loc_0041E3D2: mov fs:[00000000h], esp
  loc_0041E3D9: sub esp, 0000008Ch
  loc_0041E3DF: push ebx
  loc_0041E3E0: push esi
  loc_0041E3E1: push edi
  loc_0041E3E2: mov var_C, esp
  loc_0041E3E5: mov var_8, 004012A8h
  loc_0041E3EC: mov esi, Me
  loc_0041E3EF: mov eax, esi
  loc_0041E3F1: and eax, 00000001h
  loc_0041E3F4: mov var_4, eax
  loc_0041E3F7: and esi, FFFFFFFEh
  loc_0041E3FA: push esi
  loc_0041E3FB: mov Me, esi
  loc_0041E3FE: mov ecx, [esi]
  loc_0041E400: call [ecx+00000004h]
  loc_0041E403: mov edi, KeyAscii
  loc_0041E406: lea edx, var_24
  loc_0041E409: xor ebx, ebx
  loc_0041E40B: push edi
  loc_0041E40C: push edx
  loc_0041E40D: mov var_24, ebx
  loc_0041E410: mov var_34, ebx
  loc_0041E413: mov var_44, ebx
  loc_0041E416: mov var_54, ebx
  loc_0041E419: mov var_64, ebx
  loc_0041E41C: mov var_74, ebx
  loc_0041E41F: mov var_84, ebx
  loc_0041E425: call 0041A480h
  loc_0041E42A: lea eax, var_24
  loc_0041E42D: push eax
  loc_0041E42E: call [004010C4h] ; __vbaI2Var
  loc_0041E434: lea ecx, var_24
  loc_0041E437: mov [edi], ax
  loc_0041E43A: call [00401010h] ; __vbaFreeVar
  loc_0041E440: push 0040DD3Ch ; "."
  loc_0041E445: call [00401024h] ; rtcAnsiValueBstr
  loc_0041E44B: mov edx, [esi]
  loc_0041E44D: xor ecx, ecx
  loc_0041E44F: cmp [edi], ax
  loc_0041E452: push esi
  loc_0041E453: mov var_84, 0000000Bh
  loc_0041E45D: setz cl
  loc_0041E460: neg ecx
  loc_0041E462: mov var_7C, cx
  loc_0041E466: call [edx+00000320h]
  loc_0041E46C: mov var_1C, eax
  loc_0041E46F: lea eax, var_84
  loc_0041E475: push eax
  loc_0041E476: lea ecx, var_24
  loc_0041E479: push 00000001h
  loc_0041E47B: lea edx, var_64
  loc_0041E47E: push ecx
  loc_0041E47F: push edx
  loc_0041E480: lea eax, var_34
  loc_0041E483: push ebx
  loc_0041E484: push eax
  loc_0041E485: mov var_24, 00000009h
  loc_0041E48C: mov var_5C, 0040DD3Ch ; "."
  loc_0041E493: mov var_64, 00000008h
  loc_0041E49A: mov var_6C, ebx
  loc_0041E49D: mov var_74, 00008002h
  loc_0041E4A4: call [004010B4h] ; __vbaInStrVar
  loc_0041E4AA: lea ecx, var_74
  loc_0041E4AD: push eax
  loc_0041E4AE: lea edx, var_44
  loc_0041E4B1: push ecx
  loc_0041E4B2: push edx
  loc_0041E4B3: call [00401060h] ; __vbaVarCmpGt
  loc_0041E4B9: push eax
  loc_0041E4BA: lea eax, var_54
  loc_0041E4BD: push eax
  loc_0041E4BE: call [00401094h] ; __vbaVarAnd
  loc_0041E4C4: push eax
  loc_0041E4C5: call [00401058h] ; __vbaBoolVarNull
  loc_0041E4CB: lea ecx, var_84
  loc_0041E4D1: mov esi, eax
  loc_0041E4D3: lea edx, var_34
  loc_0041E4D6: push ecx
  loc_0041E4D7: lea eax, var_24
  loc_0041E4DA: push edx
  loc_0041E4DB: push eax
  loc_0041E4DC: push 00000003h
  loc_0041E4DE: call [00401018h] ; __vbaFreeVarList
  loc_0041E4E4: add esp, 00000010h
  loc_0041E4E7: cmp si, bx
  loc_0041E4EA: jz 0041E4EFh
  loc_0041E4EC: mov [edi], bx
  loc_0041E4EF: push 0040DD44h ; "-"
  loc_0041E4F4: call [00401024h] ; rtcAnsiValueBstr
  loc_0041E4FA: cmp [edi], ax
  loc_0041E4FD: jnz 0041E56Ah
  loc_0041E4FF: mov ecx, 80020004h
  loc_0041E504: mov eax, 0000000Ah
  loc_0041E509: mov var_4C, ecx
  loc_0041E50C: mov var_3C, ecx
  loc_0041E50F: mov var_2C, ecx
  loc_0041E512: lea edx, var_64
  loc_0041E515: lea ecx, var_24
  loc_0041E518: mov var_54, eax
  loc_0041E51B: mov var_44, eax
  loc_0041E51E: mov var_34, eax
  loc_0041E521: mov var_5C, 0040FBECh ; "Standard Errors are always positive."
  loc_0041E528: mov var_64, 00000008h
  loc_0041E52F: call [004010F4h] ; __vbaVarDup
  loc_0041E535: lea ecx, var_54
  loc_0041E538: lea edx, var_44
  loc_0041E53B: push ecx
  loc_0041E53C: lea eax, var_34
  loc_0041E53F: push edx
  loc_0041E540: push eax
  loc_0041E541: lea ecx, var_24
  loc_0041E544: push ebx
  loc_0041E545: push ecx
  loc_0041E546: call [00401038h] ; rtcMsgBox
  loc_0041E54C: lea edx, var_54
  loc_0041E54F: lea eax, var_44
  loc_0041E552: push edx
  loc_0041E553: lea ecx, var_34
  loc_0041E556: push eax
  loc_0041E557: lea edx, var_24
  loc_0041E55A: push ecx
  loc_0041E55B: push edx
  loc_0041E55C: push 00000004h
  loc_0041E55E: call [00401018h] ; __vbaFreeVarList
  loc_0041E564: add esp, 00000014h
  loc_0041E567: mov [edi], bx
  loc_0041E56A: mov var_4, ebx
  loc_0041E56D: push 0041E591h
  loc_0041E572: jmp 0041E590h
  loc_0041E574: lea eax, var_54
  loc_0041E577: lea ecx, var_44
  loc_0041E57A: push eax
  loc_0041E57B: lea edx, var_34
  loc_0041E57E: push ecx
  loc_0041E57F: lea eax, var_24
  loc_0041E582: push edx
  loc_0041E583: push eax
  loc_0041E584: push 00000004h
  loc_0041E586: call [00401018h] ; __vbaFreeVarList
  loc_0041E58C: add esp, 00000014h
  loc_0041E58F: ret
  loc_0041E590: ret
  loc_0041E591: mov eax, Me
  loc_0041E594: push eax
  loc_0041E595: mov ecx, [eax]
  loc_0041E597: call [ecx+00000008h]
  loc_0041E59A: mov eax, var_4
  loc_0041E59D: mov ecx, var_14
  loc_0041E5A0: pop edi
  loc_0041E5A1: pop esi
  loc_0041E5A2: mov fs:[00000000h], ecx
  loc_0041E5A9: pop ebx
  loc_0041E5AA: mov esp, ebp
  loc_0041E5AC: pop ebp
  loc_0041E5AD: retn 0008h
End Sub

Private Sub txtSyhatg_KeyPress(KeyAscii As Integer) '41E050
  loc_0041E050: push ebp
  loc_0041E051: mov ebp, esp
  loc_0041E053: sub esp, 0000000Ch
  loc_0041E056: push 00401926h ; __vbaExceptHandler
  loc_0041E05B: mov eax, fs:[00000000h]
  loc_0041E061: push eax
  loc_0041E062: mov fs:[00000000h], esp
  loc_0041E069: sub esp, 0000008Ch
  loc_0041E06F: push ebx
  loc_0041E070: push esi
  loc_0041E071: push edi
  loc_0041E072: mov var_C, esp
  loc_0041E075: mov var_8, 00401288h
  loc_0041E07C: mov esi, Me
  loc_0041E07F: mov eax, esi
  loc_0041E081: and eax, 00000001h
  loc_0041E084: mov var_4, eax
  loc_0041E087: and esi, FFFFFFFEh
  loc_0041E08A: push esi
  loc_0041E08B: mov Me, esi
  loc_0041E08E: mov ecx, [esi]
  loc_0041E090: call [ecx+00000004h]
  loc_0041E093: mov edi, KeyAscii
  loc_0041E096: lea edx, var_24
  loc_0041E099: xor ebx, ebx
  loc_0041E09B: push edi
  loc_0041E09C: push edx
  loc_0041E09D: mov var_24, ebx
  loc_0041E0A0: mov var_34, ebx
  loc_0041E0A3: mov var_44, ebx
  loc_0041E0A6: mov var_54, ebx
  loc_0041E0A9: mov var_64, ebx
  loc_0041E0AC: mov var_74, ebx
  loc_0041E0AF: mov var_84, ebx
  loc_0041E0B5: call 0041A480h
  loc_0041E0BA: lea eax, var_24
  loc_0041E0BD: push eax
  loc_0041E0BE: call [004010C4h] ; __vbaI2Var
  loc_0041E0C4: lea ecx, var_24
  loc_0041E0C7: mov [edi], ax
  loc_0041E0CA: call [00401010h] ; __vbaFreeVar
  loc_0041E0D0: push 0040DD3Ch ; "."
  loc_0041E0D5: call [00401024h] ; rtcAnsiValueBstr
  loc_0041E0DB: mov edx, [esi]
  loc_0041E0DD: xor ecx, ecx
  loc_0041E0DF: cmp [edi], ax
  loc_0041E0E2: push esi
  loc_0041E0E3: mov var_84, 0000000Bh
  loc_0041E0ED: setz cl
  loc_0041E0F0: neg ecx
  loc_0041E0F2: mov var_7C, cx
  loc_0041E0F6: call [edx+00000304h]
  loc_0041E0FC: mov var_1C, eax
  loc_0041E0FF: lea eax, var_84
  loc_0041E105: push eax
  loc_0041E106: lea ecx, var_24
  loc_0041E109: push 00000001h
  loc_0041E10B: lea edx, var_64
  loc_0041E10E: push ecx
  loc_0041E10F: push edx
  loc_0041E110: lea eax, var_34
  loc_0041E113: push ebx
  loc_0041E114: push eax
  loc_0041E115: mov var_24, 00000009h
  loc_0041E11C: mov var_5C, 0040DD3Ch ; "."
  loc_0041E123: mov var_64, 00000008h
  loc_0041E12A: mov var_6C, ebx
  loc_0041E12D: mov var_74, 00008002h
  loc_0041E134: call [004010B4h] ; __vbaInStrVar
  loc_0041E13A: lea ecx, var_74
  loc_0041E13D: push eax
  loc_0041E13E: lea edx, var_44
  loc_0041E141: push ecx
  loc_0041E142: push edx
  loc_0041E143: call [00401060h] ; __vbaVarCmpGt
  loc_0041E149: push eax
  loc_0041E14A: lea eax, var_54
  loc_0041E14D: push eax
  loc_0041E14E: call [00401094h] ; __vbaVarAnd
  loc_0041E154: push eax
  loc_0041E155: call [00401058h] ; __vbaBoolVarNull
  loc_0041E15B: lea ecx, var_84
  loc_0041E161: mov esi, eax
  loc_0041E163: lea edx, var_34
  loc_0041E166: push ecx
  loc_0041E167: lea eax, var_24
  loc_0041E16A: push edx
  loc_0041E16B: push eax
  loc_0041E16C: push 00000003h
  loc_0041E16E: call [00401018h] ; __vbaFreeVarList
  loc_0041E174: add esp, 00000010h
  loc_0041E177: cmp si, bx
  loc_0041E17A: jz 0041E17Fh
  loc_0041E17C: mov [edi], bx
  loc_0041E17F: push 0040DD44h ; "-"
  loc_0041E184: call [00401024h] ; rtcAnsiValueBstr
  loc_0041E18A: cmp [edi], ax
  loc_0041E18D: jnz 0041E1FAh
  loc_0041E18F: mov ecx, 80020004h
  loc_0041E194: mov eax, 0000000Ah
  loc_0041E199: mov var_4C, ecx
  loc_0041E19C: mov var_3C, ecx
  loc_0041E19F: mov var_2C, ecx
  loc_0041E1A2: lea edx, var_64
  loc_0041E1A5: lea ecx, var_24
  loc_0041E1A8: mov var_54, eax
  loc_0041E1AB: mov var_44, eax
  loc_0041E1AE: mov var_34, eax
  loc_0041E1B1: mov var_5C, 0040F98Ch ; "Standard errors are always positive."
  loc_0041E1B8: mov var_64, 00000008h
  loc_0041E1BF: call [004010F4h] ; __vbaVarDup
  loc_0041E1C5: lea ecx, var_54
  loc_0041E1C8: lea edx, var_44
  loc_0041E1CB: push ecx
  loc_0041E1CC: lea eax, var_34
  loc_0041E1CF: push edx
  loc_0041E1D0: push eax
  loc_0041E1D1: lea ecx, var_24
  loc_0041E1D4: push ebx
  loc_0041E1D5: push ecx
  loc_0041E1D6: call [00401038h] ; rtcMsgBox
  loc_0041E1DC: lea edx, var_54
  loc_0041E1DF: lea eax, var_44
  loc_0041E1E2: push edx
  loc_0041E1E3: lea ecx, var_34
  loc_0041E1E6: push eax
  loc_0041E1E7: lea edx, var_24
  loc_0041E1EA: push ecx
  loc_0041E1EB: push edx
  loc_0041E1EC: push 00000004h
  loc_0041E1EE: call [00401018h] ; __vbaFreeVarList
  loc_0041E1F4: add esp, 00000014h
  loc_0041E1F7: mov [edi], bx
  loc_0041E1FA: mov var_4, ebx
  loc_0041E1FD: push 0041E221h
  loc_0041E202: jmp 0041E220h
  loc_0041E204: lea eax, var_54
  loc_0041E207: lea ecx, var_44
  loc_0041E20A: push eax
  loc_0041E20B: lea edx, var_34
  loc_0041E20E: push ecx
  loc_0041E20F: lea eax, var_24
  loc_0041E212: push edx
  loc_0041E213: push eax
  loc_0041E214: push 00000004h
  loc_0041E216: call [00401018h] ; __vbaFreeVarList
  loc_0041E21C: add esp, 00000014h
  loc_0041E21F: ret
  loc_0041E220: ret
  loc_0041E221: mov eax, Me
  loc_0041E224: push eax
  loc_0041E225: mov ecx, [eax]
  loc_0041E227: call [ecx+00000008h]
  loc_0041E22A: mov eax, var_4
  loc_0041E22D: mov ecx, var_14
  loc_0041E230: pop edi
  loc_0041E231: pop esi
  loc_0041E232: mov fs:[00000000h], ecx
  loc_0041E239: pop ebx
  loc_0041E23A: mov esp, ebp
  loc_0041E23C: pop ebp
  loc_0041E23D: retn 0008h
End Sub

Private Sub cmdChange_Click() '41D8B0
  loc_0041D8B0: push ebp
  loc_0041D8B1: mov ebp, esp
  loc_0041D8B3: sub esp, 0000000Ch
  loc_0041D8B6: push 00401926h ; __vbaExceptHandler
  loc_0041D8BB: mov eax, fs:[00000000h]
  loc_0041D8C1: push eax
  loc_0041D8C2: mov fs:[00000000h], esp
  loc_0041D8C9: sub esp, 000000A4h
  loc_0041D8CF: push ebx
  loc_0041D8D0: push esi
  loc_0041D8D1: push edi
  loc_0041D8D2: mov var_C, esp
  loc_0041D8D5: mov var_8, 00401268h
  loc_0041D8DC: mov esi, Me
  loc_0041D8DF: mov eax, esi
  loc_0041D8E1: and eax, 00000001h
  loc_0041D8E4: mov var_4, eax
  loc_0041D8E7: and esi, FFFFFFFEh
  loc_0041D8EA: push esi
  loc_0041D8EB: mov Me, esi
  loc_0041D8EE: mov ecx, [esi]
  loc_0041D8F0: call [ecx+00000004h]
  loc_0041D8F3: mov edx, [esi]
  loc_0041D8F5: xor eax, eax
  loc_0041D8F7: push esi
  loc_0041D8F8: mov var_18, eax
  loc_0041D8FB: mov var_1C, eax
  loc_0041D8FE: mov var_2C, eax
  loc_0041D901: mov var_3C, eax
  loc_0041D904: mov var_4C, eax
  loc_0041D907: mov var_5C, eax
  loc_0041D90A: mov var_6C, eax
  loc_0041D90D: mov var_7C, eax
  loc_0041D910: call [edx+0000030Ch]
  loc_0041D916: mov ebx, [00401050h] ; rtcTrimVar
  loc_0041D91C: mov var_24, eax
  loc_0041D91F: lea eax, var_2C
  loc_0041D922: lea ecx, var_3C
  loc_0041D925: push eax
  loc_0041D926: push ecx
  loc_0041D927: mov var_2C, 00000009h
  loc_0041D92E: call ebx
  loc_0041D930: lea edx, var_3C
  loc_0041D933: lea eax, var_6C
  loc_0041D936: push edx
  loc_0041D937: push eax
  loc_0041D938: mov var_64, 0040F040h
  loc_0041D93F: mov var_6C, 00008008h
  loc_0041D946: call [00401070h] ; __vbaVarTstEq
  loc_0041D94C: mov edi, [00401018h] ; __vbaFreeVarList
  loc_0041D952: lea ecx, var_3C
  loc_0041D955: lea edx, var_2C
  loc_0041D958: push ecx
  loc_0041D959: push edx
  loc_0041D95A: push 00000002h
  loc_0041D95C: mov var_A0, ax
  loc_0041D963: call edi
  loc_0041D965: add esp, 0000000Ch
  loc_0041D968: cmp var_A0, 0000h
  loc_0041D970: jz 0041D9F1h
  loc_0041D972: mov eax, [esi]
  loc_0041D974: push esi
  loc_0041D975: call [eax+0000030Ch]
  loc_0041D97B: lea ecx, var_1C
  loc_0041D97E: push eax
  loc_0041D97F: push ecx
  loc_0041D980: call [0040103Ch] ; __vbaObjSet
  loc_0041D986: mov esi, eax
  loc_0041D988: push esi
  loc_0041D989: mov edx, [esi]
  loc_0041D98B: call [edx+00000204h]
  loc_0041D991: test eax, eax
  loc_0041D993: fnclex
  loc_0041D995: jge 0041D9A9h
  loc_0041D997: push 00000204h
  loc_0041D99C: push 0040F02Ch
  loc_0041D9A1: push esi
  loc_0041D9A2: push eax
  loc_0041D9A3: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D9A9: lea ecx, var_1C
  loc_0041D9AC: call [00401114h] ; __vbaFreeObj
  loc_0041D9B2: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041D9B8: mov ecx, 80020004h
  loc_0041D9BD: mov var_54, ecx
  loc_0041D9C0: mov eax, 0000000Ah
  loc_0041D9C5: mov var_44, ecx
  loc_0041D9C8: mov ebx, 00000008h
  loc_0041D9CD: lea edx, var_7C
  loc_0041D9D0: lea ecx, var_3C
  loc_0041D9D3: mov var_5C, eax
  loc_0041D9D6: mov var_4C, eax
  loc_0041D9D9: mov var_74, 0040F970h ; "Check Yhatg"
  loc_0041D9E0: mov var_7C, ebx
  loc_0041D9E3: call __vbaVarDup
  loc_0041D9E5: mov var_64, 0040F92Ch ; "Please enter a value for Yhatg."
  loc_0041D9EC: jmp 0041DAC8h
  loc_0041D9F1: mov edx, [esi]
  loc_0041D9F3: push esi
  loc_0041D9F4: call [edx+00000304h]
  loc_0041D9FA: mov var_24, eax
  loc_0041D9FD: lea eax, var_2C
  loc_0041DA00: lea ecx, var_3C
  loc_0041DA03: push eax
  loc_0041DA04: push ecx
  loc_0041DA05: mov var_2C, 00000009h
  loc_0041DA0C: call ebx
  loc_0041DA0E: lea edx, var_3C
  loc_0041DA11: lea eax, var_6C
  loc_0041DA14: push edx
  loc_0041DA15: push eax
  loc_0041DA16: mov var_64, 0040F040h
  loc_0041DA1D: mov var_6C, 00008008h
  loc_0041DA24: call [00401070h] ; __vbaVarTstEq
  loc_0041DA2A: lea ecx, var_3C
  loc_0041DA2D: lea edx, var_2C
  loc_0041DA30: push ecx
  loc_0041DA31: push edx
  loc_0041DA32: push 00000002h
  loc_0041DA34: mov var_A0, ax
  loc_0041DA3B: call edi
  loc_0041DA3D: add esp, 0000000Ch
  loc_0041DA40: cmp var_A0, 0000h
  loc_0041DA48: jz 0041DB07h
  loc_0041DA4E: mov eax, [esi]
  loc_0041DA50: push esi
  loc_0041DA51: call [eax+00000304h]
  loc_0041DA57: lea ecx, var_1C
  loc_0041DA5A: push eax
  loc_0041DA5B: push ecx
  loc_0041DA5C: call [0040103Ch] ; __vbaObjSet
  loc_0041DA62: mov esi, eax
  loc_0041DA64: push esi
  loc_0041DA65: mov edx, [esi]
  loc_0041DA67: call [edx+00000204h]
  loc_0041DA6D: test eax, eax
  loc_0041DA6F: fnclex
  loc_0041DA71: jge 0041DA85h
  loc_0041DA73: push 00000204h
  loc_0041DA78: push 0040F02Ch
  loc_0041DA7D: push esi
  loc_0041DA7E: push eax
  loc_0041DA7F: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DA85: lea ecx, var_1C
  loc_0041DA88: call [00401114h] ; __vbaFreeObj
  loc_0041DA8E: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041DA94: mov ecx, 80020004h
  loc_0041DA99: mov var_54, ecx
  loc_0041DA9C: mov eax, 0000000Ah
  loc_0041DAA1: mov var_44, ecx
  loc_0041DAA4: mov ebx, 00000008h
  loc_0041DAA9: lea edx, var_7C
  loc_0041DAAC: lea ecx, var_3C
  loc_0041DAAF: mov var_5C, eax
  loc_0041DAB2: mov var_4C, eax
  loc_0041DAB5: mov var_74, 0040FAD8h ; "Check Standard Error"
  loc_0041DABC: mov var_7C, ebx
  loc_0041DABF: call __vbaVarDup
  loc_0041DAC1: mov var_64, 0040F9F8h ; "Please enter a value from the estimated standard used when estimating the mean of Y given the value of X: Xg"
  loc_0041DAC8: lea edx, var_6C
  loc_0041DACB: lea ecx, var_2C
  loc_0041DACE: mov var_6C, ebx
  loc_0041DAD1: call __vbaVarDup
  loc_0041DAD3: lea eax, var_5C
  loc_0041DAD6: lea ecx, var_4C
  loc_0041DAD9: push eax
  loc_0041DADA: lea edx, var_3C
  loc_0041DADD: push ecx
  loc_0041DADE: push edx
  loc_0041DADF: lea eax, var_2C
  loc_0041DAE2: push 00000000h
  loc_0041DAE4: push eax
  loc_0041DAE5: call [00401038h] ; rtcMsgBox
  loc_0041DAEB: lea ecx, var_5C
  loc_0041DAEE: lea edx, var_4C
  loc_0041DAF1: push ecx
  loc_0041DAF2: lea eax, var_3C
  loc_0041DAF5: push edx
  loc_0041DAF6: lea ecx, var_2C
  loc_0041DAF9: push eax
  loc_0041DAFA: push ecx
  loc_0041DAFB: push 00000004h
  loc_0041DAFD: call edi
  loc_0041DAFF: add esp, 00000014h
  loc_0041DB02: jmp 0041DE03h
  loc_0041DB07: mov edx, [esi]
  loc_0041DB09: push esi
  loc_0041DB0A: call [edx+00000320h]
  loc_0041DB10: mov var_24, eax
  loc_0041DB13: lea eax, var_2C
  loc_0041DB16: lea ecx, var_3C
  loc_0041DB19: push eax
  loc_0041DB1A: push ecx
  loc_0041DB1B: mov var_2C, 00000009h
  loc_0041DB22: call ebx
  loc_0041DB24: lea edx, var_3C
  loc_0041DB27: lea eax, var_6C
  loc_0041DB2A: push edx
  loc_0041DB2B: push eax
  loc_0041DB2C: mov var_64, 0040F040h
  loc_0041DB33: mov var_6C, 00008008h
  loc_0041DB3A: call [00401070h] ; __vbaVarTstEq
  loc_0041DB40: lea ecx, var_3C
  loc_0041DB43: lea edx, var_2C
  loc_0041DB46: push ecx
  loc_0041DB47: push edx
  loc_0041DB48: push 00000002h
  loc_0041DB4A: mov var_A0, ax
  loc_0041DB51: call edi
  loc_0041DB53: add esp, 0000000Ch
  loc_0041DB56: cmp var_A0, 0000h
  loc_0041DB5E: jz 0041DC20h
  loc_0041DB64: mov ebx, [004010F4h] ; __vbaVarDup
  loc_0041DB6A: mov ecx, 80020004h
  loc_0041DB6F: mov var_54, ecx
  loc_0041DB72: mov eax, 0000000Ah
  loc_0041DB77: mov var_44, ecx
  loc_0041DB7A: lea edx, var_7C
  loc_0041DB7D: lea ecx, var_3C
  loc_0041DB80: mov var_5C, eax
  loc_0041DB83: mov var_4C, eax
  loc_0041DB86: mov var_74, 0040FAD8h ; "Check Standard Error"
  loc_0041DB8D: mov var_7C, 00000008h
  loc_0041DB94: call ebx
  loc_0041DB96: lea edx, var_6C
  loc_0041DB99: lea ecx, var_2C
  loc_0041DB9C: mov var_64, 0040FB08h ; "Please enter a value from the estimated standard used when predicting the value of Y given the value of X: Xg"
  loc_0041DBA3: mov var_6C, 00000008h
  loc_0041DBAA: call ebx
  loc_0041DBAC: lea eax, var_5C
  loc_0041DBAF: lea ecx, var_4C
  loc_0041DBB2: push eax
  loc_0041DBB3: lea edx, var_3C
  loc_0041DBB6: push ecx
  loc_0041DBB7: push edx
  loc_0041DBB8: lea eax, var_2C
  loc_0041DBBB: push 00000000h
  loc_0041DBBD: push eax
  loc_0041DBBE: call [00401038h] ; rtcMsgBox
  loc_0041DBC4: lea ecx, var_5C
  loc_0041DBC7: lea edx, var_4C
  loc_0041DBCA: push ecx
  loc_0041DBCB: lea eax, var_3C
  loc_0041DBCE: push edx
  loc_0041DBCF: lea ecx, var_2C
  loc_0041DBD2: push eax
  loc_0041DBD3: push ecx
  loc_0041DBD4: push 00000004h
  loc_0041DBD6: call edi
  loc_0041DBD8: mov edx, [esi]
  loc_0041DBDA: add esp, 00000014h
  loc_0041DBDD: push esi
  loc_0041DBDE: call [edx+00000320h]
  loc_0041DBE4: push eax
  loc_0041DBE5: lea eax, var_1C
  loc_0041DBE8: push eax
  loc_0041DBE9: call [0040103Ch] ; __vbaObjSet
  loc_0041DBEF: mov esi, eax
  loc_0041DBF1: push esi
  loc_0041DBF2: mov ecx, [esi]
  loc_0041DBF4: call [ecx+00000204h]
  loc_0041DBFA: test eax, eax
  loc_0041DBFC: fnclex
  loc_0041DBFE: jge 0041DC12h
  loc_0041DC00: push 00000204h
  loc_0041DC05: push 0040F02Ch
  loc_0041DC0A: push esi
  loc_0041DC0B: push eax
  loc_0041DC0C: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DC12: lea ecx, var_1C
  loc_0041DC15: call [00401114h] ; __vbaFreeObj
  loc_0041DC1B: jmp 0041DE03h
  loc_0041DC20: mov edx, [esi]
  loc_0041DC22: push esi
  loc_0041DC23: call [edx+0000030Ch]
  loc_0041DC29: push eax
  loc_0041DC2A: lea eax, var_1C
  loc_0041DC2D: push eax
  loc_0041DC2E: call [0040103Ch] ; __vbaObjSet
  loc_0041DC34: mov ecx, [eax]
  loc_0041DC36: lea edx, var_18
  loc_0041DC39: push edx
  loc_0041DC3A: push eax
  loc_0041DC3B: mov var_A0, eax
  loc_0041DC41: call [ecx+000000A0h]
  loc_0041DC47: test eax, eax
  loc_0041DC49: fnclex
  loc_0041DC4B: jge 0041DC65h
  loc_0041DC4D: mov ecx, var_A0
  loc_0041DC53: push 000000A0h
  loc_0041DC58: push 0040F02Ch
  loc_0041DC5D: push ecx
  loc_0041DC5E: push eax
  loc_0041DC5F: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DC65: mov eax, var_18
  loc_0041DC68: lea edx, var_2C
  loc_0041DC6B: mov var_24, eax
  loc_0041DC6E: lea eax, var_3C
  loc_0041DC71: push edx
  loc_0041DC72: push eax
  loc_0041DC73: mov var_18, 00000000h
  loc_0041DC7A: mov var_2C, 00000008h
  loc_0041DC81: call ebx
  loc_0041DC83: lea ecx, var_3C
  loc_0041DC86: push ecx
  loc_0041DC87: call [004010A0h] ; __vbaR4ErrVar
  loc_0041DC8D: fstp real4 ptr [00430038h]
  loc_0041DC93: lea ecx, var_1C
  loc_0041DC96: call [00401114h] ; __vbaFreeObj
  loc_0041DC9C: lea edx, var_3C
  loc_0041DC9F: lea eax, var_3C
  loc_0041DCA2: push edx
  loc_0041DCA3: lea ecx, var_2C
  loc_0041DCA6: push eax
  loc_0041DCA7: push ecx
  loc_0041DCA8: push 00000003h
  loc_0041DCAA: call edi
  loc_0041DCAC: mov edx, [esi]
  loc_0041DCAE: add esp, 00000010h
  loc_0041DCB1: push esi
  loc_0041DCB2: call [edx+00000320h]
  loc_0041DCB8: push eax
  loc_0041DCB9: lea eax, var_1C
  loc_0041DCBC: push eax
  loc_0041DCBD: call [0040103Ch] ; __vbaObjSet
  loc_0041DCC3: mov ecx, [eax]
  loc_0041DCC5: lea edx, var_18
  loc_0041DCC8: push edx
  loc_0041DCC9: push eax
  loc_0041DCCA: mov var_A0, eax
  loc_0041DCD0: call [ecx+000000A0h]
  loc_0041DCD6: test eax, eax
  loc_0041DCD8: fnclex
  loc_0041DCDA: jge 0041DCF4h
  loc_0041DCDC: mov ecx, var_A0
  loc_0041DCE2: push 000000A0h
  loc_0041DCE7: push 0040F02Ch
  loc_0041DCEC: push ecx
  loc_0041DCED: push eax
  loc_0041DCEE: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DCF4: mov eax, var_18
  loc_0041DCF7: lea edx, var_2C
  loc_0041DCFA: mov var_24, eax
  loc_0041DCFD: lea eax, var_3C
  loc_0041DD00: push edx
  loc_0041DD01: push eax
  loc_0041DD02: mov var_18, 00000000h
  loc_0041DD09: mov var_2C, 00000008h
  loc_0041DD10: call ebx
  loc_0041DD12: lea ecx, var_3C
  loc_0041DD15: push ecx
  loc_0041DD16: call [004010A0h] ; __vbaR4ErrVar
  loc_0041DD1C: fstp real4 ptr [00430034h]
  loc_0041DD22: lea ecx, var_1C
  loc_0041DD25: call [00401114h] ; __vbaFreeObj
  loc_0041DD2B: lea edx, var_3C
  loc_0041DD2E: lea eax, var_3C
  loc_0041DD31: push edx
  loc_0041DD32: lea ecx, var_2C
  loc_0041DD35: push eax
  loc_0041DD36: push ecx
  loc_0041DD37: push 00000003h
  loc_0041DD39: call edi
  loc_0041DD3B: mov edx, [esi]
  loc_0041DD3D: add esp, 00000010h
  loc_0041DD40: push esi
  loc_0041DD41: call [edx+00000304h]
  loc_0041DD47: push eax
  loc_0041DD48: lea eax, var_1C
  loc_0041DD4B: push eax
  loc_0041DD4C: call [0040103Ch] ; __vbaObjSet
  loc_0041DD52: mov esi, eax
  loc_0041DD54: lea edx, var_18
  loc_0041DD57: push edx
  loc_0041DD58: push esi
  loc_0041DD59: mov ecx, [esi]
  loc_0041DD5B: call [ecx+000000A0h]
  loc_0041DD61: test eax, eax
  loc_0041DD63: fnclex
  loc_0041DD65: jge 0041DD79h
  loc_0041DD67: push 000000A0h
  loc_0041DD6C: push 0040F02Ch
  loc_0041DD71: push esi
  loc_0041DD72: push eax
  loc_0041DD73: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DD79: mov eax, var_18
  loc_0041DD7C: lea ecx, var_3C
  loc_0041DD7F: mov var_24, eax
  loc_0041DD82: lea eax, var_2C
  loc_0041DD85: push eax
  loc_0041DD86: push ecx
  loc_0041DD87: mov var_18, 00000000h
  loc_0041DD8E: mov var_2C, 00000008h
  loc_0041DD95: call ebx
  loc_0041DD97: lea edx, var_3C
  loc_0041DD9A: push edx
  loc_0041DD9B: call [004010A0h] ; __vbaR4ErrVar
  loc_0041DDA1: fstp real4 ptr [0043003Ch]
  loc_0041DDA7: lea ecx, var_1C
  loc_0041DDAA: call [00401114h] ; __vbaFreeObj
  loc_0041DDB0: lea eax, var_3C
  loc_0041DDB3: lea ecx, var_3C
  loc_0041DDB6: push eax
  loc_0041DDB7: lea edx, var_2C
  loc_0041DDBA: push ecx
  loc_0041DDBB: push edx
  loc_0041DDBC: push 00000003h
  loc_0041DDBE: call edi
  loc_0041DDC0: mov eax, [00430128h]
  loc_0041DDC5: add esp, 00000010h
  loc_0041DDC8: test eax, eax
  loc_0041DDCA: jnz 0041DDDCh
  loc_0041DDCC: push 00430128h
  loc_0041DDD1: push 00404AE0h
  loc_0041DDD6: call [004010D4h] ; __vbaNew2
  loc_0041DDDC: mov esi, [00430128h]
  loc_0041DDE2: push esi
  loc_0041DDE3: mov eax, [esi]
  loc_0041DDE5: call [eax+000002B4h]
  loc_0041DDEB: test eax, eax
  loc_0041DDED: fnclex
  loc_0041DDEF: jge 0041DE03h
  loc_0041DDF1: push 000002B4h
  loc_0041DDF6: push 0040F8B4h
  loc_0041DDFB: push esi
  loc_0041DDFC: push eax
  loc_0041DDFD: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DE03: mov var_4, 00000000h
  loc_0041DE0A: fwait
  loc_0041DE0B: push 0041DE41h
  loc_0041DE10: jmp 0041DE40h
  loc_0041DE12: lea ecx, var_18
  loc_0041DE15: call [00401110h] ; __vbaFreeStr
  loc_0041DE1B: lea ecx, var_1C
  loc_0041DE1E: call [00401114h] ; __vbaFreeObj
  loc_0041DE24: lea ecx, var_5C
  loc_0041DE27: lea edx, var_4C
  loc_0041DE2A: push ecx
  loc_0041DE2B: lea eax, var_3C
  loc_0041DE2E: push edx
  loc_0041DE2F: lea ecx, var_2C
  loc_0041DE32: push eax
  loc_0041DE33: push ecx
  loc_0041DE34: push 00000004h
  loc_0041DE36: call [00401018h] ; __vbaFreeVarList
  loc_0041DE3C: add esp, 00000014h
  loc_0041DE3F: ret
  loc_0041DE40: ret
  loc_0041DE41: mov eax, Me
  loc_0041DE44: push eax
  loc_0041DE45: mov edx, [eax]
  loc_0041DE47: call [edx+00000008h]
  loc_0041DE4A: mov eax, var_4
  loc_0041DE4D: mov ecx, var_14
  loc_0041DE50: pop edi
  loc_0041DE51: pop esi
  loc_0041DE52: mov fs:[00000000h], ecx
  loc_0041DE59: pop ebx
  loc_0041DE5A: mov esp, ebp
  loc_0041DE5C: pop ebp
  loc_0041DE5D: retn 0004h
End Sub

Private Sub Form_Activate() '41DE60
  loc_0041DE60: push ebp
  loc_0041DE61: mov ebp, esp
  loc_0041DE63: sub esp, 0000000Ch
  loc_0041DE66: push 00401926h ; __vbaExceptHandler
  loc_0041DE6B: mov eax, fs:[00000000h]
  loc_0041DE71: push eax
  loc_0041DE72: mov fs:[00000000h], esp
  loc_0041DE79: sub esp, 00000018h
  loc_0041DE7C: push ebx
  loc_0041DE7D: push esi
  loc_0041DE7E: push edi
  loc_0041DE7F: mov var_C, esp
  loc_0041DE82: mov var_8, 00401278h
  loc_0041DE89: mov esi, Me
  loc_0041DE8C: mov eax, esi
  loc_0041DE8E: and eax, 00000001h
  loc_0041DE91: mov var_4, eax
  loc_0041DE94: and esi, FFFFFFFEh
  loc_0041DE97: push esi
  loc_0041DE98: mov Me, esi
  loc_0041DE9B: mov ecx, [esi]
  loc_0041DE9D: call [ecx+00000004h]
  loc_0041DEA0: mov edx, [esi]
  loc_0041DEA2: xor eax, eax
  loc_0041DEA4: push esi
  loc_0041DEA5: mov var_18, eax
  loc_0041DEA8: mov var_1C, eax
  loc_0041DEAB: call [edx+0000030Ch]
  loc_0041DEB1: push eax
  loc_0041DEB2: lea eax, var_1C
  loc_0041DEB5: push eax
  loc_0041DEB6: call [0040103Ch] ; __vbaObjSet
  loc_0041DEBC: mov ecx, [00430038h]
  loc_0041DEC2: mov edi, eax
  loc_0041DEC4: push ecx
  loc_0041DEC5: mov ebx, [edi]
  loc_0041DEC7: call [0040107Ch] ; __vbaStrR4
  loc_0041DECD: mov edx, eax
  loc_0041DECF: lea ecx, var_18
  loc_0041DED2: call [004010FCh] ; __vbaStrMove
  loc_0041DED8: push eax
  loc_0041DED9: push edi
  loc_0041DEDA: call [ebx+000000A4h]
  loc_0041DEE0: test eax, eax
  loc_0041DEE2: fnclex
  loc_0041DEE4: jge 0041DEF8h
  loc_0041DEE6: push 000000A4h
  loc_0041DEEB: push 0040F02Ch
  loc_0041DEF0: push edi
  loc_0041DEF1: push eax
  loc_0041DEF2: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DEF8: lea ecx, var_18
  loc_0041DEFB: call [00401110h] ; __vbaFreeStr
  loc_0041DF01: lea ecx, var_1C
  loc_0041DF04: call [00401114h] ; __vbaFreeObj
  loc_0041DF0A: mov edx, [esi]
  loc_0041DF0C: push esi
  loc_0041DF0D: call [edx+00000320h]
  loc_0041DF13: push eax
  loc_0041DF14: lea eax, var_1C
  loc_0041DF17: push eax
  loc_0041DF18: call [0040103Ch] ; __vbaObjSet
  loc_0041DF1E: mov ecx, [00430034h]
  loc_0041DF24: mov edi, eax
  loc_0041DF26: push ecx
  loc_0041DF27: mov ebx, [edi]
  loc_0041DF29: call [0040107Ch] ; __vbaStrR4
  loc_0041DF2F: mov edx, eax
  loc_0041DF31: lea ecx, var_18
  loc_0041DF34: call [004010FCh] ; __vbaStrMove
  loc_0041DF3A: push eax
  loc_0041DF3B: push edi
  loc_0041DF3C: call [ebx+000000A4h]
  loc_0041DF42: test eax, eax
  loc_0041DF44: fnclex
  loc_0041DF46: jge 0041DF5Ah
  loc_0041DF48: push 000000A4h
  loc_0041DF4D: push 0040F02Ch
  loc_0041DF52: push edi
  loc_0041DF53: push eax
  loc_0041DF54: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DF5A: lea ecx, var_18
  loc_0041DF5D: call [00401110h] ; __vbaFreeStr
  loc_0041DF63: lea ecx, var_1C
  loc_0041DF66: call [00401114h] ; __vbaFreeObj
  loc_0041DF6C: mov edx, [esi]
  loc_0041DF6E: push esi
  loc_0041DF6F: call [edx+00000304h]
  loc_0041DF75: push eax
  loc_0041DF76: lea eax, var_1C
  loc_0041DF79: push eax
  loc_0041DF7A: call [0040103Ch] ; __vbaObjSet
  loc_0041DF80: mov ecx, [0043003Ch]
  loc_0041DF86: mov edi, eax
  loc_0041DF88: push ecx
  loc_0041DF89: mov ebx, [edi]
  loc_0041DF8B: call [0040107Ch] ; __vbaStrR4
  loc_0041DF91: mov edx, eax
  loc_0041DF93: lea ecx, var_18
  loc_0041DF96: call [004010FCh] ; __vbaStrMove
  loc_0041DF9C: push eax
  loc_0041DF9D: push edi
  loc_0041DF9E: call [ebx+000000A4h]
  loc_0041DFA4: test eax, eax
  loc_0041DFA6: fnclex
  loc_0041DFA8: jge 0041DFBCh
  loc_0041DFAA: push 000000A4h
  loc_0041DFAF: push 0040F02Ch
  loc_0041DFB4: push edi
  loc_0041DFB5: push eax
  loc_0041DFB6: call [00401030h] ; __vbaHresultCheckObj
  loc_0041DFBC: lea ecx, var_18
  loc_0041DFBF: call [00401110h] ; __vbaFreeStr
  loc_0041DFC5: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041DFCB: lea ecx, var_1C
  loc_0041DFCE: call edi
  loc_0041DFD0: mov edx, [esi]
  loc_0041DFD2: push esi
  loc_0041DFD3: call [edx+0000030Ch]
  loc_0041DFD9: push eax
  loc_0041DFDA: lea eax, var_1C
  loc_0041DFDD: push eax
  loc_0041DFDE: call [0040103Ch] ; __vbaObjSet
  loc_0041DFE4: mov esi, eax
  loc_0041DFE6: push esi
  loc_0041DFE7: mov ecx, [esi]
  loc_0041DFE9: call [ecx+00000204h]
  loc_0041DFEF: test eax, eax
  loc_0041DFF1: fnclex
  loc_0041DFF3: jge 0041E007h
  loc_0041DFF5: push 00000204h
  loc_0041DFFA: push 0040F02Ch
  loc_0041DFFF: push esi
  loc_0041E000: push eax
  loc_0041E001: call [00401030h] ; __vbaHresultCheckObj
  loc_0041E007: lea ecx, var_1C
  loc_0041E00A: call edi
  loc_0041E00C: mov var_4, 00000000h
  loc_0041E013: fwait
  loc_0041E014: push 0041E02Fh
  loc_0041E019: jmp 0041E02Eh
  loc_0041E01B: lea ecx, var_18
  loc_0041E01E: call [00401110h] ; __vbaFreeStr
  loc_0041E024: lea ecx, var_1C
  loc_0041E027: call [00401114h] ; __vbaFreeObj
  loc_0041E02D: ret
  loc_0041E02E: ret
  loc_0041E02F: mov eax, Me
  loc_0041E032: push eax
  loc_0041E033: mov edx, [eax]
  loc_0041E035: call [edx+00000008h]
  loc_0041E038: mov eax, var_4
  loc_0041E03B: mov ecx, var_14
  loc_0041E03E: pop edi
  loc_0041E03F: pop esi
  loc_0041E040: mov fs:[00000000h], ecx
  loc_0041E047: pop ebx
  loc_0041E048: mov esp, ebp
  loc_0041E04A: pop ebp
  loc_0041E04B: retn 0004h
End Sub

Private Sub cmdRestore_Click() '41D6E0
  loc_0041D6E0: push ebp
  loc_0041D6E1: mov ebp, esp
  loc_0041D6E3: sub esp, 0000000Ch
  loc_0041D6E6: push 00401926h ; __vbaExceptHandler
  loc_0041D6EB: mov eax, fs:[00000000h]
  loc_0041D6F1: push eax
  loc_0041D6F2: mov fs:[00000000h], esp
  loc_0041D6F9: sub esp, 00000018h
  loc_0041D6FC: push ebx
  loc_0041D6FD: push esi
  loc_0041D6FE: push edi
  loc_0041D6FF: mov var_C, esp
  loc_0041D702: mov var_8, 00401258h
  loc_0041D709: mov esi, Me
  loc_0041D70C: mov eax, esi
  loc_0041D70E: and eax, 00000001h
  loc_0041D711: mov var_4, eax
  loc_0041D714: and esi, FFFFFFFEh
  loc_0041D717: push esi
  loc_0041D718: mov Me, esi
  loc_0041D71B: mov ecx, [esi]
  loc_0041D71D: call [ecx+00000004h]
  loc_0041D720: mov [00430038h], 447A8000h
  loc_0041D72A: mov [00430030h], 42F00000h
  loc_0041D734: mov [00430038h], 41C00000h
  loc_0041D73E: mov edx, [esi]
  loc_0041D740: xor eax, eax
  loc_0041D742: push esi
  loc_0041D743: mov var_18, eax
  loc_0041D746: mov var_1C, eax
  loc_0041D749: call [edx+0000030Ch]
  loc_0041D74F: push eax
  loc_0041D750: lea eax, var_1C
  loc_0041D753: push eax
  loc_0041D754: call [0040103Ch] ; __vbaObjSet
  loc_0041D75A: mov ecx, [00430038h]
  loc_0041D760: mov edi, eax
  loc_0041D762: push ecx
  loc_0041D763: mov ebx, [edi]
  loc_0041D765: call [0040107Ch] ; __vbaStrR4
  loc_0041D76B: mov edx, eax
  loc_0041D76D: lea ecx, var_18
  loc_0041D770: call [004010FCh] ; __vbaStrMove
  loc_0041D776: push eax
  loc_0041D777: push edi
  loc_0041D778: call [ebx+000000A4h]
  loc_0041D77E: test eax, eax
  loc_0041D780: fnclex
  loc_0041D782: jge 0041D796h
  loc_0041D784: push 000000A4h
  loc_0041D789: push 0040F02Ch
  loc_0041D78E: push edi
  loc_0041D78F: push eax
  loc_0041D790: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D796: lea ecx, var_18
  loc_0041D799: call [00401110h] ; __vbaFreeStr
  loc_0041D79F: lea ecx, var_1C
  loc_0041D7A2: call [00401114h] ; __vbaFreeObj
  loc_0041D7A8: mov edx, [esi]
  loc_0041D7AA: push esi
  loc_0041D7AB: call [edx+00000320h]
  loc_0041D7B1: push eax
  loc_0041D7B2: lea eax, var_1C
  loc_0041D7B5: push eax
  loc_0041D7B6: call [0040103Ch] ; __vbaObjSet
  loc_0041D7BC: mov ecx, [00430034h]
  loc_0041D7C2: mov edi, eax
  loc_0041D7C4: push ecx
  loc_0041D7C5: mov ebx, [edi]
  loc_0041D7C7: call [0040107Ch] ; __vbaStrR4
  loc_0041D7CD: mov edx, eax
  loc_0041D7CF: lea ecx, var_18
  loc_0041D7D2: call [004010FCh] ; __vbaStrMove
  loc_0041D7D8: push eax
  loc_0041D7D9: push edi
  loc_0041D7DA: call [ebx+000000A4h]
  loc_0041D7E0: test eax, eax
  loc_0041D7E2: fnclex
  loc_0041D7E4: jge 0041D7F8h
  loc_0041D7E6: push 000000A4h
  loc_0041D7EB: push 0040F02Ch
  loc_0041D7F0: push edi
  loc_0041D7F1: push eax
  loc_0041D7F2: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D7F8: mov edi, [00401110h] ; __vbaFreeStr
  loc_0041D7FE: lea ecx, var_18
  loc_0041D801: call edi
  loc_0041D803: lea ecx, var_1C
  loc_0041D806: call [00401114h] ; __vbaFreeObj
  loc_0041D80C: mov edx, [esi]
  loc_0041D80E: push esi
  loc_0041D80F: call [edx+00000304h]
  loc_0041D815: push eax
  loc_0041D816: lea eax, var_1C
  loc_0041D819: push eax
  loc_0041D81A: call [0040103Ch] ; __vbaObjSet
  loc_0041D820: mov ecx, [0043003Ch]
  loc_0041D826: mov esi, eax
  loc_0041D828: push ecx
  loc_0041D829: mov ebx, [esi]
  loc_0041D82B: call [0040107Ch] ; __vbaStrR4
  loc_0041D831: mov edx, eax
  loc_0041D833: lea ecx, var_18
  loc_0041D836: call [004010FCh] ; __vbaStrMove
  loc_0041D83C: push eax
  loc_0041D83D: push esi
  loc_0041D83E: call [ebx+000000A4h]
  loc_0041D844: test eax, eax
  loc_0041D846: fnclex
  loc_0041D848: jge 0041D85Ch
  loc_0041D84A: push 000000A4h
  loc_0041D84F: push 0040F02Ch
  loc_0041D854: push esi
  loc_0041D855: push eax
  loc_0041D856: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D85C: lea ecx, var_18
  loc_0041D85F: call edi
  loc_0041D861: lea ecx, var_1C
  loc_0041D864: call [00401114h] ; __vbaFreeObj
  loc_0041D86A: mov var_4, 00000000h
  loc_0041D871: fwait
  loc_0041D872: push 0041D88Dh
  loc_0041D877: jmp 0041D88Ch
  loc_0041D879: lea ecx, var_18
  loc_0041D87C: call [00401110h] ; __vbaFreeStr
  loc_0041D882: lea ecx, var_1C
  loc_0041D885: call [00401114h] ; __vbaFreeObj
  loc_0041D88B: ret
  loc_0041D88C: ret
  loc_0041D88D: mov eax, Me
  loc_0041D890: push eax
  loc_0041D891: mov edx, [eax]
  loc_0041D893: call [edx+00000008h]
  loc_0041D896: mov eax, var_4
  loc_0041D899: mov ecx, var_14
  loc_0041D89C: pop edi
  loc_0041D89D: pop esi
  loc_0041D89E: mov fs:[00000000h], ecx
  loc_0041D8A5: pop ebx
  loc_0041D8A6: mov esp, ebp
  loc_0041D8A8: pop ebp
  loc_0041D8A9: retn 0004h
End Sub
