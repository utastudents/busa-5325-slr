VERSION 5.00
Begin VB.Form frmChangeCorr
  Caption = "Change R and R-squared"
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 405
  ClientWidth = 7665
  ClientHeight = 5130
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fraChange
    Caption = "Calculated by hand or found on a printout"
    Left = 0
    Top = 0
    Width = 7575
    Height = 5085
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
      Left = 690
      Top = 4080
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
      Left = 2520
      Top = 4080
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
      Left = 4320
      Top = 4080
      Width = 2535
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
      Caption = "Pearson's Correlation Coefficient: r"
      Left = 720
      Top = 1080
      Width = 6495
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
      Begin VB.TextBox txtCorr
        Left = 360
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
      Caption = "Coefficient of Determination: r-squared"
      Left = 720
      Top = 2520
      Width = 6495
      Height = 1335
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
      Begin VB.TextBox txtRSquared
        Left = 360
        Top = 600
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

Attribute VB_Name = "frmChangeCorr"


Private Sub txtCorr_KeyPress(KeyAscii As Integer) '41D1E0
  loc_0041D1E0: push ebp
  loc_0041D1E1: mov ebp, esp
  loc_0041D1E3: sub esp, 0000000Ch
  loc_0041D1E6: push 00401926h ; __vbaExceptHandler
  loc_0041D1EB: mov eax, fs:[00000000h]
  loc_0041D1F1: push eax
  loc_0041D1F2: mov fs:[00000000h], esp
  loc_0041D1F9: sub esp, 0000007Ch
  loc_0041D1FC: push ebx
  loc_0041D1FD: push esi
  loc_0041D1FE: push edi
  loc_0041D1FF: mov var_C, esp
  loc_0041D202: mov var_8, 00401228h
  loc_0041D209: mov esi, Me
  loc_0041D20C: mov eax, esi
  loc_0041D20E: and eax, 00000001h
  loc_0041D211: mov var_4, eax
  loc_0041D214: and esi, FFFFFFFEh
  loc_0041D217: push esi
  loc_0041D218: mov Me, esi
  loc_0041D21B: mov ecx, [esi]
  loc_0041D21D: call [ecx+00000004h]
  loc_0041D220: mov ebx, KeyAscii
  loc_0041D223: lea edx, var_24
  loc_0041D226: xor edi, edi
  loc_0041D228: push ebx
  loc_0041D229: push edx
  loc_0041D22A: mov var_24, edi
  loc_0041D22D: mov var_34, edi
  loc_0041D230: mov var_44, edi
  loc_0041D233: mov var_54, edi
  loc_0041D236: mov var_64, edi
  loc_0041D239: mov var_74, edi
  loc_0041D23C: mov var_84, edi
  loc_0041D242: call 0041A480h
  loc_0041D247: lea eax, var_24
  loc_0041D24A: push eax
  loc_0041D24B: call [004010C4h] ; __vbaI2Var
  loc_0041D251: lea ecx, var_24
  loc_0041D254: mov [ebx], ax
  loc_0041D257: call [00401010h] ; __vbaFreeVar
  loc_0041D25D: push 0040DD3Ch ; "."
  loc_0041D262: call [00401024h] ; rtcAnsiValueBstr
  loc_0041D268: mov edx, [esi]
  loc_0041D26A: xor ecx, ecx
  loc_0041D26C: cmp [ebx], ax
  loc_0041D26F: push esi
  loc_0041D270: mov var_84, 0000000Bh
  loc_0041D27A: setz cl
  loc_0041D27D: neg ecx
  loc_0041D27F: mov var_7C, cx
  loc_0041D283: call [edx+00000310h]
  loc_0041D289: mov var_1C, eax
  loc_0041D28C: lea eax, var_84
  loc_0041D292: push eax
  loc_0041D293: lea ecx, var_24
  loc_0041D296: push 00000001h
  loc_0041D298: lea edx, var_64
  loc_0041D29B: push ecx
  loc_0041D29C: push edx
  loc_0041D29D: lea eax, var_34
  loc_0041D2A0: push edi
  loc_0041D2A1: push eax
  loc_0041D2A2: mov var_24, 00000009h
  loc_0041D2A9: mov var_5C, 0040DD3Ch ; "."
  loc_0041D2B0: mov var_64, 00000008h
  loc_0041D2B7: mov var_6C, edi
  loc_0041D2BA: mov var_74, 00008002h
  loc_0041D2C1: call [004010B4h] ; __vbaInStrVar
  loc_0041D2C7: lea ecx, var_74
  loc_0041D2CA: push eax
  loc_0041D2CB: lea edx, var_44
  loc_0041D2CE: push ecx
  loc_0041D2CF: push edx
  loc_0041D2D0: call [00401060h] ; __vbaVarCmpGt
  loc_0041D2D6: push eax
  loc_0041D2D7: lea eax, var_54
  loc_0041D2DA: push eax
  loc_0041D2DB: call [00401094h] ; __vbaVarAnd
  loc_0041D2E1: push eax
  loc_0041D2E2: call [00401058h] ; __vbaBoolVarNull
  loc_0041D2E8: lea ecx, var_84
  loc_0041D2EE: mov esi, eax
  loc_0041D2F0: lea edx, var_34
  loc_0041D2F3: push ecx
  loc_0041D2F4: lea eax, var_24
  loc_0041D2F7: push edx
  loc_0041D2F8: push eax
  loc_0041D2F9: push 00000003h
  loc_0041D2FB: call [00401018h] ; __vbaFreeVarList
  loc_0041D301: add esp, 00000010h
  loc_0041D304: cmp si, di
  loc_0041D307: jz 0041D30Ch
  loc_0041D309: mov [ebx], di
  loc_0041D30C: mov var_4, edi
  loc_0041D30F: push 0041D333h
  loc_0041D314: jmp 0041D332h
  loc_0041D316: lea ecx, var_54
  loc_0041D319: lea edx, var_44
  loc_0041D31C: push ecx
  loc_0041D31D: lea eax, var_34
  loc_0041D320: push edx
  loc_0041D321: lea ecx, var_24
  loc_0041D324: push eax
  loc_0041D325: push ecx
  loc_0041D326: push 00000004h
  loc_0041D328: call [00401018h] ; __vbaFreeVarList
  loc_0041D32E: add esp, 00000014h
  loc_0041D331: ret
  loc_0041D332: ret
  loc_0041D333: mov eax, Me
  loc_0041D336: push eax
  loc_0041D337: mov edx, [eax]
  loc_0041D339: call [edx+00000008h]
  loc_0041D33C: mov eax, var_4
  loc_0041D33F: mov ecx, var_14
  loc_0041D342: pop edi
  loc_0041D343: pop esi
  loc_0041D344: mov fs:[00000000h], ecx
  loc_0041D34B: pop ebx
  loc_0041D34C: mov esp, ebp
  loc_0041D34E: pop ebp
  loc_0041D34F: retn 0008h
End Sub

Private Sub txtRSquared_KeyPress(KeyAscii As Integer) '41D360
  loc_0041D360: push ebp
  loc_0041D361: mov ebp, esp
  loc_0041D363: sub esp, 0000000Ch
  loc_0041D366: push 00401926h ; __vbaExceptHandler
  loc_0041D36B: mov eax, fs:[00000000h]
  loc_0041D371: push eax
  loc_0041D372: mov fs:[00000000h], esp
  loc_0041D379: sub esp, 0000008Ch
  loc_0041D37F: push ebx
  loc_0041D380: push esi
  loc_0041D381: push edi
  loc_0041D382: mov var_C, esp
  loc_0041D385: mov var_8, 00401238h
  loc_0041D38C: mov esi, Me
  loc_0041D38F: mov eax, esi
  loc_0041D391: and eax, 00000001h
  loc_0041D394: mov var_4, eax
  loc_0041D397: and esi, FFFFFFFEh
  loc_0041D39A: push esi
  loc_0041D39B: mov Me, esi
  loc_0041D39E: mov ecx, [esi]
  loc_0041D3A0: call [ecx+00000004h]
  loc_0041D3A3: mov edi, KeyAscii
  loc_0041D3A6: lea edx, var_24
  loc_0041D3A9: xor ebx, ebx
  loc_0041D3AB: push edi
  loc_0041D3AC: push edx
  loc_0041D3AD: mov var_24, ebx
  loc_0041D3B0: mov var_34, ebx
  loc_0041D3B3: mov var_44, ebx
  loc_0041D3B6: mov var_54, ebx
  loc_0041D3B9: mov var_64, ebx
  loc_0041D3BC: mov var_74, ebx
  loc_0041D3BF: mov var_84, ebx
  loc_0041D3C5: call 0041A480h
  loc_0041D3CA: lea eax, var_24
  loc_0041D3CD: push eax
  loc_0041D3CE: call [004010C4h] ; __vbaI2Var
  loc_0041D3D4: lea ecx, var_24
  loc_0041D3D7: mov [edi], ax
  loc_0041D3DA: call [00401010h] ; __vbaFreeVar
  loc_0041D3E0: push 0040DD3Ch ; "."
  loc_0041D3E5: call [00401024h] ; rtcAnsiValueBstr
  loc_0041D3EB: mov edx, [esi]
  loc_0041D3ED: xor ecx, ecx
  loc_0041D3EF: cmp [edi], ax
  loc_0041D3F2: push esi
  loc_0041D3F3: mov var_84, 0000000Bh
  loc_0041D3FD: setz cl
  loc_0041D400: neg ecx
  loc_0041D402: mov var_7C, cx
  loc_0041D406: call [edx+00000318h]
  loc_0041D40C: mov var_1C, eax
  loc_0041D40F: lea eax, var_84
  loc_0041D415: push eax
  loc_0041D416: lea ecx, var_24
  loc_0041D419: push 00000001h
  loc_0041D41B: lea edx, var_64
  loc_0041D41E: push ecx
  loc_0041D41F: push edx
  loc_0041D420: lea eax, var_34
  loc_0041D423: push ebx
  loc_0041D424: push eax
  loc_0041D425: mov var_24, 00000009h
  loc_0041D42C: mov var_5C, 0040DD3Ch ; "."
  loc_0041D433: mov var_64, 00000008h
  loc_0041D43A: mov var_6C, ebx
  loc_0041D43D: mov var_74, 00008002h
  loc_0041D444: call [004010B4h] ; __vbaInStrVar
  loc_0041D44A: lea ecx, var_74
  loc_0041D44D: push eax
  loc_0041D44E: lea edx, var_44
  loc_0041D451: push ecx
  loc_0041D452: push edx
  loc_0041D453: call [00401060h] ; __vbaVarCmpGt
  loc_0041D459: push eax
  loc_0041D45A: lea eax, var_54
  loc_0041D45D: push eax
  loc_0041D45E: call [00401094h] ; __vbaVarAnd
  loc_0041D464: push eax
  loc_0041D465: call [00401058h] ; __vbaBoolVarNull
  loc_0041D46B: lea ecx, var_84
  loc_0041D471: mov esi, eax
  loc_0041D473: lea edx, var_34
  loc_0041D476: push ecx
  loc_0041D477: lea eax, var_24
  loc_0041D47A: push edx
  loc_0041D47B: push eax
  loc_0041D47C: push 00000003h
  loc_0041D47E: call [00401018h] ; __vbaFreeVarList
  loc_0041D484: add esp, 00000010h
  loc_0041D487: cmp si, bx
  loc_0041D48A: jz 0041D48Fh
  loc_0041D48C: mov [edi], bx
  loc_0041D48F: push 0040DD44h ; "-"
  loc_0041D494: call [00401024h] ; rtcAnsiValueBstr
  loc_0041D49A: cmp [edi], ax
  loc_0041D49D: jnz 0041D50Ah
  loc_0041D49F: mov ecx, 80020004h
  loc_0041D4A4: mov eax, 0000000Ah
  loc_0041D4A9: mov var_4C, ecx
  loc_0041D4AC: mov var_3C, ecx
  loc_0041D4AF: mov var_2C, ecx
  loc_0041D4B2: lea edx, var_64
  loc_0041D4B5: lea ecx, var_24
  loc_0041D4B8: mov var_54, eax
  loc_0041D4BB: mov var_44, eax
  loc_0041D4BE: mov var_34, eax
  loc_0041D4C1: mov var_5C, 0040F7E4h ; "The coefficient of determination is always positive."
  loc_0041D4C8: mov var_64, 00000008h
  loc_0041D4CF: call [004010F4h] ; __vbaVarDup
  loc_0041D4D5: lea ecx, var_54
  loc_0041D4D8: lea edx, var_44
  loc_0041D4DB: push ecx
  loc_0041D4DC: lea eax, var_34
  loc_0041D4DF: push edx
  loc_0041D4E0: push eax
  loc_0041D4E1: lea ecx, var_24
  loc_0041D4E4: push ebx
  loc_0041D4E5: push ecx
  loc_0041D4E6: call [00401038h] ; rtcMsgBox
  loc_0041D4EC: lea edx, var_54
  loc_0041D4EF: lea eax, var_44
  loc_0041D4F2: push edx
  loc_0041D4F3: lea ecx, var_34
  loc_0041D4F6: push eax
  loc_0041D4F7: lea edx, var_24
  loc_0041D4FA: push ecx
  loc_0041D4FB: push edx
  loc_0041D4FC: push 00000004h
  loc_0041D4FE: call [00401018h] ; __vbaFreeVarList
  loc_0041D504: add esp, 00000014h
  loc_0041D507: mov [edi], bx
  loc_0041D50A: mov var_4, ebx
  loc_0041D50D: push 0041D531h
  loc_0041D512: jmp 0041D530h
  loc_0041D514: lea eax, var_54
  loc_0041D517: lea ecx, var_44
  loc_0041D51A: push eax
  loc_0041D51B: lea edx, var_34
  loc_0041D51E: push ecx
  loc_0041D51F: lea eax, var_24
  loc_0041D522: push edx
  loc_0041D523: push eax
  loc_0041D524: push 00000004h
  loc_0041D526: call [00401018h] ; __vbaFreeVarList
  loc_0041D52C: add esp, 00000014h
  loc_0041D52F: ret
  loc_0041D530: ret
  loc_0041D531: mov eax, Me
  loc_0041D534: push eax
  loc_0041D535: mov ecx, [eax]
  loc_0041D537: call [ecx+00000008h]
  loc_0041D53A: mov eax, var_4
  loc_0041D53D: mov ecx, var_14
  loc_0041D540: pop edi
  loc_0041D541: pop esi
  loc_0041D542: mov fs:[00000000h], ecx
  loc_0041D549: pop ebx
  loc_0041D54A: mov esp, ebp
  loc_0041D54C: pop ebp
  loc_0041D54D: retn 0008h
End Sub

Private Sub cmdClearChange_Click() '41C340
  loc_0041C340: push ebp
  loc_0041C341: mov ebp, esp
  loc_0041C343: sub esp, 0000000Ch
  loc_0041C346: push 00401926h ; __vbaExceptHandler
  loc_0041C34B: mov eax, fs:[00000000h]
  loc_0041C351: push eax
  loc_0041C352: mov fs:[00000000h], esp
  loc_0041C359: sub esp, 00000014h
  loc_0041C35C: push ebx
  loc_0041C35D: push esi
  loc_0041C35E: push edi
  loc_0041C35F: mov var_C, esp
  loc_0041C362: mov var_8, 004011C8h
  loc_0041C369: mov esi, Me
  loc_0041C36C: mov eax, esi
  loc_0041C36E: and eax, 00000001h
  loc_0041C371: mov var_4, eax
  loc_0041C374: and esi, FFFFFFFEh
  loc_0041C377: push esi
  loc_0041C378: mov Me, esi
  loc_0041C37B: mov ecx, [esi]
  loc_0041C37D: call [ecx+00000004h]
  loc_0041C380: mov edx, [esi]
  loc_0041C382: push esi
  loc_0041C383: mov var_18, 00000000h
  loc_0041C38A: call [edx+00000310h]
  loc_0041C390: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0041C396: push eax
  loc_0041C397: lea eax, var_18
  loc_0041C39A: push eax
  loc_0041C39B: call ebx
  loc_0041C39D: mov edi, eax
  loc_0041C39F: push 0040F040h
  loc_0041C3A4: push edi
  loc_0041C3A5: mov ecx, [edi]
  loc_0041C3A7: call [ecx+000000A4h]
  loc_0041C3AD: test eax, eax
  loc_0041C3AF: fnclex
  loc_0041C3B1: jge 0041C3C5h
  loc_0041C3B3: push 000000A4h
  loc_0041C3B8: push 0040F02Ch
  loc_0041C3BD: push edi
  loc_0041C3BE: push eax
  loc_0041C3BF: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C3C5: lea ecx, var_18
  loc_0041C3C8: call [00401114h] ; __vbaFreeObj
  loc_0041C3CE: mov edx, [esi]
  loc_0041C3D0: push esi
  loc_0041C3D1: call [edx+00000318h]
  loc_0041C3D7: push eax
  loc_0041C3D8: lea eax, var_18
  loc_0041C3DB: push eax
  loc_0041C3DC: call ebx
  loc_0041C3DE: mov edi, eax
  loc_0041C3E0: push 0040F040h
  loc_0041C3E5: push edi
  loc_0041C3E6: mov ecx, [edi]
  loc_0041C3E8: call [ecx+000000A4h]
  loc_0041C3EE: test eax, eax
  loc_0041C3F0: fnclex
  loc_0041C3F2: jge 0041C406h
  loc_0041C3F4: push 000000A4h
  loc_0041C3F9: push 0040F02Ch
  loc_0041C3FE: push edi
  loc_0041C3FF: push eax
  loc_0041C400: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C406: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041C40C: lea ecx, var_18
  loc_0041C40F: call edi
  loc_0041C411: mov edx, [esi]
  loc_0041C413: push esi
  loc_0041C414: call [edx+00000310h]
  loc_0041C41A: push eax
  loc_0041C41B: lea eax, var_18
  loc_0041C41E: push eax
  loc_0041C41F: call ebx
  loc_0041C421: mov esi, eax
  loc_0041C423: push esi
  loc_0041C424: mov ecx, [esi]
  loc_0041C426: call [ecx+00000204h]
  loc_0041C42C: test eax, eax
  loc_0041C42E: fnclex
  loc_0041C430: jge 0041C444h
  loc_0041C432: push 00000204h
  loc_0041C437: push 0040F02Ch
  loc_0041C43C: push esi
  loc_0041C43D: push eax
  loc_0041C43E: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C444: lea ecx, var_18
  loc_0041C447: call edi
  loc_0041C449: mov var_4, 00000000h
  loc_0041C450: push 0041C462h
  loc_0041C455: jmp 0041C461h
  loc_0041C457: lea ecx, var_18
  loc_0041C45A: call [00401114h] ; __vbaFreeObj
  loc_0041C460: ret
  loc_0041C461: ret
  loc_0041C462: mov eax, Me
  loc_0041C465: push eax
  loc_0041C466: mov edx, [eax]
  loc_0041C468: call [edx+00000008h]
  loc_0041C46B: mov eax, var_4
  loc_0041C46E: mov ecx, var_14
  loc_0041C471: pop edi
  loc_0041C472: pop esi
  loc_0041C473: mov fs:[00000000h], ecx
  loc_0041C47A: pop ebx
  loc_0041C47B: mov esp, ebp
  loc_0041C47D: pop ebp
  loc_0041C47E: retn 0004h
End Sub

Private Sub cmdChange_Click() '41C5F0
  loc_0041C5F0: push ebp
  loc_0041C5F1: mov ebp, esp
  loc_0041C5F3: sub esp, 0000000Ch
  loc_0041C5F6: push 00401926h ; __vbaExceptHandler
  loc_0041C5FB: mov eax, fs:[00000000h]
  loc_0041C601: push eax
  loc_0041C602: mov fs:[00000000h], esp
  loc_0041C609: sub esp, 000000D0h
  loc_0041C60F: push ebx
  loc_0041C610: push esi
  loc_0041C611: push edi
  loc_0041C612: mov var_C, esp
  loc_0041C615: mov var_8, 00401208h
  loc_0041C61C: mov esi, Me
  loc_0041C61F: mov eax, esi
  loc_0041C621: and eax, 00000001h
  loc_0041C624: mov var_4, eax
  loc_0041C627: and esi, FFFFFFFEh
  loc_0041C62A: push esi
  loc_0041C62B: mov Me, esi
  loc_0041C62E: mov ecx, [esi]
  loc_0041C630: call [ecx+00000004h]
  loc_0041C633: mov edx, [esi]
  loc_0041C635: xor ebx, ebx
  loc_0041C637: push esi
  loc_0041C638: mov var_1C, ebx
  loc_0041C63B: mov var_20, ebx
  loc_0041C63E: mov var_24, ebx
  loc_0041C641: mov var_28, ebx
  loc_0041C644: mov var_2C, ebx
  loc_0041C647: mov var_30, ebx
  loc_0041C64A: mov var_34, ebx
  loc_0041C64D: mov var_44, ebx
  loc_0041C650: mov var_54, ebx
  loc_0041C653: mov var_64, ebx
  loc_0041C656: mov var_74, ebx
  loc_0041C659: mov var_84, ebx
  loc_0041C65F: mov var_94, ebx
  loc_0041C665: call [edx+00000310h]
  loc_0041C66B: mov var_3C, eax
  loc_0041C66E: lea eax, var_44
  loc_0041C671: lea ecx, var_54
  loc_0041C674: push eax
  loc_0041C675: push ecx
  loc_0041C676: mov var_44, 00000009h
  loc_0041C67D: call [00401050h] ; rtcTrimVar
  loc_0041C683: lea edx, var_54
  loc_0041C686: lea eax, var_84
  loc_0041C68C: push edx
  loc_0041C68D: push eax
  loc_0041C68E: mov var_7C, 0040F040h
  loc_0041C695: mov var_84, 00008008h
  loc_0041C69F: call [00401070h] ; __vbaVarTstEq
  loc_0041C6A5: mov edi, [00401018h] ; __vbaFreeVarList
  loc_0041C6AB: lea ecx, var_54
  loc_0041C6AE: lea edx, var_44
  loc_0041C6B1: push ecx
  loc_0041C6B2: push edx
  loc_0041C6B3: push 00000002h
  loc_0041C6B5: mov var_B8, ax
  loc_0041C6BC: call edi
  loc_0041C6BE: add esp, 0000000Ch
  loc_0041C6C1: cmp var_B8, bx
  loc_0041C6C8: jz 0041C755h
  loc_0041C6CE: mov eax, [esi]
  loc_0041C6D0: push esi
  loc_0041C6D1: call [eax+00000310h]
  loc_0041C6D7: lea ecx, var_34
  loc_0041C6DA: push eax
  loc_0041C6DB: push ecx
  loc_0041C6DC: call [0040103Ch] ; __vbaObjSet
  loc_0041C6E2: mov esi, eax
  loc_0041C6E4: push esi
  loc_0041C6E5: mov edx, [esi]
  loc_0041C6E7: call [edx+00000204h]
  loc_0041C6ED: cmp eax, ebx
  loc_0041C6EF: fnclex
  loc_0041C6F1: jge 0041C705h
  loc_0041C6F3: push 00000204h
  loc_0041C6F8: push 0040F02Ch
  loc_0041C6FD: push esi
  loc_0041C6FE: push eax
  loc_0041C6FF: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C705: lea ecx, var_34
  loc_0041C708: call [00401114h] ; __vbaFreeObj
  loc_0041C70E: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041C714: mov ecx, 80020004h
  loc_0041C719: mov var_6C, ecx
  loc_0041C71C: mov eax, 0000000Ah
  loc_0041C721: mov var_5C, ecx
  loc_0041C724: lea edx, var_94
  loc_0041C72A: lea ecx, var_54
  loc_0041C72D: mov var_74, eax
  loc_0041C730: mov var_64, eax
  loc_0041C733: mov var_8C, 0040F100h ; "Check Name"
  loc_0041C73D: mov var_94, 00000008h
  loc_0041C747: call __vbaVarDup
  loc_0041C749: mov var_7C, 0040F3A0h ; "Please enter a value for the estimate of the slope."
  loc_0041C750: jmp 0041CB9Dh
  loc_0041C755: mov edx, [esi]
  loc_0041C757: push esi
  loc_0041C758: call [edx+00000318h]
  loc_0041C75E: mov var_3C, eax
  loc_0041C761: lea eax, var_44
  loc_0041C764: lea ecx, var_54
  loc_0041C767: push eax
  loc_0041C768: push ecx
  loc_0041C769: mov var_44, 00000009h
  loc_0041C770: call [00401050h] ; rtcTrimVar
  loc_0041C776: lea edx, var_54
  loc_0041C779: lea eax, var_84
  loc_0041C77F: push edx
  loc_0041C780: push eax
  loc_0041C781: mov var_7C, 0040F040h
  loc_0041C788: mov var_84, 00008008h
  loc_0041C792: call [00401070h] ; __vbaVarTstEq
  loc_0041C798: lea ecx, var_54
  loc_0041C79B: lea edx, var_44
  loc_0041C79E: push ecx
  loc_0041C79F: push edx
  loc_0041C7A0: push 00000002h
  loc_0041C7A2: mov var_B8, ax
  loc_0041C7A9: call edi
  loc_0041C7AB: add esp, 0000000Ch
  loc_0041C7AE: cmp var_B8, bx
  loc_0041C7B5: jz 0041C887h
  loc_0041C7BB: mov ecx, 80020004h
  loc_0041C7C0: mov eax, 0000000Ah
  loc_0041C7C5: mov var_6C, ecx
  loc_0041C7C8: mov var_5C, ecx
  loc_0041C7CB: lea edx, var_94
  loc_0041C7D1: lea ecx, var_54
  loc_0041C7D4: mov var_74, eax
  loc_0041C7D7: mov var_64, eax
  loc_0041C7DA: mov var_8C, 0040F100h ; "Check Name"
  loc_0041C7E4: mov var_94, 00000008h
  loc_0041C7EE: call [004010F4h] ; __vbaVarDup
  loc_0041C7F4: lea edx, var_84
  loc_0041C7FA: lea ecx, var_44
  loc_0041C7FD: mov var_7C, 0040F40Ch ; "Please enter a value for the estimate of the standard error."
  loc_0041C804: mov var_84, 00000008h
  loc_0041C80E: call [004010F4h] ; __vbaVarDup
  loc_0041C814: lea eax, var_74
  loc_0041C817: lea ecx, var_64
  loc_0041C81A: push eax
  loc_0041C81B: lea edx, var_54
  loc_0041C81E: push ecx
  loc_0041C81F: push edx
  loc_0041C820: lea eax, var_44
  loc_0041C823: push ebx
  loc_0041C824: push eax
  loc_0041C825: call [00401038h] ; rtcMsgBox
  loc_0041C82B: lea ecx, var_74
  loc_0041C82E: lea edx, var_64
  loc_0041C831: push ecx
  loc_0041C832: lea eax, var_54
  loc_0041C835: push edx
  loc_0041C836: lea ecx, var_44
  loc_0041C839: push eax
  loc_0041C83A: push ecx
  loc_0041C83B: push 00000004h
  loc_0041C83D: call edi
  loc_0041C83F: mov edx, [esi]
  loc_0041C841: add esp, 00000014h
  loc_0041C844: push esi
  loc_0041C845: call [edx+00000318h]
  loc_0041C84B: push eax
  loc_0041C84C: lea eax, var_34
  loc_0041C84F: push eax
  loc_0041C850: call [0040103Ch] ; __vbaObjSet
  loc_0041C856: mov esi, eax
  loc_0041C858: push esi
  loc_0041C859: mov ecx, [esi]
  loc_0041C85B: call [ecx+00000204h]
  loc_0041C861: cmp eax, ebx
  loc_0041C863: fnclex
  loc_0041C865: jge 0041C879h
  loc_0041C867: push 00000204h
  loc_0041C86C: push 0040F02Ch
  loc_0041C871: push esi
  loc_0041C872: push eax
  loc_0041C873: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C879: lea ecx, var_34
  loc_0041C87C: call [00401114h] ; __vbaFreeObj
  loc_0041C882: jmp 0041CFD4h
  loc_0041C887: mov edx, [esi]
  loc_0041C889: push esi
  loc_0041C88A: call [edx+00000310h]
  loc_0041C890: mov var_3C, eax
  loc_0041C893: lea eax, var_44
  loc_0041C896: lea ecx, var_54
  loc_0041C899: push eax
  loc_0041C89A: push ecx
  loc_0041C89B: mov var_44, 00000009h
  loc_0041C8A2: call [00401050h] ; rtcTrimVar
  loc_0041C8A8: lea edx, var_54
  loc_0041C8AB: lea eax, var_1C
  loc_0041C8AE: push edx
  loc_0041C8AF: push eax
  loc_0041C8B0: call [004010B8h] ; __vbaStrVarVal
  loc_0041C8B6: push eax
  loc_0041C8B7: call [00401118h] ; rtcR8ValFromBstr
  loc_0041C8BD: call [00401054h] ; __vbaFpR8
  loc_0041C8C3: fcomp real8 ptr [00401200h]
  loc_0041C8C9: mov var_D0, 00000001h
  loc_0041C8D3: fnstsw ax
  loc_0041C8D5: test ah, 01h
  loc_0041C8D8: jnz 0041C8E0h
  loc_0041C8DA: mov var_D0, ebx
  loc_0041C8E0: lea ecx, var_1C
  loc_0041C8E3: call [00401110h] ; __vbaFreeStr
  loc_0041C8E9: lea ecx, var_54
  loc_0041C8EC: lea edx, var_44
  loc_0041C8EF: push ecx
  loc_0041C8F0: push edx
  loc_0041C8F1: push 00000002h
  loc_0041C8F3: call edi
  loc_0041C8F5: mov eax, var_D0
  loc_0041C8FB: add esp, 0000000Ch
  loc_0041C8FE: neg eax
  loc_0041C900: test ax, ax
  loc_0041C903: jz 0041C990h
  loc_0041C909: mov eax, [esi]
  loc_0041C90B: push esi
  loc_0041C90C: call [eax+00000310h]
  loc_0041C912: lea ecx, var_34
  loc_0041C915: push eax
  loc_0041C916: push ecx
  loc_0041C917: call [0040103Ch] ; __vbaObjSet
  loc_0041C91D: mov esi, eax
  loc_0041C91F: push esi
  loc_0041C920: mov edx, [esi]
  loc_0041C922: call [edx+00000204h]
  loc_0041C928: cmp eax, ebx
  loc_0041C92A: fnclex
  loc_0041C92C: jge 0041C940h
  loc_0041C92E: push 00000204h
  loc_0041C933: push 0040F02Ch
  loc_0041C938: push esi
  loc_0041C939: push eax
  loc_0041C93A: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C940: lea ecx, var_34
  loc_0041C943: call [00401114h] ; __vbaFreeObj
  loc_0041C949: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041C94F: mov ecx, 80020004h
  loc_0041C954: mov var_6C, ecx
  loc_0041C957: mov eax, 0000000Ah
  loc_0041C95C: mov var_5C, ecx
  loc_0041C95F: lea edx, var_94
  loc_0041C965: lea ecx, var_54
  loc_0041C968: mov var_74, eax
  loc_0041C96B: mov var_64, eax
  loc_0041C96E: mov var_8C, 0040F50Ch ; "Check Corr"
  loc_0041C978: mov var_94, 00000008h
  loc_0041C982: call __vbaVarDup
  loc_0041C984: mov var_7C, 0040F49Ch ; "The correlation coefficient can not be less than -1."
  loc_0041C98B: jmp 0041CB9Dh
  loc_0041C990: mov edx, [esi]
  loc_0041C992: push esi
  loc_0041C993: call [edx+00000310h]
  loc_0041C999: mov var_3C, eax
  loc_0041C99C: lea eax, var_44
  loc_0041C99F: lea ecx, var_54
  loc_0041C9A2: push eax
  loc_0041C9A3: push ecx
  loc_0041C9A4: mov var_44, 00000009h
  loc_0041C9AB: call [00401050h] ; rtcTrimVar
  loc_0041C9B1: lea edx, var_54
  loc_0041C9B4: lea eax, var_1C
  loc_0041C9B7: push edx
  loc_0041C9B8: push eax
  loc_0041C9B9: call [004010B8h] ; __vbaStrVarVal
  loc_0041C9BF: push eax
  loc_0041C9C0: call [00401118h] ; rtcR8ValFromBstr
  loc_0041C9C6: call [00401054h] ; __vbaFpR8
  loc_0041C9CC: fcomp real8 ptr [004011F8h]
  loc_0041C9D2: mov var_D4, 00000001h
  loc_0041C9DC: fnstsw ax
  loc_0041C9DE: test ah, 41h
  loc_0041C9E1: jz 0041C9E9h
  loc_0041C9E3: mov var_D4, ebx
  loc_0041C9E9: lea ecx, var_1C
  loc_0041C9EC: call [00401110h] ; __vbaFreeStr
  loc_0041C9F2: lea ecx, var_54
  loc_0041C9F5: lea edx, var_44
  loc_0041C9F8: push ecx
  loc_0041C9F9: push edx
  loc_0041C9FA: push 00000002h
  loc_0041C9FC: call edi
  loc_0041C9FE: mov eax, var_D4
  loc_0041CA04: add esp, 0000000Ch
  loc_0041CA07: neg eax
  loc_0041CA09: test ax, ax
  loc_0041CA0C: jz 0041CA99h
  loc_0041CA12: mov eax, [esi]
  loc_0041CA14: push esi
  loc_0041CA15: call [eax+00000310h]
  loc_0041CA1B: lea ecx, var_34
  loc_0041CA1E: push eax
  loc_0041CA1F: push ecx
  loc_0041CA20: call [0040103Ch] ; __vbaObjSet
  loc_0041CA26: mov esi, eax
  loc_0041CA28: push esi
  loc_0041CA29: mov edx, [esi]
  loc_0041CA2B: call [edx+00000204h]
  loc_0041CA31: cmp eax, ebx
  loc_0041CA33: fnclex
  loc_0041CA35: jge 0041CA49h
  loc_0041CA37: push 00000204h
  loc_0041CA3C: push 0040F02Ch
  loc_0041CA41: push esi
  loc_0041CA42: push eax
  loc_0041CA43: call [00401030h] ; __vbaHresultCheckObj
  loc_0041CA49: lea ecx, var_34
  loc_0041CA4C: call [00401114h] ; __vbaFreeObj
  loc_0041CA52: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041CA58: mov ecx, 80020004h
  loc_0041CA5D: mov var_6C, ecx
  loc_0041CA60: mov eax, 0000000Ah
  loc_0041CA65: mov var_5C, ecx
  loc_0041CA68: lea edx, var_94
  loc_0041CA6E: lea ecx, var_54
  loc_0041CA71: mov var_74, eax
  loc_0041CA74: mov var_64, eax
  loc_0041CA77: mov var_8C, 0040F59Ch ; "Check Correlation"
  loc_0041CA81: mov var_94, 00000008h
  loc_0041CA8B: call __vbaVarDup
  loc_0041CA8D: mov var_7C, 0040F528h ; "The correlation coefficient can not be greater than 1."
  loc_0041CA94: jmp 0041CB9Dh
  loc_0041CA99: mov edx, [esi]
  loc_0041CA9B: push esi
  loc_0041CA9C: call [edx+00000318h]
  loc_0041CAA2: mov var_3C, eax
  loc_0041CAA5: lea eax, var_44
  loc_0041CAA8: lea ecx, var_54
  loc_0041CAAB: push eax
  loc_0041CAAC: push ecx
  loc_0041CAAD: mov var_44, 00000009h
  loc_0041CAB4: call [00401050h] ; rtcTrimVar
  loc_0041CABA: lea edx, var_54
  loc_0041CABD: lea eax, var_1C
  loc_0041CAC0: push edx
  loc_0041CAC1: push eax
  loc_0041CAC2: call [004010B8h] ; __vbaStrVarVal
  loc_0041CAC8: push eax
  loc_0041CAC9: call [00401118h] ; rtcR8ValFromBstr
  loc_0041CACF: call [00401054h] ; __vbaFpR8
  loc_0041CAD5: fcomp real8 ptr [004011F8h]
  loc_0041CADB: mov var_D8, 00000001h
  loc_0041CAE5: fnstsw ax
  loc_0041CAE7: test ah, 41h
  loc_0041CAEA: jz 0041CAF2h
  loc_0041CAEC: mov var_D8, ebx
  loc_0041CAF2: lea ecx, var_1C
  loc_0041CAF5: call [00401110h] ; __vbaFreeStr
  loc_0041CAFB: lea ecx, var_54
  loc_0041CAFE: lea edx, var_44
  loc_0041CB01: push ecx
  loc_0041CB02: push edx
  loc_0041CB03: push 00000002h
  loc_0041CB05: call edi
  loc_0041CB07: mov eax, var_D8
  loc_0041CB0D: add esp, 0000000Ch
  loc_0041CB10: neg eax
  loc_0041CB12: test ax, ax
  loc_0041CB15: jz 0041CBE5h
  loc_0041CB1B: mov eax, [esi]
  loc_0041CB1D: push esi
  loc_0041CB1E: call [eax+00000310h]
  loc_0041CB24: lea ecx, var_34
  loc_0041CB27: push eax
  loc_0041CB28: push ecx
  loc_0041CB29: call [0040103Ch] ; __vbaObjSet
  loc_0041CB2F: mov esi, eax
  loc_0041CB31: push esi
  loc_0041CB32: mov edx, [esi]
  loc_0041CB34: call [edx+00000204h]
  loc_0041CB3A: cmp eax, ebx
  loc_0041CB3C: fnclex
  loc_0041CB3E: jge 0041CB52h
  loc_0041CB40: push 00000204h
  loc_0041CB45: push 0040F02Ch
  loc_0041CB4A: push esi
  loc_0041CB4B: push eax
  loc_0041CB4C: call [00401030h] ; __vbaHresultCheckObj
  loc_0041CB52: lea ecx, var_34
  loc_0041CB55: call [00401114h] ; __vbaFreeObj
  loc_0041CB5B: mov esi, [004010F4h] ; __vbaVarDup
  loc_0041CB61: mov ecx, 80020004h
  loc_0041CB66: mov var_6C, ecx
  loc_0041CB69: mov eax, 0000000Ah
  loc_0041CB6E: mov var_5C, ecx
  loc_0041CB71: lea edx, var_94
  loc_0041CB77: lea ecx, var_54
  loc_0041CB7A: mov var_74, eax
  loc_0041CB7D: mov var_64, eax
  loc_0041CB80: mov var_8C, 0040F640h ; "Check R-squared"
  loc_0041CB8A: mov var_94, 00000008h
  loc_0041CB94: call __vbaVarDup
  loc_0041CB96: mov var_7C, 0040F5C4h ; "The coefficient of determination can not be greater than 1."
  loc_0041CB9D: lea edx, var_84
  loc_0041CBA3: lea ecx, var_44
  loc_0041CBA6: mov var_84, 00000008h
  loc_0041CBB0: call __vbaVarDup
  loc_0041CBB2: lea eax, var_74
  loc_0041CBB5: lea ecx, var_64
  loc_0041CBB8: push eax
  loc_0041CBB9: lea edx, var_54
  loc_0041CBBC: push ecx
  loc_0041CBBD: push edx
  loc_0041CBBE: lea eax, var_44
  loc_0041CBC1: push ebx
  loc_0041CBC2: push eax
  loc_0041CBC3: call [00401038h] ; rtcMsgBox
  loc_0041CBC9: lea ecx, var_74
  loc_0041CBCC: lea edx, var_64
  loc_0041CBCF: push ecx
  loc_0041CBD0: lea eax, var_54
  loc_0041CBD3: push edx
  loc_0041CBD4: lea ecx, var_44
  loc_0041CBD7: push eax
  loc_0041CBD8: push ecx
  loc_0041CBD9: push 00000004h
  loc_0041CBDB: call edi
  loc_0041CBDD: add esp, 00000014h
  loc_0041CBE0: jmp 0041CFD4h
  loc_0041CBE5: mov edx, [esi]
  loc_0041CBE7: push esi
  loc_0041CBE8: call [edx+00000310h]
  loc_0041CBEE: mov var_3C, eax
  loc_0041CBF1: lea eax, var_44
  loc_0041CBF4: lea ecx, var_54
  loc_0041CBF7: push eax
  loc_0041CBF8: push ecx
  loc_0041CBF9: mov var_44, 00000009h
  loc_0041CC00: call [00401050h] ; rtcTrimVar
  loc_0041CC06: lea edx, var_54
  loc_0041CC09: lea eax, var_1C
  loc_0041CC0C: push edx
  loc_0041CC0D: push eax
  loc_0041CC0E: call [004010B8h] ; __vbaStrVarVal
  loc_0041CC14: push eax
  loc_0041CC15: call [00401118h] ; rtcR8ValFromBstr
  loc_0041CC1B: call [00401054h] ; __vbaFpR8
  loc_0041CC21: fcomp real8 ptr [004011F0h]
  loc_0041CC27: mov var_DC, 00000001h
  loc_0041CC31: fnstsw ax
  loc_0041CC33: test ah, 41h
  loc_0041CC36: jz 0041CC3Eh
  loc_0041CC38: mov var_DC, ebx
  loc_0041CC3E: fld real4 ptr [00430060h]
  loc_0041CC44: fcomp real4 ptr [004011E8h]
  loc_0041CC4A: mov var_E0, 00000001h
  loc_0041CC54: fnstsw ax
  loc_0041CC56: test ah, 01h
  loc_0041CC59: jnz 0041CC61h
  loc_0041CC5B: mov var_E0, ebx
  loc_0041CC61: mov ecx, [esi]
  loc_0041CC63: push esi
  loc_0041CC64: call [ecx+00000310h]
  loc_0041CC6A: mov var_5C, eax
  loc_0041CC6D: lea edx, var_64
  loc_0041CC70: lea eax, var_74
  loc_0041CC73: push edx
  loc_0041CC74: push eax
  loc_0041CC75: mov var_64, 00000009h
  loc_0041CC7C: call [00401050h] ; rtcTrimVar
  loc_0041CC82: lea ecx, var_74
  loc_0041CC85: lea edx, var_20
  loc_0041CC88: push ecx
  loc_0041CC89: push edx
  loc_0041CC8A: call [004010B8h] ; __vbaStrVarVal
  loc_0041CC90: push eax
  loc_0041CC91: call [00401118h] ; rtcR8ValFromBstr
  loc_0041CC97: call [00401054h] ; __vbaFpR8
  loc_0041CC9D: fcomp real8 ptr [004011F0h]
  loc_0041CCA3: mov var_E4, 00000001h
  loc_0041CCAD: fnstsw ax
  loc_0041CCAF: test ah, 01h
  loc_0041CCB2: jnz 0041CCBAh
  loc_0041CCB4: mov var_E4, ebx
  loc_0041CCBA: fld real4 ptr [00430060h]
  loc_0041CCC0: fcomp real4 ptr [004011E8h]
  loc_0041CCC6: fnstsw ax
  loc_0041CCC8: test ah, 41h
  loc_0041CCCB: jnz 0041CCD2h
  loc_0041CCCD: mov ebx, 00000001h
  loc_0041CCD2: lea eax, var_20
  loc_0041CCD5: lea ecx, var_1C
  loc_0041CCD8: push eax
  loc_0041CCD9: push ecx
  loc_0041CCDA: push 00000002h
  loc_0041CCDC: call [004010E4h] ; __vbaFreeStrList
  loc_0041CCE2: lea edx, var_74
  loc_0041CCE5: lea eax, var_64
  loc_0041CCE8: push edx
  loc_0041CCE9: lea ecx, var_54
  loc_0041CCEC: push eax
  loc_0041CCED: lea edx, var_44
  loc_0041CCF0: push ecx
  loc_0041CCF1: push edx
  loc_0041CCF2: push 00000004h
  loc_0041CCF4: call edi
  loc_0041CCF6: mov eax, var_E4
  loc_0041CCFC: mov ecx, var_DC
  loc_0041CD02: neg ebx
  loc_0041CD04: neg eax
  loc_0041CD06: and ebx, eax
  loc_0041CD08: mov eax, var_E0
  loc_0041CD0E: neg eax
  loc_0041CD10: neg ecx
  loc_0041CD12: and eax, ecx
  loc_0041CD14: add esp, 00000020h
  loc_0041CD17: or ebx, eax
  loc_0041CD19: test bx, bx
  loc_0041CD1C: jz 0041CE82h
  loc_0041CD22: mov ecx, 80020004h
  loc_0041CD27: mov eax, 0000000Ah
  loc_0041CD2C: mov var_6C, ecx
  loc_0041CD2F: mov var_5C, ecx
  loc_0041CD32: lea edx, var_84
  loc_0041CD38: lea ecx, var_54
  loc_0041CD3B: mov var_74, eax
  loc_0041CD3E: mov var_64, eax
  loc_0041CD41: mov var_7C, 0040F50Ch ; "Check Corr"
  loc_0041CD48: mov var_84, 00000008h
  loc_0041CD52: call [004010F4h] ; __vbaVarDup
  loc_0041CD58: mov edi, [0040102Ch] ; __vbaStrCat
  loc_0041CD5E: push 0040F674h ; "Warning, the correlation coefficient MUST have the same sign as the estimated slope"
  loc_0041CD63: push 0040F720h ; vbCrLf
  loc_0041CD68: call edi
  loc_0041CD6A: mov ebx, [004010FCh] ; __vbaStrMove
  loc_0041CD70: mov edx, eax
  loc_0041CD72: lea ecx, var_1C
  loc_0041CD75: call ebx
  loc_0041CD77: push eax
  loc_0041CD78: push 0040F72Ch ; "Currently they do not. The estimated slope = "
  loc_0041CD7D: call edi
  loc_0041CD7F: mov edx, eax
  loc_0041CD81: lea ecx, var_20
  loc_0041CD84: call ebx
  loc_0041CD86: push eax
  loc_0041CD87: mov eax, [00430060h]
  loc_0041CD8C: push eax
  loc_0041CD8D: call [0040107Ch] ; __vbaStrR4
  loc_0041CD93: mov edx, eax
  loc_0041CD95: lea ecx, var_24
  loc_0041CD98: call ebx
  loc_0041CD9A: push eax
  loc_0041CD9B: call edi
  loc_0041CD9D: mov edx, eax
  loc_0041CD9F: lea ecx, var_28
  loc_0041CDA2: call ebx
  loc_0041CDA4: push eax
  loc_0041CDA5: push 0040F720h ; vbCrLf
  loc_0041CDAA: call edi
  loc_0041CDAC: mov edx, eax
  loc_0041CDAE: lea ecx, var_2C
  loc_0041CDB1: call ebx
  loc_0041CDB3: push eax
  loc_0041CDB4: push 0040F720h ; vbCrLf
  loc_0041CDB9: call edi
  loc_0041CDBB: mov edx, eax
  loc_0041CDBD: lea ecx, var_30
  loc_0041CDC0: call ebx
  loc_0041CDC2: push eax
  loc_0041CDC3: push 0040F78Ch ; "Do you wish to continue with this value?"
  loc_0041CDC8: call edi
  loc_0041CDCA: lea ecx, var_74
  loc_0041CDCD: mov var_3C, eax
  loc_0041CDD0: lea edx, var_64
  loc_0041CDD3: push ecx
  loc_0041CDD4: lea eax, var_54
  loc_0041CDD7: push edx
  loc_0041CDD8: push eax
  loc_0041CDD9: lea ecx, var_44
  loc_0041CDDC: push 00000004h
  loc_0041CDDE: push ecx
  loc_0041CDDF: mov var_44, 00000008h
  loc_0041CDE6: call [00401038h] ; rtcMsgBox
  loc_0041CDEC: mov ecx, eax
  loc_0041CDEE: call [00401078h] ; __vbaI2I4
  loc_0041CDF4: mov edi, eax
  loc_0041CDF6: lea edx, var_30
  loc_0041CDF9: lea eax, var_2C
  loc_0041CDFC: push edx
  loc_0041CDFD: lea ecx, var_28
  loc_0041CE00: push eax
  loc_0041CE01: lea edx, var_24
  loc_0041CE04: push ecx
  loc_0041CE05: lea eax, var_20
  loc_0041CE08: push edx
  loc_0041CE09: lea ecx, var_1C
  loc_0041CE0C: push eax
  loc_0041CE0D: push ecx
  loc_0041CE0E: push 00000006h
  loc_0041CE10: call [004010E4h] ; __vbaFreeStrList
  loc_0041CE16: lea edx, var_74
  loc_0041CE19: lea eax, var_64
  loc_0041CE1C: push edx
  loc_0041CE1D: lea ecx, var_54
  loc_0041CE20: push eax
  loc_0041CE21: lea edx, var_44
  loc_0041CE24: push ecx
  loc_0041CE25: push edx
  loc_0041CE26: push 00000004h
  loc_0041CE28: call [00401018h] ; __vbaFreeVarList
  loc_0041CE2E: add esp, 00000030h
  loc_0041CE31: cmp di, 0007h
  loc_0041CE35: jnz 0041CE7Ch
  loc_0041CE37: mov eax, [esi]
  loc_0041CE39: push esi
  loc_0041CE3A: call [eax+00000310h]
  loc_0041CE40: lea ecx, var_34
  loc_0041CE43: push eax
  loc_0041CE44: push ecx
  loc_0041CE45: call [0040103Ch] ; __vbaObjSet
  loc_0041CE4B: mov esi, eax
  loc_0041CE4D: push esi
  loc_0041CE4E: mov edx, [esi]
  loc_0041CE50: call [edx+00000204h]
  loc_0041CE56: test eax, eax
  loc_0041CE58: fnclex
  loc_0041CE5A: jge 0041CE6Eh
  loc_0041CE5C: push 00000204h
  loc_0041CE61: push 0040F02Ch
  loc_0041CE66: push esi
  loc_0041CE67: push eax
  loc_0041CE68: call [00401030h] ; __vbaHresultCheckObj
  loc_0041CE6E: lea ecx, var_34
  loc_0041CE71: call [00401114h] ; __vbaFreeObj
  loc_0041CE77: jmp 0041CFD2h
  loc_0041CE7C: mov edi, [00401018h] ; __vbaFreeVarList
  loc_0041CE82: mov eax, [esi]
  loc_0041CE84: push esi
  loc_0041CE85: call [eax+00000310h]
  loc_0041CE8B: lea ecx, var_34
  loc_0041CE8E: push eax
  loc_0041CE8F: push ecx
  loc_0041CE90: call [0040103Ch] ; __vbaObjSet
  loc_0041CE96: mov ebx, eax
  loc_0041CE98: lea eax, var_1C
  loc_0041CE9B: push eax
  loc_0041CE9C: push ebx
  loc_0041CE9D: mov edx, [ebx]
  loc_0041CE9F: call [edx+000000A0h]
  loc_0041CEA5: test eax, eax
  loc_0041CEA7: fnclex
  loc_0041CEA9: jge 0041CEBDh
  loc_0041CEAB: push 000000A0h
  loc_0041CEB0: push 0040F02Ch
  loc_0041CEB5: push ebx
  loc_0041CEB6: push eax
  loc_0041CEB7: call [00401030h] ; __vbaHresultCheckObj
  loc_0041CEBD: mov eax, var_1C
  loc_0041CEC0: lea ecx, var_44
  loc_0041CEC3: lea edx, var_54
  loc_0041CEC6: push ecx
  loc_0041CEC7: push edx
  loc_0041CEC8: mov var_1C, 00000000h
  loc_0041CECF: mov var_3C, eax
  loc_0041CED2: mov var_44, 00000008h
  loc_0041CED9: call [00401050h] ; rtcTrimVar
  loc_0041CEDF: lea eax, var_54
  loc_0041CEE2: push eax
  loc_0041CEE3: call [004010A0h] ; __vbaR4ErrVar
  loc_0041CEE9: mov ebx, [00401114h] ; __vbaFreeObj
  loc_0041CEEF: lea ecx, var_34
  loc_0041CEF2: fstp real4 ptr [0043006Ch]
  loc_0041CEF8: call ebx
  loc_0041CEFA: lea ecx, var_54
  loc_0041CEFD: lea edx, var_54
  loc_0041CF00: push ecx
  loc_0041CF01: lea eax, var_44
  loc_0041CF04: push edx
  loc_0041CF05: push eax
  loc_0041CF06: push 00000003h
  loc_0041CF08: call edi
  loc_0041CF0A: mov ecx, [esi]
  loc_0041CF0C: add esp, 00000010h
  loc_0041CF0F: push esi
  loc_0041CF10: call [ecx+00000318h]
  loc_0041CF16: lea edx, var_34
  loc_0041CF19: push eax
  loc_0041CF1A: push edx
  loc_0041CF1B: call [0040103Ch] ; __vbaObjSet
  loc_0041CF21: mov esi, eax
  loc_0041CF23: lea ecx, var_1C
  loc_0041CF26: push ecx
  loc_0041CF27: push esi
  loc_0041CF28: mov eax, [esi]
  loc_0041CF2A: call [eax+000000A0h]
  loc_0041CF30: test eax, eax
  loc_0041CF32: fnclex
  loc_0041CF34: jge 0041CF48h
  loc_0041CF36: push 000000A0h
  loc_0041CF3B: push 0040F02Ch
  loc_0041CF40: push esi
  loc_0041CF41: push eax
  loc_0041CF42: call [00401030h] ; __vbaHresultCheckObj
  loc_0041CF48: mov eax, var_1C
  loc_0041CF4B: lea edx, var_44
  loc_0041CF4E: mov var_3C, eax
  loc_0041CF51: lea eax, var_54
  loc_0041CF54: push edx
  loc_0041CF55: push eax
  loc_0041CF56: mov var_1C, 00000000h
  loc_0041CF5D: mov var_44, 00000008h
  loc_0041CF64: call [00401050h] ; rtcTrimVar
  loc_0041CF6A: lea ecx, var_54
  loc_0041CF6D: push ecx
  loc_0041CF6E: call [004010A0h] ; __vbaR4ErrVar
  loc_0041CF74: fstp real4 ptr [00430070h]
  loc_0041CF7A: lea ecx, var_34
  loc_0041CF7D: call ebx
  loc_0041CF7F: lea edx, var_54
  loc_0041CF82: lea eax, var_54
  loc_0041CF85: push edx
  loc_0041CF86: lea ecx, var_44
  loc_0041CF89: push eax
  loc_0041CF8A: push ecx
  loc_0041CF8B: push 00000003h
  loc_0041CF8D: call edi
  loc_0041CF8F: mov eax, [00430114h]
  loc_0041CF94: add esp, 00000010h
  loc_0041CF97: test eax, eax
  loc_0041CF99: jnz 0041CFABh
  loc_0041CF9B: push 00430114h
  loc_0041CFA0: push 00404514h
  loc_0041CFA5: call [004010D4h] ; __vbaNew2
  loc_0041CFAB: mov esi, [00430114h]
  loc_0041CFB1: push esi
  loc_0041CFB2: mov edx, [esi]
  loc_0041CFB4: call [edx+000002B4h]
  loc_0041CFBA: test eax, eax
  loc_0041CFBC: fnclex
  loc_0041CFBE: jge 0041CFD2h
  loc_0041CFC0: push 000002B4h
  loc_0041CFC5: push 0040F348h
  loc_0041CFCA: push esi
  loc_0041CFCB: push eax
  loc_0041CFCC: call [00401030h] ; __vbaHresultCheckObj
  loc_0041CFD2: xor ebx, ebx
  loc_0041CFD4: mov var_4, ebx
  loc_0041CFD7: fwait
  loc_0041CFD8: push 0041D028h
  loc_0041CFDD: jmp 0041D027h
  loc_0041CFDF: lea eax, var_30
  loc_0041CFE2: lea ecx, var_2C
  loc_0041CFE5: push eax
  loc_0041CFE6: lea edx, var_28
  loc_0041CFE9: push ecx
  loc_0041CFEA: lea eax, var_24
  loc_0041CFED: push edx
  loc_0041CFEE: lea ecx, var_20
  loc_0041CFF1: push eax
  loc_0041CFF2: lea edx, var_1C
  loc_0041CFF5: push ecx
  loc_0041CFF6: push edx
  loc_0041CFF7: push 00000006h
  loc_0041CFF9: call [004010E4h] ; __vbaFreeStrList
  loc_0041CFFF: add esp, 0000001Ch
  loc_0041D002: lea ecx, var_34
  loc_0041D005: call [00401114h] ; __vbaFreeObj
  loc_0041D00B: lea eax, var_74
  loc_0041D00E: lea ecx, var_64
  loc_0041D011: push eax
  loc_0041D012: lea edx, var_54
  loc_0041D015: push ecx
  loc_0041D016: lea eax, var_44
  loc_0041D019: push edx
  loc_0041D01A: push eax
  loc_0041D01B: push 00000004h
  loc_0041D01D: call [00401018h] ; __vbaFreeVarList
  loc_0041D023: add esp, 00000014h
  loc_0041D026: ret
  loc_0041D027: ret
  loc_0041D028: mov eax, Me
  loc_0041D02B: push eax
  loc_0041D02C: mov ecx, [eax]
  loc_0041D02E: call [ecx+00000008h]
  loc_0041D031: mov eax, var_4
  loc_0041D034: mov ecx, var_14
  loc_0041D037: pop edi
  loc_0041D038: pop esi
  loc_0041D039: mov fs:[00000000h], ecx
  loc_0041D040: pop ebx
  loc_0041D041: mov esp, ebp
  loc_0041D043: pop ebp
  loc_0041D044: retn 0004h
End Sub

Private Sub cmdRestore_Click() '41C490
  loc_0041C490: push ebp
  loc_0041C491: mov ebp, esp
  loc_0041C493: sub esp, 0000000Ch
  loc_0041C496: push 00401926h ; __vbaExceptHandler
  loc_0041C49B: mov eax, fs:[00000000h]
  loc_0041C4A1: push eax
  loc_0041C4A2: mov fs:[00000000h], esp
  loc_0041C4A9: sub esp, 00000018h
  loc_0041C4AC: push ebx
  loc_0041C4AD: push esi
  loc_0041C4AE: push edi
  loc_0041C4AF: mov var_C, esp
  loc_0041C4B2: mov var_8, 004011D8h
  loc_0041C4B9: mov esi, Me
  loc_0041C4BC: mov eax, esi
  loc_0041C4BE: and eax, 00000001h
  loc_0041C4C1: mov var_4, eax
  loc_0041C4C4: and esi, FFFFFFFEh
  loc_0041C4C7: push esi
  loc_0041C4C8: mov Me, esi
  loc_0041C4CB: mov ecx, [esi]
  loc_0041C4CD: call [ecx+00000004h]
  loc_0041C4D0: mov [0043006Ch], 3F666666h
  loc_0041C4DA: mov [00430070h], 3F4F5C29h
  loc_0041C4E4: mov edx, [esi]
  loc_0041C4E6: xor eax, eax
  loc_0041C4E8: push esi
  loc_0041C4E9: mov var_18, eax
  loc_0041C4EC: mov var_1C, eax
  loc_0041C4EF: call [edx+00000310h]
  loc_0041C4F5: push eax
  loc_0041C4F6: lea eax, var_1C
  loc_0041C4F9: push eax
  loc_0041C4FA: call [0040103Ch] ; __vbaObjSet
  loc_0041C500: mov ecx, [0043006Ch]
  loc_0041C506: mov edi, eax
  loc_0041C508: push ecx
  loc_0041C509: mov ebx, [edi]
  loc_0041C50B: call [0040107Ch] ; __vbaStrR4
  loc_0041C511: mov edx, eax
  loc_0041C513: lea ecx, var_18
  loc_0041C516: call [004010FCh] ; __vbaStrMove
  loc_0041C51C: push eax
  loc_0041C51D: push edi
  loc_0041C51E: call [ebx+000000A4h]
  loc_0041C524: test eax, eax
  loc_0041C526: fnclex
  loc_0041C528: jge 0041C53Ch
  loc_0041C52A: push 000000A4h
  loc_0041C52F: push 0040F02Ch
  loc_0041C534: push edi
  loc_0041C535: push eax
  loc_0041C536: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C53C: mov edi, [00401110h] ; __vbaFreeStr
  loc_0041C542: lea ecx, var_18
  loc_0041C545: call edi
  loc_0041C547: lea ecx, var_1C
  loc_0041C54A: call [00401114h] ; __vbaFreeObj
  loc_0041C550: mov edx, [esi]
  loc_0041C552: push esi
  loc_0041C553: call [edx+00000318h]
  loc_0041C559: push eax
  loc_0041C55A: lea eax, var_1C
  loc_0041C55D: push eax
  loc_0041C55E: call [0040103Ch] ; __vbaObjSet
  loc_0041C564: mov ecx, [00430070h]
  loc_0041C56A: mov esi, eax
  loc_0041C56C: push ecx
  loc_0041C56D: mov ebx, [esi]
  loc_0041C56F: call [0040107Ch] ; __vbaStrR4
  loc_0041C575: mov edx, eax
  loc_0041C577: lea ecx, var_18
  loc_0041C57A: call [004010FCh] ; __vbaStrMove
  loc_0041C580: push eax
  loc_0041C581: push esi
  loc_0041C582: call [ebx+000000A4h]
  loc_0041C588: test eax, eax
  loc_0041C58A: fnclex
  loc_0041C58C: jge 0041C5A0h
  loc_0041C58E: push 000000A4h
  loc_0041C593: push 0040F02Ch
  loc_0041C598: push esi
  loc_0041C599: push eax
  loc_0041C59A: call [00401030h] ; __vbaHresultCheckObj
  loc_0041C5A0: lea ecx, var_18
  loc_0041C5A3: call edi
  loc_0041C5A5: lea ecx, var_1C
  loc_0041C5A8: call [00401114h] ; __vbaFreeObj
  loc_0041C5AE: mov var_4, 00000000h
  loc_0041C5B5: fwait
  loc_0041C5B6: push 0041C5D1h
  loc_0041C5BB: jmp 0041C5D0h
  loc_0041C5BD: lea ecx, var_18
  loc_0041C5C0: call [00401110h] ; __vbaFreeStr
  loc_0041C5C6: lea ecx, var_1C
  loc_0041C5C9: call [00401114h] ; __vbaFreeObj
  loc_0041C5CF: ret
  loc_0041C5D0: ret
  loc_0041C5D1: mov eax, Me
  loc_0041C5D4: push eax
  loc_0041C5D5: mov edx, [eax]
  loc_0041C5D7: call [edx+00000008h]
  loc_0041C5DA: mov eax, var_4
  loc_0041C5DD: mov ecx, var_14
  loc_0041C5E0: pop edi
  loc_0041C5E1: pop esi
  loc_0041C5E2: mov fs:[00000000h], ecx
  loc_0041C5E9: pop ebx
  loc_0041C5EA: mov esp, ebp
  loc_0041C5EC: pop ebp
  loc_0041C5ED: retn 0004h
End Sub

Private Sub Form_Activate() '41D050
  loc_0041D050: push ebp
  loc_0041D051: mov ebp, esp
  loc_0041D053: sub esp, 0000000Ch
  loc_0041D056: push 00401926h ; __vbaExceptHandler
  loc_0041D05B: mov eax, fs:[00000000h]
  loc_0041D061: push eax
  loc_0041D062: mov fs:[00000000h], esp
  loc_0041D069: sub esp, 00000018h
  loc_0041D06C: push ebx
  loc_0041D06D: push esi
  loc_0041D06E: push edi
  loc_0041D06F: mov var_C, esp
  loc_0041D072: mov var_8, 00401218h
  loc_0041D079: mov esi, Me
  loc_0041D07C: mov eax, esi
  loc_0041D07E: and eax, 00000001h
  loc_0041D081: mov var_4, eax
  loc_0041D084: and esi, FFFFFFFEh
  loc_0041D087: push esi
  loc_0041D088: mov Me, esi
  loc_0041D08B: mov ecx, [esi]
  loc_0041D08D: call [ecx+00000004h]
  loc_0041D090: mov edx, [esi]
  loc_0041D092: xor eax, eax
  loc_0041D094: push esi
  loc_0041D095: mov var_18, eax
  loc_0041D098: mov var_1C, eax
  loc_0041D09B: call [edx+00000310h]
  loc_0041D0A1: push eax
  loc_0041D0A2: lea eax, var_1C
  loc_0041D0A5: push eax
  loc_0041D0A6: call [0040103Ch] ; __vbaObjSet
  loc_0041D0AC: mov ecx, [0043006Ch]
  loc_0041D0B2: mov edi, eax
  loc_0041D0B4: push ecx
  loc_0041D0B5: mov ebx, [edi]
  loc_0041D0B7: call [0040107Ch] ; __vbaStrR4
  loc_0041D0BD: mov edx, eax
  loc_0041D0BF: lea ecx, var_18
  loc_0041D0C2: call [004010FCh] ; __vbaStrMove
  loc_0041D0C8: push eax
  loc_0041D0C9: push edi
  loc_0041D0CA: call [ebx+000000A4h]
  loc_0041D0D0: test eax, eax
  loc_0041D0D2: fnclex
  loc_0041D0D4: jge 0041D0E8h
  loc_0041D0D6: push 000000A4h
  loc_0041D0DB: push 0040F02Ch
  loc_0041D0E0: push edi
  loc_0041D0E1: push eax
  loc_0041D0E2: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D0E8: lea ecx, var_18
  loc_0041D0EB: call [00401110h] ; __vbaFreeStr
  loc_0041D0F1: lea ecx, var_1C
  loc_0041D0F4: call [00401114h] ; __vbaFreeObj
  loc_0041D0FA: mov edx, [esi]
  loc_0041D0FC: push esi
  loc_0041D0FD: call [edx+00000318h]
  loc_0041D103: push eax
  loc_0041D104: lea eax, var_1C
  loc_0041D107: push eax
  loc_0041D108: call [0040103Ch] ; __vbaObjSet
  loc_0041D10E: mov ecx, [00430070h]
  loc_0041D114: mov edi, eax
  loc_0041D116: push ecx
  loc_0041D117: mov ebx, [edi]
  loc_0041D119: call [0040107Ch] ; __vbaStrR4
  loc_0041D11F: mov edx, eax
  loc_0041D121: lea ecx, var_18
  loc_0041D124: call [004010FCh] ; __vbaStrMove
  loc_0041D12A: push eax
  loc_0041D12B: push edi
  loc_0041D12C: call [ebx+000000A4h]
  loc_0041D132: test eax, eax
  loc_0041D134: fnclex
  loc_0041D136: jge 0041D14Ah
  loc_0041D138: push 000000A4h
  loc_0041D13D: push 0040F02Ch
  loc_0041D142: push edi
  loc_0041D143: push eax
  loc_0041D144: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D14A: lea ecx, var_18
  loc_0041D14D: call [00401110h] ; __vbaFreeStr
  loc_0041D153: mov edi, [00401114h] ; __vbaFreeObj
  loc_0041D159: lea ecx, var_1C
  loc_0041D15C: call edi
  loc_0041D15E: mov edx, [esi]
  loc_0041D160: push esi
  loc_0041D161: call [edx+00000310h]
  loc_0041D167: push eax
  loc_0041D168: lea eax, var_1C
  loc_0041D16B: push eax
  loc_0041D16C: call [0040103Ch] ; __vbaObjSet
  loc_0041D172: mov esi, eax
  loc_0041D174: push esi
  loc_0041D175: mov ecx, [esi]
  loc_0041D177: call [ecx+00000204h]
  loc_0041D17D: test eax, eax
  loc_0041D17F: fnclex
  loc_0041D181: jge 0041D195h
  loc_0041D183: push 00000204h
  loc_0041D188: push 0040F02Ch
  loc_0041D18D: push esi
  loc_0041D18E: push eax
  loc_0041D18F: call [00401030h] ; __vbaHresultCheckObj
  loc_0041D195: lea ecx, var_1C
  loc_0041D198: call edi
  loc_0041D19A: mov var_4, 00000000h
  loc_0041D1A1: fwait
  loc_0041D1A2: push 0041D1BDh
  loc_0041D1A7: jmp 0041D1BCh
  loc_0041D1A9: lea ecx, var_18
  loc_0041D1AC: call [00401110h] ; __vbaFreeStr
  loc_0041D1B2: lea ecx, var_1C
  loc_0041D1B5: call [00401114h] ; __vbaFreeObj
  loc_0041D1BB: ret
  loc_0041D1BC: ret
  loc_0041D1BD: mov eax, Me
  loc_0041D1C0: push eax
  loc_0041D1C1: mov edx, [eax]
  loc_0041D1C3: call [edx+00000008h]
  loc_0041D1C6: mov eax, var_4
  loc_0041D1C9: mov ecx, var_14
  loc_0041D1CC: pop edi
  loc_0041D1CD: pop esi
  loc_0041D1CE: mov fs:[00000000h], ecx
  loc_0041D1D5: pop ebx
  loc_0041D1D6: mov esp, ebp
  loc_0041D1D8: pop ebp
  loc_0041D1D9: retn 0004h
End Sub
