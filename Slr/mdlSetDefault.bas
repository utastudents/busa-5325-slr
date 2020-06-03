

Private Sub Proc_0_0_41A480(arg_C) '41A480
  loc_0041A480: push ebp
  loc_0041A481: mov ebp, esp
  loc_0041A483: sub esp, 00000008h
  loc_0041A486: push 00401926h ; __vbaExceptHandler
  loc_0041A48B: mov eax, fs:[00000000h]
  loc_0041A491: push eax
  loc_0041A492: mov fs:[00000000h], esp
  loc_0041A499: sub esp, 00000028h
  loc_0041A49C: push ebx
  loc_0041A49D: push esi
  loc_0041A49E: push edi
  loc_0041A49F: mov var_8, esp
  loc_0041A4A2: mov var_4, 00401120h
  loc_0041A4A9: mov edi, [00401024h] ; rtcAnsiValueBstr
  loc_0041A4AF: xor eax, eax
  loc_0041A4B1: mov var_20, eax
  loc_0041A4B4: mov var_30, eax
  loc_0041A4B7: mov eax, arg_C
  loc_0041A4BA: push 0040DD34h
  loc_0041A4BF: mov si, [eax]
  loc_0041A4C2: call edi
  loc_0041A4C4: xor ebx, ebx
  loc_0041A4C6: cmp si, ax
  loc_0041A4C9: push 0040DD2Ch
  loc_0041A4CE: setg bl
  loc_0041A4D1: call edi
  loc_0041A4D3: xor ecx, ecx
  loc_0041A4D5: cmp si, ax
  loc_0041A4D8: setl cl
  loc_0041A4DB: or ebx, ecx
  loc_0041A4DD: jz 0041A505h
  loc_0041A4DF: cmp si, 0008h
  loc_0041A4E3: jz 0041A505h
  loc_0041A4E5: push 0040DD3Ch ; "."
  loc_0041A4EA: call edi
  loc_0041A4EC: cmp si, ax
  loc_0041A4EF: jz 0041A505h
  loc_0041A4F1: push 0040DD44h ; "-"
  loc_0041A4F6: call edi
  loc_0041A4F8: cmp si, ax
  loc_0041A4FB: jz 0041A505h
  loc_0041A4FD: mov edx, arg_C
  loc_0041A500: mov [edx], 0000h
  loc_0041A505: mov eax, arg_C
  loc_0041A508: lea edx, var_30
  loc_0041A50B: mov var_30, 00000002h
  loc_0041A512: mov cx, [eax]
  loc_0041A515: mov var_28, cx
  loc_0041A519: lea ecx, var_20
  loc_0041A51C: call [0040100Ch] ; __vbaVarMove
  loc_0041A522: push 0041A534h
  loc_0041A527: jmp 0041A533h
  loc_0041A529: lea ecx, var_20
  loc_0041A52C: call [00401010h] ; __vbaFreeVar
  loc_0041A532: ret
  loc_0041A533: ret
  loc_0041A534: mov eax, Me
  loc_0041A537: mov ecx, var_20
  loc_0041A53A: mov edx, eax
  loc_0041A53C: pop edi
  loc_0041A53D: pop esi
  loc_0041A53E: pop ebx
  loc_0041A53F: mov [edx], ecx
  loc_0041A541: mov ecx, var_1C
  loc_0041A544: mov [edx+00000004h], ecx
  loc_0041A547: mov ecx, var_18
  loc_0041A54A: mov [edx+00000008h], ecx
  loc_0041A54D: mov ecx, var_14
  loc_0041A550: mov [edx+0000000Ch], ecx
  loc_0041A553: mov ecx, var_10
  loc_0041A556: mov fs:[00000000h], ecx
  loc_0041A55D: mov esp, ebp
  loc_0041A55F: pop ebp
  loc_0041A560: retn 0008h
End Sub
