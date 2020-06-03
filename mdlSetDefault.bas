

Private Sub Proc_0_0_408F40(arg_C) '408F40
  loc_00408F40: push ebp
  loc_00408F41: mov ebp, esp
  loc_00408F43: sub esp, 00000008h
  loc_00408F46: push 00401546h ; __vbaExceptHandler
  loc_00408F4B: mov eax, fs:[00000000h]
  loc_00408F51: push eax
  loc_00408F52: mov fs:[00000000h], esp
  loc_00408F59: sub esp, 00000028h
  loc_00408F5C: push ebx
  loc_00408F5D: push esi
  loc_00408F5E: push edi
  loc_00408F5F: mov var_8, esp
  loc_00408F62: mov var_4, 00401190h
  loc_00408F69: mov edi, [00401030h] ; rtcAnsiValueBstr
  loc_00408F6F: xor eax, eax
  loc_00408F71: mov var_20, eax
  loc_00408F74: mov var_30, eax
  loc_00408F77: mov eax, arg_C
  loc_00408F7A: push 004035A8h
  loc_00408F7F: mov si, [eax]
  loc_00408F82: call edi
  loc_00408F84: xor ebx, ebx
  loc_00408F86: cmp si, ax
  loc_00408F89: push 004035A0h
  loc_00408F8E: setg bl
  loc_00408F91: call edi
  loc_00408F93: xor ecx, ecx
  loc_00408F95: cmp si, ax
  loc_00408F98: setl cl
  loc_00408F9B: or ebx, ecx
  loc_00408F9D: jz 00408FC5h
  loc_00408F9F: cmp si, 0008h
  loc_00408FA3: jz 00408FC5h
  loc_00408FA5: push 004035B0h ; "."
  loc_00408FAA: call edi
  loc_00408FAC: cmp si, ax
  loc_00408FAF: jz 00408FC5h
  loc_00408FB1: push 004035B8h ; "-"
  loc_00408FB6: call edi
  loc_00408FB8: cmp si, ax
  loc_00408FBB: jz 00408FC5h
  loc_00408FBD: mov edx, arg_C
  loc_00408FC0: mov [edx], 0000h
  loc_00408FC5: mov eax, arg_C
  loc_00408FC8: lea edx, var_30
  loc_00408FCB: mov var_30, 00000002h
  loc_00408FD2: mov cx, [eax]
  loc_00408FD5: mov var_28, cx
  loc_00408FD9: lea ecx, var_20
  loc_00408FDC: call [0040100Ch] ; __vbaVarMove
  loc_00408FE2: push 00408FF4h
  loc_00408FE7: jmp 00408FF3h
  loc_00408FE9: lea ecx, var_20
  loc_00408FEC: call [00401014h] ; __vbaFreeVar
  loc_00408FF2: ret
  loc_00408FF3: ret
  loc_00408FF4: mov eax, Me
  loc_00408FF7: mov ecx, var_20
  loc_00408FFA: mov edx, eax
  loc_00408FFC: pop edi
  loc_00408FFD: pop esi
  loc_00408FFE: pop ebx
  loc_00408FFF: mov [edx], ecx
  loc_00409001: mov ecx, var_1C
  loc_00409004: mov [edx+00000004h], ecx
  loc_00409007: mov ecx, var_18
  loc_0040900A: mov [edx+00000008h], ecx
  loc_0040900D: mov ecx, var_14
  loc_00409010: mov [edx+0000000Ch], ecx
  loc_00409013: mov ecx, var_10
  loc_00409016: mov fs:[00000000h], ecx
  loc_0040901D: mov esp, ebp
  loc_0040901F: pop ebp
  loc_00409020: retn 0008h
End Sub

Private Sub Proc_0_1_409030(arg_C, arg_10, arg_14) '409030
  loc_00409030: push ebp
  loc_00409031: mov ebp, esp
  loc_00409033: sub esp, 00000008h
  loc_00409036: push 00401546h ; __vbaExceptHandler
  loc_0040903B: mov eax, fs:[00000000h]
  loc_00409041: push eax
  loc_00409042: mov fs:[00000000h], esp
  loc_00409049: sub esp, 00000138h
  loc_0040904F: push ebx
  loc_00409050: push esi
  loc_00409051: push edi
  loc_00409052: mov var_8, esp
  loc_00409055: mov var_4, 004011A0h
  loc_0040905C: mov eax, Me
  loc_0040905F: xor ebx, ebx
  loc_00409061: push ebx
  loc_00409062: push 00403664h ; "Cls"
  loc_00409067: mov ecx, [eax]
  loc_00409069: mov var_14, ebx
  loc_0040906C: push ecx
  loc_0040906D: mov var_24, ebx
  loc_00409070: mov var_34, ebx
  loc_00409073: mov var_44, ebx
  loc_00409076: mov var_54, ebx
  loc_00409079: mov var_64, ebx
  loc_0040907C: mov var_74, ebx
  loc_0040907F: mov var_84, ebx
  loc_00409085: mov var_A4, ebx
  loc_0040908B: mov var_C4, ebx
  loc_00409091: mov var_E4, ebx
  loc_00409097: mov var_104, ebx
  loc_0040909D: mov var_11C, ebx
  loc_004090A3: mov var_118, ebx
  loc_004090A9: mov var_124, ebx
  loc_004090AF: mov var_120, ebx
  loc_004090B5: call [00401148h] ; __vbaLateMemCall
  loc_004090BB: mov edi, arg_C
  loc_004090BE: add esp, 0000000Ch
  loc_004090C1: mov edx, [edi]
  loc_004090C3: push edx
  loc_004090C4: push 00000001h
  loc_004090C6: call [004010FCh] ; __vbaUbound
  loc_004090CC: mov ecx, eax
  loc_004090CE: call [004010A4h] ; __vbaI2I4
  loc_004090D4: mov [0041509Ch], ax
  loc_004090DA: mov esi, [004010B8h] ; __vbaRedim
  loc_004090E0: movsx eax, ax
  loc_004090E3: push ebx
  loc_004090E4: push eax
  loc_004090E5: push 00000001h
  loc_004090E7: push 00000005h
  loc_004090E9: push 004150B4h
  loc_004090EE: push 00000008h
  loc_004090F0: push 00000080h
  loc_004090F5: call __vbaRedim
  loc_004090F7: movsx ecx, [0041509Ch]
  loc_004090FE: push ebx
  loc_004090FF: push ecx
  loc_00409100: push 00000001h
  loc_00409102: push 00000005h
  loc_00409104: push 004150B8h
  loc_00409109: push 00000008h
  loc_0040910B: push 00000080h
  loc_00409110: call __vbaRedim
  loc_00409112: mov eax, [edi]
  loc_00409114: add esp, 00000038h
  loc_00409117: cmp eax, ebx
  loc_00409119: jz 0040916Dh
  loc_0040911B: cmp [eax], 0002h
  loc_0040911F: jnz 0040916Dh
  loc_00409121: mov esi, [eax+0000001Ch]
  loc_00409124: mov edx, [eax+00000018h]
  loc_00409127: mov ecx, 00000001h
  loc_0040912C: sub ecx, esi
  loc_0040912E: cmp ecx, edx
  loc_00409130: mov var_12C, ecx
  loc_00409136: jb 0040913Eh
  loc_00409138: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040913E: mov edx, arg_14
  loc_00409141: mov eax, [edi]
  loc_00409143: movsx esi, [edx]
  loc_00409146: mov edx, [eax+00000014h]
  loc_00409149: mov ecx, [eax+00000010h]
  loc_0040914C: sub esi, edx
  loc_0040914E: cmp esi, ecx
  loc_00409150: jb 00409158h
  loc_00409152: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409158: mov eax, [edi]
  loc_0040915A: mov edx, var_12C
  loc_00409160: mov eax, [eax+00000018h]
  loc_00409163: imul eax, esi
  loc_00409166: add eax, edx
  loc_00409168: shl eax, 03h
  loc_0040916B: jmp 00409173h
  loc_0040916D: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409173: mov ecx, [edi]
  loc_00409175: mov edx, [ecx+0000000Ch]
  loc_00409178: mov ecx, [edx+eax]
  loc_0040917B: mov [004150DCh], ecx
  loc_00409181: mov edx, [edx+eax+00000004h]
  loc_00409185: mov [004150E0h], edx
  loc_0040918B: mov eax, [edi]
  loc_0040918D: cmp eax, ebx
  loc_0040918F: jz 004091E3h
  loc_00409191: cmp [eax], 0002h
  loc_00409195: jnz 004091E3h
  loc_00409197: mov esi, [eax+0000001Ch]
  loc_0040919A: mov edx, [eax+00000018h]
  loc_0040919D: mov ecx, 00000001h
  loc_004091A2: sub ecx, esi
  loc_004091A4: cmp ecx, edx
  loc_004091A6: mov var_12C, ecx
  loc_004091AC: jb 004091B4h
  loc_004091AE: call [00401094h] ; __vbaGenerateBoundsError
  loc_004091B4: mov ecx, arg_10
  loc_004091B7: mov eax, [edi]
  loc_004091B9: movsx esi, [ecx]
  loc_004091BC: mov edx, [eax+00000014h]
  loc_004091BF: mov ecx, [eax+00000010h]
  loc_004091C2: sub esi, edx
  loc_004091C4: cmp esi, ecx
  loc_004091C6: jb 004091CEh
  loc_004091C8: call [00401094h] ; __vbaGenerateBoundsError
  loc_004091CE: mov edx, [edi]
  loc_004091D0: mov eax, [edx+00000018h]
  loc_004091D3: mov edx, var_12C
  loc_004091D9: imul eax, esi
  loc_004091DC: add eax, edx
  loc_004091DE: shl eax, 03h
  loc_004091E1: jmp 004091E9h
  loc_004091E3: call [00401094h] ; __vbaGenerateBoundsError
  loc_004091E9: mov ecx, [edi]
  loc_004091EB: mov var_134, 00000001h
  loc_004091F5: mov edx, [ecx+0000000Ch]
  loc_004091F8: mov ecx, [edx+eax]
  loc_004091FB: mov [004150E4h], ecx
  loc_00409201: mov eax, [edx+eax+00000004h]
  loc_00409205: mov edx, [004150DCh]
  loc_0040920B: mov [004150E8h], eax
  loc_00409210: mov [004150ECh], edx
  loc_00409216: mov edx, [004150E0h]
  loc_0040921C: mov [004150F4h], ecx
  loc_00409222: mov [004150F8h], eax
  loc_00409227: mov eax, [0041509Ch]
  loc_0040922C: mov cx, 0002h
  loc_00409230: mov [004150F0h], edx
  loc_00409236: mov var_138, eax
  loc_0040923C: mov [0041509Eh], cx
  loc_00409243: cmp cx, var_138
  loc_0040924A: jg 004095BAh
  loc_00409250: mov eax, [edi]
  loc_00409252: cmp eax, ebx
  loc_00409254: jz 0040929Ah
  loc_00409256: cmp [eax], 0002h
  loc_0040925A: jnz 0040929Ah
  loc_0040925C: mov edx, [eax+0000001Ch]
  loc_0040925F: movsx ebx, cx
  loc_00409262: mov ecx, [eax+00000018h]
  loc_00409265: sub ebx, edx
  loc_00409267: cmp ebx, ecx
  loc_00409269: jb 00409271h
  loc_0040926B: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409271: mov ecx, arg_14
  loc_00409274: mov eax, [edi]
  loc_00409276: movsx esi, [ecx]
  loc_00409279: mov edx, [eax+00000014h]
  loc_0040927C: mov ecx, [eax+00000010h]
  loc_0040927F: sub esi, edx
  loc_00409281: cmp esi, ecx
  loc_00409283: jb 0040928Bh
  loc_00409285: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040928B: mov edx, [edi]
  loc_0040928D: mov eax, [edx+00000018h]
  loc_00409290: imul eax, esi
  loc_00409293: add eax, ebx
  loc_00409295: shl eax, 03h
  loc_00409298: jmp 004092A0h
  loc_0040929A: call [00401094h] ; __vbaGenerateBoundsError
  loc_004092A0: mov ecx, [edi]
  loc_004092A2: fld real8 ptr [004150DCh]
  loc_004092A8: mov edx, [ecx+0000000Ch]
  loc_004092AB: fcomp real8 ptr [edx+eax]
  loc_004092AE: fnstsw ax
  loc_004092B0: test ah, 01h
  loc_004092B3: jz 0040931Fh
  loc_004092B5: test ecx, ecx
  loc_004092B7: jz 00409301h
  loc_004092B9: cmp [ecx], 0002h
  loc_004092BD: jnz 00409301h
  loc_004092BF: movsx ebx, [0041509Eh]
  loc_004092C6: mov edx, [ecx+0000001Ch]
  loc_004092C9: mov eax, [ecx+00000018h]
  loc_004092CC: sub ebx, edx
  loc_004092CE: cmp ebx, eax
  loc_004092D0: jb 004092D8h
  loc_004092D2: call [00401094h] ; __vbaGenerateBoundsError
  loc_004092D8: mov ecx, arg_14
  loc_004092DB: mov eax, [edi]
  loc_004092DD: movsx esi, [ecx]
  loc_004092E0: mov edx, [eax+00000014h]
  loc_004092E3: mov ecx, [eax+00000010h]
  loc_004092E6: sub esi, edx
  loc_004092E8: cmp esi, ecx
  loc_004092EA: jb 004092F2h
  loc_004092EC: call [00401094h] ; __vbaGenerateBoundsError
  loc_004092F2: mov edx, [edi]
  loc_004092F4: mov eax, [edx+00000018h]
  loc_004092F7: imul eax, esi
  loc_004092FA: add eax, ebx
  loc_004092FC: shl eax, 03h
  loc_004092FF: jmp 00409307h
  loc_00409301: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409307: mov ecx, [edi]
  loc_00409309: mov edx, [ecx+0000000Ch]
  loc_0040930C: mov ecx, [edx+eax]
  loc_0040930F: mov [004150DCh], ecx
  loc_00409315: mov edx, [edx+eax+00000004h]
  loc_00409319: mov [004150E0h], edx
  loc_0040931F: mov eax, [edi]
  loc_00409321: test eax, eax
  loc_00409323: jz 0040936Dh
  loc_00409325: cmp [eax], 0002h
  loc_00409329: jnz 0040936Dh
  loc_0040932B: movsx ebx, [0041509Eh]
  loc_00409332: mov edx, [eax+0000001Ch]
  loc_00409335: mov ecx, [eax+00000018h]
  loc_00409338: sub ebx, edx
  loc_0040933A: cmp ebx, ecx
  loc_0040933C: jb 00409344h
  loc_0040933E: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409344: mov ecx, arg_14
  loc_00409347: mov eax, [edi]
  loc_00409349: movsx esi, [ecx]
  loc_0040934C: mov edx, [eax+00000014h]
  loc_0040934F: mov ecx, [eax+00000010h]
  loc_00409352: sub esi, edx
  loc_00409354: cmp esi, ecx
  loc_00409356: jb 0040935Eh
  loc_00409358: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040935E: mov edx, [edi]
  loc_00409360: mov eax, [edx+00000018h]
  loc_00409363: imul eax, esi
  loc_00409366: add eax, ebx
  loc_00409368: shl eax, 03h
  loc_0040936B: jmp 00409373h
  loc_0040936D: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409373: mov ecx, [edi]
  loc_00409375: fld real8 ptr [004150ECh]
  loc_0040937B: mov edx, [ecx+0000000Ch]
  loc_0040937E: fcomp real8 ptr [edx+eax]
  loc_00409381: fnstsw ax
  loc_00409383: test ah, 41h
  loc_00409386: jnz 004093F2h
  loc_00409388: test ecx, ecx
  loc_0040938A: jz 004093D4h
  loc_0040938C: cmp [ecx], 0002h
  loc_00409390: jnz 004093D4h
  loc_00409392: movsx ebx, [0041509Eh]
  loc_00409399: mov edx, [ecx+0000001Ch]
  loc_0040939C: mov eax, [ecx+00000018h]
  loc_0040939F: sub ebx, edx
  loc_004093A1: cmp ebx, eax
  loc_004093A3: jb 004093ABh
  loc_004093A5: call [00401094h] ; __vbaGenerateBoundsError
  loc_004093AB: mov ecx, arg_14
  loc_004093AE: mov eax, [edi]
  loc_004093B0: movsx esi, [ecx]
  loc_004093B3: mov edx, [eax+00000014h]
  loc_004093B6: mov ecx, [eax+00000010h]
  loc_004093B9: sub esi, edx
  loc_004093BB: cmp esi, ecx
  loc_004093BD: jb 004093C5h
  loc_004093BF: call [00401094h] ; __vbaGenerateBoundsError
  loc_004093C5: mov edx, [edi]
  loc_004093C7: mov eax, [edx+00000018h]
  loc_004093CA: imul eax, esi
  loc_004093CD: add eax, ebx
  loc_004093CF: shl eax, 03h
  loc_004093D2: jmp 004093DAh
  loc_004093D4: call [00401094h] ; __vbaGenerateBoundsError
  loc_004093DA: mov ecx, [edi]
  loc_004093DC: mov edx, [ecx+0000000Ch]
  loc_004093DF: mov ecx, [edx+eax]
  loc_004093E2: mov [004150ECh], ecx
  loc_004093E8: mov edx, [edx+eax+00000004h]
  loc_004093EC: mov [004150F0h], edx
  loc_004093F2: mov eax, [edi]
  loc_004093F4: test eax, eax
  loc_004093F6: jz 00409440h
  loc_004093F8: cmp [eax], 0002h
  loc_004093FC: jnz 00409440h
  loc_004093FE: movsx ebx, [0041509Eh]
  loc_00409405: mov edx, [eax+0000001Ch]
  loc_00409408: mov ecx, [eax+00000018h]
  loc_0040940B: sub ebx, edx
  loc_0040940D: cmp ebx, ecx
  loc_0040940F: jb 00409417h
  loc_00409411: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409417: mov ecx, arg_10
  loc_0040941A: mov eax, [edi]
  loc_0040941C: movsx esi, [ecx]
  loc_0040941F: mov edx, [eax+00000014h]
  loc_00409422: mov ecx, [eax+00000010h]
  loc_00409425: sub esi, edx
  loc_00409427: cmp esi, ecx
  loc_00409429: jb 00409431h
  loc_0040942B: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409431: mov edx, [edi]
  loc_00409433: mov eax, [edx+00000018h]
  loc_00409436: imul eax, esi
  loc_00409439: add eax, ebx
  loc_0040943B: shl eax, 03h
  loc_0040943E: jmp 00409446h
  loc_00409440: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409446: mov ecx, [edi]
  loc_00409448: fld real8 ptr [004150E4h]
  loc_0040944E: mov edx, [ecx+0000000Ch]
  loc_00409451: fcomp real8 ptr [edx+eax]
  loc_00409454: fnstsw ax
  loc_00409456: test ah, 01h
  loc_00409459: jz 004094C5h
  loc_0040945B: test ecx, ecx
  loc_0040945D: jz 004094A7h
  loc_0040945F: cmp [ecx], 0002h
  loc_00409463: jnz 004094A7h
  loc_00409465: movsx ebx, [0041509Eh]
  loc_0040946C: mov edx, [ecx+0000001Ch]
  loc_0040946F: mov eax, [ecx+00000018h]
  loc_00409472: sub ebx, edx
  loc_00409474: cmp ebx, eax
  loc_00409476: jb 0040947Eh
  loc_00409478: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040947E: mov ecx, arg_10
  loc_00409481: mov eax, [edi]
  loc_00409483: movsx esi, [ecx]
  loc_00409486: mov edx, [eax+00000014h]
  loc_00409489: mov ecx, [eax+00000010h]
  loc_0040948C: sub esi, edx
  loc_0040948E: cmp esi, ecx
  loc_00409490: jb 00409498h
  loc_00409492: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409498: mov edx, [edi]
  loc_0040949A: mov eax, [edx+00000018h]
  loc_0040949D: imul eax, esi
  loc_004094A0: add eax, ebx
  loc_004094A2: shl eax, 03h
  loc_004094A5: jmp 004094ADh
  loc_004094A7: call [00401094h] ; __vbaGenerateBoundsError
  loc_004094AD: mov ecx, [edi]
  loc_004094AF: mov edx, [ecx+0000000Ch]
  loc_004094B2: mov ecx, [edx+eax]
  loc_004094B5: mov [004150E4h], ecx
  loc_004094BB: mov edx, [edx+eax+00000004h]
  loc_004094BF: mov [004150E8h], edx
  loc_004094C5: mov eax, [edi]
  loc_004094C7: test eax, eax
  loc_004094C9: jz 00409513h
  loc_004094CB: cmp [eax], 0002h
  loc_004094CF: jnz 00409513h
  loc_004094D1: movsx ebx, [0041509Eh]
  loc_004094D8: mov edx, [eax+0000001Ch]
  loc_004094DB: mov ecx, [eax+00000018h]
  loc_004094DE: sub ebx, edx
  loc_004094E0: cmp ebx, ecx
  loc_004094E2: jb 004094EAh
  loc_004094E4: call [00401094h] ; __vbaGenerateBoundsError
  loc_004094EA: mov ecx, arg_10
  loc_004094ED: mov eax, [edi]
  loc_004094EF: movsx esi, [ecx]
  loc_004094F2: mov edx, [eax+00000014h]
  loc_004094F5: mov ecx, [eax+00000010h]
  loc_004094F8: sub esi, edx
  loc_004094FA: cmp esi, ecx
  loc_004094FC: jb 00409504h
  loc_004094FE: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409504: mov edx, [edi]
  loc_00409506: mov eax, [edx+00000018h]
  loc_00409509: imul eax, esi
  loc_0040950C: add eax, ebx
  loc_0040950E: shl eax, 03h
  loc_00409511: jmp 00409519h
  loc_00409513: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409519: mov ecx, [edi]
  loc_0040951B: fld real8 ptr [004150F4h]
  loc_00409521: mov edx, [ecx+0000000Ch]
  loc_00409524: fcomp real8 ptr [edx+eax]
  loc_00409527: fnstsw ax
  loc_00409529: test ah, 41h
  loc_0040952C: jnz 00409598h
  loc_0040952E: test ecx, ecx
  loc_00409530: jz 0040957Ah
  loc_00409532: cmp [ecx], 0002h
  loc_00409536: jnz 0040957Ah
  loc_00409538: movsx ebx, [0041509Eh]
  loc_0040953F: mov edx, [ecx+0000001Ch]
  loc_00409542: mov eax, [ecx+00000018h]
  loc_00409545: sub ebx, edx
  loc_00409547: cmp ebx, eax
  loc_00409549: jb 00409551h
  loc_0040954B: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409551: mov ecx, arg_10
  loc_00409554: mov eax, [edi]
  loc_00409556: movsx esi, [ecx]
  loc_00409559: mov edx, [eax+00000014h]
  loc_0040955C: mov ecx, [eax+00000010h]
  loc_0040955F: sub esi, edx
  loc_00409561: cmp esi, ecx
  loc_00409563: jb 0040956Bh
  loc_00409565: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040956B: mov edx, [edi]
  loc_0040956D: mov eax, [edx+00000018h]
  loc_00409570: imul eax, esi
  loc_00409573: add eax, ebx
  loc_00409575: shl eax, 03h
  loc_00409578: jmp 00409580h
  loc_0040957A: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409580: mov ecx, [edi]
  loc_00409582: mov edx, [ecx+0000000Ch]
  loc_00409585: mov ecx, [edx+eax]
  loc_00409588: mov [004150F4h], ecx
  loc_0040958E: mov edx, [edx+eax+00000004h]
  loc_00409592: mov [004150F8h], edx
  loc_00409598: mov cx, var_134
  loc_0040959F: add cx, [0041509Eh]
  loc_004095A6: jo 00409BCFh
  loc_004095AC: mov [0041509Eh], cx
  loc_004095B3: xor ebx, ebx
  loc_004095B5: jmp 00409243h
  loc_004095BA: fld real8 ptr [004150DCh]
  loc_004095C0: fsub st0, real8 ptr [004150ECh]
  loc_004095C6: mov edx, 00000005h
  loc_004095CB: mov ebx, [0040104Ch] ; __vbaLateMemSt
  loc_004095D1: mov var_44, edx
  loc_004095D4: fstp real8 ptr [004150A4h]
  loc_004095DA: fnstsw ax
  loc_004095DC: test al, 0Dh
  loc_004095DE: jnz 00409BCAh
  loc_004095E4: fld real8 ptr [004150E4h]
  loc_004095EA: fsub st0, real8 ptr [004150F4h]
  loc_004095F0: mov ecx, [004150A8h]
  loc_004095F6: mov var_38, ecx
  loc_004095F9: fstp real8 ptr [004150ACh]
  loc_004095FF: fnstsw ax
  loc_00409601: test al, 0Dh
  loc_00409603: jnz 00409BCAh
  loc_00409609: sub esp, 00000010h
  loc_0040960C: mov eax, [004150A4h]
  loc_00409611: mov esi, esp
  loc_00409613: mov var_3C, eax
  loc_00409616: push 0040366Ch ; "ScaleWidth"
  loc_0040961B: mov [esi], edx
  loc_0040961D: mov edx, var_40
  loc_00409620: mov [esi+00000004h], edx
  loc_00409623: mov [esi+00000008h], eax
  loc_00409626: mov eax, Me
  loc_00409629: mov [esi+0000000Ch], ecx
  loc_0040962C: mov ecx, [eax]
  loc_0040962E: push ecx
  loc_0040962F: call ebx
  loc_00409631: fld real8 ptr [004150A4h]
  loc_00409637: mov esi, [00401158h] ; __vbaFpI2
  loc_0040963D: call __vbaFpI2
  loc_0040963F: mov edx, [004150B0h]
  loc_00409645: sub esp, 00000010h
  loc_00409648: mov var_38, edx
  loc_0040964B: mov edx, esp
  loc_0040964D: mov ecx, 00000005h
  loc_00409652: mov [004150FCh], ax
  loc_00409658: mov eax, [004150ACh]
  loc_0040965D: mov var_44, ecx
  loc_00409660: mov [edx], ecx
  loc_00409662: mov ecx, var_40
  loc_00409665: mov var_3C, eax
  loc_00409668: push 00403684h ; "ScaleHeight"
  loc_0040966D: mov [edx+00000004h], ecx
  loc_00409670: mov ecx, Me
  loc_00409673: mov [edx+00000008h], eax
  loc_00409676: mov eax, var_38
  loc_00409679: mov [edx+0000000Ch], eax
  loc_0040967C: mov edx, [ecx]
  loc_0040967E: push edx
  loc_0040967F: call ebx
  loc_00409681: fld real8 ptr [004150ACh]
  loc_00409687: call __vbaFpI2
  loc_00409689: push 004150ECh
  loc_0040968E: mov [004150FEh], ax
  loc_00409694: mov var_3C, 00000004h
  loc_0040969B: mov var_44, 00000002h
  loc_004096A2: call 00409BE0h
  loc_004096A7: fstp real8 ptr var_5C
  loc_004096AA: xor eax, eax
  loc_004096AC: mov esi, 00000005h
  loc_004096B1: mov var_11C, eax
  loc_004096B7: mov var_118, eax
  loc_004096BD: lea eax, var_11C
  loc_004096C3: push eax
  loc_004096C4: call 00409C30h
  loc_004096C9: mov ecx, [004150DCh]
  loc_004096CF: mov edx, [004150E0h]
  loc_004096D5: fstp real8 ptr var_7C
  loc_004096D8: xor eax, eax
  loc_004096DA: mov ebx, esi
  loc_004096DC: mov var_124, eax
  loc_004096E2: mov var_120, eax
  loc_004096E8: lea eax, var_124
  loc_004096EE: mov var_9C, ecx
  loc_004096F4: push eax
  loc_004096F5: mov var_98, edx
  loc_004096FB: call 00409C30h
  loc_00409700: mov edx, var_44
  loc_00409703: sub esp, 00000010h
  loc_00409706: mov ecx, esp
  loc_00409708: sub esp, 00000010h
  loc_0040970B: fstp real8 ptr var_BC
  loc_00409711: mov [ecx], edx
  loc_00409713: mov edx, var_40
  loc_00409716: mov eax, esi
  loc_00409718: mov [ecx+00000004h], edx
  loc_0040971B: mov edx, var_3C
  loc_0040971E: mov [ecx+00000008h], edx
  loc_00409721: mov edx, var_38
  loc_00409724: mov [ecx+0000000Ch], edx
  loc_00409727: mov edx, var_60
  loc_0040972A: mov ecx, esp
  loc_0040972C: sub esp, 00000010h
  loc_0040972F: mov [ecx], esi
  loc_00409731: mov [ecx+00000004h], edx
  loc_00409734: mov edx, var_5C
  loc_00409737: mov [ecx+00000008h], edx
  loc_0040973A: mov edx, var_58
  loc_0040973D: mov [ecx+0000000Ch], edx
  loc_00409740: mov edx, var_80
  loc_00409743: mov ecx, esp
  loc_00409745: sub esp, 00000010h
  loc_00409748: mov [ecx], ebx
  loc_0040974A: mov [ecx+00000004h], edx
  loc_0040974D: mov edx, var_7C
  loc_00409750: mov [ecx+00000008h], edx
  loc_00409753: mov edx, var_78
  loc_00409756: mov [ecx+0000000Ch], edx
  loc_00409759: mov edx, esp
  loc_0040975B: mov ecx, esi
  loc_0040975D: sub esp, 00000010h
  loc_00409760: mov [edx], ecx
  loc_00409762: mov ecx, var_A0
  loc_00409768: mov [edx+00000004h], ecx
  loc_0040976B: mov ecx, var_9C
  loc_00409771: mov [edx+00000008h], ecx
  loc_00409774: mov ecx, var_98
  loc_0040977A: mov [edx+0000000Ch], ecx
  loc_0040977D: mov ecx, var_BC
  loc_00409783: mov edx, esp
  loc_00409785: sub esp, 00000010h
  loc_00409788: mov [edx], eax
  loc_0040978A: mov eax, var_C0
  loc_00409790: mov [edx+00000004h], eax
  loc_00409793: mov eax, var_B8
  loc_00409799: mov [edx+00000008h], ecx
  loc_0040979C: mov ecx, esp
  loc_0040979E: push 00000006h
  loc_004097A0: push 0040369Ch ; "Line"
  loc_004097A5: mov [edx+0000000Ch], eax
  loc_004097A8: mov edx, var_E0
  loc_004097AE: mov eax, 00000002h
  loc_004097B3: mov [ecx], eax
  loc_004097B5: xor eax, eax
  loc_004097B7: mov [ecx+00000004h], edx
  loc_004097BA: mov [ecx+00000008h], eax
  loc_004097BD: mov eax, var_D8
  loc_004097C3: mov [ecx+0000000Ch], eax
  loc_004097C6: mov ecx, Me
  loc_004097C9: mov edx, [ecx]
  loc_004097CB: push edx
  loc_004097CC: call [00401148h] ; __vbaLateMemCall
  loc_004097D2: mov eax, [0041509Ch]
  loc_004097D7: add esp, 0000006Ch
  loc_004097DA: mov var_140, eax
  loc_004097E0: mov ax, 0001h
  loc_004097E4: mov [0041509Eh], ax
  loc_004097EA: cmp ax, var_140
  loc_004097F1: jg 00409B90h
  loc_004097F7: mov ecx, [edi]
  loc_004097F9: lea edx, var_14
  loc_004097FC: push ecx
  loc_004097FD: push edx
  loc_004097FE: call [0040114Ch] ; __vbaAryLock
  loc_00409804: mov ecx, var_14
  loc_00409807: test ecx, ecx
  loc_00409809: jz 00409855h
  loc_0040980B: cmp [ecx], 0002h
  loc_0040980F: jnz 00409855h
  loc_00409811: movsx esi, [0041509Eh]
  loc_00409818: mov edx, [ecx+0000001Ch]
  loc_0040981B: mov eax, [ecx+00000018h]
  loc_0040981E: sub esi, edx
  loc_00409820: cmp esi, eax
  loc_00409822: jb 0040982Dh
  loc_00409824: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040982A: mov ecx, var_14
  loc_0040982D: mov eax, arg_14
  loc_00409830: mov edx, [ecx+00000014h]
  loc_00409833: movsx ebx, [eax]
  loc_00409836: mov eax, [ecx+00000010h]
  loc_00409839: sub ebx, edx
  loc_0040983B: cmp ebx, eax
  loc_0040983D: jb 00409848h
  loc_0040983F: call [00401094h] ; __vbaGenerateBoundsError
  loc_00409845: mov ecx, var_14
  loc_00409848: mov eax, [ecx+00000018h]
  loc_0040984B: imul eax, ebx
  loc_0040984E: add eax, esi
  loc_00409850: shl eax, 03h
  loc_00409853: jmp 0040985Eh
  loc_00409855: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040985B: mov ecx, var_14
  loc_0040985E: mov ecx, [ecx+0000000Ch]
  loc_00409861: add ecx, eax
  loc_00409863: push ecx
  loc_00409864: call 00409BE0h
  loc_00409869: fstp real8 ptr var_11C
  loc_0040986F: lea edx, var_14
  loc_00409872: push edx
  loc_00409873: call [00401174h] ; __vbaAryUnlock
  loc_00409879: mov ecx, [004150B4h]
  loc_0040987F: test ecx, ecx
  loc_00409881: jz 004098B3h
  loc_00409883: cmp [ecx], 0001h
  loc_00409887: jnz 004098B3h
  loc_00409889: movsx ebx, [0041509Eh]
  loc_00409890: mov edx, [ecx+00000014h]
  loc_00409893: mov eax, [ecx+00000010h]
  loc_00409896: mov esi, [00401094h] ; __vbaGenerateBoundsError
  loc_0040989C: sub ebx, edx
  loc_0040989E: cmp ebx, eax
  loc_004098A0: jb 004098AAh
  loc_004098A2: call __vbaGenerateBoundsError
  loc_004098A4: mov ecx, [004150B4h]
  loc_004098AA: lea eax, [ebx*8]
  loc_004098B1: jmp 004098C5h
  loc_004098B3: call [00401094h] ; __vbaGenerateBoundsError
  loc_004098B9: mov ecx, [004150B4h]
  loc_004098BF: mov esi, [00401094h] ; __vbaGenerateBoundsError
  loc_004098C5: mov ecx, [ecx+0000000Ch]
  loc_004098C8: mov edx, var_11C
  loc_004098CE: mov [ecx+eax], edx
  loc_004098D1: mov edx, var_118
  loc_004098D7: mov [ecx+eax+00000004h], edx
  loc_004098DB: mov eax, [edi]
  loc_004098DD: lea ecx, var_14
  loc_004098E0: push eax
  loc_004098E1: push ecx
  loc_004098E2: call [0040114Ch] ; __vbaAryLock
  loc_004098E8: mov ecx, var_14
  loc_004098EB: test ecx, ecx
  loc_004098ED: jz 0040992Fh
  loc_004098EF: cmp [ecx], 0002h
  loc_004098F3: jnz 0040992Fh
  loc_004098F5: movsx ebx, [0041509Eh]
  loc_004098FC: mov edx, [ecx+0000001Ch]
  loc_004098FF: mov eax, [ecx+00000018h]
  loc_00409902: sub ebx, edx
  loc_00409904: cmp ebx, eax
  loc_00409906: jb 0040990Dh
  loc_00409908: call __vbaGenerateBoundsError
  loc_0040990A: mov ecx, var_14
  loc_0040990D: mov edx, arg_10
  loc_00409910: mov eax, [ecx+00000010h]
  loc_00409913: movsx edi, [edx]
  loc_00409916: sub edi, [ecx+00000014h]
  loc_00409919: cmp edi, eax
  loc_0040991B: jb 00409922h
  loc_0040991D: call __vbaGenerateBoundsError
  loc_0040991F: mov ecx, var_14
  loc_00409922: mov eax, [ecx+00000018h]
  loc_00409925: imul eax, edi
  loc_00409928: add eax, ebx
  loc_0040992A: shl eax, 03h
  loc_0040992D: jmp 00409934h
  loc_0040992F: call __vbaGenerateBoundsError
  loc_00409931: mov ecx, var_14
  loc_00409934: mov ecx, [ecx+0000000Ch]
  loc_00409937: add ecx, eax
  loc_00409939: push ecx
  loc_0040993A: call 00409C30h
  loc_0040993F: fstp real8 ptr var_11C
  loc_00409945: lea edx, var_14
  loc_00409948: push edx
  loc_00409949: call [00401174h] ; __vbaAryUnlock
  loc_0040994F: mov ecx, [004150B8h]
  loc_00409955: test ecx, ecx
  loc_00409957: jz 00409983h
  loc_00409959: cmp [ecx], 0001h
  loc_0040995D: jnz 00409983h
  loc_0040995F: movsx edi, [0041509Eh]
  loc_00409966: mov edx, [ecx+00000014h]
  loc_00409969: mov eax, [ecx+00000010h]
  loc_0040996C: sub edi, edx
  loc_0040996E: cmp edi, eax
  loc_00409970: jb 0040997Ah
  loc_00409972: call __vbaGenerateBoundsError
  loc_00409974: mov ecx, [004150B8h]
  loc_0040997A: lea eax, [edi*8]
  loc_00409981: jmp 0040998Bh
  loc_00409983: call __vbaGenerateBoundsError
  loc_00409985: mov ecx, [004150B8h]
  loc_0040998B: mov ecx, [ecx+0000000Ch]
  loc_0040998E: mov edx, var_11C
  loc_00409994: mov ebx, 00000002h
  loc_00409999: mov [ecx+eax], edx
  loc_0040999C: mov edx, var_118
  loc_004099A2: mov [ecx+eax+00000004h], edx
  loc_004099A6: mov ecx, [004150B4h]
  loc_004099AC: test ecx, ecx
  loc_004099AE: jz 004099DAh
  loc_004099B0: cmp [ecx], 0001h
  loc_004099B4: jnz 004099DAh
  loc_004099B6: movsx edi, [0041509Eh]
  loc_004099BD: mov edx, [ecx+00000014h]
  loc_004099C0: mov eax, [ecx+00000010h]
  loc_004099C3: sub edi, edx
  loc_004099C5: cmp edi, eax
  loc_004099C7: jb 004099D1h
  loc_004099C9: call __vbaGenerateBoundsError
  loc_004099CB: mov ecx, [004150B4h]
  loc_004099D1: lea eax, [edi*8]
  loc_004099D8: jmp 004099E2h
  loc_004099DA: call __vbaGenerateBoundsError
  loc_004099DC: mov ecx, [004150B4h]
  loc_004099E2: mov ecx, [ecx+0000000Ch]
  loc_004099E5: add ecx, eax
  loc_004099E7: mov var_6C, ecx
  loc_004099EA: mov ecx, [004150B8h]
  loc_004099F0: test ecx, ecx
  loc_004099F2: jz 00409A1Eh
  loc_004099F4: cmp [ecx], 0001h
  loc_004099F8: jnz 00409A1Eh
  loc_004099FA: movsx edi, [0041509Eh]
  loc_00409A01: mov edx, [ecx+00000014h]
  loc_00409A04: mov eax, [ecx+00000010h]
  loc_00409A07: sub edi, edx
  loc_00409A09: cmp edi, eax
  loc_00409A0B: jb 00409A15h
  loc_00409A0D: call __vbaGenerateBoundsError
  loc_00409A0F: mov ecx, [004150B8h]
  loc_00409A15: lea eax, [edi*8]
  loc_00409A1C: jmp 00409A26h
  loc_00409A1E: call __vbaGenerateBoundsError
  loc_00409A20: mov ecx, [004150B8h]
  loc_00409A26: sub esp, 00000010h
  loc_00409A29: mov ecx, [ecx+0000000Ch]
  loc_00409A2C: mov edx, esp
  loc_00409A2E: add ecx, eax
  loc_00409A30: mov eax, 00004005h
  loc_00409A35: xor esi, esi
  loc_00409A37: mov [edx], ebx
  loc_00409A39: mov ebx, var_50
  loc_00409A3C: mov var_3C, 00000064h
  loc_00409A43: mov var_44, 00000002h
  loc_00409A4A: mov [edx+00000004h], ebx
  loc_00409A4D: mov ebx, esp
  loc_00409A4F: mov var_148, edx
  loc_00409A55: xor edx, edx
  loc_00409A57: mov [ebx+00000008h], edx
  loc_00409A5A: mov edx, var_48
  loc_00409A5D: sub esp, 00000010h
  loc_00409A60: mov edi, 00000002h
  loc_00409A65: mov [ebx+0000000Ch], edx
  loc_00409A68: mov ebx, esp
  loc_00409A6A: mov edx, eax
  loc_00409A6C: sub esp, 00000010h
  loc_00409A6F: mov [ebx], edx
  loc_00409A71: mov edx, var_70
  loc_00409A74: mov [ebx+00000004h], edx
  loc_00409A77: mov edx, var_6C
  loc_00409A7A: mov [ebx+00000008h], edx
  loc_00409A7D: mov edx, var_68
  loc_00409A80: mov [ebx+0000000Ch], edx
  loc_00409A83: mov edx, esp
  loc_00409A85: mov ebx, Me
  loc_00409A88: push esi
  loc_00409A89: mov [edx], eax
  loc_00409A8B: mov eax, var_80
  loc_00409A8E: push 0040366Ch ; "ScaleWidth"
  loc_00409A93: mov [edx+00000004h], eax
  loc_00409A96: lea eax, var_24
  loc_00409A99: mov [edx+00000008h], ecx
  loc_00409A9C: mov ecx, var_78
  loc_00409A9F: mov [edx+0000000Ch], ecx
  loc_00409AA2: mov edx, [ebx]
  loc_00409AA4: push edx
  loc_00409AA5: push eax
  loc_00409AA6: call [0040115Ch] ; __vbaLateMemCallLd
  loc_00409AAC: add esp, 00000010h
  loc_00409AAF: lea ecx, var_44
  loc_00409AB2: lea edx, var_34
  loc_00409AB5: push eax
  loc_00409AB6: push ecx
  loc_00409AB7: push edx
  loc_00409AB8: call [004010ECh] ; __vbaVarDiv
  loc_00409ABE: mov edx, [eax]
  loc_00409AC0: sub esp, 00000010h
  loc_00409AC3: mov ecx, esp
  loc_00409AC5: sub esp, 00000010h
  loc_00409AC8: mov [ecx], edx
  loc_00409ACA: mov edx, [eax+00000004h]
  loc_00409ACD: mov [ecx+00000004h], edx
  loc_00409AD0: mov edx, [eax+00000008h]
  loc_00409AD3: mov eax, [eax+0000000Ch]
  loc_00409AD6: mov [ecx+00000008h], edx
  loc_00409AD9: mov edx, var_A0
  loc_00409ADF: mov [ecx+0000000Ch], eax
  loc_00409AE2: mov ecx, esp
  loc_00409AE4: mov eax, var_98
  loc_00409AEA: sub esp, 00000010h
  loc_00409AED: mov [ecx], edi
  loc_00409AEF: mov [ecx+00000004h], edx
  loc_00409AF2: mov edx, var_C0
  loc_00409AF8: mov [ecx+00000008h], esi
  loc_00409AFB: mov [ecx+0000000Ch], eax
  loc_00409AFE: mov ecx, esp
  loc_00409B00: mov eax, edi
  loc_00409B02: mov [ecx], eax
  loc_00409B04: xor eax, eax
  loc_00409B06: sub esp, 00000010h
  loc_00409B09: mov [ecx+00000004h], edx
  loc_00409B0C: mov [ecx+00000008h], eax
  loc_00409B0F: mov eax, var_B8
  loc_00409B15: mov [ecx+0000000Ch], eax
  loc_00409B18: mov ecx, esp
  loc_00409B1A: mov edx, var_E0
  loc_00409B20: mov eax, edi
  loc_00409B22: mov [ecx], eax
  loc_00409B24: xor eax, eax
  loc_00409B26: sub esp, 00000010h
  loc_00409B29: mov [ecx+00000004h], edx
  loc_00409B2C: mov edx, var_100
  loc_00409B32: mov [ecx+00000008h], eax
  loc_00409B35: mov eax, var_D8
  loc_00409B3B: mov [ecx+0000000Ch], eax
  loc_00409B3E: mov ecx, esp
  loc_00409B40: mov eax, edi
  loc_00409B42: push 00000008h
  loc_00409B44: mov [ecx], eax
  loc_00409B46: xor eax, eax
  loc_00409B48: push 004036A8h ; "Circle"
  loc_00409B4D: mov [ecx+00000004h], edx
  loc_00409B50: mov [ecx+00000008h], eax
  loc_00409B53: mov eax, var_F8
  loc_00409B59: mov [ecx+0000000Ch], eax
  loc_00409B5C: mov ecx, [ebx]
  loc_00409B5E: push ecx
  loc_00409B5F: call [00401148h] ; __vbaLateMemCall
  loc_00409B65: add esp, 0000008Ch
  loc_00409B6B: lea ecx, var_24
  loc_00409B6E: call [00401014h] ; __vbaFreeVar
  loc_00409B74: mov edi, arg_C
  loc_00409B77: mov eax, 00000001h
  loc_00409B7C: add ax, [0041509Eh]
  loc_00409B83: jo 00409BCFh
  loc_00409B85: mov [0041509Eh], ax
  loc_00409B8B: jmp 004097EAh
  loc_00409B90: fwait
  loc_00409B91: push 00409BB7h
  loc_00409B96: jmp 00409BB6h
  loc_00409B98: lea edx, var_14
  loc_00409B9B: push edx
  loc_00409B9C: call [00401174h] ; __vbaAryUnlock
  loc_00409BA2: lea eax, var_34
  loc_00409BA5: lea ecx, var_24
  loc_00409BA8: push eax
  loc_00409BA9: push ecx
  loc_00409BAA: push 00000002h
  loc_00409BAC: call [00401024h] ; __vbaFreeVarList
  loc_00409BB2: add esp, 0000000Ch
  loc_00409BB5: ret
  loc_00409BB6: ret
  loc_00409BB7: mov ecx, var_10
  loc_00409BBA: pop edi
  loc_00409BBB: pop esi
  loc_00409BBC: mov fs:[00000000h], ecx
  loc_00409BC3: pop ebx
  loc_00409BC4: mov esp, ebp
  loc_00409BC6: pop ebp
  loc_00409BC7: retn 0010h
End Sub
