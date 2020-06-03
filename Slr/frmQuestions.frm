VERSION 5.00
Begin VB.Form frmQuestions
  Caption = "SLR - Questions Associated with Specific Parameters"
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 510
  ClientWidth = 8535
  ClientHeight = 5055
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame Frame2
    Left = 120
    Top = 6120
    Width = 11655
    Height = 2055
    TabIndex = 10
    Begin VB.CommandButton cmdGoEstimation
      Caption = "Estimation and Prediction"
      Left = 3600
      Top = 120
      Width = 4215
      Height = 495
      TabIndex = 16
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "Click here to go to the estimation interval procedure."
    End
    Begin VB.Label Label2
      Caption = "II. Prediction"
      Left = 5520
      Top = 600
      Width = 2295
      Height = 495
      TabIndex = 19
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
    Begin VB.Label Label1
      Caption = "I. Estimation:"
      Left = 360
      Top = 600
      Width = 2535
      Height = 375
      TabIndex = 18
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
    Begin VB.Label lblPredict
      Caption = "What is the value of Y when X equals Xg?"
      Left = 5760
      Top = 1080
      Width = 5415
      Height = 855
      TabIndex = 17
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
    Begin VB.Label lblEstimation
      Caption = "What is the average value of Y given X equals Xg?"
      Left = 480
      Top = 1080
      Width = 4815
      Height = 855
      TabIndex = 11
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
    Left = 120
    Top = 0
    Width = 11775
    Height = 6135
    TabIndex = 0
    Begin VB.CommandButton cmdGoSlope
      Caption = "Slope"
      Left = 4800
      Top = 0
      Width = 2055
      Height = 495
      TabIndex = 15
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "Click here to go to procedures for inferring about the slope"
    End
    Begin VB.Label lblSlopeLeft
      Caption = "Are increases in X associated with decreases in the mean of Y?"
      Left = 1080
      Top = 2040
      Width = 10455
      Height = 855
      TabIndex = 14
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
    Begin VB.Label Label6
      Caption = "A. Left-Sided "
      Left = 720
      Top = 1680
      Width = 2655
      Height = 495
      TabIndex = 13
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
    Begin VB.Label lblSlopeTestLabel
      Caption = "II. Test of Hypothesis"
      Left = 360
      Top = 1320
      Width = 3375
      Height = 375
      TabIndex = 12
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
    Begin VB.Label lblSlope2d
      Caption = "4. Is variation in X associated with variation in Y?"
      Left = 1080
      Top = 5640
      Width = 9615
      Height = 375
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
    Begin VB.Label lblSlope2c
      Caption = "3. Does X help estimate the mean of Y?"
      Left = 1080
      Top = 5280
      Width = 7575
      Height = 375
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
      ToolTipText = "what"
    End
    Begin VB.Label lblSlope2b
      Caption = "2. Does X help predict Y?"
      Left = 1080
      Top = 4920
      Width = 7695
      Height = 375
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
    Begin VB.Label lblSlope2a
      Caption = "1. Are changes in X associated with changes in the mean of Y?"
      Left = 1080
      Top = 4440
      Width = 10125
      Height = 435
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
    Begin VB.Label Label5
      Caption = "C. Two-Sided "
      Left = 720
      Top = 4080
      Width = 2775
      Height = 375
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
    Begin VB.Label lblSlope1
      Caption = "Are increases in X associated with increases in the mean of Y?"
      Left = 1080
      Top = 3240
      Width = 10335
      Height = 855
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
    Begin VB.Label Label3
      Caption = "B. Right-Sided "
      Left = 720
      Top = 2880
      Width = 2415
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
    Begin VB.Label lblSlopeCI
      Caption = "What is the change in the mean value of Y when X increases by one unit?"
      Left = 720
      Top = 480
      Width = 10335
      Height = 855
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
    Begin VB.Label lblSlopeCILabel
      Caption = "I. Confidence Interval:"
      Left = 360
      Top = 120
      Width = 3975
      Height = 375
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
  Begin VB.Menu mnuInstructions
    Caption = "&Instructions"
    Begin VB.Menu mnuContext
      Caption = "Co&ntext"
    End
    Begin VB.Menu mnuProcedures
      Caption = "&Procedures"
    End
  End
  Begin VB.Menu mnuChange
    Caption = "&Change"
    Begin VB.Menu mnuChangeNames
      Caption = "&Names"
    End
    Begin VB.Menu mnuUnits
      Caption = "&Units"
    End
    Begin VB.Menu mnuChangeXg
      Caption = "&Xg"
    End
    Begin VB.Menu mnuChangeHypValue
      Caption = "&Hypothesized value of the slope"
    End
  End
  Begin VB.Menu mnuGoTo
    Caption = "&Go To"
    Begin VB.Menu mnuIntoduction
      Caption = "In&troduction"
    End
    Begin VB.Menu mnuStatistics
      Caption = "Statistics and Point Esti&mates"
    End
    Begin VB.Menu mnuInferences
      Caption = "In&ferences"
      Begin VB.Menu mnuSlope
        Caption = "&Slope"
      End
      Begin VB.Menu mnuEstPred
        Caption = "&Estimation and Prediction"
      End
    End
    Begin VB.Menu mnuAssumptions
      Caption = "A&ssumptions"
    End
  End
  Begin VB.Menu mnuExit
    Caption = "&Exit"
  End
End

Attribute VB_Name = "frmQuestions"


Private Sub mnuChangeXg_Click() '4271D0
  loc_004271D0: push ebp
  loc_004271D1: mov ebp, esp
  loc_004271D3: sub esp, 0000000Ch
  loc_004271D6: push 00401926h ; __vbaExceptHandler
  loc_004271DB: mov eax, fs:[00000000h]
  loc_004271E1: push eax
  loc_004271E2: mov fs:[00000000h], esp
  loc_004271E9: sub esp, 00000030h
  loc_004271EC: push ebx
  loc_004271ED: push esi
  loc_004271EE: push edi
  loc_004271EF: mov var_C, esp
  loc_004271F2: mov var_8, 00401600h
  loc_004271F9: mov eax, Me
  loc_004271FC: mov ecx, eax
  loc_004271FE: and ecx, 00000001h
  loc_00427201: mov var_4, ecx
  loc_00427204: and al, FEh
  loc_00427206: push eax
  loc_00427207: mov Me, eax
  loc_0042720A: mov edx, [eax]
  loc_0042720C: call [edx+00000004h]
  loc_0042720F: mov eax, [004301B4h]
  loc_00427214: test eax, eax
  loc_00427216: jnz 00427228h
  loc_00427218: push 004301B4h
  loc_0042721D: push 00402480h
  loc_00427222: call [004010D4h] ; __vbaNew2
  loc_00427228: sub esp, 00000010h
  loc_0042722B: mov ecx, 0000000Ah
  loc_00427230: mov ebx, esp
  loc_00427232: mov var_24, ecx
  loc_00427235: mov eax, 80020004h
  loc_0042723A: sub esp, 00000010h
  loc_0042723D: mov [ebx], ecx
  loc_0042723F: mov ecx, var_30
  loc_00427242: mov edx, eax
  loc_00427244: mov esi, [004301B4h]
  loc_0042724A: mov [ebx+00000004h], ecx
  loc_0042724D: mov ecx, esp
  loc_0042724F: mov edi, [esi]
  loc_00427251: push esi
  loc_00427252: mov [ebx+00000008h], eax
  loc_00427255: mov eax, var_28
  loc_00427258: mov [ebx+0000000Ch], eax
  loc_0042725B: mov eax, var_24
  loc_0042725E: mov [ecx], eax
  loc_00427260: mov eax, var_20
  loc_00427263: mov [ecx+00000004h], eax
  loc_00427266: mov [ecx+00000008h], edx
  loc_00427269: mov edx, var_18
  loc_0042726C: mov [ecx+0000000Ch], edx
  loc_0042726F: call [edi+000002B0h]
  loc_00427275: test eax, eax
  loc_00427277: fnclex
  loc_00427279: jge 0042728Dh
  loc_0042727B: push 000002B0h
  loc_00427280: push 00410574h
  loc_00427285: push esi
  loc_00427286: push eax
  loc_00427287: call [00401030h] ; __vbaHresultCheckObj
  loc_0042728D: mov esi, Me
  loc_00427290: push esi
  loc_00427291: mov eax, [esi]
  loc_00427293: call [eax+000006F8h]
  loc_00427299: test eax, eax
  loc_0042729B: jge 004272AFh
  loc_0042729D: push 000006F8h
  loc_004272A2: push 0040FF78h
  loc_004272A7: push esi
  loc_004272A8: push eax
  loc_004272A9: call [00401030h] ; __vbaHresultCheckObj
  loc_004272AF: mov var_4, 00000000h
  loc_004272B6: mov eax, Me
  loc_004272B9: push eax
  loc_004272BA: mov ecx, [eax]
  loc_004272BC: call [ecx+00000008h]
  loc_004272BF: mov eax, var_4
  loc_004272C2: mov ecx, var_14
  loc_004272C5: pop edi
  loc_004272C6: pop esi
  loc_004272C7: mov fs:[00000000h], ecx
  loc_004272CE: pop ebx
  loc_004272CF: mov esp, ebp
  loc_004272D1: pop ebp
  loc_004272D2: retn 0004h
End Sub

Private Sub mnuChangeNames_Click() '4270C0
  loc_004270C0: push ebp
  loc_004270C1: mov ebp, esp
  loc_004270C3: sub esp, 0000000Ch
  loc_004270C6: push 00401926h ; __vbaExceptHandler
  loc_004270CB: mov eax, fs:[00000000h]
  loc_004270D1: push eax
  loc_004270D2: mov fs:[00000000h], esp
  loc_004270D9: sub esp, 00000030h
  loc_004270DC: push ebx
  loc_004270DD: push esi
  loc_004270DE: push edi
  loc_004270DF: mov var_C, esp
  loc_004270E2: mov var_8, 004015F8h
  loc_004270E9: mov eax, Me
  loc_004270EC: mov ecx, eax
  loc_004270EE: and ecx, 00000001h
  loc_004270F1: mov var_4, ecx
  loc_004270F4: and al, FEh
  loc_004270F6: push eax
  loc_004270F7: mov Me, eax
  loc_004270FA: mov edx, [eax]
  loc_004270FC: call [edx+00000004h]
  loc_004270FF: mov eax, [004300D8h]
  loc_00427104: test eax, eax
  loc_00427106: jnz 00427118h
  loc_00427108: push 004300D8h
  loc_0042710D: push 00402E04h
  loc_00427112: call [004010D4h] ; __vbaNew2
  loc_00427118: sub esp, 00000010h
  loc_0042711B: mov ecx, 0000000Ah
  loc_00427120: mov ebx, esp
  loc_00427122: mov var_24, ecx
  loc_00427125: mov eax, 80020004h
  loc_0042712A: sub esp, 00000010h
  loc_0042712D: mov [ebx], ecx
  loc_0042712F: mov ecx, var_30
  loc_00427132: mov edx, eax
  loc_00427134: mov esi, [004300D8h]
  loc_0042713A: mov [ebx+00000004h], ecx
  loc_0042713D: mov ecx, esp
  loc_0042713F: mov edi, [esi]
  loc_00427141: push esi
  loc_00427142: mov [ebx+00000008h], eax
  loc_00427145: mov eax, var_28
  loc_00427148: mov [ebx+0000000Ch], eax
  loc_0042714B: mov eax, var_24
  loc_0042714E: mov [ecx], eax
  loc_00427150: mov eax, var_20
  loc_00427153: mov [ecx+00000004h], eax
  loc_00427156: mov [ecx+00000008h], edx
  loc_00427159: mov edx, var_18
  loc_0042715C: mov [ecx+0000000Ch], edx
  loc_0042715F: call [edi+000002B0h]
  loc_00427165: test eax, eax
  loc_00427167: fnclex
  loc_00427169: jge 0042717Dh
  loc_0042716B: push 000002B0h
  loc_00427170: push 0040E260h
  loc_00427175: push esi
  loc_00427176: push eax
  loc_00427177: call [00401030h] ; __vbaHresultCheckObj
  loc_0042717D: mov esi, Me
  loc_00427180: push esi
  loc_00427181: mov eax, [esi]
  loc_00427183: call [eax+000006F8h]
  loc_00427189: test eax, eax
  loc_0042718B: jge 0042719Fh
  loc_0042718D: push 000006F8h
  loc_00427192: push 0040FF78h
  loc_00427197: push esi
  loc_00427198: push eax
  loc_00427199: call [00401030h] ; __vbaHresultCheckObj
  loc_0042719F: mov var_4, 00000000h
  loc_004271A6: mov eax, Me
  loc_004271A9: push eax
  loc_004271AA: mov ecx, [eax]
  loc_004271AC: call [ecx+00000008h]
  loc_004271AF: mov eax, var_4
  loc_004271B2: mov ecx, var_14
  loc_004271B5: pop edi
  loc_004271B6: pop esi
  loc_004271B7: mov fs:[00000000h], ecx
  loc_004271BE: pop ebx
  loc_004271BF: mov esp, ebp
  loc_004271C1: pop ebp
  loc_004271C2: retn 0004h
End Sub

Private Sub mnuProcedures_Click() '427550
  loc_00427550: push ebp
  loc_00427551: mov ebp, esp
  loc_00427553: sub esp, 0000000Ch
  loc_00427556: push 00401926h ; __vbaExceptHandler
  loc_0042755B: mov eax, fs:[00000000h]
  loc_00427561: push eax
  loc_00427562: mov fs:[00000000h], esp
  loc_00427569: sub esp, 00000088h
  loc_0042756F: push ebx
  loc_00427570: push esi
  loc_00427571: push edi
  loc_00427572: mov var_C, esp
  loc_00427575: mov var_8, 00401620h
  loc_0042757C: mov eax, Me
  loc_0042757F: mov ecx, eax
  loc_00427581: and ecx, 00000001h
  loc_00427584: mov var_4, ecx
  loc_00427587: and al, FEh
  loc_00427589: push eax
  loc_0042758A: mov Me, eax
  loc_0042758D: mov edx, [eax]
  loc_0042758F: call [edx+00000004h]
  loc_00427592: mov edi, [004010F4h] ; __vbaVarDup
  loc_00427598: mov ecx, 80020004h
  loc_0042759D: xor esi, esi
  loc_0042759F: mov var_4C, ecx
  loc_004275A2: mov eax, 0000000Ah
  loc_004275A7: mov var_3C, ecx
  loc_004275AA: mov ebx, 00000008h
  loc_004275AF: mov var_44, esi
  loc_004275B2: mov var_54, esi
  loc_004275B5: mov var_74, esi
  loc_004275B8: lea edx, var_74
  loc_004275BB: lea ecx, var_34
  loc_004275BE: mov var_24, esi
  loc_004275C1: mov var_34, esi
  loc_004275C4: mov var_64, esi
  loc_004275C7: mov var_54, eax
  loc_004275CA: mov var_44, eax
  loc_004275CD: mov var_6C, 0041180Ch ; "Statistical Procedure"
  loc_004275D4: mov var_74, ebx
  loc_004275D7: call edi
  loc_004275D9: lea edx, var_64
  loc_004275DC: lea ecx, var_24
  loc_004275DF: mov var_5C, 00411780h ; "To go to a specific procedure, click on one of the command buttons."
  loc_004275E6: mov var_64, ebx
  loc_004275E9: call edi
  loc_004275EB: lea eax, var_54
  loc_004275EE: lea ecx, var_44
  loc_004275F1: push eax
  loc_004275F2: lea edx, var_34
  loc_004275F5: push ecx
  loc_004275F6: push edx
  loc_004275F7: lea eax, var_24
  loc_004275FA: push esi
  loc_004275FB: push eax
  loc_004275FC: call [00401038h] ; rtcMsgBox
  loc_00427602: lea ecx, var_54
  loc_00427605: lea edx, var_44
  loc_00427608: push ecx
  loc_00427609: lea eax, var_34
  loc_0042760C: push edx
  loc_0042760D: lea ecx, var_24
  loc_00427610: push eax
  loc_00427611: push ecx
  loc_00427612: push 00000004h
  loc_00427614: call [00401018h] ; __vbaFreeVarList
  loc_0042761A: add esp, 00000014h
  loc_0042761D: mov var_4, esi
  loc_00427620: push 00427644h
  loc_00427625: jmp 00427643h
  loc_00427627: lea edx, var_54
  loc_0042762A: lea eax, var_44
  loc_0042762D: push edx
  loc_0042762E: lea ecx, var_34
  loc_00427631: push eax
  loc_00427632: lea edx, var_24
  loc_00427635: push ecx
  loc_00427636: push edx
  loc_00427637: push 00000004h
  loc_00427639: call [00401018h] ; __vbaFreeVarList
  loc_0042763F: add esp, 00000014h
  loc_00427642: ret
  loc_00427643: ret
  loc_00427644: mov eax, Me
  loc_00427647: push eax
  loc_00427648: mov ecx, [eax]
  loc_0042764A: call [ecx+00000008h]
  loc_0042764D: mov eax, var_4
  loc_00427650: mov ecx, var_14
  loc_00427653: pop edi
  loc_00427654: pop esi
  loc_00427655: mov fs:[00000000h], ecx
  loc_0042765C: pop ebx
  loc_0042765D: mov esp, ebp
  loc_0042765F: pop ebp
  loc_00427660: retn 0004h
End Sub

Private Sub mnuAssumptions_Click() '425270
  loc_00425270: push ebp
  loc_00425271: mov ebp, esp
  loc_00425273: sub esp, 0000000Ch
  loc_00425276: push 00401926h ; __vbaExceptHandler
  loc_0042527B: mov eax, fs:[00000000h]
  loc_00425281: push eax
  loc_00425282: mov fs:[00000000h], esp
  loc_00425289: sub esp, 00000030h
  loc_0042528C: push ebx
  loc_0042528D: push esi
  loc_0042528E: push edi
  loc_0042528F: mov var_C, esp
  loc_00425292: mov var_8, 004015B0h
  loc_00425299: mov eax, Me
  loc_0042529C: mov ecx, eax
  loc_0042529E: and ecx, 00000001h
  loc_004252A1: mov var_4, ecx
  loc_004252A4: and al, FEh
  loc_004252A6: push eax
  loc_004252A7: mov Me, eax
  loc_004252AA: mov edx, [eax]
  loc_004252AC: call [edx+00000004h]
  loc_004252AF: mov eax, [0043009Ch]
  loc_004252B4: test eax, eax
  loc_004252B6: jnz 004252C8h
  loc_004252B8: push 0043009Ch
  loc_004252BD: push 00405FC0h
  loc_004252C2: call [004010D4h] ; __vbaNew2
  loc_004252C8: sub esp, 00000010h
  loc_004252CB: mov ecx, 0000000Ah
  loc_004252D0: mov ebx, esp
  loc_004252D2: mov var_24, ecx
  loc_004252D5: mov eax, 80020004h
  loc_004252DA: sub esp, 00000010h
  loc_004252DD: mov [ebx], ecx
  loc_004252DF: mov ecx, var_30
  loc_004252E2: mov edx, eax
  loc_004252E4: mov esi, [0043009Ch]
  loc_004252EA: mov [ebx+00000004h], ecx
  loc_004252ED: mov ecx, esp
  loc_004252EF: mov edi, [esi]
  loc_004252F1: push esi
  loc_004252F2: mov [ebx+00000008h], eax
  loc_004252F5: mov eax, var_28
  loc_004252F8: mov [ebx+0000000Ch], eax
  loc_004252FB: mov eax, var_24
  loc_004252FE: mov [ecx], eax
  loc_00425300: mov eax, var_20
  loc_00425303: mov [ecx+00000004h], eax
  loc_00425306: mov [ecx+00000008h], edx
  loc_00425309: mov edx, var_18
  loc_0042530C: mov [ecx+0000000Ch], edx
  loc_0042530F: call [edi+000002B0h]
  loc_00425315: test eax, eax
  loc_00425317: fnclex
  loc_00425319: jge 0042532Dh
  loc_0042531B: push 000002B0h
  loc_00425320: push 0040DDE0h
  loc_00425325: push esi
  loc_00425326: push eax
  loc_00425327: call [00401030h] ; __vbaHresultCheckObj
  loc_0042532D: mov eax, [00430164h]
  loc_00425332: test eax, eax
  loc_00425334: jnz 00425346h
  loc_00425336: push 00430164h
  loc_0042533B: push 00408108h
  loc_00425340: call [004010D4h] ; __vbaNew2
  loc_00425346: mov esi, [00430164h]
  loc_0042534C: push esi
  loc_0042534D: mov eax, [esi]
  loc_0042534F: call [eax+000002B4h]
  loc_00425355: test eax, eax
  loc_00425357: fnclex
  loc_00425359: jge 0042536Dh
  loc_0042535B: push 000002B4h
  loc_00425360: push 0040FF58h
  loc_00425365: push esi
  loc_00425366: push eax
  loc_00425367: call [00401030h] ; __vbaHresultCheckObj
  loc_0042536D: mov var_4, 00000000h
  loc_00425374: mov eax, Me
  loc_00425377: push eax
  loc_00425378: mov ecx, [eax]
  loc_0042537A: call [ecx+00000008h]
  loc_0042537D: mov eax, var_4
  loc_00425380: mov ecx, var_14
  loc_00425383: pop edi
  loc_00425384: pop esi
  loc_00425385: mov fs:[00000000h], ecx
  loc_0042538C: pop ebx
  loc_0042538D: mov esp, ebp
  loc_0042538F: pop ebp
  loc_00425390: retn 0004h
End Sub

Private Sub mnuEstPred_Click() '425490
  loc_00425490: push ebp
  loc_00425491: mov ebp, esp
  loc_00425493: sub esp, 0000000Ch
  loc_00425496: push 00401926h ; __vbaExceptHandler
  loc_0042549B: mov eax, fs:[00000000h]
  loc_004254A1: push eax
  loc_004254A2: mov fs:[00000000h], esp
  loc_004254A9: sub esp, 00000030h
  loc_004254AC: push ebx
  loc_004254AD: push esi
  loc_004254AE: push edi
  loc_004254AF: mov var_C, esp
  loc_004254B2: mov var_8, 004015C0h
  loc_004254B9: mov eax, Me
  loc_004254BC: mov ecx, eax
  loc_004254BE: and ecx, 00000001h
  loc_004254C1: mov var_4, ecx
  loc_004254C4: and al, FEh
  loc_004254C6: push eax
  loc_004254C7: mov Me, eax
  loc_004254CA: mov edx, [eax]
  loc_004254CC: call [edx+00000004h]
  loc_004254CF: mov eax, [00430164h]
  loc_004254D4: test eax, eax
  loc_004254D6: jnz 004254E8h
  loc_004254D8: push 00430164h
  loc_004254DD: push 00408108h
  loc_004254E2: call [004010D4h] ; __vbaNew2
  loc_004254E8: mov esi, [00430164h]
  loc_004254EE: push esi
  loc_004254EF: mov eax, [esi]
  loc_004254F1: call [eax+000002B4h]
  loc_004254F7: test eax, eax
  loc_004254F9: fnclex
  loc_004254FB: jge 0042550Fh
  loc_004254FD: push 000002B4h
  loc_00425502: push 0040FF58h
  loc_00425507: push esi
  loc_00425508: push eax
  loc_00425509: call [00401030h] ; __vbaHresultCheckObj
  loc_0042550F: mov eax, [004300B0h]
  loc_00425514: test eax, eax
  loc_00425516: jnz 00425528h
  loc_00425518: push 004300B0h
  loc_0042551D: push 00407528h
  loc_00425522: call [004010D4h] ; __vbaNew2
  loc_00425528: sub esp, 00000010h
  loc_0042552B: mov ecx, 0000000Ah
  loc_00425530: mov ebx, esp
  loc_00425532: mov var_24, ecx
  loc_00425535: mov eax, 80020004h
  loc_0042553A: sub esp, 00000010h
  loc_0042553D: mov [ebx], ecx
  loc_0042553F: mov ecx, var_30
  loc_00425542: mov edx, eax
  loc_00425544: mov esi, [004300B0h]
  loc_0042554A: mov [ebx+00000004h], ecx
  loc_0042554D: mov ecx, esp
  loc_0042554F: mov edi, [esi]
  loc_00425551: push esi
  loc_00425552: mov [ebx+00000008h], eax
  loc_00425555: mov eax, var_28
  loc_00425558: mov [ebx+0000000Ch], eax
  loc_0042555B: mov eax, var_24
  loc_0042555E: mov [ecx], eax
  loc_00425560: mov eax, var_20
  loc_00425563: mov [ecx+00000004h], eax
  loc_00425566: mov [ecx+00000008h], edx
  loc_00425569: mov edx, var_18
  loc_0042556C: mov [ecx+0000000Ch], edx
  loc_0042556F: call [edi+000002B0h]
  loc_00425575: test eax, eax
  loc_00425577: fnclex
  loc_00425579: jge 0042558Dh
  loc_0042557B: push 000002B0h
  loc_00425580: push 0040DE98h
  loc_00425585: push esi
  loc_00425586: push eax
  loc_00425587: call [00401030h] ; __vbaHresultCheckObj
  loc_0042558D: mov var_4, 00000000h
  loc_00425594: mov eax, Me
  loc_00425597: push eax
  loc_00425598: mov ecx, [eax]
  loc_0042559A: call [ecx+00000008h]
  loc_0042559D: mov eax, var_4
  loc_004255A0: mov ecx, var_14
  loc_004255A3: pop edi
  loc_004255A4: pop esi
  loc_004255A5: mov fs:[00000000h], ecx
  loc_004255AC: pop ebx
  loc_004255AD: mov esp, ebp
  loc_004255AF: pop ebp
  loc_004255B0: retn 0004h
End Sub

Private Sub mnuStatistics_Click() '427940
  loc_00427940: push ebp
  loc_00427941: mov ebp, esp
  loc_00427943: sub esp, 0000000Ch
  loc_00427946: push 00401926h ; __vbaExceptHandler
  loc_0042794B: mov eax, fs:[00000000h]
  loc_00427951: push eax
  loc_00427952: mov fs:[00000000h], esp
  loc_00427959: sub esp, 00000030h
  loc_0042795C: push ebx
  loc_0042795D: push esi
  loc_0042795E: push edi
  loc_0042795F: mov var_C, esp
  loc_00427962: mov var_8, 00401648h
  loc_00427969: mov eax, Me
  loc_0042796C: mov ecx, eax
  loc_0042796E: and ecx, 00000001h
  loc_00427971: mov var_4, ecx
  loc_00427974: and al, FEh
  loc_00427976: push eax
  loc_00427977: mov Me, eax
  loc_0042797A: mov edx, [eax]
  loc_0042797C: call [edx+00000004h]
  loc_0042797F: mov eax, [004300ECh]
  loc_00427984: test eax, eax
  loc_00427986: jnz 00427998h
  loc_00427988: push 004300ECh
  loc_0042798D: push 0040A624h
  loc_00427992: call [004010D4h] ; __vbaNew2
  loc_00427998: sub esp, 00000010h
  loc_0042799B: mov ecx, 0000000Ah
  loc_004279A0: mov ebx, esp
  loc_004279A2: mov var_24, ecx
  loc_004279A5: mov eax, 80020004h
  loc_004279AA: sub esp, 00000010h
  loc_004279AD: mov [ebx], ecx
  loc_004279AF: mov ecx, var_30
  loc_004279B2: mov edx, eax
  loc_004279B4: mov esi, [004300ECh]
  loc_004279BA: mov [ebx+00000004h], ecx
  loc_004279BD: mov ecx, esp
  loc_004279BF: mov edi, [esi]
  loc_004279C1: push esi
  loc_004279C2: mov [ebx+00000008h], eax
  loc_004279C5: mov eax, var_28
  loc_004279C8: mov [ebx+0000000Ch], eax
  loc_004279CB: mov eax, var_24
  loc_004279CE: mov [ecx], eax
  loc_004279D0: mov eax, var_20
  loc_004279D3: mov [ecx+00000004h], eax
  loc_004279D6: mov [ecx+00000008h], edx
  loc_004279D9: mov edx, var_18
  loc_004279DC: mov [ecx+0000000Ch], edx
  loc_004279DF: call [edi+000002B0h]
  loc_004279E5: test eax, eax
  loc_004279E7: fnclex
  loc_004279E9: jge 004279FDh
  loc_004279EB: push 000002B0h
  loc_004279F0: push 0040ECECh
  loc_004279F5: push esi
  loc_004279F6: push eax
  loc_004279F7: call [00401030h] ; __vbaHresultCheckObj
  loc_004279FD: mov eax, [00430164h]
  loc_00427A02: test eax, eax
  loc_00427A04: jnz 00427A16h
  loc_00427A06: push 00430164h
  loc_00427A0B: push 00408108h
  loc_00427A10: call [004010D4h] ; __vbaNew2
  loc_00427A16: mov esi, [00430164h]
  loc_00427A1C: push esi
  loc_00427A1D: mov eax, [esi]
  loc_00427A1F: call [eax+000002B4h]
  loc_00427A25: test eax, eax
  loc_00427A27: fnclex
  loc_00427A29: jge 00427A3Dh
  loc_00427A2B: push 000002B4h
  loc_00427A30: push 0040FF58h
  loc_00427A35: push esi
  loc_00427A36: push eax
  loc_00427A37: call [00401030h] ; __vbaHresultCheckObj
  loc_00427A3D: mov var_4, 00000000h
  loc_00427A44: mov eax, Me
  loc_00427A47: push eax
  loc_00427A48: mov ecx, [eax]
  loc_00427A4A: call [ecx+00000008h]
  loc_00427A4D: mov eax, var_4
  loc_00427A50: mov ecx, var_14
  loc_00427A53: pop edi
  loc_00427A54: pop esi
  loc_00427A55: mov fs:[00000000h], ecx
  loc_00427A5C: pop ebx
  loc_00427A5D: mov esp, ebp
  loc_00427A5F: pop ebp
  loc_00427A60: retn 0004h
End Sub

Private Sub mnuContext_Click() '4272E0
  loc_004272E0: push ebp
  loc_004272E1: mov ebp, esp
  loc_004272E3: sub esp, 0000000Ch
  loc_004272E6: push 00401926h ; __vbaExceptHandler
  loc_004272EB: mov eax, fs:[00000000h]
  loc_004272F1: push eax
  loc_004272F2: mov fs:[00000000h], esp
  loc_004272F9: sub esp, 0000007Ch
  loc_004272FC: push ebx
  loc_004272FD: push esi
  loc_004272FE: push edi
  loc_004272FF: mov var_C, esp
  loc_00427302: mov var_8, 00401608h
  loc_00427309: mov eax, Me
  loc_0042730C: mov ecx, eax
  loc_0042730E: and ecx, 00000001h
  loc_00427311: mov var_4, ecx
  loc_00427314: and al, FEh
  loc_00427316: push eax
  loc_00427317: mov Me, eax
  loc_0042731A: mov edx, [eax]
  loc_0042731C: call [edx+00000004h]
  loc_0042731F: mov ecx, 80020004h
  loc_00427324: xor esi, esi
  loc_00427326: mov var_50, ecx
  loc_00427329: mov eax, 0000000Ah
  loc_0042732E: mov var_40, ecx
  loc_00427331: mov ebx, 00000008h
  loc_00427336: mov var_48, esi
  loc_00427339: mov var_58, esi
  loc_0042733C: mov var_68, esi
  loc_0042733F: lea edx, var_68
  loc_00427342: lea ecx, var_38
  loc_00427345: mov var_18, esi
  loc_00427348: mov var_28, esi
  loc_0042734B: mov var_38, esi
  loc_0042734E: mov var_58, eax
  loc_00427351: mov var_48, eax
  loc_00427354: mov var_60, 00411750h ; "Context Instructions"
  loc_0042735B: mov var_68, ebx
  loc_0042735E: call [004010F4h] ; __vbaVarDup
  loc_00427364: mov edi, [0040102Ch] ; __vbaStrCat
  loc_0042736A: push 004113D8h ; "Hold mouse point over question to see question in context of the example."
  loc_0042736F: push 0040F720h ; vbCrLf
  loc_00427374: call edi
  loc_00427376: mov edx, eax
  loc_00427378: lea ecx, var_18
  loc_0042737B: call [004010FCh] ; __vbaStrMove
  loc_00427381: push eax
  loc_00427382: push 00411698h ; "Change the hypothesized value of the slope and notice the differences in the statements."
  loc_00427387: call edi
  loc_00427389: mov var_20, eax
  loc_0042738C: lea eax, var_58
  loc_0042738F: lea ecx, var_48
  loc_00427392: push eax
  loc_00427393: lea edx, var_38
  loc_00427396: push ecx
  loc_00427397: push edx
  loc_00427398: lea eax, var_28
  loc_0042739B: push esi
  loc_0042739C: push eax
  loc_0042739D: mov var_28, ebx
  loc_004273A0: call [00401038h] ; rtcMsgBox
  loc_004273A6: lea ecx, var_18
  loc_004273A9: call [00401110h] ; __vbaFreeStr
  loc_004273AF: lea ecx, var_58
  loc_004273B2: lea edx, var_48
  loc_004273B5: push ecx
  loc_004273B6: lea eax, var_38
  loc_004273B9: push edx
  loc_004273BA: lea ecx, var_28
  loc_004273BD: push eax
  loc_004273BE: push ecx
  loc_004273BF: push 00000004h
  loc_004273C1: call [00401018h] ; __vbaFreeVarList
  loc_004273C7: add esp, 00000014h
  loc_004273CA: mov var_4, esi
  loc_004273CD: push 004273FAh
  loc_004273D2: jmp 004273F9h
  loc_004273D4: lea ecx, var_18
  loc_004273D7: call [00401110h] ; __vbaFreeStr
  loc_004273DD: lea edx, var_58
  loc_004273E0: lea eax, var_48
  loc_004273E3: push edx
  loc_004273E4: lea ecx, var_38
  loc_004273E7: push eax
  loc_004273E8: lea edx, var_28
  loc_004273EB: push ecx
  loc_004273EC: push edx
  loc_004273ED: push 00000004h
  loc_004273EF: call [00401018h] ; __vbaFreeVarList
  loc_004273F5: add esp, 00000014h
  loc_004273F8: ret
  loc_004273F9: ret
  loc_004273FA: mov eax, Me
  loc_004273FD: push eax
  loc_004273FE: mov ecx, [eax]
  loc_00427400: call [ecx+00000008h]
  loc_00427403: mov eax, var_4
  loc_00427406: mov ecx, var_14
  loc_00427409: pop edi
  loc_0042740A: pop esi
  loc_0042740B: mov fs:[00000000h], ecx
  loc_00427412: pop ebx
  loc_00427413: mov esp, ebp
  loc_00427415: pop ebp
  loc_00427416: retn 0004h
End Sub

Private Sub mnuIntoduction_Click() '427420
  loc_00427420: push ebp
  loc_00427421: mov ebp, esp
  loc_00427423: sub esp, 0000000Ch
  loc_00427426: push 00401926h ; __vbaExceptHandler
  loc_0042742B: mov eax, fs:[00000000h]
  loc_00427431: push eax
  loc_00427432: mov fs:[00000000h], esp
  loc_00427439: sub esp, 00000030h
  loc_0042743C: push ebx
  loc_0042743D: push esi
  loc_0042743E: push edi
  loc_0042743F: mov var_C, esp
  loc_00427442: mov var_8, 00401618h
  loc_00427449: mov eax, Me
  loc_0042744C: mov ecx, eax
  loc_0042744E: and ecx, 00000001h
  loc_00427451: mov var_4, ecx
  loc_00427454: and al, FEh
  loc_00427456: push eax
  loc_00427457: mov Me, eax
  loc_0042745A: mov edx, [eax]
  loc_0042745C: call [edx+00000004h]
  loc_0042745F: mov eax, [00430088h]
  loc_00427464: test eax, eax
  loc_00427466: jnz 00427478h
  loc_00427468: push 00430088h
  loc_0042746D: push 004058C0h
  loc_00427472: call [004010D4h] ; __vbaNew2
  loc_00427478: sub esp, 00000010h
  loc_0042747B: mov ecx, 0000000Ah
  loc_00427480: mov ebx, esp
  loc_00427482: mov var_24, ecx
  loc_00427485: mov eax, 80020004h
  loc_0042748A: sub esp, 00000010h
  loc_0042748D: mov [ebx], ecx
  loc_0042748F: mov ecx, var_30
  loc_00427492: mov edx, eax
  loc_00427494: mov esi, [00430088h]
  loc_0042749A: mov [ebx+00000004h], ecx
  loc_0042749D: mov ecx, esp
  loc_0042749F: mov edi, [esi]
  loc_004274A1: push esi
  loc_004274A2: mov [ebx+00000008h], eax
  loc_004274A5: mov eax, var_28
  loc_004274A8: mov [ebx+0000000Ch], eax
  loc_004274AB: mov eax, var_24
  loc_004274AE: mov [ecx], eax
  loc_004274B0: mov eax, var_20
  loc_004274B3: mov [ecx+00000004h], eax
  loc_004274B6: mov [ecx+00000008h], edx
  loc_004274B9: mov edx, var_18
  loc_004274BC: mov [ecx+0000000Ch], edx
  loc_004274BF: call [edi+000002B0h]
  loc_004274C5: test eax, eax
  loc_004274C7: fnclex
  loc_004274C9: jge 004274DDh
  loc_004274CB: push 000002B0h
  loc_004274D0: push 0040DB64h
  loc_004274D5: push esi
  loc_004274D6: push eax
  loc_004274D7: call [00401030h] ; __vbaHresultCheckObj
  loc_004274DD: mov eax, [00430164h]
  loc_004274E2: test eax, eax
  loc_004274E4: jnz 004274F6h
  loc_004274E6: push 00430164h
  loc_004274EB: push 00408108h
  loc_004274F0: call [004010D4h] ; __vbaNew2
  loc_004274F6: mov esi, [00430164h]
  loc_004274FC: push esi
  loc_004274FD: mov eax, [esi]
  loc_004274FF: call [eax+000002B4h]
  loc_00427505: test eax, eax
  loc_00427507: fnclex
  loc_00427509: jge 0042751Dh
  loc_0042750B: push 000002B4h
  loc_00427510: push 0040FF58h
  loc_00427515: push esi
  loc_00427516: push eax
  loc_00427517: call [00401030h] ; __vbaHresultCheckObj
  loc_0042751D: mov var_4, 00000000h
  loc_00427524: mov eax, Me
  loc_00427527: push eax
  loc_00427528: mov ecx, [eax]
  loc_0042752A: call [ecx+00000008h]
  loc_0042752D: mov eax, var_4
  loc_00427530: mov ecx, var_14
  loc_00427533: pop edi
  loc_00427534: pop esi
  loc_00427535: mov fs:[00000000h], ecx
  loc_0042753C: pop ebx
  loc_0042753D: mov esp, ebp
  loc_0042753F: pop ebp
  loc_00427540: retn 0004h
End Sub

Private Sub mnuUnits_Click() '427A70
  loc_00427A70: push ebp
  loc_00427A71: mov ebp, esp
  loc_00427A73: sub esp, 0000000Ch
  loc_00427A76: push 00401926h ; __vbaExceptHandler
  loc_00427A7B: mov eax, fs:[00000000h]
  loc_00427A81: push eax
  loc_00427A82: mov fs:[00000000h], esp
  loc_00427A89: sub esp, 00000030h
  loc_00427A8C: push ebx
  loc_00427A8D: push esi
  loc_00427A8E: push edi
  loc_00427A8F: mov var_C, esp
  loc_00427A92: mov var_8, 00401650h
  loc_00427A99: mov eax, Me
  loc_00427A9C: mov ecx, eax
  loc_00427A9E: and ecx, 00000001h
  loc_00427AA1: mov var_4, ecx
  loc_00427AA4: and al, FEh
  loc_00427AA6: push eax
  loc_00427AA7: mov Me, eax
  loc_00427AAA: mov edx, [eax]
  loc_00427AAC: call [edx+00000004h]
  loc_00427AAF: mov eax, [004301A0h]
  loc_00427AB4: test eax, eax
  loc_00427AB6: jnz 00427AC8h
  loc_00427AB8: push 004301A0h
  loc_00427ABD: push 004033C0h
  loc_00427AC2: call [004010D4h] ; __vbaNew2
  loc_00427AC8: sub esp, 00000010h
  loc_00427ACB: mov ecx, 0000000Ah
  loc_00427AD0: mov ebx, esp
  loc_00427AD2: mov var_24, ecx
  loc_00427AD5: mov eax, 80020004h
  loc_00427ADA: sub esp, 00000010h
  loc_00427ADD: mov [ebx], ecx
  loc_00427ADF: mov ecx, var_30
  loc_00427AE2: mov edx, eax
  loc_00427AE4: mov esi, [004301A0h]
  loc_00427AEA: mov [ebx+00000004h], ecx
  loc_00427AED: mov ecx, esp
  loc_00427AEF: mov edi, [esi]
  loc_00427AF1: push esi
  loc_00427AF2: mov [ebx+00000008h], eax
  loc_00427AF5: mov eax, var_28
  loc_00427AF8: mov [ebx+0000000Ch], eax
  loc_00427AFB: mov eax, var_24
  loc_00427AFE: mov [ecx], eax
  loc_00427B00: mov eax, var_20
  loc_00427B03: mov [ecx+00000004h], eax
  loc_00427B06: mov [ecx+00000008h], edx
  loc_00427B09: mov edx, var_18
  loc_00427B0C: mov [ecx+0000000Ch], edx
  loc_00427B0F: call [edi+000002B0h]
  loc_00427B15: test eax, eax
  loc_00427B17: fnclex
  loc_00427B19: jge 00427B2Dh
  loc_00427B1B: push 000002B0h
  loc_00427B20: push 00410454h
  loc_00427B25: push esi
  loc_00427B26: push eax
  loc_00427B27: call [00401030h] ; __vbaHresultCheckObj
  loc_00427B2D: mov var_4, 00000000h
  loc_00427B34: mov eax, Me
  loc_00427B37: push eax
  loc_00427B38: mov ecx, [eax]
  loc_00427B3A: call [ecx+00000008h]
  loc_00427B3D: mov eax, var_4
  loc_00427B40: mov ecx, var_14
  loc_00427B43: pop edi
  loc_00427B44: pop esi
  loc_00427B45: mov fs:[00000000h], ecx
  loc_00427B4C: pop ebx
  loc_00427B4D: mov esp, ebp
  loc_00427B4F: pop ebp
  loc_00427B50: retn 0004h
End Sub

Private Sub cmdGoEstimation_Click() '425010
  loc_00425010: push ebp
  loc_00425011: mov ebp, esp
  loc_00425013: sub esp, 0000000Ch
  loc_00425016: push 00401926h ; __vbaExceptHandler
  loc_0042501B: mov eax, fs:[00000000h]
  loc_00425021: push eax
  loc_00425022: mov fs:[00000000h], esp
  loc_00425029: sub esp, 00000030h
  loc_0042502C: push ebx
  loc_0042502D: push esi
  loc_0042502E: push edi
  loc_0042502F: mov var_C, esp
  loc_00425032: mov var_8, 004015A0h
  loc_00425039: mov eax, Me
  loc_0042503C: mov ecx, eax
  loc_0042503E: and ecx, 00000001h
  loc_00425041: mov var_4, ecx
  loc_00425044: and al, FEh
  loc_00425046: push eax
  loc_00425047: mov Me, eax
  loc_0042504A: mov edx, [eax]
  loc_0042504C: call [edx+00000004h]
  loc_0042504F: mov eax, [004300B0h]
  loc_00425054: test eax, eax
  loc_00425056: jnz 00425068h
  loc_00425058: push 004300B0h
  loc_0042505D: push 00407528h
  loc_00425062: call [004010D4h] ; __vbaNew2
  loc_00425068: sub esp, 00000010h
  loc_0042506B: mov ecx, 0000000Ah
  loc_00425070: mov ebx, esp
  loc_00425072: mov var_24, ecx
  loc_00425075: mov eax, 80020004h
  loc_0042507A: sub esp, 00000010h
  loc_0042507D: mov [ebx], ecx
  loc_0042507F: mov ecx, var_30
  loc_00425082: mov edx, eax
  loc_00425084: mov esi, [004300B0h]
  loc_0042508A: mov [ebx+00000004h], ecx
  loc_0042508D: mov ecx, esp
  loc_0042508F: mov edi, [esi]
  loc_00425091: push esi
  loc_00425092: mov [ebx+00000008h], eax
  loc_00425095: mov eax, var_28
  loc_00425098: mov [ebx+0000000Ch], eax
  loc_0042509B: mov eax, var_24
  loc_0042509E: mov [ecx], eax
  loc_004250A0: mov eax, var_20
  loc_004250A3: mov [ecx+00000004h], eax
  loc_004250A6: mov [ecx+00000008h], edx
  loc_004250A9: mov edx, var_18
  loc_004250AC: mov [ecx+0000000Ch], edx
  loc_004250AF: call [edi+000002B0h]
  loc_004250B5: test eax, eax
  loc_004250B7: fnclex
  loc_004250B9: jge 004250CDh
  loc_004250BB: push 000002B0h
  loc_004250C0: push 0040DE98h
  loc_004250C5: push esi
  loc_004250C6: push eax
  loc_004250C7: call [00401030h] ; __vbaHresultCheckObj
  loc_004250CD: mov eax, [00430164h]
  loc_004250D2: test eax, eax
  loc_004250D4: jnz 004250E6h
  loc_004250D6: push 00430164h
  loc_004250DB: push 00408108h
  loc_004250E0: call [004010D4h] ; __vbaNew2
  loc_004250E6: mov esi, [00430164h]
  loc_004250EC: push esi
  loc_004250ED: mov eax, [esi]
  loc_004250EF: call [eax+000002B4h]
  loc_004250F5: test eax, eax
  loc_004250F7: fnclex
  loc_004250F9: jge 0042510Dh
  loc_004250FB: push 000002B4h
  loc_00425100: push 0040FF58h
  loc_00425105: push esi
  loc_00425106: push eax
  loc_00425107: call [00401030h] ; __vbaHresultCheckObj
  loc_0042510D: mov var_4, 00000000h
  loc_00425114: mov eax, Me
  loc_00425117: push eax
  loc_00425118: mov ecx, [eax]
  loc_0042511A: call [ecx+00000008h]
  loc_0042511D: mov eax, var_4
  loc_00425120: mov ecx, var_14
  loc_00425123: pop edi
  loc_00425124: pop esi
  loc_00425125: mov fs:[00000000h], ecx
  loc_0042512C: pop ebx
  loc_0042512D: mov esp, ebp
  loc_0042512F: pop ebp
  loc_00425130: retn 0004h
End Sub

Private Sub mnuChangeHypValue_Click() '4253A0
  loc_004253A0: push ebp
  loc_004253A1: mov ebp, esp
  loc_004253A3: sub esp, 0000000Ch
  loc_004253A6: push 00401926h ; __vbaExceptHandler
  loc_004253AB: mov eax, fs:[00000000h]
  loc_004253B1: push eax
  loc_004253B2: mov fs:[00000000h], esp
  loc_004253B9: sub esp, 00000030h
  loc_004253BC: push ebx
  loc_004253BD: push esi
  loc_004253BE: push edi
  loc_004253BF: mov var_C, esp
  loc_004253C2: mov var_8, 004015B8h
  loc_004253C9: mov eax, Me
  loc_004253CC: mov ecx, eax
  loc_004253CE: and ecx, 00000001h
  loc_004253D1: mov var_4, ecx
  loc_004253D4: and al, FEh
  loc_004253D6: push eax
  loc_004253D7: mov Me, eax
  loc_004253DA: mov edx, [eax]
  loc_004253DC: call [edx+00000004h]
  loc_004253DF: mov eax, [0043013Ch]
  loc_004253E4: test eax, eax
  loc_004253E6: jnz 004253F8h
  loc_004253E8: push 0043013Ch
  loc_004253ED: push 0040204Ch
  loc_004253F2: call [004010D4h] ; __vbaNew2
  loc_004253F8: sub esp, 00000010h
  loc_004253FB: mov ecx, 0000000Ah
  loc_00425400: mov ebx, esp
  loc_00425402: mov var_24, ecx
  loc_00425405: mov eax, 80020004h
  loc_0042540A: sub esp, 00000010h
  loc_0042540D: mov [ebx], ecx
  loc_0042540F: mov ecx, var_30
  loc_00425412: mov edx, eax
  loc_00425414: mov esi, [0043013Ch]
  loc_0042541A: mov [ebx+00000004h], ecx
  loc_0042541D: mov ecx, esp
  loc_0042541F: mov edi, [esi]
  loc_00425421: push esi
  loc_00425422: mov [ebx+00000008h], eax
  loc_00425425: mov eax, var_28
  loc_00425428: mov [ebx+0000000Ch], eax
  loc_0042542B: mov eax, var_24
  loc_0042542E: mov [ecx], eax
  loc_00425430: mov eax, var_20
  loc_00425433: mov [ecx+00000004h], eax
  loc_00425436: mov [ecx+00000008h], edx
  loc_00425439: mov edx, var_18
  loc_0042543C: mov [ecx+0000000Ch], edx
  loc_0042543F: call [edi+000002B0h]
  loc_00425445: test eax, eax
  loc_00425447: fnclex
  loc_00425449: jge 0042545Dh
  loc_0042544B: push 000002B0h
  loc_00425450: push 0040FC38h
  loc_00425455: push esi
  loc_00425456: push eax
  loc_00425457: call [00401030h] ; __vbaHresultCheckObj
  loc_0042545D: mov var_4, 00000000h
  loc_00425464: mov eax, Me
  loc_00425467: push eax
  loc_00425468: mov ecx, [eax]
  loc_0042546A: call [ecx+00000008h]
  loc_0042546D: mov eax, var_4
  loc_00425470: mov ecx, var_14
  loc_00425473: pop edi
  loc_00425474: pop esi
  loc_00425475: mov fs:[00000000h], ecx
  loc_0042547C: pop ebx
  loc_0042547D: mov esp, ebp
  loc_0042547F: pop ebp
  loc_00425480: retn 0004h
End Sub

Private Sub mnuExit_Click() '4255C0
  loc_004255C0: push ebp
  loc_004255C1: mov ebp, esp
  loc_004255C3: sub esp, 0000000Ch
  loc_004255C6: push 00401926h ; __vbaExceptHandler
  loc_004255CB: mov eax, fs:[00000000h]
  loc_004255D1: push eax
  loc_004255D2: mov fs:[00000000h], esp
  loc_004255D9: sub esp, 00000018h
  loc_004255DC: push ebx
  loc_004255DD: push esi
  loc_004255DE: push edi
  loc_004255DF: mov var_C, esp
  loc_004255E2: mov var_8, 004015C8h
  loc_004255E9: mov edi, Me
  loc_004255EC: mov eax, edi
  loc_004255EE: and eax, 00000001h
  loc_004255F1: mov var_4, eax
  loc_004255F4: and edi, FFFFFFFEh
  loc_004255F7: push edi
  loc_004255F8: mov Me, edi
  loc_004255FB: mov ecx, [edi]
  loc_004255FD: call [ecx+00000004h]
  loc_00425600: mov eax, [00430A24h]
  loc_00425605: xor ebx, ebx
  loc_00425607: cmp eax, ebx
  loc_00425609: mov var_18, ebx
  loc_0042560C: jnz 0042561Eh
  loc_0042560E: push 00430A24h
  loc_00425613: push 0040EC7Ch
  loc_00425618: call [004010D4h] ; __vbaNew2
  loc_0042561E: mov esi, [00430A24h]
  loc_00425624: lea eax, var_18
  loc_00425627: push edi
  loc_00425628: push eax
  loc_00425629: mov edx, [esi]
  loc_0042562B: mov var_2C, edx
  loc_0042562E: call [00401044h] ; __vbaObjSetAddref
  loc_00425634: mov ecx, var_2C
  loc_00425637: push eax
  loc_00425638: push esi
  loc_00425639: call [ecx+00000010h]
  loc_0042563C: cmp eax, ebx
  loc_0042563E: fnclex
  loc_00425640: jge 00425651h
  loc_00425642: push 00000010h
  loc_00425644: push 0040EC6Ch
  loc_00425649: push esi
  loc_0042564A: push eax
  loc_0042564B: call [00401030h] ; __vbaHresultCheckObj
  loc_00425651: lea ecx, var_18
  loc_00425654: call [00401114h] ; __vbaFreeObj
  loc_0042565A: mov var_4, ebx
  loc_0042565D: push 0042566Fh
  loc_00425662: jmp 0042566Eh
  loc_00425664: lea ecx, var_18
  loc_00425667: call [00401114h] ; __vbaFreeObj
  loc_0042566D: ret
  loc_0042566E: ret
  loc_0042566F: mov eax, Me
  loc_00425672: push eax
  loc_00425673: mov edx, [eax]
  loc_00425675: call [edx+00000008h]
  loc_00425678: mov eax, var_4
  loc_0042567B: mov ecx, var_14
  loc_0042567E: pop edi
  loc_0042567F: pop esi
  loc_00425680: mov fs:[00000000h], ecx
  loc_00425687: pop ebx
  loc_00425688: mov esp, ebp
  loc_0042568A: pop ebp
  loc_0042568B: retn 0004h
End Sub

Private Sub Form_Load() '425720
  loc_00425720: push ebp
  loc_00425721: mov ebp, esp
  loc_00425723: sub esp, 0000000Ch
  loc_00425726: push 00401926h ; __vbaExceptHandler
  loc_0042572B: mov eax, fs:[00000000h]
  loc_00425731: push eax
  loc_00425732: mov fs:[00000000h], esp
  loc_00425739: sub esp, 0000000Ch
  loc_0042573C: push ebx
  loc_0042573D: push esi
  loc_0042573E: push edi
  loc_0042573F: mov var_C, esp
  loc_00425742: mov var_8, 004015E0h
  loc_00425749: mov esi, Me
  loc_0042574C: mov eax, esi
  loc_0042574E: and eax, 00000001h
  loc_00425751: mov var_4, eax
  loc_00425754: and esi, FFFFFFFEh
  loc_00425757: push esi
  loc_00425758: mov Me, esi
  loc_0042575B: mov ecx, [esi]
  loc_0042575D: call [ecx+00000004h]
  loc_00425760: mov edx, [esi]
  loc_00425762: push esi
  loc_00425763: call [edx+000006F8h]
  loc_00425769: test eax, eax
  loc_0042576B: jge 0042577Fh
  loc_0042576D: push 000006F8h
  loc_00425772: push 0040FF78h
  loc_00425777: push esi
  loc_00425778: push eax
  loc_00425779: call [00401030h] ; __vbaHresultCheckObj
  loc_0042577F: mov var_4, 00000000h
  loc_00425786: mov eax, Me
  loc_00425789: push eax
  loc_0042578A: mov ecx, [eax]
  loc_0042578C: call [ecx+00000008h]
  loc_0042578F: mov eax, var_4
  loc_00425792: mov ecx, var_14
  loc_00425795: pop edi
  loc_00425796: pop esi
  loc_00425797: mov fs:[00000000h], ecx
  loc_0042579E: pop ebx
  loc_0042579F: mov esp, ebp
  loc_004257A1: pop ebp
  loc_004257A2: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '427670
  loc_00427670: push ebp
  loc_00427671: mov ebp, esp
  loc_00427673: sub esp, 0000000Ch
  loc_00427676: push 00401926h ; __vbaExceptHandler
  loc_0042767B: mov eax, fs:[00000000h]
  loc_00427681: push eax
  loc_00427682: mov fs:[00000000h], esp
  loc_00427689: sub esp, 0000009Ch
  loc_0042768F: push ebx
  loc_00427690: push esi
  loc_00427691: push edi
  loc_00427692: mov var_C, esp
  loc_00427695: mov var_8, 00401630h
  loc_0042769C: mov edi, Me
  loc_0042769F: mov eax, edi
  loc_004276A1: and eax, 00000001h
  loc_004276A4: mov var_4, eax
  loc_004276A7: and edi, FFFFFFFEh
  loc_004276AA: push edi
  loc_004276AB: mov Me, edi
  loc_004276AE: mov ecx, [edi]
  loc_004276B0: call [ecx+00000004h]
  loc_004276B3: mov ebx, [004010F4h] ; __vbaVarDup
  loc_004276B9: mov ecx, 80020004h
  loc_004276BE: xor esi, esi
  loc_004276C0: mov var_54, ecx
  loc_004276C3: mov eax, 0000000Ah
  loc_004276C8: mov var_44, ecx
  loc_004276CB: mov var_4C, esi
  loc_004276CE: mov var_5C, esi
  loc_004276D1: mov var_7C, esi
  loc_004276D4: lea edx, var_7C
  loc_004276D7: lea ecx, var_3C
  loc_004276DA: mov var_1C, esi
  loc_004276DD: mov var_2C, esi
  loc_004276E0: mov var_3C, esi
  loc_004276E3: mov var_6C, esi
  loc_004276E6: mov var_5C, eax
  loc_004276E9: mov var_4C, eax
  loc_004276EC: mov var_74, 0040ECD4h ; "Exit Check"
  loc_004276F3: mov var_7C, 00000008h
  loc_004276FA: call ebx
  loc_004276FC: lea edx, var_6C
  loc_004276FF: lea ecx, var_2C
  loc_00427702: mov var_64, 0040EC90h ; "Are you sure you want to exit?"
  loc_00427709: mov var_6C, 00000008h
  loc_00427710: call ebx
  loc_00427712: lea edx, var_5C
  loc_00427715: lea eax, var_4C
  loc_00427718: push edx
  loc_00427719: lea ecx, var_3C
  loc_0042771C: push eax
  loc_0042771D: push ecx
  loc_0042771E: lea edx, var_2C
  loc_00427721: push 00000004h
  loc_00427723: push edx
  loc_00427724: call [00401038h] ; rtcMsgBox
  loc_0042772A: mov ecx, eax
  loc_0042772C: call [00401078h] ; __vbaI2I4
  loc_00427732: mov ebx, eax
  loc_00427734: lea eax, var_5C
  loc_00427737: lea ecx, var_4C
  loc_0042773A: push eax
  loc_0042773B: lea edx, var_3C
  loc_0042773E: push ecx
  loc_0042773F: lea eax, var_2C
  loc_00427742: push edx
  loc_00427743: push eax
  loc_00427744: push 00000004h
  loc_00427746: call [00401018h] ; __vbaFreeVarList
  loc_0042774C: add esp, 00000014h
  loc_0042774F: cmp bx, 0007h
  loc_00427753: jnz 0042775Fh
  loc_00427755: mov ecx, Cancel
  loc_00427758: mov [ecx], FFFFFFh
  loc_0042775D: jmp 004277BFh
  loc_0042775F: cmp [00430A24h], esi
  loc_00427765: jnz 00427777h
  loc_00427767: push 00430A24h
  loc_0042776C: push 0040EC7Ch
  loc_00427771: call [004010D4h] ; __vbaNew2
  loc_00427777: mov ebx, [00430A24h]
  loc_0042777D: lea eax, var_1C
  loc_00427780: push edi
  loc_00427781: push eax
  loc_00427782: mov edx, [ebx]
  loc_00427784: mov var_B0, edx
  loc_0042778A: call [00401044h] ; __vbaObjSetAddref
  loc_00427790: mov ecx, var_B0
  loc_00427796: push eax
  loc_00427797: push ebx
  loc_00427798: call [ecx+00000010h]
  loc_0042779B: cmp eax, esi
  loc_0042779D: fnclex
  loc_0042779F: jge 004277B0h
  loc_004277A1: push 00000010h
  loc_004277A3: push 0040EC6Ch
  loc_004277A8: push ebx
  loc_004277A9: push eax
  loc_004277AA: call [00401030h] ; __vbaHresultCheckObj
  loc_004277B0: lea ecx, var_1C
  loc_004277B3: call [00401114h] ; __vbaFreeObj
  loc_004277B9: call [0040101Ch] ; __vbaEnd
  loc_004277BF: mov var_4, esi
  loc_004277C2: push 004277EFh
  loc_004277C7: jmp 004277EEh
  loc_004277C9: lea ecx, var_1C
  loc_004277CC: call [00401114h] ; __vbaFreeObj
  loc_004277D2: lea edx, var_5C
  loc_004277D5: lea eax, var_4C
  loc_004277D8: push edx
  loc_004277D9: lea ecx, var_3C
  loc_004277DC: push eax
  loc_004277DD: lea edx, var_2C
  loc_004277E0: push ecx
  loc_004277E1: push edx
  loc_004277E2: push 00000004h
  loc_004277E4: call [00401018h] ; __vbaFreeVarList
  loc_004277EA: add esp, 00000014h
  loc_004277ED: ret
  loc_004277EE: ret
  loc_004277EF: mov eax, Me
  loc_004277F2: push eax
  loc_004277F3: mov ecx, [eax]
  loc_004277F5: call [ecx+00000008h]
  loc_004277F8: mov eax, var_4
  loc_004277FB: mov ecx, var_14
  loc_004277FE: pop edi
  loc_004277FF: pop esi
  loc_00427800: mov fs:[00000000h], ecx
  loc_00427807: pop ebx
  loc_00427808: mov esp, ebp
  loc_0042780A: pop ebp
  loc_0042780B: retn 000Ch
End Sub

Private Sub Form_Activate() '425690
  loc_00425690: push ebp
  loc_00425691: mov ebp, esp
  loc_00425693: sub esp, 0000000Ch
  loc_00425696: push 00401926h ; __vbaExceptHandler
  loc_0042569B: mov eax, fs:[00000000h]
  loc_004256A1: push eax
  loc_004256A2: mov fs:[00000000h], esp
  loc_004256A9: sub esp, 0000000Ch
  loc_004256AC: push ebx
  loc_004256AD: push esi
  loc_004256AE: push edi
  loc_004256AF: mov var_C, esp
  loc_004256B2: mov var_8, 004015D8h
  loc_004256B9: mov esi, Me
  loc_004256BC: mov eax, esi
  loc_004256BE: and eax, 00000001h
  loc_004256C1: mov var_4, eax
  loc_004256C4: and esi, FFFFFFFEh
  loc_004256C7: push esi
  loc_004256C8: mov Me, esi
  loc_004256CB: mov ecx, [esi]
  loc_004256CD: call [ecx+00000004h]
  loc_004256D0: mov edx, [esi]
  loc_004256D2: push esi
  loc_004256D3: call [edx+000006F8h]
  loc_004256D9: test eax, eax
  loc_004256DB: jge 004256EFh
  loc_004256DD: push 000006F8h
  loc_004256E2: push 0040FF78h
  loc_004256E7: push esi
  loc_004256E8: push eax
  loc_004256E9: call [00401030h] ; __vbaHresultCheckObj
  loc_004256EF: mov var_4, 00000000h
  loc_004256F6: mov eax, Me
  loc_004256F9: push eax
  loc_004256FA: mov ecx, [eax]
  loc_004256FC: call [ecx+00000008h]
  loc_004256FF: mov eax, var_4
  loc_00425702: mov ecx, var_14
  loc_00425705: pop edi
  loc_00425706: pop esi
  loc_00425707: mov fs:[00000000h], ecx
  loc_0042570E: pop ebx
  loc_0042570F: mov esp, ebp
  loc_00425711: pop ebp
  loc_00425712: retn 0004h
End Sub

Private Sub cmdGoSlope_Click() '425140
  loc_00425140: push ebp
  loc_00425141: mov ebp, esp
  loc_00425143: sub esp, 0000000Ch
  loc_00425146: push 00401926h ; __vbaExceptHandler
  loc_0042514B: mov eax, fs:[00000000h]
  loc_00425151: push eax
  loc_00425152: mov fs:[00000000h], esp
  loc_00425159: sub esp, 00000030h
  loc_0042515C: push ebx
  loc_0042515D: push esi
  loc_0042515E: push edi
  loc_0042515F: mov var_C, esp
  loc_00425162: mov var_8, 004015A8h
  loc_00425169: mov eax, Me
  loc_0042516C: mov ecx, eax
  loc_0042516E: and ecx, 00000001h
  loc_00425171: mov var_4, ecx
  loc_00425174: and al, FEh
  loc_00425176: push eax
  loc_00425177: mov Me, eax
  loc_0042517A: mov edx, [eax]
  loc_0042517C: call [edx+00000004h]
  loc_0042517F: mov eax, [00430164h]
  loc_00425184: test eax, eax
  loc_00425186: jnz 00425198h
  loc_00425188: push 00430164h
  loc_0042518D: push 00408108h
  loc_00425192: call [004010D4h] ; __vbaNew2
  loc_00425198: mov esi, [00430164h]
  loc_0042519E: push esi
  loc_0042519F: mov eax, [esi]
  loc_004251A1: call [eax+000002B4h]
  loc_004251A7: test eax, eax
  loc_004251A9: fnclex
  loc_004251AB: jge 004251BFh
  loc_004251AD: push 000002B4h
  loc_004251B2: push 0040FF58h
  loc_004251B7: push esi
  loc_004251B8: push eax
  loc_004251B9: call [00401030h] ; __vbaHresultCheckObj
  loc_004251BF: mov eax, [004300C4h]
  loc_004251C4: test eax, eax
  loc_004251C6: jnz 004251D8h
  loc_004251C8: push 004300C4h
  loc_004251CD: push 00409228h
  loc_004251D2: call [004010D4h] ; __vbaNew2
  loc_004251D8: sub esp, 00000010h
  loc_004251DB: mov ecx, 0000000Ah
  loc_004251E0: mov ebx, esp
  loc_004251E2: mov var_24, ecx
  loc_004251E5: mov eax, 80020004h
  loc_004251EA: sub esp, 00000010h
  loc_004251ED: mov [ebx], ecx
  loc_004251EF: mov ecx, var_30
  loc_004251F2: mov edx, eax
  loc_004251F4: mov esi, [004300C4h]
  loc_004251FA: mov [ebx+00000004h], ecx
  loc_004251FD: mov ecx, esp
  loc_004251FF: mov edi, [esi]
  loc_00425201: push esi
  loc_00425202: mov [ebx+00000008h], eax
  loc_00425205: mov eax, var_28
  loc_00425208: mov [ebx+0000000Ch], eax
  loc_0042520B: mov eax, var_24
  loc_0042520E: mov [ecx], eax
  loc_00425210: mov eax, var_20
  loc_00425213: mov [ecx+00000004h], eax
  loc_00425216: mov [ecx+00000008h], edx
  loc_00425219: mov edx, var_18
  loc_0042521C: mov [ecx+0000000Ch], edx
  loc_0042521F: call [edi+000002B0h]
  loc_00425225: test eax, eax
  loc_00425227: fnclex
  loc_00425229: jge 0042523Dh
  loc_0042522B: push 000002B0h
  loc_00425230: push 0040E0ECh
  loc_00425235: push esi
  loc_00425236: push eax
  loc_00425237: call [00401030h] ; __vbaHresultCheckObj
  loc_0042523D: mov var_4, 00000000h
  loc_00425244: mov eax, Me
  loc_00425247: push eax
  loc_00425248: mov ecx, [eax]
  loc_0042524A: call [ecx+00000008h]
  loc_0042524D: mov eax, var_4
  loc_00425250: mov ecx, var_14
  loc_00425253: pop edi
  loc_00425254: pop esi
  loc_00425255: mov fs:[00000000h], ecx
  loc_0042525C: pop ebx
  loc_0042525D: mov esp, ebp
  loc_0042525F: pop ebp
  loc_00425260: retn 0004h
End Sub

Private Sub mnuSlope_Click() '427810
  loc_00427810: push ebp
  loc_00427811: mov ebp, esp
  loc_00427813: sub esp, 0000000Ch
  loc_00427816: push 00401926h ; __vbaExceptHandler
  loc_0042781B: mov eax, fs:[00000000h]
  loc_00427821: push eax
  loc_00427822: mov fs:[00000000h], esp
  loc_00427829: sub esp, 00000030h
  loc_0042782C: push ebx
  loc_0042782D: push esi
  loc_0042782E: push edi
  loc_0042782F: mov var_C, esp
  loc_00427832: mov var_8, 00401640h
  loc_00427839: mov eax, Me
  loc_0042783C: mov ecx, eax
  loc_0042783E: and ecx, 00000001h
  loc_00427841: mov var_4, ecx
  loc_00427844: and al, FEh
  loc_00427846: push eax
  loc_00427847: mov Me, eax
  loc_0042784A: mov edx, [eax]
  loc_0042784C: call [edx+00000004h]
  loc_0042784F: mov eax, [004300C4h]
  loc_00427854: test eax, eax
  loc_00427856: jnz 00427868h
  loc_00427858: push 004300C4h
  loc_0042785D: push 00409228h
  loc_00427862: call [004010D4h] ; __vbaNew2
  loc_00427868: sub esp, 00000010h
  loc_0042786B: mov ecx, 0000000Ah
  loc_00427870: mov ebx, esp
  loc_00427872: mov var_24, ecx
  loc_00427875: mov eax, 80020004h
  loc_0042787A: sub esp, 00000010h
  loc_0042787D: mov [ebx], ecx
  loc_0042787F: mov ecx, var_30
  loc_00427882: mov edx, eax
  loc_00427884: mov esi, [004300C4h]
  loc_0042788A: mov [ebx+00000004h], ecx
  loc_0042788D: mov ecx, esp
  loc_0042788F: mov edi, [esi]
  loc_00427891: push esi
  loc_00427892: mov [ebx+00000008h], eax
  loc_00427895: mov eax, var_28
  loc_00427898: mov [ebx+0000000Ch], eax
  loc_0042789B: mov eax, var_24
  loc_0042789E: mov [ecx], eax
  loc_004278A0: mov eax, var_20
  loc_004278A3: mov [ecx+00000004h], eax
  loc_004278A6: mov [ecx+00000008h], edx
  loc_004278A9: mov edx, var_18
  loc_004278AC: mov [ecx+0000000Ch], edx
  loc_004278AF: call [edi+000002B0h]
  loc_004278B5: test eax, eax
  loc_004278B7: fnclex
  loc_004278B9: jge 004278CDh
  loc_004278BB: push 000002B0h
  loc_004278C0: push 0040E0ECh
  loc_004278C5: push esi
  loc_004278C6: push eax
  loc_004278C7: call [00401030h] ; __vbaHresultCheckObj
  loc_004278CD: mov eax, [00430164h]
  loc_004278D2: test eax, eax
  loc_004278D4: jnz 004278E6h
  loc_004278D6: push 00430164h
  loc_004278DB: push 00408108h
  loc_004278E0: call [004010D4h] ; __vbaNew2
  loc_004278E6: mov esi, [00430164h]
  loc_004278EC: push esi
  loc_004278ED: mov eax, [esi]
  loc_004278EF: call [eax+000002B4h]
  loc_004278F5: test eax, eax
  loc_004278F7: fnclex
  loc_004278F9: jge 0042790Dh
  loc_004278FB: push 000002B4h
  loc_00427900: push 0040FF58h
  loc_00427905: push esi
  loc_00427906: push eax
  loc_00427907: call [00401030h] ; __vbaHresultCheckObj
  loc_0042790D: mov var_4, 00000000h
  loc_00427914: mov eax, Me
  loc_00427917: push eax
  loc_00427918: mov ecx, [eax]
  loc_0042791A: call [ecx+00000008h]
  loc_0042791D: mov eax, var_4
  loc_00427920: mov ecx, var_14
  loc_00427923: pop edi
  loc_00427924: pop esi
  loc_00427925: mov fs:[00000000h], ecx
  loc_0042792C: pop ebx
  loc_0042792D: mov esp, ebp
  loc_0042792F: pop ebp
  loc_00427930: retn 0004h
End Sub

Public Sub changelabels() '4257B0
  loc_004257B0: push ebp
  loc_004257B1: mov ebp, esp
  loc_004257B3: sub esp, 0000000Ch
  loc_004257B6: push 00401926h ; __vbaExceptHandler
  loc_004257BB: mov eax, fs:[00000000h]
  loc_004257C1: push eax
  loc_004257C2: mov fs:[00000000h], esp
  loc_004257C9: sub esp, 000000A0h
  loc_004257CF: push ebx
  loc_004257D0: push esi
  loc_004257D1: push edi
  loc_004257D2: mov var_C, esp
  loc_004257D5: mov var_8, 004015E8h
  loc_004257DC: xor esi, esi
  loc_004257DE: mov var_4, esi
  loc_004257E1: mov edi, Me
  loc_004257E4: push edi
  loc_004257E5: mov eax, [edi]
  loc_004257E7: call [eax+00000004h]
  loc_004257EA: mov ecx, [edi]
  loc_004257EC: push edi
  loc_004257ED: mov var_18, esi
  loc_004257F0: mov var_1C, esi
  loc_004257F3: mov var_20, esi
  loc_004257F6: mov var_24, esi
  loc_004257F9: mov var_28, esi
  loc_004257FC: mov var_2C, esi
  loc_004257FF: mov var_30, esi
  loc_00425802: mov var_34, esi
  loc_00425805: mov var_38, esi
  loc_00425808: mov var_3C, esi
  loc_0042580B: mov var_40, esi
  loc_0042580E: mov var_44, esi
  loc_00425811: call [ecx+0000030Ch]
  loc_00425817: lea edx, var_44
  loc_0042581A: push eax
  loc_0042581B: push edx
  loc_0042581C: call [0040103Ch] ; __vbaObjSet
  loc_00425822: mov ebx, [eax]
  loc_00425824: mov edi, [0040102Ch] ; __vbaStrCat
  loc_0042582A: mov var_48, eax
  loc_0042582D: mov eax, [00430010h]
  loc_00425832: push 004107E4h ; "What is the value of"
  loc_00425837: push eax
  loc_00425838: call edi
  loc_0042583A: mov esi, [004010FCh] ; __vbaStrMove
  loc_00425840: mov edx, eax
  loc_00425842: lea ecx, var_18
  loc_00425845: call __vbaStrMove
  loc_00425847: push eax
  loc_00425848: push 00410814h ; " when"
  loc_0042584D: call edi
  loc_0042584F: mov edx, eax
  loc_00425851: lea ecx, var_1C
  loc_00425854: call __vbaStrMove
  loc_00425856: mov ecx, [00430018h]
  loc_0042585C: push eax
  loc_0042585D: push ecx
  loc_0042585E: call edi
  loc_00425860: mov edx, eax
  loc_00425862: lea ecx, var_20
  loc_00425865: call __vbaStrMove
  loc_00425867: push eax
  loc_00425868: push 00410824h ; " equals "
  loc_0042586D: call edi
  loc_0042586F: mov edx, eax
  loc_00425871: lea ecx, var_24
  loc_00425874: call __vbaStrMove
  loc_00425876: mov edx, [0043002Ch]
  loc_0042587C: push eax
  loc_0042587D: push edx
  loc_0042587E: call edi
  loc_00425880: mov edx, eax
  loc_00425882: lea ecx, var_28
  loc_00425885: call __vbaStrMove
  loc_00425887: push eax
  loc_00425888: mov eax, [0043001Ch]
  loc_0042588D: push eax
  loc_0042588E: call edi
  loc_00425890: mov edx, eax
  loc_00425892: lea ecx, var_2C
  loc_00425895: call __vbaStrMove
  loc_00425897: push eax
  loc_00425898: push 0041083Ch
  loc_0042589D: call edi
  loc_0042589F: mov edx, eax
  loc_004258A1: lea ecx, var_30
  loc_004258A4: call __vbaStrMove
  loc_004258A6: push eax
  loc_004258A7: mov ecx, ebx
  loc_004258A9: mov ebx, var_48
  loc_004258AC: push ebx
  loc_004258AD: call [ecx+0000019Ch]
  loc_004258B3: test eax, eax
  loc_004258B5: fnclex
  loc_004258B7: jge 004258CBh
  loc_004258B9: push 0000019Ch
  loc_004258BE: push 0040E390h
  loc_004258C3: push ebx
  loc_004258C4: push eax
  loc_004258C5: call [00401030h] ; __vbaHresultCheckObj
  loc_004258CB: lea edx, var_30
  loc_004258CE: lea eax, var_2C
  loc_004258D1: push edx
  loc_004258D2: lea ecx, var_28
  loc_004258D5: push eax
  loc_004258D6: lea edx, var_24
  loc_004258D9: push ecx
  loc_004258DA: lea eax, var_20
  loc_004258DD: push edx
  loc_004258DE: lea ecx, var_1C
  loc_004258E1: push eax
  loc_004258E2: lea edx, var_18
  loc_004258E5: push ecx
  loc_004258E6: push edx
  loc_004258E7: push 00000007h
  loc_004258E9: call [004010E4h] ; __vbaFreeStrList
  loc_004258EF: add esp, 00000020h
  loc_004258F2: lea ecx, var_44
  loc_004258F5: call [00401114h] ; __vbaFreeObj
  loc_004258FB: mov eax, Me
  loc_004258FE: push eax
  loc_004258FF: mov ecx, [eax]
  loc_00425901: call [ecx+00000310h]
  loc_00425907: lea edx, var_44
  loc_0042590A: push eax
  loc_0042590B: push edx
  loc_0042590C: call [0040103Ch] ; __vbaObjSet
  loc_00425912: mov ebx, [eax]
  loc_00425914: mov var_48, eax
  loc_00425917: mov eax, [00430010h]
  loc_0042591C: push 00410844h ; "What is the mean value of"
  loc_00425921: push eax
  loc_00425922: call edi
  loc_00425924: mov edx, eax
  loc_00425926: lea ecx, var_18
  loc_00425929: call __vbaStrMove
  loc_0042592B: push eax
  loc_0042592C: push 00410814h ; " when"
  loc_00425931: call edi
  loc_00425933: mov edx, eax
  loc_00425935: lea ecx, var_1C
  loc_00425938: call __vbaStrMove
  loc_0042593A: mov ecx, [00430018h]
  loc_00425940: push eax
  loc_00425941: push ecx
  loc_00425942: call edi
  loc_00425944: mov edx, eax
  loc_00425946: lea ecx, var_20
  loc_00425949: call __vbaStrMove
  loc_0042594B: push eax
  loc_0042594C: push 00410824h ; " equals "
  loc_00425951: call edi
  loc_00425953: mov edx, eax
  loc_00425955: lea ecx, var_24
  loc_00425958: call __vbaStrMove
  loc_0042595A: mov edx, [0043002Ch]
  loc_00425960: push eax
  loc_00425961: push edx
  loc_00425962: call edi
  loc_00425964: mov edx, eax
  loc_00425966: lea ecx, var_28
  loc_00425969: call __vbaStrMove
  loc_0042596B: push eax
  loc_0042596C: mov eax, [0043001Ch]
  loc_00425971: push eax
  loc_00425972: call edi
  loc_00425974: mov edx, eax
  loc_00425976: lea ecx, var_2C
  loc_00425979: call __vbaStrMove
  loc_0042597B: push eax
  loc_0042597C: push 0041083Ch
  loc_00425981: call edi
  loc_00425983: mov edx, eax
  loc_00425985: lea ecx, var_30
  loc_00425988: call __vbaStrMove
  loc_0042598A: mov ecx, ebx
  loc_0042598C: mov ebx, var_48
  loc_0042598F: push eax
  loc_00425990: push ebx
  loc_00425991: call [ecx+0000019Ch]
  loc_00425997: test eax, eax
  loc_00425999: fnclex
  loc_0042599B: jge 004259AFh
  loc_0042599D: push 0000019Ch
  loc_004259A2: push 0040E390h
  loc_004259A7: push ebx
  loc_004259A8: push eax
  loc_004259A9: call [00401030h] ; __vbaHresultCheckObj
  loc_004259AF: lea edx, var_30
  loc_004259B2: lea eax, var_2C
  loc_004259B5: push edx
  loc_004259B6: lea ecx, var_28
  loc_004259B9: push eax
  loc_004259BA: lea edx, var_24
  loc_004259BD: push ecx
  loc_004259BE: lea eax, var_20
  loc_004259C1: push edx
  loc_004259C2: lea ecx, var_1C
  loc_004259C5: push eax
  loc_004259C6: lea edx, var_18
  loc_004259C9: push ecx
  loc_004259CA: push edx
  loc_004259CB: push 00000007h
  loc_004259CD: call [004010E4h] ; __vbaFreeStrList
  loc_004259D3: add esp, 00000020h
  loc_004259D6: lea ecx, var_44
  loc_004259D9: call [00401114h] ; __vbaFreeObj
  loc_004259DF: mov eax, Me
  loc_004259E2: push eax
  loc_004259E3: mov ecx, [eax]
  loc_004259E5: call [ecx+00000344h]
  loc_004259EB: lea edx, var_44
  loc_004259EE: push eax
  loc_004259EF: push edx
  loc_004259F0: call [0040103Ch] ; __vbaObjSet
  loc_004259F6: mov ebx, [eax]
  loc_004259F8: mov var_48, eax
  loc_004259FB: mov eax, [00430010h]
  loc_00425A00: push 0041087Ch ; "What is the change in the mean of"
  loc_00425A05: push eax
  loc_00425A06: call edi
  loc_00425A08: mov edx, eax
  loc_00425A0A: lea ecx, var_18
  loc_00425A0D: call __vbaStrMove
  loc_00425A0F: push eax
  loc_00425A10: push 004108C4h ; " with a one"
  loc_00425A15: call edi
  loc_00425A17: mov edx, eax
  loc_00425A19: lea ecx, var_1C
  loc_00425A1C: call __vbaStrMove
  loc_00425A1E: mov ecx, [0043001Ch]
  loc_00425A24: push eax
  loc_00425A25: push ecx
  loc_00425A26: call edi
  loc_00425A28: mov edx, eax
  loc_00425A2A: lea ecx, var_20
  loc_00425A2D: call __vbaStrMove
  loc_00425A2F: push eax
  loc_00425A30: push 004108E0h ; " increase in"
  loc_00425A35: call edi
  loc_00425A37: mov edx, eax
  loc_00425A39: lea ecx, var_24
  loc_00425A3C: call __vbaStrMove
  loc_00425A3E: mov edx, [00430018h]
  loc_00425A44: push eax
  loc_00425A45: push edx
  loc_00425A46: call edi
  loc_00425A48: mov edx, eax
  loc_00425A4A: lea ecx, var_28
  loc_00425A4D: call __vbaStrMove
  loc_00425A4F: push eax
  loc_00425A50: push 0041083Ch
  loc_00425A55: call edi
  loc_00425A57: mov edx, eax
  loc_00425A59: lea ecx, var_2C
  loc_00425A5C: call __vbaStrMove
  loc_00425A5E: mov var_60, ebx
  loc_00425A61: mov ebx, var_48
  loc_00425A64: push eax
  loc_00425A65: mov eax, var_60
  loc_00425A68: push ebx
  loc_00425A69: call [eax+0000019Ch]
  loc_00425A6F: test eax, eax
  loc_00425A71: fnclex
  loc_00425A73: jge 00425A87h
  loc_00425A75: push 0000019Ch
  loc_00425A7A: push 0040E390h
  loc_00425A7F: push ebx
  loc_00425A80: push eax
  loc_00425A81: call [00401030h] ; __vbaHresultCheckObj
  loc_00425A87: lea ecx, var_2C
  loc_00425A8A: lea edx, var_28
  loc_00425A8D: push ecx
  loc_00425A8E: lea eax, var_24
  loc_00425A91: push edx
  loc_00425A92: lea ecx, var_20
  loc_00425A95: push eax
  loc_00425A96: lea edx, var_1C
  loc_00425A99: push ecx
  loc_00425A9A: lea eax, var_18
  loc_00425A9D: push edx
  loc_00425A9E: push eax
  loc_00425A9F: push 00000006h
  loc_00425AA1: call [004010E4h] ; __vbaFreeStrList
  loc_00425AA7: add esp, 0000001Ch
  loc_00425AAA: lea ecx, var_44
  loc_00425AAD: call [00401114h] ; __vbaFreeObj
  loc_00425AB3: fld real4 ptr [0043005Ch]
  loc_00425AB9: fcomp real4 ptr [004011E8h]
  loc_00425ABF: fnstsw ax
  loc_00425AC1: test ah, 40h
  loc_00425AC4: jz 004261A5h
  loc_00425ACA: mov ebx, Me
  loc_00425ACD: push ebx
  loc_00425ACE: mov ecx, [ebx]
  loc_00425AD0: call [ecx+00000330h]
  loc_00425AD6: lea edx, var_44
  loc_00425AD9: push eax
  loc_00425ADA: push edx
  loc_00425ADB: call [0040103Ch] ; __vbaObjSet
  loc_00425AE1: mov ecx, [eax]
  loc_00425AE3: push FFFFFFFFh
  loc_00425AE5: push eax
  loc_00425AE6: mov var_48, eax
  loc_00425AE9: call [ecx+0000009Ch]
  loc_00425AEF: test eax, eax
  loc_00425AF1: fnclex
  loc_00425AF3: jge 00425B0Ah
  loc_00425AF5: mov edx, var_48
  loc_00425AF8: push 0000009Ch
  loc_00425AFD: push 0040E390h
  loc_00425B02: push edx
  loc_00425B03: push eax
  loc_00425B04: call [00401030h] ; __vbaHresultCheckObj
  loc_00425B0A: lea ecx, var_44
  loc_00425B0D: call [00401114h] ; __vbaFreeObj
  loc_00425B13: mov eax, [ebx]
  loc_00425B15: push ebx
  loc_00425B16: call [eax+0000032Ch]
  loc_00425B1C: lea ecx, var_44
  loc_00425B1F: push eax
  loc_00425B20: push ecx
  loc_00425B21: call [0040103Ch] ; __vbaObjSet
  loc_00425B27: mov edx, [eax]
  loc_00425B29: push FFFFFFFFh
  loc_00425B2B: push eax
  loc_00425B2C: mov var_48, eax
  loc_00425B2F: call [edx+0000009Ch]
  loc_00425B35: test eax, eax
  loc_00425B37: fnclex
  loc_00425B39: jge 00425B50h
  loc_00425B3B: mov ecx, var_48
  loc_00425B3E: push 0000009Ch
  loc_00425B43: push 0040E390h
  loc_00425B48: push ecx
  loc_00425B49: push eax
  loc_00425B4A: call [00401030h] ; __vbaHresultCheckObj
  loc_00425B50: lea ecx, var_44
  loc_00425B53: call [00401114h] ; __vbaFreeObj
  loc_00425B59: mov edx, [ebx]
  loc_00425B5B: push ebx
  loc_00425B5C: call [edx+00000328h]
  loc_00425B62: push eax
  loc_00425B63: lea eax, var_44
  loc_00425B66: push eax
  loc_00425B67: call [0040103Ch] ; __vbaObjSet
  loc_00425B6D: mov ecx, [eax]
  loc_00425B6F: push FFFFFFFFh
  loc_00425B71: push eax
  loc_00425B72: mov var_48, eax
  loc_00425B75: call [ecx+0000009Ch]
  loc_00425B7B: test eax, eax
  loc_00425B7D: fnclex
  loc_00425B7F: jge 00425B96h
  loc_00425B81: mov edx, var_48
  loc_00425B84: push 0000009Ch
  loc_00425B89: push 0040E390h
  loc_00425B8E: push edx
  loc_00425B8F: push eax
  loc_00425B90: call [00401030h] ; __vbaHresultCheckObj
  loc_00425B96: lea ecx, var_44
  loc_00425B99: call [00401114h] ; __vbaFreeObj
  loc_00425B9F: mov eax, [ebx]
  loc_00425BA1: push ebx
  loc_00425BA2: call [eax+00000334h]
  loc_00425BA8: lea ecx, var_44
  loc_00425BAB: push eax
  loc_00425BAC: push ecx
  loc_00425BAD: call [0040103Ch] ; __vbaObjSet
  loc_00425BB3: mov edx, [eax]
  loc_00425BB5: push 43D98000h
  loc_00425BBA: push eax
  loc_00425BBB: mov var_48, eax
  loc_00425BBE: call [edx+0000008Ch]
  loc_00425BC4: test eax, eax
  loc_00425BC6: fnclex
  loc_00425BC8: jge 00425BDFh
  loc_00425BCA: mov ecx, var_48
  loc_00425BCD: push 0000008Ch
  loc_00425BD2: push 0040E390h
  loc_00425BD7: push ecx
  loc_00425BD8: push eax
  loc_00425BD9: call [00401030h] ; __vbaHresultCheckObj
  loc_00425BDF: lea ecx, var_44
  loc_00425BE2: call [00401114h] ; __vbaFreeObj
  loc_00425BE8: mov edx, [ebx]
  loc_00425BEA: push ebx
  loc_00425BEB: call [edx+0000033Ch]
  loc_00425BF1: push eax
  loc_00425BF2: lea eax, var_44
  loc_00425BF5: push eax
  loc_00425BF6: call [0040103Ch] ; __vbaObjSet
  loc_00425BFC: mov ecx, [0043001Ch]
  loc_00425C02: mov ebx, [eax]
  loc_00425C04: push 00410900h ; "Is a one"
  loc_00425C09: push ecx
  loc_00425C0A: mov var_48, eax
  loc_00425C0D: call edi
  loc_00425C0F: mov edx, eax
  loc_00425C11: lea ecx, var_18
  loc_00425C14: call __vbaStrMove
  loc_00425C16: push eax
  loc_00425C17: push 004108E0h ; " increase in"
  loc_00425C1C: call edi
  loc_00425C1E: mov edx, eax
  loc_00425C20: lea ecx, var_1C
  loc_00425C23: call __vbaStrMove
  loc_00425C25: mov edx, [00430018h]
  loc_00425C2B: push eax
  loc_00425C2C: push edx
  loc_00425C2D: call edi
  loc_00425C2F: mov edx, eax
  loc_00425C31: lea ecx, var_20
  loc_00425C34: call __vbaStrMove
  loc_00425C36: push eax
  loc_00425C37: push 00410918h ; " associated with an increase in the mean of"
  loc_00425C3C: call edi
  loc_00425C3E: mov edx, eax
  loc_00425C40: lea ecx, var_24
  loc_00425C43: call __vbaStrMove
  loc_00425C45: push eax
  loc_00425C46: mov eax, [00430010h]
  loc_00425C4B: push eax
  loc_00425C4C: call edi
  loc_00425C4E: mov edx, eax
  loc_00425C50: lea ecx, var_28
  loc_00425C53: call __vbaStrMove
  loc_00425C55: push eax
  loc_00425C56: push 0041083Ch
  loc_00425C5B: call edi
  loc_00425C5D: mov edx, eax
  loc_00425C5F: lea ecx, var_2C
  loc_00425C62: call __vbaStrMove
  loc_00425C64: mov ecx, ebx
  loc_00425C66: mov ebx, var_48
  loc_00425C69: push eax
  loc_00425C6A: push ebx
  loc_00425C6B: call [ecx+0000019Ch]
  loc_00425C71: test eax, eax
  loc_00425C73: fnclex
  loc_00425C75: jge 00425C89h
  loc_00425C77: push 0000019Ch
  loc_00425C7C: push 0040E390h
  loc_00425C81: push ebx
  loc_00425C82: push eax
  loc_00425C83: call [00401030h] ; __vbaHresultCheckObj
  loc_00425C89: lea edx, var_2C
  loc_00425C8C: lea eax, var_28
  loc_00425C8F: push edx
  loc_00425C90: lea ecx, var_24
  loc_00425C93: push eax
  loc_00425C94: lea edx, var_20
  loc_00425C97: push ecx
  loc_00425C98: lea eax, var_1C
  loc_00425C9B: push edx
  loc_00425C9C: lea ecx, var_18
  loc_00425C9F: push eax
  loc_00425CA0: push ecx
  loc_00425CA1: push 00000006h
  loc_00425CA3: call [004010E4h] ; __vbaFreeStrList
  loc_00425CA9: add esp, 0000001Ch
  loc_00425CAC: lea ecx, var_44
  loc_00425CAF: call [00401114h] ; __vbaFreeObj
  loc_00425CB5: mov eax, Me
  loc_00425CB8: push eax
  loc_00425CB9: mov edx, [eax]
  loc_00425CBB: call [edx+00000334h]
  loc_00425CC1: push eax
  loc_00425CC2: lea eax, var_44
  loc_00425CC5: push eax
  loc_00425CC6: call [0040103Ch] ; __vbaObjSet
  loc_00425CCC: mov ecx, [00430018h]
  loc_00425CD2: mov ebx, [eax]
  loc_00425CD4: push 00410974h ; "Are changes in"
  loc_00425CD9: push ecx
  loc_00425CDA: mov var_48, eax
  loc_00425CDD: call edi
  loc_00425CDF: mov edx, eax
  loc_00425CE1: lea ecx, var_18
  loc_00425CE4: call __vbaStrMove
  loc_00425CE6: push eax
  loc_00425CE7: push 00410998h ; " associated with changes in the mean of"
  loc_00425CEC: call edi
  loc_00425CEE: mov edx, eax
  loc_00425CF0: lea ecx, var_1C
  loc_00425CF3: call __vbaStrMove
  loc_00425CF5: mov edx, [00430010h]
  loc_00425CFB: push eax
  loc_00425CFC: push edx
  loc_00425CFD: call edi
  loc_00425CFF: mov edx, eax
  loc_00425D01: lea ecx, var_20
  loc_00425D04: call __vbaStrMove
  loc_00425D06: push eax
  loc_00425D07: push 0040DD3Ch ; "."
  loc_00425D0C: call edi
  loc_00425D0E: mov edx, eax
  loc_00425D10: lea ecx, var_24
  loc_00425D13: call __vbaStrMove
  loc_00425D15: mov var_68, ebx
  loc_00425D18: mov ebx, var_48
  loc_00425D1B: push eax
  loc_00425D1C: mov eax, var_68
  loc_00425D1F: push ebx
  loc_00425D20: call [eax+0000019Ch]
  loc_00425D26: test eax, eax
  loc_00425D28: fnclex
  loc_00425D2A: jge 00425D3Eh
  loc_00425D2C: push 0000019Ch
  loc_00425D31: push 0040E390h
  loc_00425D36: push ebx
  loc_00425D37: push eax
  loc_00425D38: call [00401030h] ; __vbaHresultCheckObj
  loc_00425D3E: lea ecx, var_24
  loc_00425D41: lea edx, var_20
  loc_00425D44: push ecx
  loc_00425D45: lea eax, var_1C
  loc_00425D48: push edx
  loc_00425D49: lea ecx, var_18
  loc_00425D4C: push eax
  loc_00425D4D: push ecx
  loc_00425D4E: push 00000004h
  loc_00425D50: call [004010E4h] ; __vbaFreeStrList
  loc_00425D56: add esp, 00000014h
  loc_00425D59: lea ecx, var_44
  loc_00425D5C: call [00401114h] ; __vbaFreeObj
  loc_00425D62: mov eax, Me
  loc_00425D65: push eax
  loc_00425D66: mov edx, [eax]
  loc_00425D68: call [edx+00000330h]
  loc_00425D6E: push eax
  loc_00425D6F: lea eax, var_44
  loc_00425D72: push eax
  loc_00425D73: call [0040103Ch] ; __vbaObjSet
  loc_00425D79: mov ecx, [00430018h]
  loc_00425D7F: mov ebx, [eax]
  loc_00425D81: push 004109ECh ; "Does"
  loc_00425D86: push ecx
  loc_00425D87: mov var_48, eax
  loc_00425D8A: call edi
  loc_00425D8C: mov edx, eax
  loc_00425D8E: lea ecx, var_18
  loc_00425D91: call __vbaStrMove
  loc_00425D93: push eax
  loc_00425D94: push 004109FCh ; " help predict"
  loc_00425D99: call edi
  loc_00425D9B: mov edx, eax
  loc_00425D9D: lea ecx, var_1C
  loc_00425DA0: call __vbaStrMove
  loc_00425DA2: mov edx, [00430010h]
  loc_00425DA8: push eax
  loc_00425DA9: push edx
  loc_00425DAA: call edi
  loc_00425DAC: mov edx, eax
  loc_00425DAE: lea ecx, var_20
  loc_00425DB1: call __vbaStrMove
  loc_00425DB3: push eax
  loc_00425DB4: push 0041083Ch
  loc_00425DB9: call edi
  loc_00425DBB: mov edx, eax
  loc_00425DBD: lea ecx, var_24
  loc_00425DC0: call __vbaStrMove
  loc_00425DC2: mov var_6C, ebx
  loc_00425DC5: mov ebx, var_48
  loc_00425DC8: push eax
  loc_00425DC9: mov eax, var_6C
  loc_00425DCC: push ebx
  loc_00425DCD: call [eax+0000019Ch]
  loc_00425DD3: test eax, eax
  loc_00425DD5: fnclex
  loc_00425DD7: jge 00425DEBh
  loc_00425DD9: push 0000019Ch
  loc_00425DDE: push 0040E390h
  loc_00425DE3: push ebx
  loc_00425DE4: push eax
  loc_00425DE5: call [00401030h] ; __vbaHresultCheckObj
  loc_00425DEB: lea ecx, var_24
  loc_00425DEE: lea edx, var_20
  loc_00425DF1: push ecx
  loc_00425DF2: lea eax, var_1C
  loc_00425DF5: push edx
  loc_00425DF6: lea ecx, var_18
  loc_00425DF9: push eax
  loc_00425DFA: push ecx
  loc_00425DFB: push 00000004h
  loc_00425DFD: call [004010E4h] ; __vbaFreeStrList
  loc_00425E03: add esp, 00000014h
  loc_00425E06: lea ecx, var_44
  loc_00425E09: call [00401114h] ; __vbaFreeObj
  loc_00425E0F: mov eax, Me
  loc_00425E12: push eax
  loc_00425E13: mov edx, [eax]
  loc_00425E15: call [edx+0000032Ch]
  loc_00425E1B: push eax
  loc_00425E1C: lea eax, var_44
  loc_00425E1F: push eax
  loc_00425E20: call [0040103Ch] ; __vbaObjSet
  loc_00425E26: mov ecx, [00430018h]
  loc_00425E2C: mov ebx, [eax]
  loc_00425E2E: push 004109ECh ; "Does"
  loc_00425E33: push ecx
  loc_00425E34: mov var_48, eax
  loc_00425E37: call edi
  loc_00425E39: mov edx, eax
  loc_00425E3B: lea ecx, var_18
  loc_00425E3E: call __vbaStrMove
  loc_00425E40: push eax
  loc_00425E41: push 00410A1Ch ; " help estimate the mean"
  loc_00425E46: call edi
  loc_00425E48: mov edx, eax
  loc_00425E4A: lea ecx, var_1C
  loc_00425E4D: call __vbaStrMove
  loc_00425E4F: mov edx, [00430010h]
  loc_00425E55: push eax
  loc_00425E56: push edx
  loc_00425E57: call edi
  loc_00425E59: mov edx, eax
  loc_00425E5B: lea ecx, var_20
  loc_00425E5E: call __vbaStrMove
  loc_00425E60: push eax
  loc_00425E61: push 0041083Ch
  loc_00425E66: call edi
  loc_00425E68: mov edx, eax
  loc_00425E6A: lea ecx, var_24
  loc_00425E6D: call __vbaStrMove
  loc_00425E6F: mov var_70, ebx
  loc_00425E72: mov ebx, var_48
  loc_00425E75: push eax
  loc_00425E76: mov eax, var_70
  loc_00425E79: push ebx
  loc_00425E7A: call [eax+0000019Ch]
  loc_00425E80: test eax, eax
  loc_00425E82: fnclex
  loc_00425E84: jge 00425E98h
  loc_00425E86: push 0000019Ch
  loc_00425E8B: push 0040E390h
  loc_00425E90: push ebx
  loc_00425E91: push eax
  loc_00425E92: call [00401030h] ; __vbaHresultCheckObj
  loc_00425E98: lea ecx, var_24
  loc_00425E9B: lea edx, var_20
  loc_00425E9E: push ecx
  loc_00425E9F: lea eax, var_1C
  loc_00425EA2: push edx
  loc_00425EA3: lea ecx, var_18
  loc_00425EA6: push eax
  loc_00425EA7: push ecx
  loc_00425EA8: push 00000004h
  loc_00425EAA: call [004010E4h] ; __vbaFreeStrList
  loc_00425EB0: add esp, 00000014h
  loc_00425EB3: lea ecx, var_44
  loc_00425EB6: call [00401114h] ; __vbaFreeObj
  loc_00425EBC: mov eax, Me
  loc_00425EBF: push eax
  loc_00425EC0: mov edx, [eax]
  loc_00425EC2: call [edx+00000328h]
  loc_00425EC8: push eax
  loc_00425EC9: lea eax, var_44
  loc_00425ECC: push eax
  loc_00425ECD: call [0040103Ch] ; __vbaObjSet
  loc_00425ED3: mov ecx, [00430018h]
  loc_00425ED9: mov ebx, [eax]
  loc_00425EDB: push 00410A50h ; "Is variation in"
  loc_00425EE0: push ecx
  loc_00425EE1: mov var_48, eax
  loc_00425EE4: call edi
  loc_00425EE6: mov edx, eax
  loc_00425EE8: lea ecx, var_18
  loc_00425EEB: call __vbaStrMove
  loc_00425EED: push eax
  loc_00425EEE: push 00410A74h ; " associated with variation in"
  loc_00425EF3: call edi
  loc_00425EF5: mov edx, eax
  loc_00425EF7: lea ecx, var_1C
  loc_00425EFA: call __vbaStrMove
  loc_00425EFC: mov edx, [00430010h]
  loc_00425F02: push eax
  loc_00425F03: push edx
  loc_00425F04: call edi
  loc_00425F06: mov edx, eax
  loc_00425F08: lea ecx, var_20
  loc_00425F0B: call __vbaStrMove
  loc_00425F0D: push eax
  loc_00425F0E: push 0041083Ch
  loc_00425F13: call edi
  loc_00425F15: mov edx, eax
  loc_00425F17: lea ecx, var_24
  loc_00425F1A: call __vbaStrMove
  loc_00425F1C: mov var_74, ebx
  loc_00425F1F: mov ebx, var_48
  loc_00425F22: push eax
  loc_00425F23: mov eax, var_74
  loc_00425F26: push ebx
  loc_00425F27: call [eax+0000019Ch]
  loc_00425F2D: test eax, eax
  loc_00425F2F: fnclex
  loc_00425F31: jge 00425F45h
  loc_00425F33: push 0000019Ch
  loc_00425F38: push 0040E390h
  loc_00425F3D: push ebx
  loc_00425F3E: push eax
  loc_00425F3F: call [00401030h] ; __vbaHresultCheckObj
  loc_00425F45: lea ecx, var_24
  loc_00425F48: lea edx, var_20
  loc_00425F4B: push ecx
  loc_00425F4C: lea eax, var_1C
  loc_00425F4F: push edx
  loc_00425F50: lea ecx, var_18
  loc_00425F53: push eax
  loc_00425F54: push ecx
  loc_00425F55: push 00000004h
  loc_00425F57: call [004010E4h] ; __vbaFreeStrList
  loc_00425F5D: add esp, 00000014h
  loc_00425F60: lea ecx, var_44
  loc_00425F63: call [00401114h] ; __vbaFreeObj
  loc_00425F69: mov eax, Me
  loc_00425F6C: push eax
  loc_00425F6D: mov edx, [eax]
  loc_00425F6F: call [edx+0000031Ch]
  loc_00425F75: push eax
  loc_00425F76: lea eax, var_44
  loc_00425F79: push eax
  loc_00425F7A: call [0040103Ch] ; __vbaObjSet
  loc_00425F80: mov ecx, [0043001Ch]
  loc_00425F86: mov ebx, [eax]
  loc_00425F88: push 00410900h ; "Is a one"
  loc_00425F8D: push ecx
  loc_00425F8E: mov var_48, eax
  loc_00425F91: call edi
  loc_00425F93: mov edx, eax
  loc_00425F95: lea ecx, var_18
  loc_00425F98: call __vbaStrMove
  loc_00425F9A: push eax
  loc_00425F9B: push 004108E0h ; " increase in"
  loc_00425FA0: call edi
  loc_00425FA2: mov edx, eax
  loc_00425FA4: lea ecx, var_1C
  loc_00425FA7: call __vbaStrMove
  loc_00425FA9: mov edx, [00430018h]
  loc_00425FAF: push eax
  loc_00425FB0: push edx
  loc_00425FB1: call edi
  loc_00425FB3: mov edx, eax
  loc_00425FB5: lea ecx, var_20
  loc_00425FB8: call __vbaStrMove
  loc_00425FBA: push eax
  loc_00425FBB: push 00410AB4h ; " associated with a decrease in the mean of"
  loc_00425FC0: call edi
  loc_00425FC2: mov edx, eax
  loc_00425FC4: lea ecx, var_24
  loc_00425FC7: call __vbaStrMove
  loc_00425FC9: push eax
  loc_00425FCA: mov eax, [00430010h]
  loc_00425FCF: push eax
  loc_00425FD0: call edi
  loc_00425FD2: mov edx, eax
  loc_00425FD4: lea ecx, var_28
  loc_00425FD7: call __vbaStrMove
  loc_00425FD9: push eax
  loc_00425FDA: push 0041083Ch
  loc_00425FDF: call edi
  loc_00425FE1: mov edx, eax
  loc_00425FE3: lea ecx, var_2C
  loc_00425FE6: call __vbaStrMove
  loc_00425FE8: mov esi, var_48
  loc_00425FEB: push eax
  loc_00425FEC: push esi
  loc_00425FED: call [ebx+0000019Ch]
  loc_00425FF3: test eax, eax
  loc_00425FF5: fnclex
  loc_00425FF7: jge 0042600Bh
  loc_00425FF9: push 0000019Ch
  loc_00425FFE: push 0040E390h
  loc_00426003: push esi
  loc_00426004: push eax
  loc_00426005: call [00401030h] ; __vbaHresultCheckObj
  loc_0042600B: lea ecx, var_2C
  loc_0042600E: lea edx, var_28
  loc_00426011: push ecx
  loc_00426012: lea eax, var_24
  loc_00426015: push edx
  loc_00426016: lea ecx, var_20
  loc_00426019: push eax
  loc_0042601A: lea edx, var_1C
  loc_0042601D: push ecx
  loc_0042601E: lea eax, var_18
  loc_00426021: push edx
  loc_00426022: push eax
  loc_00426023: push 00000006h
  loc_00426025: call [004010E4h] ; __vbaFreeStrList
  loc_0042602B: add esp, 0000001Ch
  loc_0042602E: lea ecx, var_44
  loc_00426031: call [00401114h] ; __vbaFreeObj
  loc_00426037: mov esi, Me
  loc_0042603A: push esi
  loc_0042603B: mov ecx, [esi]
  loc_0042603D: call [ecx+0000033Ch]
  loc_00426043: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_00426049: lea edx, var_44
  loc_0042604C: push eax
  loc_0042604D: push edx
  loc_0042604E: call ebx
  loc_00426050: mov edi, eax
  loc_00426052: push 00410B78h ; "Is a one unit increase in X associated with an increase in the mean of Y?"
  loc_00426057: push edi
  loc_00426058: mov eax, [edi]
  loc_0042605A: call [eax+00000054h]
  loc_0042605D: test eax, eax
  loc_0042605F: fnclex
  loc_00426061: jge 00426072h
  loc_00426063: push 00000054h
  loc_00426065: push 0040E390h
  loc_0042606A: push edi
  loc_0042606B: push eax
  loc_0042606C: call [00401030h] ; __vbaHresultCheckObj
  loc_00426072: lea ecx, var_44
  loc_00426075: call [00401114h] ; __vbaFreeObj
  loc_0042607B: mov ecx, [esi]
  loc_0042607D: push esi
  loc_0042607E: call [ecx+00000334h]
  loc_00426084: lea edx, var_44
  loc_00426087: push eax
  loc_00426088: push edx
  loc_00426089: call ebx
  loc_0042608B: mov edi, eax
  loc_0042608D: push 00410C10h ; "a. Are changes in X associated with changes in the mean of Y."
  loc_00426092: push edi
  loc_00426093: mov eax, [edi]
  loc_00426095: call [eax+00000054h]
  loc_00426098: test eax, eax
  loc_0042609A: fnclex
  loc_0042609C: jge 004260ADh
  loc_0042609E: push 00000054h
  loc_004260A0: push 0040E390h
  loc_004260A5: push edi
  loc_004260A6: push eax
  loc_004260A7: call [00401030h] ; __vbaHresultCheckObj
  loc_004260AD: lea ecx, var_44
  loc_004260B0: call [00401114h] ; __vbaFreeObj
  loc_004260B6: mov ecx, [esi]
  loc_004260B8: push esi
  loc_004260B9: call [ecx+00000330h]
  loc_004260BF: lea edx, var_44
  loc_004260C2: push eax
  loc_004260C3: push edx
  loc_004260C4: call ebx
  loc_004260C6: mov edi, eax
  loc_004260C8: push 00410C90h ; "b. Does X help predict Y?"
  loc_004260CD: push edi
  loc_004260CE: mov eax, [edi]
  loc_004260D0: call [eax+00000054h]
  loc_004260D3: test eax, eax
  loc_004260D5: fnclex
  loc_004260D7: jge 004260E8h
  loc_004260D9: push 00000054h
  loc_004260DB: push 0040E390h
  loc_004260E0: push edi
  loc_004260E1: push eax
  loc_004260E2: call [00401030h] ; __vbaHresultCheckObj
  loc_004260E8: lea ecx, var_44
  loc_004260EB: call [00401114h] ; __vbaFreeObj
  loc_004260F1: mov ecx, [esi]
  loc_004260F3: push esi
  loc_004260F4: call [ecx+0000032Ch]
  loc_004260FA: lea edx, var_44
  loc_004260FD: push eax
  loc_004260FE: push edx
  loc_004260FF: call ebx
  loc_00426101: mov edi, eax
  loc_00426103: push 00410CC8h ; "c. Does X help estimate the mean Y?"
  loc_00426108: push edi
  loc_00426109: mov eax, [edi]
  loc_0042610B: call [eax+00000054h]
  loc_0042610E: test eax, eax
  loc_00426110: fnclex
  loc_00426112: jge 00426123h
  loc_00426114: push 00000054h
  loc_00426116: push 0040E390h
  loc_0042611B: push edi
  loc_0042611C: push eax
  loc_0042611D: call [00401030h] ; __vbaHresultCheckObj
  loc_00426123: lea ecx, var_44
  loc_00426126: call [00401114h] ; __vbaFreeObj
  loc_0042612C: mov ecx, [esi]
  loc_0042612E: push esi
  loc_0042612F: call [ecx+00000328h]
  loc_00426135: lea edx, var_44
  loc_00426138: push eax
  loc_00426139: push edx
  loc_0042613A: call ebx
  loc_0042613C: mov edi, eax
  loc_0042613E: push 00410D54h ; "d. Is variation in X associated with variation in Y?"
  loc_00426143: push edi
  loc_00426144: mov eax, [edi]
  loc_00426146: call [eax+00000054h]
  loc_00426149: test eax, eax
  loc_0042614B: fnclex
  loc_0042614D: jge 0042615Eh
  loc_0042614F: push 00000054h
  loc_00426151: push 0040E390h
  loc_00426156: push edi
  loc_00426157: push eax
  loc_00426158: call [00401030h] ; __vbaHresultCheckObj
  loc_0042615E: mov edi, [00401114h] ; __vbaFreeObj
  loc_00426164: lea ecx, var_44
  loc_00426167: call edi
  loc_00426169: mov ecx, [esi]
  loc_0042616B: push esi
  loc_0042616C: call [ecx+0000031Ch]
  loc_00426172: lea edx, var_44
  loc_00426175: push eax
  loc_00426176: push edx
  loc_00426177: call ebx
  loc_00426179: mov esi, eax
  loc_0042617B: push 00410DC4h ; "Is a one unit increase in X associated with a decrease in the mean of Y?"
  loc_00426180: push esi
  loc_00426181: mov eax, [esi]
  loc_00426183: call [eax+00000054h]
  loc_00426186: test eax, eax
  loc_00426188: fnclex
  loc_0042618A: jge 0042619Bh
  loc_0042618C: push 00000054h
  loc_0042618E: push 0040E390h
  loc_00426193: push esi
  loc_00426194: push eax
  loc_00426195: call [00401030h] ; __vbaHresultCheckObj
  loc_0042619B: lea ecx, var_44
  loc_0042619E: call edi
  loc_004261A0: jmp 00427051h
  loc_004261A5: fld real4 ptr [0043005Ch]
  loc_004261AB: fcomp real4 ptr [004011E8h]
  loc_004261B1: mov ebx, Me
  loc_004261B4: fnstsw ax
  loc_004261B6: test ah, 41h
  loc_004261B9: jnz 00426893h
  loc_004261BF: mov ecx, [ebx]
  loc_004261C1: push ebx
  loc_004261C2: call [ecx+00000330h]
  loc_004261C8: lea edx, var_44
  loc_004261CB: push eax
  loc_004261CC: push edx
  loc_004261CD: call [0040103Ch] ; __vbaObjSet
  loc_004261D3: mov ecx, [eax]
  loc_004261D5: push 00000000h
  loc_004261D7: push eax
  loc_004261D8: mov var_48, eax
  loc_004261DB: call [ecx+0000009Ch]
  loc_004261E1: test eax, eax
  loc_004261E3: fnclex
  loc_004261E5: jge 004261FCh
  loc_004261E7: mov edx, var_48
  loc_004261EA: push 0000009Ch
  loc_004261EF: push 0040E390h
  loc_004261F4: push edx
  loc_004261F5: push eax
  loc_004261F6: call [00401030h] ; __vbaHresultCheckObj
  loc_004261FC: lea ecx, var_44
  loc_004261FF: call [00401114h] ; __vbaFreeObj
  loc_00426205: mov eax, [ebx]
  loc_00426207: push ebx
  loc_00426208: call [eax+0000032Ch]
  loc_0042620E: lea ecx, var_44
  loc_00426211: push eax
  loc_00426212: push ecx
  loc_00426213: call [0040103Ch] ; __vbaObjSet
  loc_00426219: mov edx, [eax]
  loc_0042621B: push 00000000h
  loc_0042621D: push eax
  loc_0042621E: mov var_48, eax
  loc_00426221: call [edx+0000009Ch]
  loc_00426227: test eax, eax
  loc_00426229: fnclex
  loc_0042622B: jge 00426242h
  loc_0042622D: mov ecx, var_48
  loc_00426230: push 0000009Ch
  loc_00426235: push 0040E390h
  loc_0042623A: push ecx
  loc_0042623B: push eax
  loc_0042623C: call [00401030h] ; __vbaHresultCheckObj
  loc_00426242: lea ecx, var_44
  loc_00426245: call [00401114h] ; __vbaFreeObj
  loc_0042624B: mov edx, [ebx]
  loc_0042624D: push ebx
  loc_0042624E: call [edx+00000328h]
  loc_00426254: push eax
  loc_00426255: lea eax, var_44
  loc_00426258: push eax
  loc_00426259: call [0040103Ch] ; __vbaObjSet
  loc_0042625F: mov ecx, [eax]
  loc_00426261: push 00000000h
  loc_00426263: push eax
  loc_00426264: mov var_48, eax
  loc_00426267: call [ecx+0000009Ch]
  loc_0042626D: test eax, eax
  loc_0042626F: fnclex
  loc_00426271: jge 00426288h
  loc_00426273: mov edx, var_48
  loc_00426276: push 0000009Ch
  loc_0042627B: push 0040E390h
  loc_00426280: push edx
  loc_00426281: push eax
  loc_00426282: call [00401030h] ; __vbaHresultCheckObj
  loc_00426288: lea ecx, var_44
  loc_0042628B: call [00401114h] ; __vbaFreeObj
  loc_00426291: mov eax, [ebx]
  loc_00426293: push ebx
  loc_00426294: call [eax+00000334h]
  loc_0042629A: lea ecx, var_44
  loc_0042629D: push eax
  loc_0042629E: push ecx
  loc_0042629F: call [0040103Ch] ; __vbaObjSet
  loc_004262A5: mov edx, [eax]
  loc_004262A7: push 44480000h
  loc_004262AC: push eax
  loc_004262AD: mov var_48, eax
  loc_004262B0: call [edx+0000008Ch]
  loc_004262B6: test eax, eax
  loc_004262B8: fnclex
  loc_004262BA: jge 004262D1h
  loc_004262BC: mov ecx, var_48
  loc_004262BF: push 0000008Ch
  loc_004262C4: push 0040E390h
  loc_004262C9: push ecx
  loc_004262CA: push eax
  loc_004262CB: call [00401030h] ; __vbaHresultCheckObj
  loc_004262D1: lea ecx, var_44
  loc_004262D4: call [00401114h] ; __vbaFreeObj
  loc_004262DA: mov edx, [ebx]
  loc_004262DC: push ebx
  loc_004262DD: call [edx+0000033Ch]
  loc_004262E3: push eax
  loc_004262E4: lea eax, var_44
  loc_004262E7: push eax
  loc_004262E8: call [0040103Ch] ; __vbaObjSet
  loc_004262EE: mov ecx, [0043001Ch]
  loc_004262F4: mov ebx, [eax]
  loc_004262F6: push 00410E5Ch ; "Is a one "
  loc_004262FB: push ecx
  loc_004262FC: mov var_48, eax
  loc_004262FF: call edi
  loc_00426301: mov edx, eax
  loc_00426303: lea ecx, var_18
  loc_00426306: call __vbaStrMove
  loc_00426308: push eax
  loc_00426309: push 00410E74h ; " increase in "
  loc_0042630E: call edi
  loc_00426310: mov edx, eax
  loc_00426312: lea ecx, var_1C
  loc_00426315: call __vbaStrMove
  loc_00426317: mov edx, [00430018h]
  loc_0042631D: push eax
  loc_0042631E: push edx
  loc_0042631F: call edi
  loc_00426321: mov edx, eax
  loc_00426323: lea ecx, var_20
  loc_00426326: call __vbaStrMove
  loc_00426328: push eax
  loc_00426329: push 00410E94h ; " associated with an increase in the mean of "
  loc_0042632E: call edi
  loc_00426330: mov edx, eax
  loc_00426332: lea ecx, var_24
  loc_00426335: call __vbaStrMove
  loc_00426337: push eax
  loc_00426338: mov eax, [00430010h]
  loc_0042633D: push eax
  loc_0042633E: call edi
  loc_00426340: mov edx, eax
  loc_00426342: lea ecx, var_28
  loc_00426345: call __vbaStrMove
  loc_00426347: push eax
  loc_00426348: push 00410EF4h ; " of more than "
  loc_0042634D: call edi
  loc_0042634F: mov edx, eax
  loc_00426351: lea ecx, var_2C
  loc_00426354: call __vbaStrMove
  loc_00426356: mov ecx, [0043005Ch]
  loc_0042635C: push eax
  loc_0042635D: push ecx
  loc_0042635E: call [0040107Ch] ; __vbaStrR4
  loc_00426364: mov edx, eax
  loc_00426366: lea ecx, var_30
  loc_00426369: call __vbaStrMove
  loc_0042636B: push eax
  loc_0042636C: call edi
  loc_0042636E: mov edx, eax
  loc_00426370: lea ecx, var_34
  loc_00426373: call __vbaStrMove
  loc_00426375: push eax
  loc_00426376: push 0040F028h
  loc_0042637B: call edi
  loc_0042637D: mov edx, eax
  loc_0042637F: lea ecx, var_38
  loc_00426382: call __vbaStrMove
  loc_00426384: mov edx, [00430014h]
  loc_0042638A: push eax
  loc_0042638B: push edx
  loc_0042638C: call edi
  loc_0042638E: mov edx, eax
  loc_00426390: lea ecx, var_3C
  loc_00426393: call __vbaStrMove
  loc_00426395: push eax
  loc_00426396: push 0041083Ch
  loc_0042639B: call edi
  loc_0042639D: mov edx, eax
  loc_0042639F: lea ecx, var_40
  loc_004263A2: call __vbaStrMove
  loc_004263A4: mov var_78, ebx
  loc_004263A7: mov ebx, var_48
  loc_004263AA: push eax
  loc_004263AB: mov eax, var_78
  loc_004263AE: push ebx
  loc_004263AF: call [eax+0000019Ch]
  loc_004263B5: test eax, eax
  loc_004263B7: fnclex
  loc_004263B9: jge 004263CDh
  loc_004263BB: push 0000019Ch
  loc_004263C0: push 0040E390h
  loc_004263C5: push ebx
  loc_004263C6: push eax
  loc_004263C7: call [00401030h] ; __vbaHresultCheckObj
  loc_004263CD: lea ecx, var_40
  loc_004263D0: lea edx, var_3C
  loc_004263D3: push ecx
  loc_004263D4: lea eax, var_38
  loc_004263D7: push edx
  loc_004263D8: lea ecx, var_34
  loc_004263DB: push eax
  loc_004263DC: lea edx, var_30
  loc_004263DF: push ecx
  loc_004263E0: lea eax, var_2C
  loc_004263E3: push edx
  loc_004263E4: lea ecx, var_28
  loc_004263E7: push eax
  loc_004263E8: lea edx, var_24
  loc_004263EB: push ecx
  loc_004263EC: lea eax, var_20
  loc_004263EF: push edx
  loc_004263F0: lea ecx, var_1C
  loc_004263F3: push eax
  loc_004263F4: lea edx, var_18
  loc_004263F7: push ecx
  loc_004263F8: push edx
  loc_004263F9: push 0000000Bh
  loc_004263FB: call [004010E4h] ; __vbaFreeStrList
  loc_00426401: add esp, 00000030h
  loc_00426404: lea ecx, var_44
  loc_00426407: call [00401114h] ; __vbaFreeObj
  loc_0042640D: mov eax, Me
  loc_00426410: push eax
  loc_00426411: mov ecx, [eax]
  loc_00426413: call [ecx+00000334h]
  loc_00426419: lea edx, var_44
  loc_0042641C: push eax
  loc_0042641D: push edx
  loc_0042641E: call [0040103Ch] ; __vbaObjSet
  loc_00426424: mov ebx, [eax]
  loc_00426426: mov var_48, eax
  loc_00426429: mov eax, [00430018h]
  loc_0042642E: push 00410D14h ; "Can we say that increases in "
  loc_00426433: push eax
  loc_00426434: call edi
  loc_00426436: mov edx, eax
  loc_00426438: lea ecx, var_18
  loc_0042643B: call __vbaStrMove
  loc_0042643D: push eax
  loc_0042643E: push 00410B10h ; " are associated with increases in the mean of "
  loc_00426443: call edi
  loc_00426445: mov edx, eax
  loc_00426447: lea ecx, var_1C
  loc_0042644A: call __vbaStrMove
  loc_0042644C: mov ecx, [00430010h]
  loc_00426452: push eax
  loc_00426453: push ecx
  loc_00426454: call edi
  loc_00426456: mov edx, eax
  loc_00426458: lea ecx, var_20
  loc_0042645B: call __vbaStrMove
  loc_0042645D: push eax
  loc_0042645E: push 00410F18h ; " by an amount different from "
  loc_00426463: call edi
  loc_00426465: mov edx, eax
  loc_00426467: lea ecx, var_24
  loc_0042646A: call __vbaStrMove
  loc_0042646C: mov edx, [0043005Ch]
  loc_00426472: push eax
  loc_00426473: push edx
  loc_00426474: call [0040107Ch] ; __vbaStrR4
  loc_0042647A: mov edx, eax
  loc_0042647C: lea ecx, var_28
  loc_0042647F: call __vbaStrMove
  loc_00426481: push eax
  loc_00426482: call edi
  loc_00426484: mov edx, eax
  loc_00426486: lea ecx, var_2C
  loc_00426489: call __vbaStrMove
  loc_0042648B: push eax
  loc_0042648C: push 0040F028h
  loc_00426491: call edi
  loc_00426493: mov edx, eax
  loc_00426495: lea ecx, var_30
  loc_00426498: call __vbaStrMove
  loc_0042649A: push eax
  loc_0042649B: mov eax, [00430014h]
  loc_004264A0: push eax
  loc_004264A1: call edi
  loc_004264A3: mov edx, eax
  loc_004264A5: lea ecx, var_34
  loc_004264A8: call __vbaStrMove
  loc_004264AA: push eax
  loc_004264AB: push 0041083Ch
  loc_004264B0: call edi
  loc_004264B2: mov edx, eax
  loc_004264B4: lea ecx, var_38
  loc_004264B7: call __vbaStrMove
  loc_004264B9: mov ecx, ebx
  loc_004264BB: mov ebx, var_48
  loc_004264BE: push eax
  loc_004264BF: push ebx
  loc_004264C0: call [ecx+0000019Ch]
  loc_004264C6: test eax, eax
  loc_004264C8: fnclex
  loc_004264CA: jge 004264DEh
  loc_004264CC: push 0000019Ch
  loc_004264D1: push 0040E390h
  loc_004264D6: push ebx
  loc_004264D7: push eax
  loc_004264D8: call [00401030h] ; __vbaHresultCheckObj
  loc_004264DE: lea edx, var_38
  loc_004264E1: lea eax, var_34
  loc_004264E4: push edx
  loc_004264E5: lea ecx, var_30
  loc_004264E8: push eax
  loc_004264E9: lea edx, var_2C
  loc_004264EC: push ecx
  loc_004264ED: lea eax, var_28
  loc_004264F0: push edx
  loc_004264F1: lea ecx, var_24
  loc_004264F4: push eax
  loc_004264F5: lea edx, var_20
  loc_004264F8: push ecx
  loc_004264F9: lea eax, var_1C
  loc_004264FC: push edx
  loc_004264FD: lea ecx, var_18
  loc_00426500: push eax
  loc_00426501: push ecx
  loc_00426502: push 00000009h
  loc_00426504: call [004010E4h] ; __vbaFreeStrList
  loc_0042650A: add esp, 00000028h
  loc_0042650D: lea ecx, var_44
  loc_00426510: call [00401114h] ; __vbaFreeObj
  loc_00426516: mov eax, Me
  loc_00426519: push eax
  loc_0042651A: mov edx, [eax]
  loc_0042651C: call [edx+0000031Ch]
  loc_00426522: push eax
  loc_00426523: lea eax, var_44
  loc_00426526: push eax
  loc_00426527: call [0040103Ch] ; __vbaObjSet
  loc_0042652D: mov ecx, [0043001Ch]
  loc_00426533: mov ebx, [eax]
  loc_00426535: push 00410E5Ch ; "Is a one "
  loc_0042653A: push ecx
  loc_0042653B: mov var_48, eax
  loc_0042653E: call edi
  loc_00426540: mov edx, eax
  loc_00426542: lea ecx, var_18
  loc_00426545: call __vbaStrMove
  loc_00426547: push eax
  loc_00426548: push 00410E74h ; " increase in "
  loc_0042654D: call edi
  loc_0042654F: mov edx, eax
  loc_00426551: lea ecx, var_1C
  loc_00426554: call __vbaStrMove
  loc_00426556: mov edx, [00430018h]
  loc_0042655C: push eax
  loc_0042655D: push edx
  loc_0042655E: call edi
  loc_00426560: mov edx, eax
  loc_00426562: lea ecx, var_20
  loc_00426565: call __vbaStrMove
  loc_00426567: push eax
  loc_00426568: push 00410E94h ; " associated with an increase in the mean of "
  loc_0042656D: call edi
  loc_0042656F: mov edx, eax
  loc_00426571: lea ecx, var_24
  loc_00426574: call __vbaStrMove
  loc_00426576: push eax
  loc_00426577: mov eax, [00430010h]
  loc_0042657C: push eax
  loc_0042657D: call edi
  loc_0042657F: mov edx, eax
  loc_00426581: lea ecx, var_28
  loc_00426584: call __vbaStrMove
  loc_00426586: push eax
  loc_00426587: push 00410F58h ; " of less than "
  loc_0042658C: call edi
  loc_0042658E: mov edx, eax
  loc_00426590: lea ecx, var_2C
  loc_00426593: call __vbaStrMove
  loc_00426595: mov ecx, [0043005Ch]
  loc_0042659B: push eax
  loc_0042659C: push ecx
  loc_0042659D: call [0040107Ch] ; __vbaStrR4
  loc_004265A3: mov edx, eax
  loc_004265A5: lea ecx, var_30
  loc_004265A8: call __vbaStrMove
  loc_004265AA: push eax
  loc_004265AB: call edi
  loc_004265AD: mov edx, eax
  loc_004265AF: lea ecx, var_34
  loc_004265B2: call __vbaStrMove
  loc_004265B4: push eax
  loc_004265B5: push 0040F028h
  loc_004265BA: call edi
  loc_004265BC: mov edx, eax
  loc_004265BE: lea ecx, var_38
  loc_004265C1: call __vbaStrMove
  loc_004265C3: mov edx, [00430014h]
  loc_004265C9: push eax
  loc_004265CA: push edx
  loc_004265CB: call edi
  loc_004265CD: mov edx, eax
  loc_004265CF: lea ecx, var_3C
  loc_004265D2: call __vbaStrMove
  loc_004265D4: push eax
  loc_004265D5: push 0041083Ch
  loc_004265DA: call edi
  loc_004265DC: mov edx, eax
  loc_004265DE: lea ecx, var_40
  loc_004265E1: call __vbaStrMove
  loc_004265E3: mov var_80, ebx
  loc_004265E6: mov ebx, var_48
  loc_004265E9: push eax
  loc_004265EA: mov eax, var_80
  loc_004265ED: push ebx
  loc_004265EE: call [eax+0000019Ch]
  loc_004265F4: test eax, eax
  loc_004265F6: fnclex
  loc_004265F8: jge 0042660Ch
  loc_004265FA: push 0000019Ch
  loc_004265FF: push 0040E390h
  loc_00426604: push ebx
  loc_00426605: push eax
  loc_00426606: call [00401030h] ; __vbaHresultCheckObj
  loc_0042660C: lea ecx, var_40
  loc_0042660F: lea edx, var_3C
  loc_00426612: push ecx
  loc_00426613: lea eax, var_38
  loc_00426616: push edx
  loc_00426617: lea ecx, var_34
  loc_0042661A: push eax
  loc_0042661B: lea edx, var_30
  loc_0042661E: push ecx
  loc_0042661F: lea eax, var_2C
  loc_00426622: push edx
  loc_00426623: lea ecx, var_28
  loc_00426626: push eax
  loc_00426627: lea edx, var_24
  loc_0042662A: push ecx
  loc_0042662B: lea eax, var_20
  loc_0042662E: push edx
  loc_0042662F: lea ecx, var_1C
  loc_00426632: push eax
  loc_00426633: lea edx, var_18
  loc_00426636: push ecx
  loc_00426637: push edx
  loc_00426638: push 0000000Bh
  loc_0042663A: call [004010E4h] ; __vbaFreeStrList
  loc_00426640: add esp, 00000030h
  loc_00426643: lea ecx, var_44
  loc_00426646: call [00401114h] ; __vbaFreeObj
  loc_0042664C: mov eax, Me
  loc_0042664F: push eax
  loc_00426650: mov ecx, [eax]
  loc_00426652: call [ecx+0000033Ch]
  loc_00426658: lea edx, var_44
  loc_0042665B: push eax
  loc_0042665C: push edx
  loc_0042665D: call [0040103Ch] ; __vbaObjSet
  loc_00426663: mov ebx, [eax]
  loc_00426665: mov var_48, eax
  loc_00426668: mov eax, [0043005Ch]
  loc_0042666D: push 00410F7Ch ; "Is a one unit increase in X associated with an increase in the mean of Y of more than "
  loc_00426672: push eax
  loc_00426673: call [0040107Ch] ; __vbaStrR4
  loc_00426679: mov edx, eax
  loc_0042667B: lea ecx, var_18
  loc_0042667E: call __vbaStrMove
  loc_00426680: push eax
  loc_00426681: call edi
  loc_00426683: mov edx, eax
  loc_00426685: lea ecx, var_1C
  loc_00426688: call __vbaStrMove
  loc_0042668A: push eax
  loc_0042668B: push 00411030h ; " units?"
  loc_00426690: call edi
  loc_00426692: mov edx, eax
  loc_00426694: lea ecx, var_20
  loc_00426697: call __vbaStrMove
  loc_00426699: mov ecx, ebx
  loc_0042669B: mov ebx, var_48
  loc_0042669E: push eax
  loc_0042669F: push ebx
  loc_004266A0: call [ecx+00000054h]
  loc_004266A3: test eax, eax
  loc_004266A5: fnclex
  loc_004266A7: jge 004266B8h
  loc_004266A9: push 00000054h
  loc_004266AB: push 0040E390h
  loc_004266B0: push ebx
  loc_004266B1: push eax
  loc_004266B2: call [00401030h] ; __vbaHresultCheckObj
  loc_004266B8: lea edx, var_20
  loc_004266BB: lea eax, var_1C
  loc_004266BE: push edx
  loc_004266BF: lea ecx, var_18
  loc_004266C2: push eax
  loc_004266C3: push ecx
  loc_004266C4: push 00000003h
  loc_004266C6: call [004010E4h] ; __vbaFreeStrList
  loc_004266CC: add esp, 00000010h
  loc_004266CF: lea ecx, var_44
  loc_004266D2: call [00401114h] ; __vbaFreeObj
  loc_004266D8: mov eax, Me
  loc_004266DB: push eax
  loc_004266DC: mov edx, [eax]
  loc_004266DE: call [edx+00000334h]
  loc_004266E4: push eax
  loc_004266E5: lea eax, var_44
  loc_004266E8: push eax
  loc_004266E9: call [0040103Ch] ; __vbaObjSet
  loc_004266EF: mov ecx, [0043005Ch]
  loc_004266F5: mov ebx, [eax]
  loc_004266F7: push 004110FCh ; "Can we say that increases in X are associated with increases in the mean of Y by an amount different from "
  loc_004266FC: push ecx
  loc_004266FD: mov var_48, eax
  loc_00426700: call [0040107Ch] ; __vbaStrR4
  loc_00426706: mov edx, eax
  loc_00426708: lea ecx, var_18
  loc_0042670B: call __vbaStrMove
  loc_0042670D: push eax
  loc_0042670E: call edi
  loc_00426710: mov edx, eax
  loc_00426712: lea ecx, var_1C
  loc_00426715: call __vbaStrMove
  loc_00426717: push eax
  loc_00426718: push 004111D8h ; " units."
  loc_0042671D: call edi
  loc_0042671F: mov edx, eax
  loc_00426721: lea ecx, var_20
  loc_00426724: call __vbaStrMove
  loc_00426726: mov edx, ebx
  loc_00426728: mov ebx, var_48
  loc_0042672B: push eax
  loc_0042672C: push ebx
  loc_0042672D: call [edx+00000054h]
  loc_00426730: test eax, eax
  loc_00426732: fnclex
  loc_00426734: jge 00426745h
  loc_00426736: push 00000054h
  loc_00426738: push 0040E390h
  loc_0042673D: push ebx
  loc_0042673E: push eax
  loc_0042673F: call [00401030h] ; __vbaHresultCheckObj
  loc_00426745: lea eax, var_20
  loc_00426748: lea ecx, var_1C
  loc_0042674B: push eax
  loc_0042674C: lea edx, var_18
  loc_0042674F: push ecx
  loc_00426750: push edx
  loc_00426751: push 00000003h
  loc_00426753: call [004010E4h] ; __vbaFreeStrList
  loc_00426759: add esp, 00000010h
  loc_0042675C: lea ecx, var_44
  loc_0042675F: call [00401114h] ; __vbaFreeObj
  loc_00426765: mov ebx, Me
  loc_00426768: push ebx
  loc_00426769: mov eax, [ebx]
  loc_0042676B: call [eax+00000330h]
  loc_00426771: lea ecx, var_44
  loc_00426774: push eax
  loc_00426775: push ecx
  loc_00426776: call [0040103Ch] ; __vbaObjSet
  loc_0042677C: mov edx, [eax]
  loc_0042677E: push 0040F040h
  loc_00426783: push eax
  loc_00426784: mov var_48, eax
  loc_00426787: call [edx+00000054h]
  loc_0042678A: test eax, eax
  loc_0042678C: fnclex
  loc_0042678E: jge 004267A2h
  loc_00426790: mov ecx, var_48
  loc_00426793: push 00000054h
  loc_00426795: push 0040E390h
  loc_0042679A: push ecx
  loc_0042679B: push eax
  loc_0042679C: call [00401030h] ; __vbaHresultCheckObj
  loc_004267A2: lea ecx, var_44
  loc_004267A5: call [00401114h] ; __vbaFreeObj
  loc_004267AB: mov edx, [ebx]
  loc_004267AD: push ebx
  loc_004267AE: call [edx+0000032Ch]
  loc_004267B4: push eax
  loc_004267B5: lea eax, var_44
  loc_004267B8: push eax
  loc_004267B9: call [0040103Ch] ; __vbaObjSet
  loc_004267BF: mov ecx, [eax]
  loc_004267C1: push 0040F040h
  loc_004267C6: push eax
  loc_004267C7: mov var_48, eax
  loc_004267CA: call [ecx+00000054h]
  loc_004267CD: test eax, eax
  loc_004267CF: fnclex
  loc_004267D1: jge 004267E5h
  loc_004267D3: mov edx, var_48
  loc_004267D6: push 00000054h
  loc_004267D8: push 0040E390h
  loc_004267DD: push edx
  loc_004267DE: push eax
  loc_004267DF: call [00401030h] ; __vbaHresultCheckObj
  loc_004267E5: lea ecx, var_44
  loc_004267E8: call [00401114h] ; __vbaFreeObj
  loc_004267EE: mov eax, [ebx]
  loc_004267F0: push ebx
  loc_004267F1: call [eax+00000328h]
  loc_004267F7: lea ecx, var_44
  loc_004267FA: push eax
  loc_004267FB: push ecx
  loc_004267FC: call [0040103Ch] ; __vbaObjSet
  loc_00426802: mov edx, [eax]
  loc_00426804: push 0040F040h
  loc_00426809: push eax
  loc_0042680A: mov var_48, eax
  loc_0042680D: call [edx+00000054h]
  loc_00426810: test eax, eax
  loc_00426812: fnclex
  loc_00426814: jge 00426828h
  loc_00426816: mov ecx, var_48
  loc_00426819: push 00000054h
  loc_0042681B: push 0040E390h
  loc_00426820: push ecx
  loc_00426821: push eax
  loc_00426822: call [00401030h] ; __vbaHresultCheckObj
  loc_00426828: lea ecx, var_44
  loc_0042682B: call [00401114h] ; __vbaFreeObj
  loc_00426831: mov edx, [ebx]
  loc_00426833: push ebx
  loc_00426834: call [edx+0000031Ch]
  loc_0042683A: push eax
  loc_0042683B: lea eax, var_44
  loc_0042683E: push eax
  loc_0042683F: call [0040103Ch] ; __vbaObjSet
  loc_00426845: mov ecx, [0043005Ch]
  loc_0042684B: mov ebx, [eax]
  loc_0042684D: push 004111ECh ; "Is a one unit increase in X associated with an increase in the mean of Y of less than "
  loc_00426852: push ecx
  loc_00426853: mov var_48, eax
  loc_00426856: call [0040107Ch] ; __vbaStrR4
  loc_0042685C: mov edx, eax
  loc_0042685E: lea ecx, var_18
  loc_00426861: call __vbaStrMove
  loc_00426863: push eax
  loc_00426864: call edi
  loc_00426866: mov edx, eax
  loc_00426868: lea ecx, var_1C
  loc_0042686B: call __vbaStrMove
  loc_0042686D: push eax
  loc_0042686E: push 00411030h ; " units?"
  loc_00426873: call edi
  loc_00426875: mov edx, eax
  loc_00426877: lea ecx, var_20
  loc_0042687A: call __vbaStrMove
  loc_0042687C: mov esi, var_48
  loc_0042687F: push eax
  loc_00426880: push esi
  loc_00426881: call [ebx+00000054h]
  loc_00426884: test eax, eax
  loc_00426886: fnclex
  loc_00426888: jge 00427031h
  loc_0042688E: jmp 00427022h
  loc_00426893: mov edx, [ebx]
  loc_00426895: push ebx
  loc_00426896: call [edx+00000330h]
  loc_0042689C: push eax
  loc_0042689D: lea eax, var_44
  loc_004268A0: push eax
  loc_004268A1: call [0040103Ch] ; __vbaObjSet
  loc_004268A7: mov ecx, [eax]
  loc_004268A9: push 00000000h
  loc_004268AB: push eax
  loc_004268AC: mov var_48, eax
  loc_004268AF: call [ecx+0000009Ch]
  loc_004268B5: test eax, eax
  loc_004268B7: fnclex
  loc_004268B9: jge 004268D0h
  loc_004268BB: mov edx, var_48
  loc_004268BE: push 0000009Ch
  loc_004268C3: push 0040E390h
  loc_004268C8: push edx
  loc_004268C9: push eax
  loc_004268CA: call [00401030h] ; __vbaHresultCheckObj
  loc_004268D0: lea ecx, var_44
  loc_004268D3: call [00401114h] ; __vbaFreeObj
  loc_004268D9: mov eax, [ebx]
  loc_004268DB: push ebx
  loc_004268DC: call [eax+0000032Ch]
  loc_004268E2: lea ecx, var_44
  loc_004268E5: push eax
  loc_004268E6: push ecx
  loc_004268E7: call [0040103Ch] ; __vbaObjSet
  loc_004268ED: mov edx, [eax]
  loc_004268EF: push 00000000h
  loc_004268F1: push eax
  loc_004268F2: mov var_48, eax
  loc_004268F5: call [edx+0000009Ch]
  loc_004268FB: test eax, eax
  loc_004268FD: fnclex
  loc_004268FF: jge 00426916h
  loc_00426901: mov ecx, var_48
  loc_00426904: push 0000009Ch
  loc_00426909: push 0040E390h
  loc_0042690E: push ecx
  loc_0042690F: push eax
  loc_00426910: call [00401030h] ; __vbaHresultCheckObj
  loc_00426916: lea ecx, var_44
  loc_00426919: call [00401114h] ; __vbaFreeObj
  loc_0042691F: mov edx, [ebx]
  loc_00426921: push ebx
  loc_00426922: call [edx+00000328h]
  loc_00426928: push eax
  loc_00426929: lea eax, var_44
  loc_0042692C: push eax
  loc_0042692D: call [0040103Ch] ; __vbaObjSet
  loc_00426933: mov ecx, [eax]
  loc_00426935: push 00000000h
  loc_00426937: push eax
  loc_00426938: mov var_48, eax
  loc_0042693B: call [ecx+0000009Ch]
  loc_00426941: test eax, eax
  loc_00426943: fnclex
  loc_00426945: jge 0042695Ch
  loc_00426947: mov edx, var_48
  loc_0042694A: push 0000009Ch
  loc_0042694F: push 0040E390h
  loc_00426954: push edx
  loc_00426955: push eax
  loc_00426956: call [00401030h] ; __vbaHresultCheckObj
  loc_0042695C: lea ecx, var_44
  loc_0042695F: call [00401114h] ; __vbaFreeObj
  loc_00426965: mov eax, [ebx]
  loc_00426967: push ebx
  loc_00426968: call [eax+00000334h]
  loc_0042696E: lea ecx, var_44
  loc_00426971: push eax
  loc_00426972: push ecx
  loc_00426973: call [0040103Ch] ; __vbaObjSet
  loc_00426979: mov edx, [eax]
  loc_0042697B: push 44480000h
  loc_00426980: push eax
  loc_00426981: mov var_48, eax
  loc_00426984: call [edx+0000008Ch]
  loc_0042698A: test eax, eax
  loc_0042698C: fnclex
  loc_0042698E: jge 004269A5h
  loc_00426990: mov ecx, var_48
  loc_00426993: push 0000008Ch
  loc_00426998: push 0040E390h
  loc_0042699D: push ecx
  loc_0042699E: push eax
  loc_0042699F: call [00401030h] ; __vbaHresultCheckObj
  loc_004269A5: lea ecx, var_44
  loc_004269A8: call [00401114h] ; __vbaFreeObj
  loc_004269AE: mov edx, [ebx]
  loc_004269B0: push ebx
  loc_004269B1: call [edx+0000033Ch]
  loc_004269B7: push eax
  loc_004269B8: lea eax, var_44
  loc_004269BB: push eax
  loc_004269BC: call [0040103Ch] ; __vbaObjSet
  loc_004269C2: mov ecx, [0043001Ch]
  loc_004269C8: mov ebx, [eax]
  loc_004269CA: push 00410E5Ch ; "Is a one "
  loc_004269CF: push ecx
  loc_004269D0: mov var_48, eax
  loc_004269D3: call edi
  loc_004269D5: mov edx, eax
  loc_004269D7: lea ecx, var_18
  loc_004269DA: call __vbaStrMove
  loc_004269DC: push eax
  loc_004269DD: push 00410E74h ; " increase in "
  loc_004269E2: call edi
  loc_004269E4: mov edx, eax
  loc_004269E6: lea ecx, var_1C
  loc_004269E9: call __vbaStrMove
  loc_004269EB: mov edx, [00430018h]
  loc_004269F1: push eax
  loc_004269F2: push edx
  loc_004269F3: call edi
  loc_004269F5: mov edx, eax
  loc_004269F7: lea ecx, var_20
  loc_004269FA: call __vbaStrMove
  loc_004269FC: push eax
  loc_004269FD: push 00411044h ; " associated with a decrease in the mean of "
  loc_00426A02: call edi
  loc_00426A04: mov edx, eax
  loc_00426A06: lea ecx, var_24
  loc_00426A09: call __vbaStrMove
  loc_00426A0B: push eax
  loc_00426A0C: mov eax, [00430010h]
  loc_00426A11: push eax
  loc_00426A12: call edi
  loc_00426A14: mov edx, eax
  loc_00426A16: lea ecx, var_28
  loc_00426A19: call __vbaStrMove
  loc_00426A1B: push eax
  loc_00426A1C: push 00410EF4h ; " of more than "
  loc_00426A21: call edi
  loc_00426A23: mov edx, eax
  loc_00426A25: lea ecx, var_2C
  loc_00426A28: call __vbaStrMove
  loc_00426A2A: fld real4 ptr [0043005Ch]
  loc_00426A30: fabs
  loc_00426A32: push eax
  loc_00426A33: push ecx
  loc_00426A34: fnstsw ax
  loc_00426A36: test al, 0Dh
  loc_00426A38: jnz 004270BAh
  loc_00426A3E: fstp real4 ptr var_8C
  loc_00426A44: fld real4 ptr var_8C
  loc_00426A4A: fstp real4 ptr [esp]
  loc_00426A4D: call [0040107Ch] ; __vbaStrR4
  loc_00426A53: mov edx, eax
  loc_00426A55: lea ecx, var_30
  loc_00426A58: call __vbaStrMove
  loc_00426A5A: push eax
  loc_00426A5B: call edi
  loc_00426A5D: mov edx, eax
  loc_00426A5F: lea ecx, var_34
  loc_00426A62: call __vbaStrMove
  loc_00426A64: push eax
  loc_00426A65: push 0040F028h
  loc_00426A6A: call edi
  loc_00426A6C: mov edx, eax
  loc_00426A6E: lea ecx, var_38
  loc_00426A71: call __vbaStrMove
  loc_00426A73: mov ecx, [00430014h]
  loc_00426A79: push eax
  loc_00426A7A: push ecx
  loc_00426A7B: call edi
  loc_00426A7D: mov edx, eax
  loc_00426A7F: lea ecx, var_3C
  loc_00426A82: call __vbaStrMove
  loc_00426A84: push eax
  loc_00426A85: push 0041083Ch
  loc_00426A8A: call edi
  loc_00426A8C: mov edx, eax
  loc_00426A8E: lea ecx, var_40
  loc_00426A91: call __vbaStrMove
  loc_00426A93: mov edx, ebx
  loc_00426A95: mov ebx, var_48
  loc_00426A98: push eax
  loc_00426A99: push ebx
  loc_00426A9A: call [edx+0000019Ch]
  loc_00426AA0: test eax, eax
  loc_00426AA2: fnclex
  loc_00426AA4: jge 00426AB8h
  loc_00426AA6: push 0000019Ch
  loc_00426AAB: push 0040E390h
  loc_00426AB0: push ebx
  loc_00426AB1: push eax
  loc_00426AB2: call [00401030h] ; __vbaHresultCheckObj
  loc_00426AB8: lea eax, var_40
  loc_00426ABB: lea ecx, var_3C
  loc_00426ABE: push eax
  loc_00426ABF: lea edx, var_38
  loc_00426AC2: push ecx
  loc_00426AC3: lea eax, var_34
  loc_00426AC6: push edx
  loc_00426AC7: lea ecx, var_30
  loc_00426ACA: push eax
  loc_00426ACB: lea edx, var_2C
  loc_00426ACE: push ecx
  loc_00426ACF: lea eax, var_28
  loc_00426AD2: push edx
  loc_00426AD3: lea ecx, var_24
  loc_00426AD6: push eax
  loc_00426AD7: lea edx, var_20
  loc_00426ADA: push ecx
  loc_00426ADB: lea eax, var_1C
  loc_00426ADE: push edx
  loc_00426ADF: lea ecx, var_18
  loc_00426AE2: push eax
  loc_00426AE3: push ecx
  loc_00426AE4: push 0000000Bh
  loc_00426AE6: call [004010E4h] ; __vbaFreeStrList
  loc_00426AEC: add esp, 00000030h
  loc_00426AEF: lea ecx, var_44
  loc_00426AF2: call [00401114h] ; __vbaFreeObj
  loc_00426AF8: mov eax, Me
  loc_00426AFB: push eax
  loc_00426AFC: mov edx, [eax]
  loc_00426AFE: call [edx+00000334h]
  loc_00426B04: push eax
  loc_00426B05: lea eax, var_44
  loc_00426B08: push eax
  loc_00426B09: call [0040103Ch] ; __vbaObjSet
  loc_00426B0F: mov ecx, [0043001Ch]
  loc_00426B15: mov ebx, [eax]
  loc_00426B17: push 004110A0h ; "Can we say that a one "
  loc_00426B1C: push ecx
  loc_00426B1D: mov var_48, eax
  loc_00426B20: call edi
  loc_00426B22: mov edx, eax
  loc_00426B24: lea ecx, var_18
  loc_00426B27: call __vbaStrMove
  loc_00426B29: push eax
  loc_00426B2A: push 00410E74h ; " increase in "
  loc_00426B2F: call edi
  loc_00426B31: mov edx, eax
  loc_00426B33: lea ecx, var_1C
  loc_00426B36: call __vbaStrMove
  loc_00426B38: mov edx, [00430018h]
  loc_00426B3E: push eax
  loc_00426B3F: push edx
  loc_00426B40: call edi
  loc_00426B42: mov edx, eax
  loc_00426B44: lea ecx, var_20
  loc_00426B47: call __vbaStrMove
  loc_00426B49: push eax
  loc_00426B4A: push 004112C4h ; " are associated with a decrease in the mean of "
  loc_00426B4F: call edi
  loc_00426B51: mov edx, eax
  loc_00426B53: lea ecx, var_24
  loc_00426B56: call __vbaStrMove
  loc_00426B58: push eax
  loc_00426B59: mov eax, [00430010h]
  loc_00426B5E: push eax
  loc_00426B5F: call edi
  loc_00426B61: mov edx, eax
  loc_00426B63: lea ecx, var_28
  loc_00426B66: call __vbaStrMove
  loc_00426B68: push eax
  loc_00426B69: push 00410F18h ; " by an amount different from "
  loc_00426B6E: call edi
  loc_00426B70: mov edx, eax
  loc_00426B72: lea ecx, var_2C
  loc_00426B75: call __vbaStrMove
  loc_00426B77: fld real4 ptr [0043005Ch]
  loc_00426B7D: fabs
  loc_00426B7F: push eax
  loc_00426B80: fnstsw ax
  loc_00426B82: test al, 0Dh
  loc_00426B84: jnz 004270BAh
  loc_00426B8A: fstp real4 ptr var_94
  loc_00426B90: fld real4 ptr var_94
  loc_00426B96: push ecx
  loc_00426B97: fstp real4 ptr [esp]
  loc_00426B9A: call [0040107Ch] ; __vbaStrR4
  loc_00426BA0: mov edx, eax
  loc_00426BA2: lea ecx, var_30
  loc_00426BA5: call __vbaStrMove
  loc_00426BA7: push eax
  loc_00426BA8: call edi
  loc_00426BAA: mov edx, eax
  loc_00426BAC: lea ecx, var_34
  loc_00426BAF: call __vbaStrMove
  loc_00426BB1: push eax
  loc_00426BB2: push 0040F028h
  loc_00426BB7: call edi
  loc_00426BB9: mov edx, eax
  loc_00426BBB: lea ecx, var_38
  loc_00426BBE: call __vbaStrMove
  loc_00426BC0: mov ecx, [00430014h]
  loc_00426BC6: push eax
  loc_00426BC7: push ecx
  loc_00426BC8: call edi
  loc_00426BCA: mov edx, eax
  loc_00426BCC: lea ecx, var_3C
  loc_00426BCF: call __vbaStrMove
  loc_00426BD1: push eax
  loc_00426BD2: push 0041083Ch
  loc_00426BD7: call edi
  loc_00426BD9: mov edx, eax
  loc_00426BDB: lea ecx, var_40
  loc_00426BDE: call __vbaStrMove
  loc_00426BE0: mov edx, ebx
  loc_00426BE2: mov ebx, var_48
  loc_00426BE5: push eax
  loc_00426BE6: push ebx
  loc_00426BE7: call [edx+0000019Ch]
  loc_00426BED: test eax, eax
  loc_00426BEF: fnclex
  loc_00426BF1: jge 00426C05h
  loc_00426BF3: push 0000019Ch
  loc_00426BF8: push 0040E390h
  loc_00426BFD: push ebx
  loc_00426BFE: push eax
  loc_00426BFF: call [00401030h] ; __vbaHresultCheckObj
  loc_00426C05: lea eax, var_40
  loc_00426C08: lea ecx, var_3C
  loc_00426C0B: push eax
  loc_00426C0C: lea edx, var_38
  loc_00426C0F: push ecx
  loc_00426C10: lea eax, var_34
  loc_00426C13: push edx
  loc_00426C14: lea ecx, var_30
  loc_00426C17: push eax
  loc_00426C18: lea edx, var_2C
  loc_00426C1B: push ecx
  loc_00426C1C: lea eax, var_28
  loc_00426C1F: push edx
  loc_00426C20: lea ecx, var_24
  loc_00426C23: push eax
  loc_00426C24: lea edx, var_20
  loc_00426C27: push ecx
  loc_00426C28: lea eax, var_1C
  loc_00426C2B: push edx
  loc_00426C2C: lea ecx, var_18
  loc_00426C2F: push eax
  loc_00426C30: push ecx
  loc_00426C31: push 0000000Bh
  loc_00426C33: call [004010E4h] ; __vbaFreeStrList
  loc_00426C39: add esp, 00000030h
  loc_00426C3C: lea ecx, var_44
  loc_00426C3F: call [00401114h] ; __vbaFreeObj
  loc_00426C45: mov eax, Me
  loc_00426C48: push eax
  loc_00426C49: mov edx, [eax]
  loc_00426C4B: call [edx+0000031Ch]
  loc_00426C51: push eax
  loc_00426C52: lea eax, var_44
  loc_00426C55: push eax
  loc_00426C56: call [0040103Ch] ; __vbaObjSet
  loc_00426C5C: mov ecx, [0043001Ch]
  loc_00426C62: mov ebx, [eax]
  loc_00426C64: push 00410E5Ch ; "Is a one "
  loc_00426C69: push ecx
  loc_00426C6A: mov var_48, eax
  loc_00426C6D: call edi
  loc_00426C6F: mov edx, eax
  loc_00426C71: lea ecx, var_18
  loc_00426C74: call __vbaStrMove
  loc_00426C76: push eax
  loc_00426C77: push 00410E74h ; " increase in "
  loc_00426C7C: call edi
  loc_00426C7E: mov edx, eax
  loc_00426C80: lea ecx, var_1C
  loc_00426C83: call __vbaStrMove
  loc_00426C85: mov edx, [00430018h]
  loc_00426C8B: push eax
  loc_00426C8C: push edx
  loc_00426C8D: call edi
  loc_00426C8F: mov edx, eax
  loc_00426C91: lea ecx, var_20
  loc_00426C94: call __vbaStrMove
  loc_00426C96: push eax
  loc_00426C97: push 00411044h ; " associated with a decrease in the mean of "
  loc_00426C9C: call edi
  loc_00426C9E: mov edx, eax
  loc_00426CA0: lea ecx, var_24
  loc_00426CA3: call __vbaStrMove
  loc_00426CA5: push eax
  loc_00426CA6: mov eax, [00430010h]
  loc_00426CAB: push eax
  loc_00426CAC: call edi
  loc_00426CAE: mov edx, eax
  loc_00426CB0: lea ecx, var_28
  loc_00426CB3: call __vbaStrMove
  loc_00426CB5: push eax
  loc_00426CB6: push 00410F58h ; " of less than "
  loc_00426CBB: call edi
  loc_00426CBD: mov edx, eax
  loc_00426CBF: lea ecx, var_2C
  loc_00426CC2: call __vbaStrMove
  loc_00426CC4: fld real4 ptr [0043005Ch]
  loc_00426CCA: fabs
  loc_00426CCC: push eax
  loc_00426CCD: fnstsw ax
  loc_00426CCF: test al, 0Dh
  loc_00426CD1: jnz 004270BAh
  loc_00426CD7: fstp real4 ptr var_9C
  loc_00426CDD: fld real4 ptr var_9C
  loc_00426CE3: push ecx
  loc_00426CE4: fstp real4 ptr [esp]
  loc_00426CE7: call [0040107Ch] ; __vbaStrR4
  loc_00426CED: mov edx, eax
  loc_00426CEF: lea ecx, var_30
  loc_00426CF2: call __vbaStrMove
  loc_00426CF4: push eax
  loc_00426CF5: call edi
  loc_00426CF7: mov edx, eax
  loc_00426CF9: lea ecx, var_34
  loc_00426CFC: call __vbaStrMove
  loc_00426CFE: push eax
  loc_00426CFF: push 0040F028h
  loc_00426D04: call edi
  loc_00426D06: mov edx, eax
  loc_00426D08: lea ecx, var_38
  loc_00426D0B: call __vbaStrMove
  loc_00426D0D: mov ecx, [00430014h]
  loc_00426D13: push eax
  loc_00426D14: push ecx
  loc_00426D15: call edi
  loc_00426D17: mov edx, eax
  loc_00426D19: lea ecx, var_3C
  loc_00426D1C: call __vbaStrMove
  loc_00426D1E: push eax
  loc_00426D1F: push 0041083Ch
  loc_00426D24: call edi
  loc_00426D26: mov edx, eax
  loc_00426D28: lea ecx, var_40
  loc_00426D2B: call __vbaStrMove
  loc_00426D2D: mov edx, ebx
  loc_00426D2F: mov ebx, var_48
  loc_00426D32: push eax
  loc_00426D33: push ebx
  loc_00426D34: call [edx+0000019Ch]
  loc_00426D3A: test eax, eax
  loc_00426D3C: fnclex
  loc_00426D3E: jge 00426D52h
  loc_00426D40: push 0000019Ch
  loc_00426D45: push 0040E390h
  loc_00426D4A: push ebx
  loc_00426D4B: push eax
  loc_00426D4C: call [00401030h] ; __vbaHresultCheckObj
  loc_00426D52: lea eax, var_40
  loc_00426D55: lea ecx, var_3C
  loc_00426D58: push eax
  loc_00426D59: lea edx, var_38
  loc_00426D5C: push ecx
  loc_00426D5D: lea eax, var_34
  loc_00426D60: push edx
  loc_00426D61: lea ecx, var_30
  loc_00426D64: push eax
  loc_00426D65: lea edx, var_2C
  loc_00426D68: push ecx
  loc_00426D69: lea eax, var_28
  loc_00426D6C: push edx
  loc_00426D6D: lea ecx, var_24
  loc_00426D70: push eax
  loc_00426D71: lea edx, var_20
  loc_00426D74: push ecx
  loc_00426D75: lea eax, var_1C
  loc_00426D78: push edx
  loc_00426D79: lea ecx, var_18
  loc_00426D7C: push eax
  loc_00426D7D: push ecx
  loc_00426D7E: push 0000000Bh
  loc_00426D80: call [004010E4h] ; __vbaFreeStrList
  loc_00426D86: add esp, 00000030h
  loc_00426D89: lea ecx, var_44
  loc_00426D8C: call [00401114h] ; __vbaFreeObj
  loc_00426D92: mov eax, Me
  loc_00426D95: push eax
  loc_00426D96: mov edx, [eax]
  loc_00426D98: call [edx+0000033Ch]
  loc_00426D9E: push eax
  loc_00426D9F: lea eax, var_44
  loc_00426DA2: push eax
  loc_00426DA3: call [0040103Ch] ; __vbaObjSet
  loc_00426DA9: fld real4 ptr [0043005Ch]
  loc_00426DAF: mov ebx, [eax]
  loc_00426DB1: mov var_48, eax
  loc_00426DB4: fabs
  loc_00426DB6: fnstsw ax
  loc_00426DB8: test al, 0Dh
  loc_00426DBA: jnz 004270BAh
  loc_00426DC0: fstp real4 ptr var_A4
  loc_00426DC6: fld real4 ptr var_A4
  loc_00426DCC: push 00411328h ; "Is a one unit increase in X associated with a decrease in the mean of Y of more than "
  loc_00426DD1: push ecx
  loc_00426DD2: fstp real4 ptr [esp]
  loc_00426DD5: call [0040107Ch] ; __vbaStrR4
  loc_00426DDB: mov edx, eax
  loc_00426DDD: lea ecx, var_18
  loc_00426DE0: call __vbaStrMove
  loc_00426DE2: push eax
  loc_00426DE3: call edi
  loc_00426DE5: mov edx, eax
  loc_00426DE7: lea ecx, var_1C
  loc_00426DEA: call __vbaStrMove
  loc_00426DEC: push eax
  loc_00426DED: push 00411030h ; " units?"
  loc_00426DF2: call edi
  loc_00426DF4: mov edx, eax
  loc_00426DF6: lea ecx, var_20
  loc_00426DF9: call __vbaStrMove
  loc_00426DFB: mov ecx, ebx
  loc_00426DFD: mov ebx, var_48
  loc_00426E00: push eax
  loc_00426E01: push ebx
  loc_00426E02: call [ecx+00000054h]
  loc_00426E05: test eax, eax
  loc_00426E07: fnclex
  loc_00426E09: jge 00426E1Ah
  loc_00426E0B: push 00000054h
  loc_00426E0D: push 0040E390h
  loc_00426E12: push ebx
  loc_00426E13: push eax
  loc_00426E14: call [00401030h] ; __vbaHresultCheckObj
  loc_00426E1A: lea edx, var_20
  loc_00426E1D: lea eax, var_1C
  loc_00426E20: push edx
  loc_00426E21: lea ecx, var_18
  loc_00426E24: push eax
  loc_00426E25: push ecx
  loc_00426E26: push 00000003h
  loc_00426E28: call [004010E4h] ; __vbaFreeStrList
  loc_00426E2E: add esp, 00000010h
  loc_00426E31: lea ecx, var_44
  loc_00426E34: call [00401114h] ; __vbaFreeObj
  loc_00426E3A: mov eax, Me
  loc_00426E3D: push eax
  loc_00426E3E: mov edx, [eax]
  loc_00426E40: call [edx+00000334h]
  loc_00426E46: push eax
  loc_00426E47: lea eax, var_44
  loc_00426E4A: push eax
  loc_00426E4B: call [0040103Ch] ; __vbaObjSet
  loc_00426E51: fld real4 ptr [0043005Ch]
  loc_00426E57: mov ebx, [eax]
  loc_00426E59: mov var_48, eax
  loc_00426E5C: fabs
  loc_00426E5E: fnstsw ax
  loc_00426E60: test al, 0Dh
  loc_00426E62: jnz 004270BAh
  loc_00426E68: fstp real4 ptr var_AC
  loc_00426E6E: fld real4 ptr var_AC
  loc_00426E74: push 004114ACh ; "Can we say that a one unit increase in X is associated with a decrease in the mean of Y by an amount different from "
  loc_00426E79: push ecx
  loc_00426E7A: fstp real4 ptr [esp]
  loc_00426E7D: call [0040107Ch] ; __vbaStrR4
  loc_00426E83: mov edx, eax
  loc_00426E85: lea ecx, var_18
  loc_00426E88: call __vbaStrMove
  loc_00426E8A: push eax
  loc_00426E8B: call edi
  loc_00426E8D: mov edx, eax
  loc_00426E8F: lea ecx, var_1C
  loc_00426E92: call __vbaStrMove
  loc_00426E94: push eax
  loc_00426E95: push 0041159Ch ; "units."
  loc_00426E9A: call edi
  loc_00426E9C: mov edx, eax
  loc_00426E9E: lea ecx, var_20
  loc_00426EA1: call __vbaStrMove
  loc_00426EA3: mov ecx, ebx
  loc_00426EA5: mov ebx, var_48
  loc_00426EA8: push eax
  loc_00426EA9: push ebx
  loc_00426EAA: call [ecx+00000054h]
  loc_00426EAD: test eax, eax
  loc_00426EAF: fnclex
  loc_00426EB1: jge 00426EC2h
  loc_00426EB3: push 00000054h
  loc_00426EB5: push 0040E390h
  loc_00426EBA: push ebx
  loc_00426EBB: push eax
  loc_00426EBC: call [00401030h] ; __vbaHresultCheckObj
  loc_00426EC2: lea edx, var_20
  loc_00426EC5: lea eax, var_1C
  loc_00426EC8: push edx
  loc_00426EC9: lea ecx, var_18
  loc_00426ECC: push eax
  loc_00426ECD: push ecx
  loc_00426ECE: push 00000003h
  loc_00426ED0: call [004010E4h] ; __vbaFreeStrList
  loc_00426ED6: add esp, 00000010h
  loc_00426ED9: lea ecx, var_44
  loc_00426EDC: call [00401114h] ; __vbaFreeObj
  loc_00426EE2: mov ebx, Me
  loc_00426EE5: push ebx
  loc_00426EE6: mov edx, [ebx]
  loc_00426EE8: call [edx+00000330h]
  loc_00426EEE: push eax
  loc_00426EEF: lea eax, var_44
  loc_00426EF2: push eax
  loc_00426EF3: call [0040103Ch] ; __vbaObjSet
  loc_00426EF9: mov ecx, [eax]
  loc_00426EFB: push 0040F040h
  loc_00426F00: push eax
  loc_00426F01: mov var_48, eax
  loc_00426F04: call [ecx+00000054h]
  loc_00426F07: test eax, eax
  loc_00426F09: fnclex
  loc_00426F0B: jge 00426F1Fh
  loc_00426F0D: mov edx, var_48
  loc_00426F10: push 00000054h
  loc_00426F12: push 0040E390h
  loc_00426F17: push edx
  loc_00426F18: push eax
  loc_00426F19: call [00401030h] ; __vbaHresultCheckObj
  loc_00426F1F: lea ecx, var_44
  loc_00426F22: call [00401114h] ; __vbaFreeObj
  loc_00426F28: mov eax, [ebx]
  loc_00426F2A: push ebx
  loc_00426F2B: call [eax+0000032Ch]
  loc_00426F31: lea ecx, var_44
  loc_00426F34: push eax
  loc_00426F35: push ecx
  loc_00426F36: call [0040103Ch] ; __vbaObjSet
  loc_00426F3C: mov edx, [eax]
  loc_00426F3E: push 0040F040h
  loc_00426F43: push eax
  loc_00426F44: mov var_48, eax
  loc_00426F47: call [edx+00000054h]
  loc_00426F4A: test eax, eax
  loc_00426F4C: fnclex
  loc_00426F4E: jge 00426F62h
  loc_00426F50: mov ecx, var_48
  loc_00426F53: push 00000054h
  loc_00426F55: push 0040E390h
  loc_00426F5A: push ecx
  loc_00426F5B: push eax
  loc_00426F5C: call [00401030h] ; __vbaHresultCheckObj
  loc_00426F62: lea ecx, var_44
  loc_00426F65: call [00401114h] ; __vbaFreeObj
  loc_00426F6B: mov edx, [ebx]
  loc_00426F6D: push ebx
  loc_00426F6E: call [edx+00000328h]
  loc_00426F74: push eax
  loc_00426F75: lea eax, var_44
  loc_00426F78: push eax
  loc_00426F79: call [0040103Ch] ; __vbaObjSet
  loc_00426F7F: mov ecx, [eax]
  loc_00426F81: push 0040F040h
  loc_00426F86: push eax
  loc_00426F87: mov var_48, eax
  loc_00426F8A: call [ecx+00000054h]
  loc_00426F8D: test eax, eax
  loc_00426F8F: fnclex
  loc_00426F91: jge 00426FA5h
  loc_00426F93: mov edx, var_48
  loc_00426F96: push 00000054h
  loc_00426F98: push 0040E390h
  loc_00426F9D: push edx
  loc_00426F9E: push eax
  loc_00426F9F: call [00401030h] ; __vbaHresultCheckObj
  loc_00426FA5: lea ecx, var_44
  loc_00426FA8: call [00401114h] ; __vbaFreeObj
  loc_00426FAE: mov eax, [ebx]
  loc_00426FB0: push ebx
  loc_00426FB1: call [eax+0000031Ch]
  loc_00426FB7: lea ecx, var_44
  loc_00426FBA: push eax
  loc_00426FBB: push ecx
  loc_00426FBC: call [0040103Ch] ; __vbaObjSet
  loc_00426FC2: fld real4 ptr [0043005Ch]
  loc_00426FC8: mov ebx, [eax]
  loc_00426FCA: mov var_48, eax
  loc_00426FCD: fabs
  loc_00426FCF: fnstsw ax
  loc_00426FD1: test al, 0Dh
  loc_00426FD3: jnz 004270BAh
  loc_00426FD9: fstp real4 ptr var_B4
  loc_00426FDF: fld real4 ptr var_B4
  loc_00426FE5: push 004115B0h ; "Is a one unit increase in X associated with a decrease in the mean of Y of less than "
  loc_00426FEA: push ecx
  loc_00426FEB: fstp real4 ptr [esp]
  loc_00426FEE: call [0040107Ch] ; __vbaStrR4
  loc_00426FF4: mov edx, eax
  loc_00426FF6: lea ecx, var_18
  loc_00426FF9: call __vbaStrMove
  loc_00426FFB: push eax
  loc_00426FFC: call edi
  loc_00426FFE: mov edx, eax
  loc_00427000: lea ecx, var_1C
  loc_00427003: call __vbaStrMove
  loc_00427005: push eax
  loc_00427006: push 00411030h ; " units?"
  loc_0042700B: call edi
  loc_0042700D: mov edx, eax
  loc_0042700F: lea ecx, var_20
  loc_00427012: call __vbaStrMove
  loc_00427014: mov esi, var_48
  loc_00427017: push eax
  loc_00427018: push esi
  loc_00427019: call [ebx+00000054h]
  loc_0042701C: test eax, eax
  loc_0042701E: fnclex
  loc_00427020: jge 00427031h
  loc_00427022: push 00000054h
  loc_00427024: push 0040E390h
  loc_00427029: push esi
  loc_0042702A: push eax
  loc_0042702B: call [00401030h] ; __vbaHresultCheckObj
  loc_00427031: lea edx, var_20
  loc_00427034: lea eax, var_1C
  loc_00427037: push edx
  loc_00427038: lea ecx, var_18
  loc_0042703B: push eax
  loc_0042703C: push ecx
  loc_0042703D: push 00000003h
  loc_0042703F: call [004010E4h] ; __vbaFreeStrList
  loc_00427045: add esp, 00000010h
  loc_00427048: lea ecx, var_44
  loc_0042704B: call [00401114h] ; __vbaFreeObj
  loc_00427051: fwait
  loc_00427052: push 0042709Bh
  loc_00427057: jmp 0042709Ah
  loc_00427059: lea edx, var_40
  loc_0042705C: lea eax, var_3C
  loc_0042705F: push edx
  loc_00427060: lea ecx, var_38
  loc_00427063: push eax
  loc_00427064: lea edx, var_34
  loc_00427067: push ecx
  loc_00427068: lea eax, var_30
  loc_0042706B: push edx
  loc_0042706C: lea ecx, var_2C
  loc_0042706F: push eax
  loc_00427070: lea edx, var_28
  loc_00427073: push ecx
  loc_00427074: lea eax, var_24
  loc_00427077: push edx
  loc_00427078: lea ecx, var_20
  loc_0042707B: push eax
  loc_0042707C: lea edx, var_1C
  loc_0042707F: push ecx
  loc_00427080: lea eax, var_18
  loc_00427083: push edx
  loc_00427084: push eax
  loc_00427085: push 0000000Bh
  loc_00427087: call [004010E4h] ; __vbaFreeStrList
  loc_0042708D: add esp, 00000030h
  loc_00427090: lea ecx, var_44
  loc_00427093: call [00401114h] ; __vbaFreeObj
  loc_00427099: ret
  loc_0042709A: ret
  loc_0042709B: mov eax, Me
  loc_0042709E: push eax
  loc_0042709F: mov ecx, [eax]
  loc_004270A1: call [ecx+00000008h]
  loc_004270A4: mov eax, var_4
  loc_004270A7: mov ecx, var_14
  loc_004270AA: pop edi
  loc_004270AB: pop esi
  loc_004270AC: mov fs:[00000000h], ecx
  loc_004270B3: pop ebx
  loc_004270B4: mov esp, ebp
  loc_004270B6: pop ebp
  loc_004270B7: retn 0004h
End Sub
