VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "TABCTL32.OCX"
Begin VB.Form frmSlope
  Caption = "Inferences on the Slope"
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 510
  ClientWidth = 8340
  ClientHeight = 4725
  StartUpPosition = 2 'CenterScreen
  Begin TabDlg.SSTab SSTab1
    Left = 120
    Top = 0
    Width = 11775
    Height = 8175
    TabIndex = 0
    OleObjectBlob = "frmSlope.frx":0000
    Begin VB.OptionButton optRight
      Caption = "Right-Sided"
      Left = 6360
      Top = 720
      Width = 2415
      Height = 435
      TabIndex = 23
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
    Begin VB.OptionButton optTwo
      Caption = "Two-Sided "
      Left = 3480
      Top = 720
      Width = 2415
      Height = 435
      TabIndex = 22
      Value = 255
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
    Begin VB.OptionButton optLeft
      Caption = "Left-Sided "
      Left = 960
      Top = 720
      Width = 2775
      Height = 435
      TabIndex = 21
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
    Begin VB.Frame Frame9
      Caption = "Conclusion:"
      Left = 120
      Top = 5400
      Width = 11175
      Height = 2535
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
      Begin VB.Label lblTestCon
        Left = 120
        Top = 480
        Width = 10695
        Height = 1815
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
    End
    Begin VB.Frame Frame8
      Caption = "R.R.: Reject Ho if"
      Left = 120
      Top = 3600
      Width = 3735
      Height = 1695
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
      Begin VB.Label lblRR
        Left = 120
        Top = 600
        Width = 3255
        Height = 975
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
    End
    Begin VB.Frame Frame7
      Caption = "Substitution"
      Left = 3960
      Top = 3600
      Width = 7575
      Height = 1695
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
      Begin VB.Label lblTSValue
        Left = 6120
        Top = 600
        Width = 1215
        Height = 615
        TabIndex = 26
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
      Begin VB.Label lblSubBottom
        Left = 2520
        Top = 1080
        Width = 2655
        Height = 495
        TabIndex = 25
        Alignment = 2 'Center
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
      Begin VB.Line Line2
        X1 = 2280
        Y1 = 840
        X2 = 6000
        Y2 = 840
      End
      Begin VB.Label Label4
        Caption = "t* ="
        Left = 240
        Top = 600
        Width = 855
        Height = 735
        TabIndex = 24
        BeginProperty Font
          Name = "MS Sans Serif"
          Size = 24
          Charset = 0
          Weight = 400
          Underline = 0 'False
          Italic = 0 'False
          Strikethrough = 0 'False
        EndProperty
      End
      Begin VB.Label lblSubtop
        Left = 2040
        Top = 360
        Width = 3975
        Height = 375
        TabIndex = 15
        Alignment = 2 'Center
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
    Begin VB.Frame Frame6
      Caption = "Test Statistic Formula"
      Left = 3960
      Top = 1440
      Width = 7575
      Height = 2175
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
      Begin VB.Label Label5
        Caption = "t*="
        Left = 120
        Top = 720
        Width = 495
        Height = 495
        TabIndex = 20
        BeginProperty Font
          Name = "MS Sans Serif"
          Size = 24
          Charset = 0
          Weight = 400
          Underline = 0 'False
          Italic = 0 'False
          Strikethrough = 0 'False
        EndProperty
      End
      Begin VB.Label Label3
        Caption = "Estimated Standard Error of Slope Estimate"
        Left = 1200
        Top = 1440
        Width = 5535
        Height = 375
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
      Begin VB.Line Line1
        X1 = 840
        Y1 = 1200
        X2 = 6840
        Y2 = 1200
      End
      Begin VB.Label Label2
        Caption = "(Estimate of Slope - Hypothsized Value)"
        Left = 960
        Top = 600
        Width = 6495
        Height = 495
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
    End
    Begin VB.Frame Frame5
      Caption = "Alternative Hypothesis"
      Left = 120
      Top = 2520
      Width = 3735
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
      Begin VB.Label lblHa
        Left = 240
        Top = 480
        Width = 3255
        Height = 495
        TabIndex = 10
        BeginProperty Font
          Name = "Symbol"
          Size = 18
          Charset = 2
          Weight = 400
          Underline = 0 'False
          Italic = 0 'False
          Strikethrough = 0 'False
        EndProperty
      End
    End
    Begin VB.Frame Frame4
      Caption = "Null Hypothesis"
      Left = 120
      Top = 1320
      Width = 3735
      Height = 1095
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
      Begin VB.Label lblHo
        Left = 240
        Top = 480
        Width = 3255
        Height = 495
        TabIndex = 9
        BeginProperty Font
          Name = "Symbol"
          Size = 18
          Charset = 2
          Weight = 400
          Underline = 0 'False
          Italic = 0 'False
          Strikethrough = 0 'False
        EndProperty
      End
    End
    Begin VB.Frame Frame3
      Caption = "Conclusion:"
      Left = -74520
      Top = 5160
      Width = 11055
      Height = 2655
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
      Begin VB.Label lblCIConclusion
        Caption = "frmSlope.frx":038C
        Left = 120
        Top = 480
        Width = 10815
        Height = 1935
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
    End
    Begin VB.Frame Frame2
      Caption = "Substitution"
      Left = -74520
      Top = 3000
      Width = 11055
      Height = 2055
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
      Begin VB.Label lblCISubstitution
        Left = 120
        Top = 360
        Width = 10695
        Height = 1455
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
    Begin VB.Frame Frame1
      Caption = "Formula"
      Left = -74520
      Top = 1320
      Width = 11055
      Height = 1455
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
      Begin VB.Label Label1
        Caption = "(Point Estimate of Slope) ±  t(n-2, 1-alpha/2) (Estimate of Standard Error of Slope Estimate)"
        Left = 120
        Top = 360
        Width = 10815
        Height = 975
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
  Begin VB.Menu mnuChange
    Caption = "&Change"
    Begin VB.Menu mnuChangeAlpha
      Caption = "&Alpha & T-Table Value"
    End
    Begin VB.Menu mnuChangePointEstimate
      Caption = "&Point Estimate and Standard Error"
    End
    Begin VB.Menu mnuChangeHypothesizedValue
      Caption = "&Hypothesized Value of Slope"
    End
    Begin VB.Menu mnuChangeNames
      Caption = "Variable &Names"
    End
    Begin VB.Menu mnuChangeUnits
      Caption = "&Units"
    End
  End
  Begin VB.Menu mnuGoTo
    Caption = "&Go To"
    Begin VB.Menu mnuIntro
      Caption = "&Introduction"
    End
    Begin VB.Menu mnustatistics
      Caption = "Statistics and Point Estimates"
    End
    Begin VB.Menu mnuQuestions
      Caption = "&Questions"
    End
    Begin VB.Menu mnuPredEst
      Caption = "Prediction and &Estimation Intervals"
    End
    Begin VB.Menu mnuAssumptions
      Caption = "&Assumptions"
    End
  End
  Begin VB.Menu mnuExit
    Caption = "&Exit"
  End
End

Attribute VB_Name = "frmSlope"


Private Sub SSTab1_UnknownEvent_9 '42AD80
  loc_0042AD80: push ebp
  loc_0042AD81: mov ebp, esp
  loc_0042AD83: sub esp, 0000000Ch
  loc_0042AD86: push 00401926h ; __vbaExceptHandler
  loc_0042AD8B: mov eax, fs:[00000000h]
  loc_0042AD91: push eax
  loc_0042AD92: mov fs:[00000000h], esp
  loc_0042AD99: sub esp, 00000008h
  loc_0042AD9C: push ebx
  loc_0042AD9D: push esi
  loc_0042AD9E: push edi
  loc_0042AD9F: mov var_C, esp
  loc_0042ADA2: mov var_8, 00401750h
  loc_0042ADA9: mov eax, Me
  loc_0042ADAC: mov ecx, eax
  loc_0042ADAE: and ecx, 00000001h
  loc_0042ADB1: mov var_4, ecx
  loc_0042ADB4: and al, FEh
  loc_0042ADB6: push eax
  loc_0042ADB7: mov Me, eax
  loc_0042ADBA: mov edx, [eax]
  loc_0042ADBC: call [edx+00000004h]
  loc_0042ADBF: mov var_4, 00000000h
  loc_0042ADC6: mov eax, Me
  loc_0042ADC9: push eax
  loc_0042ADCA: mov ecx, [eax]
  loc_0042ADCC: call [ecx+00000008h]
  loc_0042ADCF: mov eax, var_4
  loc_0042ADD2: mov ecx, var_14
  loc_0042ADD5: pop edi
  loc_0042ADD6: pop esi
  loc_0042ADD7: mov fs:[00000000h], ecx
  loc_0042ADDE: pop ebx
  loc_0042ADDF: mov esp, ebp
  loc_0042ADE1: pop ebp
  loc_0042ADE2: retn 0004h
End Sub

Private Sub SSTab1_UnknownEvent_10 '42AC40
  loc_0042AC40: push ebp
  loc_0042AC41: mov ebp, esp
  loc_0042AC43: sub esp, 0000000Ch
  loc_0042AC46: push 00401926h ; __vbaExceptHandler
  loc_0042AC4B: mov eax, fs:[00000000h]
  loc_0042AC51: push eax
  loc_0042AC52: mov fs:[00000000h], esp
  loc_0042AC59: sub esp, 00000024h
  loc_0042AC5C: push ebx
  loc_0042AC5D: push esi
  loc_0042AC5E: push edi
  loc_0042AC5F: mov var_C, esp
  loc_0042AC62: mov var_8, 00401740h
  loc_0042AC69: mov edi, Me
  loc_0042AC6C: mov eax, edi
  loc_0042AC6E: and eax, 00000001h
  loc_0042AC71: mov var_4, eax
  loc_0042AC74: and edi, FFFFFFFEh
  loc_0042AC77: push edi
  loc_0042AC78: mov Me, edi
  loc_0042AC7B: mov ecx, [edi]
  loc_0042AC7D: call [ecx+00000004h]
  loc_0042AC80: mov edx, [edi]
  loc_0042AC82: xor eax, eax
  loc_0042AC84: push eax
  loc_0042AC85: push 00000004h
  loc_0042AC87: push edi
  loc_0042AC88: mov var_18, eax
  loc_0042AC8B: mov var_28, eax
  loc_0042AC8E: call [edx+000003A0h]
  loc_0042AC94: mov ebx, [0040103Ch] ; __vbaObjSet
  loc_0042AC9A: push eax
  loc_0042AC9B: lea eax, var_18
  loc_0042AC9E: push eax
  loc_0042AC9F: call ebx
  loc_0042ACA1: lea ecx, var_28
  loc_0042ACA4: push eax
  loc_0042ACA5: push ecx
  loc_0042ACA6: call [00401088h] ; __vbaLateIdCallLd
  loc_0042ACAC: add esp, 00000010h
  loc_0042ACAF: push eax
  loc_0042ACB0: call [004010C4h] ; __vbaI2Var
  loc_0042ACB6: mov si, ax
  loc_0042ACB9: lea ecx, var_18
  loc_0042ACBC: dec si
  loc_0042ACBE: neg si
  loc_0042ACC1: sbb esi, esi
  loc_0042ACC3: inc esi
  loc_0042ACC4: neg esi
  loc_0042ACC6: call [00401114h] ; __vbaFreeObj
  loc_0042ACCC: lea ecx, var_28
  loc_0042ACCF: call [00401010h] ; __vbaFreeVar
  loc_0042ACD5: mov edx, [edi]
  loc_0042ACD7: push edi
  loc_0042ACD8: test si, si
  loc_0042ACDB: jz 0042ACFCh
  loc_0042ACDD: call [edx+00000378h]
  loc_0042ACE3: push eax
  loc_0042ACE4: lea eax, var_18
  loc_0042ACE7: push eax
  loc_0042ACE8: call ebx
  loc_0042ACEA: mov esi, eax
  loc_0042ACEC: push FFFFFFFFh
  loc_0042ACEE: push esi
  loc_0042ACEF: mov ecx, [esi]
  loc_0042ACF1: call [ecx+00000074h]
  loc_0042ACF4: test eax, eax
  loc_0042ACF6: fnclex
  loc_0042ACF8: jge 0042AD28h
  loc_0042ACFA: jmp 0042AD19h
  loc_0042ACFC: call [edx+00000378h]
  loc_0042AD02: push eax
  loc_0042AD03: lea eax, var_18
  loc_0042AD06: push eax
  loc_0042AD07: call ebx
  loc_0042AD09: mov esi, eax
  loc_0042AD0B: push 00000000h
  loc_0042AD0D: push esi
  loc_0042AD0E: mov ecx, [esi]
  loc_0042AD10: call [ecx+00000074h]
  loc_0042AD13: test eax, eax
  loc_0042AD15: fnclex
  loc_0042AD17: jge 0042AD28h
  loc_0042AD19: push 00000074h
  loc_0042AD1B: push 0041148Ch
  loc_0042AD20: push esi
  loc_0042AD21: push eax
  loc_0042AD22: call [00401030h] ; __vbaHresultCheckObj
  loc_0042AD28: lea ecx, var_18
  loc_0042AD2B: call [00401114h] ; __vbaFreeObj
  loc_0042AD31: mov edx, [edi]
  loc_0042AD33: push edi
  loc_0042AD34: call [edx+00000700h]
  loc_0042AD3A: mov var_4, 00000000h
  loc_0042AD41: push 0042AD5Ch
  loc_0042AD46: jmp 0042AD5Bh
  loc_0042AD48: lea ecx, var_18
  loc_0042AD4B: call [00401114h] ; __vbaFreeObj
  loc_0042AD51: lea ecx, var_28
  loc_0042AD54: call [00401010h] ; __vbaFreeVar
  loc_0042AD5A: ret
  loc_0042AD5B: ret
  loc_0042AD5C: mov eax, Me
  loc_0042AD5F: push eax
  loc_0042AD60: mov ecx, [eax]
  loc_0042AD62: call [ecx+00000008h]
  loc_0042AD65: mov eax, var_4
  loc_0042AD68: mov ecx, var_14
  loc_0042AD6B: pop edi
  loc_0042AD6C: pop esi
  loc_0042AD6D: mov fs:[00000000h], ecx
  loc_0042AD74: pop ebx
  loc_0042AD75: mov esp, ebp
  loc_0042AD77: pop ebp
  loc_0042AD78: retn 0008h
End Sub

Private Sub Form_Load() '427BD0
  loc_00427BD0: push ebp
  loc_00427BD1: mov ebp, esp
  loc_00427BD3: sub esp, 0000000Ch
  loc_00427BD6: push 00401926h ; __vbaExceptHandler
  loc_00427BDB: mov eax, fs:[00000000h]
  loc_00427BE1: push eax
  loc_00427BE2: mov fs:[00000000h], esp
  loc_00427BE9: sub esp, 00000018h
  loc_00427BEC: push ebx
  loc_00427BED: push esi
  loc_00427BEE: push edi
  loc_00427BEF: mov var_C, esp
  loc_00427BF2: mov var_8, 00401660h
  loc_00427BF9: mov ebx, Me
  loc_00427BFC: mov eax, ebx
  loc_00427BFE: and eax, 00000001h
  loc_00427C01: mov var_4, eax
  loc_00427C04: and ebx, FFFFFFFEh
  loc_00427C07: push ebx
  loc_00427C08: mov Me, ebx
  loc_00427C0B: mov ecx, [ebx]
  loc_00427C0D: call [ecx+00000004h]
  loc_00427C10: mov edx, [0043005Ch]
  loc_00427C16: xor eax, eax
  loc_00427C18: push 0041183Ch ; "= "
  loc_00427C1D: push edx
  loc_00427C1E: mov var_18, eax
  loc_00427C21: mov var_1C, eax
  loc_00427C24: mov var_20, eax
  loc_00427C27: mov var_24, eax
  loc_00427C2A: call [0040107Ch] ; __vbaStrR4
  loc_00427C30: mov esi, [004010FCh] ; __vbaStrMove
  loc_00427C36: mov edx, eax
  loc_00427C38: lea ecx, var_18
  loc_00427C3B: call __vbaStrMove
  loc_00427C3D: mov edi, [0040102Ch] ; __vbaStrCat
  loc_00427C43: push eax
  loc_00427C44: call edi
  loc_00427C46: mov edx, eax
  loc_00427C48: mov ecx, 00430054h
  loc_00427C4D: call __vbaStrMove
  loc_00427C4F: lea ecx, var_18
  loc_00427C52: call [00401110h] ; __vbaFreeStr
  loc_00427C58: mov eax, [0043005Ch]
  loc_00427C5D: push 00411848h ; "¹ "
  loc_00427C62: push eax
  loc_00427C63: call [0040107Ch] ; __vbaStrR4
  loc_00427C69: mov edx, eax
  loc_00427C6B: lea ecx, var_18
  loc_00427C6E: call __vbaStrMove
  loc_00427C70: push eax
  loc_00427C71: call edi
  loc_00427C73: mov edx, eax
  loc_00427C75: mov ecx, 00430058h
  loc_00427C7A: call __vbaStrMove
  loc_00427C7C: lea ecx, var_18
  loc_00427C7F: call [00401110h] ; __vbaFreeStr
  loc_00427C85: mov ecx, [00430028h]
  loc_00427C8B: push 00411854h ; " t* > "
  loc_00427C90: push ecx
  loc_00427C91: call [0040107Ch] ; __vbaStrR4
  loc_00427C97: mov edx, eax
  loc_00427C99: lea ecx, var_18
  loc_00427C9C: call __vbaStrMove
  loc_00427C9E: push eax
  loc_00427C9F: call edi
  loc_00427CA1: mov edx, eax
  loc_00427CA3: lea ecx, var_1C
  loc_00427CA6: call __vbaStrMove
  loc_00427CA8: push eax
  loc_00427CA9: push 00411470h ; " or if t* < -"
  loc_00427CAE: call edi
  loc_00427CB0: mov edx, eax
  loc_00427CB2: lea ecx, var_20
  loc_00427CB5: call __vbaStrMove
  loc_00427CB7: mov edx, [00430028h]
  loc_00427CBD: push eax
  loc_00427CBE: push edx
  loc_00427CBF: call [0040107Ch] ; __vbaStrR4
  loc_00427CC5: mov edx, eax
  loc_00427CC7: lea ecx, var_24
  loc_00427CCA: call __vbaStrMove
  loc_00427CCC: push eax
  loc_00427CCD: call edi
  loc_00427CCF: mov edx, eax
  loc_00427CD1: mov ecx, 0043004Ch
  loc_00427CD6: call __vbaStrMove
  loc_00427CD8: lea eax, var_24
  loc_00427CDB: push eax
  loc_00427CDC: lea ecx, var_20
  loc_00427CDF: lea edx, var_1C
  loc_00427CE2: push ecx
  loc_00427CE3: lea eax, var_18
  loc_00427CE6: push edx
  loc_00427CE7: push eax
  loc_00427CE8: push 00000004h
  loc_00427CEA: call [004010E4h] ; __vbaFreeStrList
  loc_00427CF0: add esp, 00000014h
  loc_00427CF3: mov [00430068h], 0001h
  loc_00427CFC: mov ecx, [ebx]
  loc_00427CFE: push ebx
  loc_00427CFF: call [ecx+00000700h]
  loc_00427D05: mov var_4, 00000000h
  loc_00427D0C: fwait
  loc_00427D0D: push 00427D31h
  loc_00427D12: jmp 00427D30h
  loc_00427D14: lea edx, var_24
  loc_00427D17: lea eax, var_20
  loc_00427D1A: push edx
  loc_00427D1B: lea ecx, var_1C
  loc_00427D1E: push eax
  loc_00427D1F: lea edx, var_18
  loc_00427D22: push ecx
  loc_00427D23: push edx
  loc_00427D24: push 00000004h
  loc_00427D26: call [004010E4h] ; __vbaFreeStrList
  loc_00427D2C: add esp, 00000014h
  loc_00427D2F: ret
  loc_00427D30: ret
  loc_00427D31: mov eax, Me
  loc_00427D34: push eax
  loc_00427D35: mov ecx, [eax]
  loc_00427D37: call [ecx+00000008h]
  loc_00427D3A: mov eax, var_4
  loc_00427D3D: mov ecx, var_14
  loc_00427D40: pop edi
  loc_00427D41: pop esi
  loc_00427D42: mov fs:[00000000h], ecx
  loc_00427D49: pop ebx
  loc_00427D4A: mov esp, ebp
  loc_00427D4C: pop ebp
  loc_00427D4D: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '42A220
  loc_0042A220: push ebp
  loc_0042A221: mov ebp, esp
  loc_0042A223: sub esp, 0000000Ch
  loc_0042A226: push 00401926h ; __vbaExceptHandler
  loc_0042A22B: mov eax, fs:[00000000h]
  loc_0042A231: push eax
  loc_0042A232: mov fs:[00000000h], esp
  loc_0042A239: sub esp, 0000009Ch
  loc_0042A23F: push ebx
  loc_0042A240: push esi
  loc_0042A241: push edi
  loc_0042A242: mov var_C, esp
  loc_0042A245: mov var_8, 004016E0h
  loc_0042A24C: mov edi, Me
  loc_0042A24F: mov eax, edi
  loc_0042A251: and eax, 00000001h
  loc_0042A254: mov var_4, eax
  loc_0042A257: and edi, FFFFFFFEh
  loc_0042A25A: push edi
  loc_0042A25B: mov Me, edi
  loc_0042A25E: mov ecx, [edi]
  loc_0042A260: call [ecx+00000004h]
  loc_0042A263: mov ebx, [004010F4h] ; __vbaVarDup
  loc_0042A269: mov ecx, 80020004h
  loc_0042A26E: xor esi, esi
  loc_0042A270: mov var_54, ecx
  loc_0042A273: mov eax, 0000000Ah
  loc_0042A278: mov var_44, ecx
  loc_0042A27B: mov var_4C, esi
  loc_0042A27E: mov var_5C, esi
  loc_0042A281: mov var_7C, esi
  loc_0042A284: lea edx, var_7C
  loc_0042A287: lea ecx, var_3C
  loc_0042A28A: mov var_1C, esi
  loc_0042A28D: mov var_2C, esi
  loc_0042A290: mov var_3C, esi
  loc_0042A293: mov var_6C, esi
  loc_0042A296: mov var_5C, eax
  loc_0042A299: mov var_4C, eax
  loc_0042A29C: mov var_74, 0040ECD4h ; "Exit Check"
  loc_0042A2A3: mov var_7C, 00000008h
  loc_0042A2AA: call ebx
  loc_0042A2AC: lea edx, var_6C
  loc_0042A2AF: lea ecx, var_2C
  loc_0042A2B2: mov var_64, 0040EC90h ; "Are you sure you want to exit?"
  loc_0042A2B9: mov var_6C, 00000008h
  loc_0042A2C0: call ebx
  loc_0042A2C2: lea edx, var_5C
  loc_0042A2C5: lea eax, var_4C
  loc_0042A2C8: push edx
  loc_0042A2C9: lea ecx, var_3C
  loc_0042A2CC: push eax
  loc_0042A2CD: push ecx
  loc_0042A2CE: lea edx, var_2C
  loc_0042A2D1: push 00000004h
  loc_0042A2D3: push edx
  loc_0042A2D4: call [00401038h] ; rtcMsgBox
  loc_0042A2DA: mov ecx, eax
  loc_0042A2DC: call [00401078h] ; __vbaI2I4
  loc_0042A2E2: mov ebx, eax
  loc_0042A2E4: lea eax, var_5C
  loc_0042A2E7: lea ecx, var_4C
  loc_0042A2EA: push eax
  loc_0042A2EB: lea edx, var_3C
  loc_0042A2EE: push ecx
  loc_0042A2EF: lea eax, var_2C
  loc_0042A2F2: push edx
  loc_0042A2F3: push eax
  loc_0042A2F4: push 00000004h
  loc_0042A2F6: call [00401018h] ; __vbaFreeVarList
  loc_0042A2FC: add esp, 00000014h
  loc_0042A2FF: cmp bx, 0007h
  loc_0042A303: jnz 0042A30Fh
  loc_0042A305: mov ecx, Cancel
  loc_0042A308: mov [ecx], FFFFFFh
  loc_0042A30D: jmp 0042A36Fh
  loc_0042A30F: cmp [00430A24h], esi
  loc_0042A315: jnz 0042A327h
  loc_0042A317: push 00430A24h
  loc_0042A31C: push 0040EC7Ch
  loc_0042A321: call [004010D4h] ; __vbaNew2
  loc_0042A327: mov ebx, [00430A24h]
  loc_0042A32D: lea eax, var_1C
  loc_0042A330: push edi
  loc_0042A331: push eax
  loc_0042A332: mov edx, [ebx]
  loc_0042A334: mov var_B0, edx
  loc_0042A33A: call [00401044h] ; __vbaObjSetAddref
  loc_0042A340: mov ecx, var_B0
  loc_0042A346: push eax
  loc_0042A347: push ebx
  loc_0042A348: call [ecx+00000010h]
  loc_0042A34B: cmp eax, esi
  loc_0042A34D: fnclex
  loc_0042A34F: jge 0042A360h
  loc_0042A351: push 00000010h
  loc_0042A353: push 0040EC6Ch
  loc_0042A358: push ebx
  loc_0042A359: push eax
  loc_0042A35A: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A360: lea ecx, var_1C
  loc_0042A363: call [00401114h] ; __vbaFreeObj
  loc_0042A369: call [0040101Ch] ; __vbaEnd
  loc_0042A36F: mov var_4, esi
  loc_0042A372: push 0042A39Fh
  loc_0042A377: jmp 0042A39Eh
  loc_0042A379: lea ecx, var_1C
  loc_0042A37C: call [00401114h] ; __vbaFreeObj
  loc_0042A382: lea edx, var_5C
  loc_0042A385: lea eax, var_4C
  loc_0042A388: push edx
  loc_0042A389: lea ecx, var_3C
  loc_0042A38C: push eax
  loc_0042A38D: lea edx, var_2C
  loc_0042A390: push ecx
  loc_0042A391: push edx
  loc_0042A392: push 00000004h
  loc_0042A394: call [00401018h] ; __vbaFreeVarList
  loc_0042A39A: add esp, 00000014h
  loc_0042A39D: ret
  loc_0042A39E: ret
  loc_0042A39F: mov eax, Me
  loc_0042A3A2: push eax
  loc_0042A3A3: mov ecx, [eax]
  loc_0042A3A5: call [ecx+00000008h]
  loc_0042A3A8: mov eax, var_4
  loc_0042A3AB: mov ecx, var_14
  loc_0042A3AE: pop edi
  loc_0042A3AF: pop esi
  loc_0042A3B0: mov fs:[00000000h], ecx
  loc_0042A3B7: pop ebx
  loc_0042A3B8: mov esp, ebp
  loc_0042A3BA: pop ebp
  loc_0042A3BB: retn 000Ch
End Sub

Private Sub Form_Activate() '427B60
  loc_00427B60: push ebp
  loc_00427B61: mov ebp, esp
  loc_00427B63: sub esp, 0000000Ch
  loc_00427B66: push 00401926h ; __vbaExceptHandler
  loc_00427B6B: mov eax, fs:[00000000h]
  loc_00427B71: push eax
  loc_00427B72: mov fs:[00000000h], esp
  loc_00427B79: sub esp, 00000008h
  loc_00427B7C: push ebx
  loc_00427B7D: push esi
  loc_00427B7E: push edi
  loc_00427B7F: mov var_C, esp
  loc_00427B82: mov var_8, 00401658h
  loc_00427B89: mov esi, Me
  loc_00427B8C: mov eax, esi
  loc_00427B8E: and eax, 00000001h
  loc_00427B91: mov var_4, eax
  loc_00427B94: and esi, FFFFFFFEh
  loc_00427B97: push esi
  loc_00427B98: mov Me, esi
  loc_00427B9B: mov ecx, [esi]
  loc_00427B9D: call [ecx+00000004h]
  loc_00427BA0: mov edx, [esi]
  loc_00427BA2: push esi
  loc_00427BA3: call [edx+00000700h]
  loc_00427BA9: mov var_4, 00000000h
  loc_00427BB0: mov eax, Me
  loc_00427BB3: push eax
  loc_00427BB4: mov ecx, [eax]
  loc_00427BB6: call [ecx+00000008h]
  loc_00427BB9: mov eax, var_4
  loc_00427BBC: mov ecx, var_14
  loc_00427BBF: pop edi
  loc_00427BC0: pop esi
  loc_00427BC1: mov fs:[00000000h], ecx
  loc_00427BC8: pop ebx
  loc_00427BC9: mov esp, ebp
  loc_00427BCB: pop ebp
  loc_00427BCC: retn 0004h
End Sub

Private Sub mnuChangeNames_Click() '429E80
  loc_00429E80: push ebp
  loc_00429E81: mov ebp, esp
  loc_00429E83: sub esp, 0000000Ch
  loc_00429E86: push 00401926h ; __vbaExceptHandler
  loc_00429E8B: mov eax, fs:[00000000h]
  loc_00429E91: push eax
  loc_00429E92: mov fs:[00000000h], esp
  loc_00429E99: sub esp, 00000030h
  loc_00429E9C: push ebx
  loc_00429E9D: push esi
  loc_00429E9E: push edi
  loc_00429E9F: mov var_C, esp
  loc_00429EA2: mov var_8, 004016B8h
  loc_00429EA9: mov eax, Me
  loc_00429EAC: mov ecx, eax
  loc_00429EAE: and ecx, 00000001h
  loc_00429EB1: mov var_4, ecx
  loc_00429EB4: and al, FEh
  loc_00429EB6: push eax
  loc_00429EB7: mov Me, eax
  loc_00429EBA: mov edx, [eax]
  loc_00429EBC: call [edx+00000004h]
  loc_00429EBF: mov eax, [004300D8h]
  loc_00429EC4: test eax, eax
  loc_00429EC6: jnz 00429ED8h
  loc_00429EC8: push 004300D8h
  loc_00429ECD: push 00402E04h
  loc_00429ED2: call [004010D4h] ; __vbaNew2
  loc_00429ED8: sub esp, 00000010h
  loc_00429EDB: mov ecx, 0000000Ah
  loc_00429EE0: mov ebx, esp
  loc_00429EE2: mov var_24, ecx
  loc_00429EE5: mov eax, 80020004h
  loc_00429EEA: sub esp, 00000010h
  loc_00429EED: mov [ebx], ecx
  loc_00429EEF: mov ecx, var_30
  loc_00429EF2: mov edx, eax
  loc_00429EF4: mov esi, [004300D8h]
  loc_00429EFA: mov [ebx+00000004h], ecx
  loc_00429EFD: mov ecx, esp
  loc_00429EFF: mov edi, [esi]
  loc_00429F01: push esi
  loc_00429F02: mov [ebx+00000008h], eax
  loc_00429F05: mov eax, var_28
  loc_00429F08: mov [ebx+0000000Ch], eax
  loc_00429F0B: mov eax, var_24
  loc_00429F0E: mov [ecx], eax
  loc_00429F10: mov eax, var_20
  loc_00429F13: mov [ecx+00000004h], eax
  loc_00429F16: mov [ecx+00000008h], edx
  loc_00429F19: mov edx, var_18
  loc_00429F1C: mov [ecx+0000000Ch], edx
  loc_00429F1F: call [edi+000002B0h]
  loc_00429F25: test eax, eax
  loc_00429F27: fnclex
  loc_00429F29: jge 00429F3Dh
  loc_00429F2B: push 000002B0h
  loc_00429F30: push 0040E260h
  loc_00429F35: push esi
  loc_00429F36: push eax
  loc_00429F37: call [00401030h] ; __vbaHresultCheckObj
  loc_00429F3D: mov var_4, 00000000h
  loc_00429F44: mov eax, Me
  loc_00429F47: push eax
  loc_00429F48: mov ecx, [eax]
  loc_00429F4A: call [ecx+00000008h]
  loc_00429F4D: mov eax, var_4
  loc_00429F50: mov ecx, var_14
  loc_00429F53: pop edi
  loc_00429F54: pop esi
  loc_00429F55: mov fs:[00000000h], ecx
  loc_00429F5C: pop ebx
  loc_00429F5D: mov esp, ebp
  loc_00429F5F: pop ebp
  loc_00429F60: retn 0004h
End Sub

Private Sub mnuAssumptions_Click() '429B70
  loc_00429B70: push ebp
  loc_00429B71: mov ebp, esp
  loc_00429B73: sub esp, 0000000Ch
  loc_00429B76: push 00401926h ; __vbaExceptHandler
  loc_00429B7B: mov eax, fs:[00000000h]
  loc_00429B81: push eax
  loc_00429B82: mov fs:[00000000h], esp
  loc_00429B89: sub esp, 00000030h
  loc_00429B8C: push ebx
  loc_00429B8D: push esi
  loc_00429B8E: push edi
  loc_00429B8F: mov var_C, esp
  loc_00429B92: mov var_8, 004016A0h
  loc_00429B99: mov eax, Me
  loc_00429B9C: mov ecx, eax
  loc_00429B9E: and ecx, 00000001h
  loc_00429BA1: mov var_4, ecx
  loc_00429BA4: and al, FEh
  loc_00429BA6: push eax
  loc_00429BA7: mov Me, eax
  loc_00429BAA: mov edx, [eax]
  loc_00429BAC: call [edx+00000004h]
  loc_00429BAF: mov eax, [0043009Ch]
  loc_00429BB4: test eax, eax
  loc_00429BB6: jnz 00429BC8h
  loc_00429BB8: push 0043009Ch
  loc_00429BBD: push 00405FC0h
  loc_00429BC2: call [004010D4h] ; __vbaNew2
  loc_00429BC8: sub esp, 00000010h
  loc_00429BCB: mov ecx, 0000000Ah
  loc_00429BD0: mov ebx, esp
  loc_00429BD2: mov var_24, ecx
  loc_00429BD5: mov eax, 80020004h
  loc_00429BDA: sub esp, 00000010h
  loc_00429BDD: mov [ebx], ecx
  loc_00429BDF: mov ecx, var_30
  loc_00429BE2: mov edx, eax
  loc_00429BE4: mov esi, [0043009Ch]
  loc_00429BEA: mov [ebx+00000004h], ecx
  loc_00429BED: mov ecx, esp
  loc_00429BEF: mov edi, [esi]
  loc_00429BF1: push esi
  loc_00429BF2: mov [ebx+00000008h], eax
  loc_00429BF5: mov eax, var_28
  loc_00429BF8: mov [ebx+0000000Ch], eax
  loc_00429BFB: mov eax, var_24
  loc_00429BFE: mov [ecx], eax
  loc_00429C00: mov eax, var_20
  loc_00429C03: mov [ecx+00000004h], eax
  loc_00429C06: mov [ecx+00000008h], edx
  loc_00429C09: mov edx, var_18
  loc_00429C0C: mov [ecx+0000000Ch], edx
  loc_00429C0F: call [edi+000002B0h]
  loc_00429C15: test eax, eax
  loc_00429C17: fnclex
  loc_00429C19: jge 00429C2Dh
  loc_00429C1B: push 000002B0h
  loc_00429C20: push 0040DDE0h
  loc_00429C25: push esi
  loc_00429C26: push eax
  loc_00429C27: call [00401030h] ; __vbaHresultCheckObj
  loc_00429C2D: mov eax, [004300C4h]
  loc_00429C32: test eax, eax
  loc_00429C34: jnz 00429C46h
  loc_00429C36: push 004300C4h
  loc_00429C3B: push 00409228h
  loc_00429C40: call [004010D4h] ; __vbaNew2
  loc_00429C46: mov esi, [004300C4h]
  loc_00429C4C: push esi
  loc_00429C4D: mov eax, [esi]
  loc_00429C4F: call [eax+000002B4h]
  loc_00429C55: test eax, eax
  loc_00429C57: fnclex
  loc_00429C59: jge 00429C6Dh
  loc_00429C5B: push 000002B4h
  loc_00429C60: push 0040E0ECh
  loc_00429C65: push esi
  loc_00429C66: push eax
  loc_00429C67: call [00401030h] ; __vbaHresultCheckObj
  loc_00429C6D: mov var_4, 00000000h
  loc_00429C74: mov eax, Me
  loc_00429C77: push eax
  loc_00429C78: mov ecx, [eax]
  loc_00429C7A: call [ecx+00000008h]
  loc_00429C7D: mov eax, var_4
  loc_00429C80: mov ecx, var_14
  loc_00429C83: pop edi
  loc_00429C84: pop esi
  loc_00429C85: mov fs:[00000000h], ecx
  loc_00429C8C: pop ebx
  loc_00429C8D: mov esp, ebp
  loc_00429C8F: pop ebp
  loc_00429C90: retn 0004h
End Sub

Private Sub mnuStatistics_Click() '42A750
  loc_0042A750: push ebp
  loc_0042A751: mov ebp, esp
  loc_0042A753: sub esp, 0000000Ch
  loc_0042A756: push 00401926h ; __vbaExceptHandler
  loc_0042A75B: mov eax, fs:[00000000h]
  loc_0042A761: push eax
  loc_0042A762: mov fs:[00000000h], esp
  loc_0042A769: sub esp, 00000030h
  loc_0042A76C: push ebx
  loc_0042A76D: push esi
  loc_0042A76E: push edi
  loc_0042A76F: mov var_C, esp
  loc_0042A772: mov var_8, 00401708h
  loc_0042A779: mov eax, Me
  loc_0042A77C: mov ecx, eax
  loc_0042A77E: and ecx, 00000001h
  loc_0042A781: mov var_4, ecx
  loc_0042A784: and al, FEh
  loc_0042A786: push eax
  loc_0042A787: mov Me, eax
  loc_0042A78A: mov edx, [eax]
  loc_0042A78C: call [edx+00000004h]
  loc_0042A78F: mov eax, [004300ECh]
  loc_0042A794: test eax, eax
  loc_0042A796: jnz 0042A7A8h
  loc_0042A798: push 004300ECh
  loc_0042A79D: push 0040A624h
  loc_0042A7A2: call [004010D4h] ; __vbaNew2
  loc_0042A7A8: sub esp, 00000010h
  loc_0042A7AB: mov ecx, 0000000Ah
  loc_0042A7B0: mov ebx, esp
  loc_0042A7B2: mov var_24, ecx
  loc_0042A7B5: mov eax, 80020004h
  loc_0042A7BA: sub esp, 00000010h
  loc_0042A7BD: mov [ebx], ecx
  loc_0042A7BF: mov ecx, var_30
  loc_0042A7C2: mov edx, eax
  loc_0042A7C4: mov esi, [004300ECh]
  loc_0042A7CA: mov [ebx+00000004h], ecx
  loc_0042A7CD: mov ecx, esp
  loc_0042A7CF: mov edi, [esi]
  loc_0042A7D1: push esi
  loc_0042A7D2: mov [ebx+00000008h], eax
  loc_0042A7D5: mov eax, var_28
  loc_0042A7D8: mov [ebx+0000000Ch], eax
  loc_0042A7DB: mov eax, var_24
  loc_0042A7DE: mov [ecx], eax
  loc_0042A7E0: mov eax, var_20
  loc_0042A7E3: mov [ecx+00000004h], eax
  loc_0042A7E6: mov [ecx+00000008h], edx
  loc_0042A7E9: mov edx, var_18
  loc_0042A7EC: mov [ecx+0000000Ch], edx
  loc_0042A7EF: call [edi+000002B0h]
  loc_0042A7F5: test eax, eax
  loc_0042A7F7: fnclex
  loc_0042A7F9: jge 0042A80Dh
  loc_0042A7FB: push 000002B0h
  loc_0042A800: push 0040ECECh
  loc_0042A805: push esi
  loc_0042A806: push eax
  loc_0042A807: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A80D: mov eax, [004300C4h]
  loc_0042A812: test eax, eax
  loc_0042A814: jnz 0042A826h
  loc_0042A816: push 004300C4h
  loc_0042A81B: push 00409228h
  loc_0042A820: call [004010D4h] ; __vbaNew2
  loc_0042A826: mov esi, [004300C4h]
  loc_0042A82C: push esi
  loc_0042A82D: mov eax, [esi]
  loc_0042A82F: call [eax+000002B4h]
  loc_0042A835: test eax, eax
  loc_0042A837: fnclex
  loc_0042A839: jge 0042A84Dh
  loc_0042A83B: push 000002B4h
  loc_0042A840: push 0040E0ECh
  loc_0042A845: push esi
  loc_0042A846: push eax
  loc_0042A847: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A84D: mov var_4, 00000000h
  loc_0042A854: mov eax, Me
  loc_0042A857: push eax
  loc_0042A858: mov ecx, [eax]
  loc_0042A85A: call [ecx+00000008h]
  loc_0042A85D: mov eax, var_4
  loc_0042A860: mov ecx, var_14
  loc_0042A863: pop edi
  loc_0042A864: pop esi
  loc_0042A865: mov fs:[00000000h], ecx
  loc_0042A86C: pop ebx
  loc_0042A86D: mov esp, ebp
  loc_0042A86F: pop ebp
  loc_0042A870: retn 0004h
End Sub

Private Sub mnuExit_Click() '42A150
  loc_0042A150: push ebp
  loc_0042A151: mov ebp, esp
  loc_0042A153: sub esp, 0000000Ch
  loc_0042A156: push 00401926h ; __vbaExceptHandler
  loc_0042A15B: mov eax, fs:[00000000h]
  loc_0042A161: push eax
  loc_0042A162: mov fs:[00000000h], esp
  loc_0042A169: sub esp, 00000018h
  loc_0042A16C: push ebx
  loc_0042A16D: push esi
  loc_0042A16E: push edi
  loc_0042A16F: mov var_C, esp
  loc_0042A172: mov var_8, 004016D0h
  loc_0042A179: mov edi, Me
  loc_0042A17C: mov eax, edi
  loc_0042A17E: and eax, 00000001h
  loc_0042A181: mov var_4, eax
  loc_0042A184: and edi, FFFFFFFEh
  loc_0042A187: push edi
  loc_0042A188: mov Me, edi
  loc_0042A18B: mov ecx, [edi]
  loc_0042A18D: call [ecx+00000004h]
  loc_0042A190: mov eax, [00430A24h]
  loc_0042A195: xor ebx, ebx
  loc_0042A197: cmp eax, ebx
  loc_0042A199: mov var_18, ebx
  loc_0042A19C: jnz 0042A1AEh
  loc_0042A19E: push 00430A24h
  loc_0042A1A3: push 0040EC7Ch
  loc_0042A1A8: call [004010D4h] ; __vbaNew2
  loc_0042A1AE: mov esi, [00430A24h]
  loc_0042A1B4: lea eax, var_18
  loc_0042A1B7: push edi
  loc_0042A1B8: push eax
  loc_0042A1B9: mov edx, [esi]
  loc_0042A1BB: mov var_2C, edx
  loc_0042A1BE: call [00401044h] ; __vbaObjSetAddref
  loc_0042A1C4: mov ecx, var_2C
  loc_0042A1C7: push eax
  loc_0042A1C8: push esi
  loc_0042A1C9: call [ecx+00000010h]
  loc_0042A1CC: cmp eax, ebx
  loc_0042A1CE: fnclex
  loc_0042A1D0: jge 0042A1E1h
  loc_0042A1D2: push 00000010h
  loc_0042A1D4: push 0040EC6Ch
  loc_0042A1D9: push esi
  loc_0042A1DA: push eax
  loc_0042A1DB: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A1E1: lea ecx, var_18
  loc_0042A1E4: call [00401114h] ; __vbaFreeObj
  loc_0042A1EA: mov var_4, ebx
  loc_0042A1ED: push 0042A1FFh
  loc_0042A1F2: jmp 0042A1FEh
  loc_0042A1F4: lea ecx, var_18
  loc_0042A1F7: call [00401114h] ; __vbaFreeObj
  loc_0042A1FD: ret
  loc_0042A1FE: ret
  loc_0042A1FF: mov eax, Me
  loc_0042A202: push eax
  loc_0042A203: mov edx, [eax]
  loc_0042A205: call [edx+00000008h]
  loc_0042A208: mov eax, var_4
  loc_0042A20B: mov ecx, var_14
  loc_0042A20E: pop edi
  loc_0042A20F: pop esi
  loc_0042A210: mov fs:[00000000h], ecx
  loc_0042A217: pop ebx
  loc_0042A218: mov esp, ebp
  loc_0042A21A: pop ebp
  loc_0042A21B: retn 0004h
End Sub

Private Sub mnuChangeHypothesizedValue_Click() '429D90
  loc_00429D90: push ebp
  loc_00429D91: mov ebp, esp
  loc_00429D93: sub esp, 0000000Ch
  loc_00429D96: push 00401926h ; __vbaExceptHandler
  loc_00429D9B: mov eax, fs:[00000000h]
  loc_00429DA1: push eax
  loc_00429DA2: mov fs:[00000000h], esp
  loc_00429DA9: sub esp, 00000030h
  loc_00429DAC: push ebx
  loc_00429DAD: push esi
  loc_00429DAE: push edi
  loc_00429DAF: mov var_C, esp
  loc_00429DB2: mov var_8, 004016B0h
  loc_00429DB9: mov eax, Me
  loc_00429DBC: mov ecx, eax
  loc_00429DBE: and ecx, 00000001h
  loc_00429DC1: mov var_4, ecx
  loc_00429DC4: and al, FEh
  loc_00429DC6: push eax
  loc_00429DC7: mov Me, eax
  loc_00429DCA: mov edx, [eax]
  loc_00429DCC: call [edx+00000004h]
  loc_00429DCF: mov eax, [0043013Ch]
  loc_00429DD4: test eax, eax
  loc_00429DD6: jnz 00429DE8h
  loc_00429DD8: push 0043013Ch
  loc_00429DDD: push 0040204Ch
  loc_00429DE2: call [004010D4h] ; __vbaNew2
  loc_00429DE8: sub esp, 00000010h
  loc_00429DEB: mov ecx, 0000000Ah
  loc_00429DF0: mov ebx, esp
  loc_00429DF2: mov var_24, ecx
  loc_00429DF5: mov eax, 80020004h
  loc_00429DFA: sub esp, 00000010h
  loc_00429DFD: mov [ebx], ecx
  loc_00429DFF: mov ecx, var_30
  loc_00429E02: mov edx, eax
  loc_00429E04: mov esi, [0043013Ch]
  loc_00429E0A: mov [ebx+00000004h], ecx
  loc_00429E0D: mov ecx, esp
  loc_00429E0F: mov edi, [esi]
  loc_00429E11: push esi
  loc_00429E12: mov [ebx+00000008h], eax
  loc_00429E15: mov eax, var_28
  loc_00429E18: mov [ebx+0000000Ch], eax
  loc_00429E1B: mov eax, var_24
  loc_00429E1E: mov [ecx], eax
  loc_00429E20: mov eax, var_20
  loc_00429E23: mov [ecx+00000004h], eax
  loc_00429E26: mov [ecx+00000008h], edx
  loc_00429E29: mov edx, var_18
  loc_00429E2C: mov [ecx+0000000Ch], edx
  loc_00429E2F: call [edi+000002B0h]
  loc_00429E35: test eax, eax
  loc_00429E37: fnclex
  loc_00429E39: jge 00429E4Dh
  loc_00429E3B: push 000002B0h
  loc_00429E40: push 0040FC38h
  loc_00429E45: push esi
  loc_00429E46: push eax
  loc_00429E47: call [00401030h] ; __vbaHresultCheckObj
  loc_00429E4D: mov var_4, 00000000h
  loc_00429E54: mov eax, Me
  loc_00429E57: push eax
  loc_00429E58: mov ecx, [eax]
  loc_00429E5A: call [ecx+00000008h]
  loc_00429E5D: mov eax, var_4
  loc_00429E60: mov ecx, var_14
  loc_00429E63: pop edi
  loc_00429E64: pop esi
  loc_00429E65: mov fs:[00000000h], ecx
  loc_00429E6C: pop ebx
  loc_00429E6D: mov esp, ebp
  loc_00429E6F: pop ebp
  loc_00429E70: retn 0004h
End Sub

Private Sub optLeft_Click() '42A880
  loc_0042A880: push ebp
  loc_0042A881: mov ebp, esp
  loc_0042A883: sub esp, 0000000Ch
  loc_0042A886: push 00401926h ; __vbaExceptHandler
  loc_0042A88B: mov eax, fs:[00000000h]
  loc_0042A891: push eax
  loc_0042A892: mov fs:[00000000h], esp
  loc_0042A899: sub esp, 0000000Ch
  loc_0042A89C: push ebx
  loc_0042A89D: push esi
  loc_0042A89E: push edi
  loc_0042A89F: mov var_C, esp
  loc_0042A8A2: mov var_8, 00401710h
  loc_0042A8A9: mov edi, Me
  loc_0042A8AC: mov eax, edi
  loc_0042A8AE: and eax, 00000001h
  loc_0042A8B1: mov var_4, eax
  loc_0042A8B4: and edi, FFFFFFFEh
  loc_0042A8B7: push edi
  loc_0042A8B8: mov Me, edi
  loc_0042A8BB: mov ecx, [edi]
  loc_0042A8BD: call [ecx+00000004h]
  loc_0042A8C0: mov edx, [0043005Ch]
  loc_0042A8C6: mov ebx, [0040107Ch] ; __vbaStrR4
  loc_0042A8CC: push 00411A18h ; "³ "
  loc_0042A8D1: push edx
  loc_0042A8D2: mov var_18, 00000000h
  loc_0042A8D9: call ebx
  loc_0042A8DB: mov esi, [004010FCh] ; __vbaStrMove
  loc_0042A8E1: mov edx, eax
  loc_0042A8E3: lea ecx, var_18
  loc_0042A8E6: call __vbaStrMove
  loc_0042A8E8: push eax
  loc_0042A8E9: call [0040102Ch] ; __vbaStrCat
  loc_0042A8EF: mov edx, eax
  loc_0042A8F1: mov ecx, 00430054h
  loc_0042A8F6: call __vbaStrMove
  loc_0042A8F8: lea ecx, var_18
  loc_0042A8FB: call [00401110h] ; __vbaFreeStr
  loc_0042A901: mov eax, [0043005Ch]
  loc_0042A906: push 00411A24h ; "< "
  loc_0042A90B: push eax
  loc_0042A90C: call ebx
  loc_0042A90E: mov edx, eax
  loc_0042A910: lea ecx, var_18
  loc_0042A913: call __vbaStrMove
  loc_0042A915: push eax
  loc_0042A916: call [0040102Ch] ; __vbaStrCat
  loc_0042A91C: mov edx, eax
  loc_0042A91E: mov ecx, 00430058h
  loc_0042A923: call __vbaStrMove
  loc_0042A925: lea ecx, var_18
  loc_0042A928: call [00401110h] ; __vbaFreeStr
  loc_0042A92E: mov ecx, [00430024h]
  loc_0042A934: push 00411A30h ; "t* < -"
  loc_0042A939: push ecx
  loc_0042A93A: call ebx
  loc_0042A93C: mov edx, eax
  loc_0042A93E: lea ecx, var_18
  loc_0042A941: call __vbaStrMove
  loc_0042A943: push eax
  loc_0042A944: call [0040102Ch] ; __vbaStrCat
  loc_0042A94A: mov edx, eax
  loc_0042A94C: mov ecx, 0043004Ch
  loc_0042A951: call __vbaStrMove
  loc_0042A953: lea ecx, var_18
  loc_0042A956: call [00401110h] ; __vbaFreeStr
  loc_0042A95C: mov edx, [edi]
  loc_0042A95E: push edi
  loc_0042A95F: call [edx+00000700h]
  loc_0042A965: mov var_4, 00000000h
  loc_0042A96C: fwait
  loc_0042A96D: push 0042A97Fh
  loc_0042A972: jmp 0042A97Eh
  loc_0042A974: lea ecx, var_18
  loc_0042A977: call [00401110h] ; __vbaFreeStr
  loc_0042A97D: ret
  loc_0042A97E: ret
  loc_0042A97F: mov eax, Me
  loc_0042A982: push eax
  loc_0042A983: mov ecx, [eax]
  loc_0042A985: call [ecx+00000008h]
  loc_0042A988: mov eax, var_4
  loc_0042A98B: mov ecx, var_14
  loc_0042A98E: pop edi
  loc_0042A98F: pop esi
  loc_0042A990: mov fs:[00000000h], ecx
  loc_0042A997: pop ebx
  loc_0042A998: mov esp, ebp
  loc_0042A99A: pop ebp
  loc_0042A99B: retn 0004h
End Sub

Private Sub mnuChangePointEstimate_Click() '429F70
  loc_00429F70: push ebp
  loc_00429F71: mov ebp, esp
  loc_00429F73: sub esp, 0000000Ch
  loc_00429F76: push 00401926h ; __vbaExceptHandler
  loc_00429F7B: mov eax, fs:[00000000h]
  loc_00429F81: push eax
  loc_00429F82: mov fs:[00000000h], esp
  loc_00429F89: sub esp, 00000030h
  loc_00429F8C: push ebx
  loc_00429F8D: push esi
  loc_00429F8E: push edi
  loc_00429F8F: mov var_C, esp
  loc_00429F92: mov var_8, 004016C0h
  loc_00429F99: mov eax, Me
  loc_00429F9C: mov ecx, eax
  loc_00429F9E: and ecx, 00000001h
  loc_00429FA1: mov var_4, ecx
  loc_00429FA4: and al, FEh
  loc_00429FA6: push eax
  loc_00429FA7: mov Me, eax
  loc_00429FAA: mov edx, [eax]
  loc_00429FAC: call [edx+00000004h]
  loc_00429FAF: mov eax, [0043018Ch]
  loc_00429FB4: test eax, eax
  loc_00429FB6: jnz 00429FC8h
  loc_00429FB8: push 0043018Ch
  loc_00429FBD: push 0040397Ch
  loc_00429FC2: call [004010D4h] ; __vbaNew2
  loc_00429FC8: sub esp, 00000010h
  loc_00429FCB: mov ecx, 0000000Ah
  loc_00429FD0: mov ebx, esp
  loc_00429FD2: mov var_24, ecx
  loc_00429FD5: mov eax, 80020004h
  loc_00429FDA: sub esp, 00000010h
  loc_00429FDD: mov [ebx], ecx
  loc_00429FDF: mov ecx, var_30
  loc_00429FE2: mov edx, eax
  loc_00429FE4: mov esi, [0043018Ch]
  loc_00429FEA: mov [ebx+00000004h], ecx
  loc_00429FED: mov ecx, esp
  loc_00429FEF: mov edi, [esi]
  loc_00429FF1: push esi
  loc_00429FF2: mov [ebx+00000008h], eax
  loc_00429FF5: mov eax, var_28
  loc_00429FF8: mov [ebx+0000000Ch], eax
  loc_00429FFB: mov eax, var_24
  loc_00429FFE: mov [ecx], eax
  loc_0042A000: mov eax, var_20
  loc_0042A003: mov [ecx+00000004h], eax
  loc_0042A006: mov [ecx+00000008h], edx
  loc_0042A009: mov edx, var_18
  loc_0042A00C: mov [ecx+0000000Ch], edx
  loc_0042A00F: call [edi+000002B0h]
  loc_0042A015: test eax, eax
  loc_0042A017: fnclex
  loc_0042A019: jge 0042A02Dh
  loc_0042A01B: push 000002B0h
  loc_0042A020: push 004103E4h
  loc_0042A025: push esi
  loc_0042A026: push eax
  loc_0042A027: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A02D: mov var_4, 00000000h
  loc_0042A034: mov eax, Me
  loc_0042A037: push eax
  loc_0042A038: mov ecx, [eax]
  loc_0042A03A: call [ecx+00000008h]
  loc_0042A03D: mov eax, var_4
  loc_0042A040: mov ecx, var_14
  loc_0042A043: pop edi
  loc_0042A044: pop esi
  loc_0042A045: mov fs:[00000000h], ecx
  loc_0042A04C: pop ebx
  loc_0042A04D: mov esp, ebp
  loc_0042A04F: pop ebp
  loc_0042A050: retn 0004h
End Sub

Private Sub mnuPredEst_Click() '42A4F0
  loc_0042A4F0: push ebp
  loc_0042A4F1: mov ebp, esp
  loc_0042A4F3: sub esp, 0000000Ch
  loc_0042A4F6: push 00401926h ; __vbaExceptHandler
  loc_0042A4FB: mov eax, fs:[00000000h]
  loc_0042A501: push eax
  loc_0042A502: mov fs:[00000000h], esp
  loc_0042A509: sub esp, 00000030h
  loc_0042A50C: push ebx
  loc_0042A50D: push esi
  loc_0042A50E: push edi
  loc_0042A50F: mov var_C, esp
  loc_0042A512: mov var_8, 004016F8h
  loc_0042A519: mov eax, Me
  loc_0042A51C: mov ecx, eax
  loc_0042A51E: and ecx, 00000001h
  loc_0042A521: mov var_4, ecx
  loc_0042A524: and al, FEh
  loc_0042A526: push eax
  loc_0042A527: mov Me, eax
  loc_0042A52A: mov edx, [eax]
  loc_0042A52C: call [edx+00000004h]
  loc_0042A52F: mov eax, [004300B0h]
  loc_0042A534: test eax, eax
  loc_0042A536: jnz 0042A548h
  loc_0042A538: push 004300B0h
  loc_0042A53D: push 00407528h
  loc_0042A542: call [004010D4h] ; __vbaNew2
  loc_0042A548: sub esp, 00000010h
  loc_0042A54B: mov ecx, 0000000Ah
  loc_0042A550: mov ebx, esp
  loc_0042A552: mov var_24, ecx
  loc_0042A555: mov eax, 80020004h
  loc_0042A55A: sub esp, 00000010h
  loc_0042A55D: mov [ebx], ecx
  loc_0042A55F: mov ecx, var_30
  loc_0042A562: mov edx, eax
  loc_0042A564: mov esi, [004300B0h]
  loc_0042A56A: mov [ebx+00000004h], ecx
  loc_0042A56D: mov ecx, esp
  loc_0042A56F: mov edi, [esi]
  loc_0042A571: push esi
  loc_0042A572: mov [ebx+00000008h], eax
  loc_0042A575: mov eax, var_28
  loc_0042A578: mov [ebx+0000000Ch], eax
  loc_0042A57B: mov eax, var_24
  loc_0042A57E: mov [ecx], eax
  loc_0042A580: mov eax, var_20
  loc_0042A583: mov [ecx+00000004h], eax
  loc_0042A586: mov [ecx+00000008h], edx
  loc_0042A589: mov edx, var_18
  loc_0042A58C: mov [ecx+0000000Ch], edx
  loc_0042A58F: call [edi+000002B0h]
  loc_0042A595: test eax, eax
  loc_0042A597: fnclex
  loc_0042A599: jge 0042A5ADh
  loc_0042A59B: push 000002B0h
  loc_0042A5A0: push 0040DE98h
  loc_0042A5A5: push esi
  loc_0042A5A6: push eax
  loc_0042A5A7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A5AD: mov eax, [004300C4h]
  loc_0042A5B2: test eax, eax
  loc_0042A5B4: jnz 0042A5C6h
  loc_0042A5B6: push 004300C4h
  loc_0042A5BB: push 00409228h
  loc_0042A5C0: call [004010D4h] ; __vbaNew2
  loc_0042A5C6: mov esi, [004300C4h]
  loc_0042A5CC: push esi
  loc_0042A5CD: mov eax, [esi]
  loc_0042A5CF: call [eax+000002B4h]
  loc_0042A5D5: test eax, eax
  loc_0042A5D7: fnclex
  loc_0042A5D9: jge 0042A5EDh
  loc_0042A5DB: push 000002B4h
  loc_0042A5E0: push 0040E0ECh
  loc_0042A5E5: push esi
  loc_0042A5E6: push eax
  loc_0042A5E7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A5ED: mov var_4, 00000000h
  loc_0042A5F4: mov eax, Me
  loc_0042A5F7: push eax
  loc_0042A5F8: mov ecx, [eax]
  loc_0042A5FA: call [ecx+00000008h]
  loc_0042A5FD: mov eax, var_4
  loc_0042A600: mov ecx, var_14
  loc_0042A603: pop edi
  loc_0042A604: pop esi
  loc_0042A605: mov fs:[00000000h], ecx
  loc_0042A60C: pop ebx
  loc_0042A60D: mov esp, ebp
  loc_0042A60F: pop ebp
  loc_0042A610: retn 0004h
End Sub

Private Sub optTwo_Click() '42AAC0
  loc_0042AAC0: push ebp
  loc_0042AAC1: mov ebp, esp
  loc_0042AAC3: sub esp, 0000000Ch
  loc_0042AAC6: push 00401926h ; __vbaExceptHandler
  loc_0042AACB: mov eax, fs:[00000000h]
  loc_0042AAD1: push eax
  loc_0042AAD2: mov fs:[00000000h], esp
  loc_0042AAD9: sub esp, 00000018h
  loc_0042AADC: push ebx
  loc_0042AADD: push esi
  loc_0042AADE: push edi
  loc_0042AADF: mov var_C, esp
  loc_0042AAE2: mov var_8, 00401730h
  loc_0042AAE9: mov ebx, Me
  loc_0042AAEC: mov eax, ebx
  loc_0042AAEE: and eax, 00000001h
  loc_0042AAF1: mov var_4, eax
  loc_0042AAF4: and ebx, FFFFFFFEh
  loc_0042AAF7: push ebx
  loc_0042AAF8: mov Me, ebx
  loc_0042AAFB: mov ecx, [ebx]
  loc_0042AAFD: call [ecx+00000004h]
  loc_0042AB00: mov edx, [0043005Ch]
  loc_0042AB06: xor eax, eax
  loc_0042AB08: push 0041183Ch ; "= "
  loc_0042AB0D: push edx
  loc_0042AB0E: mov var_18, eax
  loc_0042AB11: mov var_1C, eax
  loc_0042AB14: mov var_20, eax
  loc_0042AB17: mov var_24, eax
  loc_0042AB1A: call [0040107Ch] ; __vbaStrR4
  loc_0042AB20: mov esi, [004010FCh] ; __vbaStrMove
  loc_0042AB26: mov edx, eax
  loc_0042AB28: lea ecx, var_18
  loc_0042AB2B: call __vbaStrMove
  loc_0042AB2D: mov edi, [0040102Ch] ; __vbaStrCat
  loc_0042AB33: push eax
  loc_0042AB34: call edi
  loc_0042AB36: mov edx, eax
  loc_0042AB38: mov ecx, 00430054h
  loc_0042AB3D: call __vbaStrMove
  loc_0042AB3F: lea ecx, var_18
  loc_0042AB42: call [00401110h] ; __vbaFreeStr
  loc_0042AB48: mov eax, [0043005Ch]
  loc_0042AB4D: push 00411848h ; "¹ "
  loc_0042AB52: push eax
  loc_0042AB53: call [0040107Ch] ; __vbaStrR4
  loc_0042AB59: mov edx, eax
  loc_0042AB5B: lea ecx, var_18
  loc_0042AB5E: call __vbaStrMove
  loc_0042AB60: push eax
  loc_0042AB61: call edi
  loc_0042AB63: mov edx, eax
  loc_0042AB65: mov ecx, 00430058h
  loc_0042AB6A: call __vbaStrMove
  loc_0042AB6C: lea ecx, var_18
  loc_0042AB6F: call [00401110h] ; __vbaFreeStr
  loc_0042AB75: mov ecx, [00430028h]
  loc_0042AB7B: push 00411854h ; " t* > "
  loc_0042AB80: push ecx
  loc_0042AB81: call [0040107Ch] ; __vbaStrR4
  loc_0042AB87: mov edx, eax
  loc_0042AB89: lea ecx, var_18
  loc_0042AB8C: call __vbaStrMove
  loc_0042AB8E: push eax
  loc_0042AB8F: call edi
  loc_0042AB91: mov edx, eax
  loc_0042AB93: lea ecx, var_1C
  loc_0042AB96: call __vbaStrMove
  loc_0042AB98: push eax
  loc_0042AB99: push 00411470h ; " or if t* < -"
  loc_0042AB9E: call edi
  loc_0042ABA0: mov edx, eax
  loc_0042ABA2: lea ecx, var_20
  loc_0042ABA5: call __vbaStrMove
  loc_0042ABA7: mov edx, [00430028h]
  loc_0042ABAD: push eax
  loc_0042ABAE: push edx
  loc_0042ABAF: call [0040107Ch] ; __vbaStrR4
  loc_0042ABB5: mov edx, eax
  loc_0042ABB7: lea ecx, var_24
  loc_0042ABBA: call __vbaStrMove
  loc_0042ABBC: push eax
  loc_0042ABBD: call edi
  loc_0042ABBF: mov edx, eax
  loc_0042ABC1: mov ecx, 0043004Ch
  loc_0042ABC6: call __vbaStrMove
  loc_0042ABC8: lea eax, var_24
  loc_0042ABCB: push eax
  loc_0042ABCC: lea ecx, var_20
  loc_0042ABCF: lea edx, var_1C
  loc_0042ABD2: push ecx
  loc_0042ABD3: lea eax, var_18
  loc_0042ABD6: push edx
  loc_0042ABD7: push eax
  loc_0042ABD8: push 00000004h
  loc_0042ABDA: call [004010E4h] ; __vbaFreeStrList
  loc_0042ABE0: mov ecx, [ebx]
  loc_0042ABE2: add esp, 00000014h
  loc_0042ABE5: push ebx
  loc_0042ABE6: call [ecx+00000700h]
  loc_0042ABEC: mov var_4, 00000000h
  loc_0042ABF3: fwait
  loc_0042ABF4: push 0042AC18h
  loc_0042ABF9: jmp 0042AC17h
  loc_0042ABFB: lea edx, var_24
  loc_0042ABFE: lea eax, var_20
  loc_0042AC01: push edx
  loc_0042AC02: lea ecx, var_1C
  loc_0042AC05: push eax
  loc_0042AC06: lea edx, var_18
  loc_0042AC09: push ecx
  loc_0042AC0A: push edx
  loc_0042AC0B: push 00000004h
  loc_0042AC0D: call [004010E4h] ; __vbaFreeStrList
  loc_0042AC13: add esp, 00000014h
  loc_0042AC16: ret
  loc_0042AC17: ret
  loc_0042AC18: mov eax, Me
  loc_0042AC1B: push eax
  loc_0042AC1C: mov ecx, [eax]
  loc_0042AC1E: call [ecx+00000008h]
  loc_0042AC21: mov eax, var_4
  loc_0042AC24: mov ecx, var_14
  loc_0042AC27: pop edi
  loc_0042AC28: pop esi
  loc_0042AC29: mov fs:[00000000h], ecx
  loc_0042AC30: pop ebx
  loc_0042AC31: mov esp, ebp
  loc_0042AC33: pop ebp
  loc_0042AC34: retn 0004h
End Sub

Private Sub optRight_Click() '42A9A0
  loc_0042A9A0: push ebp
  loc_0042A9A1: mov ebp, esp
  loc_0042A9A3: sub esp, 0000000Ch
  loc_0042A9A6: push 00401926h ; __vbaExceptHandler
  loc_0042A9AB: mov eax, fs:[00000000h]
  loc_0042A9B1: push eax
  loc_0042A9B2: mov fs:[00000000h], esp
  loc_0042A9B9: sub esp, 0000000Ch
  loc_0042A9BC: push ebx
  loc_0042A9BD: push esi
  loc_0042A9BE: push edi
  loc_0042A9BF: mov var_C, esp
  loc_0042A9C2: mov var_8, 00401720h
  loc_0042A9C9: mov edi, Me
  loc_0042A9CC: mov eax, edi
  loc_0042A9CE: and eax, 00000001h
  loc_0042A9D1: mov var_4, eax
  loc_0042A9D4: and edi, FFFFFFFEh
  loc_0042A9D7: push edi
  loc_0042A9D8: mov Me, edi
  loc_0042A9DB: mov ecx, [edi]
  loc_0042A9DD: call [ecx+00000004h]
  loc_0042A9E0: mov edx, [0043005Ch]
  loc_0042A9E6: mov ebx, [0040107Ch] ; __vbaStrR4
  loc_0042A9EC: push 004114A0h ; "£ "
  loc_0042A9F1: push edx
  loc_0042A9F2: mov var_18, 00000000h
  loc_0042A9F9: call ebx
  loc_0042A9FB: mov esi, [004010FCh] ; __vbaStrMove
  loc_0042AA01: mov edx, eax
  loc_0042AA03: lea ecx, var_18
  loc_0042AA06: call __vbaStrMove
  loc_0042AA08: push eax
  loc_0042AA09: call [0040102Ch] ; __vbaStrCat
  loc_0042AA0F: mov edx, eax
  loc_0042AA11: mov ecx, 00430054h
  loc_0042AA16: call __vbaStrMove
  loc_0042AA18: lea ecx, var_18
  loc_0042AA1B: call [00401110h] ; __vbaFreeStr
  loc_0042AA21: mov eax, [0043005Ch]
  loc_0042AA26: push 00411868h ; "> "
  loc_0042AA2B: push eax
  loc_0042AA2C: call ebx
  loc_0042AA2E: mov edx, eax
  loc_0042AA30: lea ecx, var_18
  loc_0042AA33: call __vbaStrMove
  loc_0042AA35: push eax
  loc_0042AA36: call [0040102Ch] ; __vbaStrCat
  loc_0042AA3C: mov edx, eax
  loc_0042AA3E: mov ecx, 00430058h
  loc_0042AA43: call __vbaStrMove
  loc_0042AA45: lea ecx, var_18
  loc_0042AA48: call [00401110h] ; __vbaFreeStr
  loc_0042AA4E: mov ecx, [00430024h]
  loc_0042AA54: push 00411660h ; "t* > "
  loc_0042AA59: push ecx
  loc_0042AA5A: call ebx
  loc_0042AA5C: mov edx, eax
  loc_0042AA5E: lea ecx, var_18
  loc_0042AA61: call __vbaStrMove
  loc_0042AA63: push eax
  loc_0042AA64: call [0040102Ch] ; __vbaStrCat
  loc_0042AA6A: mov edx, eax
  loc_0042AA6C: mov ecx, 0043004Ch
  loc_0042AA71: call __vbaStrMove
  loc_0042AA73: lea ecx, var_18
  loc_0042AA76: call [00401110h] ; __vbaFreeStr
  loc_0042AA7C: mov edx, [edi]
  loc_0042AA7E: push edi
  loc_0042AA7F: call [edx+00000700h]
  loc_0042AA85: mov var_4, 00000000h
  loc_0042AA8C: fwait
  loc_0042AA8D: push 0042AA9Fh
  loc_0042AA92: jmp 0042AA9Eh
  loc_0042AA94: lea ecx, var_18
  loc_0042AA97: call [00401110h] ; __vbaFreeStr
  loc_0042AA9D: ret
  loc_0042AA9E: ret
  loc_0042AA9F: mov eax, Me
  loc_0042AAA2: push eax
  loc_0042AAA3: mov ecx, [eax]
  loc_0042AAA5: call [ecx+00000008h]
  loc_0042AAA8: mov eax, var_4
  loc_0042AAAB: mov ecx, var_14
  loc_0042AAAE: pop edi
  loc_0042AAAF: pop esi
  loc_0042AAB0: mov fs:[00000000h], ecx
  loc_0042AAB7: pop ebx
  loc_0042AAB8: mov esp, ebp
  loc_0042AABA: pop ebp
  loc_0042AABB: retn 0004h
End Sub

Private Sub mnuQuestions_Click() '42A620
  loc_0042A620: push ebp
  loc_0042A621: mov ebp, esp
  loc_0042A623: sub esp, 0000000Ch
  loc_0042A626: push 00401926h ; __vbaExceptHandler
  loc_0042A62B: mov eax, fs:[00000000h]
  loc_0042A631: push eax
  loc_0042A632: mov fs:[00000000h], esp
  loc_0042A639: sub esp, 00000030h
  loc_0042A63C: push ebx
  loc_0042A63D: push esi
  loc_0042A63E: push edi
  loc_0042A63F: mov var_C, esp
  loc_0042A642: mov var_8, 00401700h
  loc_0042A649: mov eax, Me
  loc_0042A64C: mov ecx, eax
  loc_0042A64E: and ecx, 00000001h
  loc_0042A651: mov var_4, ecx
  loc_0042A654: and al, FEh
  loc_0042A656: push eax
  loc_0042A657: mov Me, eax
  loc_0042A65A: mov edx, [eax]
  loc_0042A65C: call [edx+00000004h]
  loc_0042A65F: mov eax, [00430164h]
  loc_0042A664: test eax, eax
  loc_0042A666: jnz 0042A678h
  loc_0042A668: push 00430164h
  loc_0042A66D: push 00408108h
  loc_0042A672: call [004010D4h] ; __vbaNew2
  loc_0042A678: sub esp, 00000010h
  loc_0042A67B: mov ecx, 0000000Ah
  loc_0042A680: mov ebx, esp
  loc_0042A682: mov var_24, ecx
  loc_0042A685: mov eax, 80020004h
  loc_0042A68A: sub esp, 00000010h
  loc_0042A68D: mov [ebx], ecx
  loc_0042A68F: mov ecx, var_30
  loc_0042A692: mov edx, eax
  loc_0042A694: mov esi, [00430164h]
  loc_0042A69A: mov [ebx+00000004h], ecx
  loc_0042A69D: mov ecx, esp
  loc_0042A69F: mov edi, [esi]
  loc_0042A6A1: push esi
  loc_0042A6A2: mov [ebx+00000008h], eax
  loc_0042A6A5: mov eax, var_28
  loc_0042A6A8: mov [ebx+0000000Ch], eax
  loc_0042A6AB: mov eax, var_24
  loc_0042A6AE: mov [ecx], eax
  loc_0042A6B0: mov eax, var_20
  loc_0042A6B3: mov [ecx+00000004h], eax
  loc_0042A6B6: mov [ecx+00000008h], edx
  loc_0042A6B9: mov edx, var_18
  loc_0042A6BC: mov [ecx+0000000Ch], edx
  loc_0042A6BF: call [edi+000002B0h]
  loc_0042A6C5: test eax, eax
  loc_0042A6C7: fnclex
  loc_0042A6C9: jge 0042A6DDh
  loc_0042A6CB: push 000002B0h
  loc_0042A6D0: push 0040FF58h
  loc_0042A6D5: push esi
  loc_0042A6D6: push eax
  loc_0042A6D7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A6DD: mov eax, [004300C4h]
  loc_0042A6E2: test eax, eax
  loc_0042A6E4: jnz 0042A6F6h
  loc_0042A6E6: push 004300C4h
  loc_0042A6EB: push 00409228h
  loc_0042A6F0: call [004010D4h] ; __vbaNew2
  loc_0042A6F6: mov esi, [004300C4h]
  loc_0042A6FC: push esi
  loc_0042A6FD: mov eax, [esi]
  loc_0042A6FF: call [eax+000002B4h]
  loc_0042A705: test eax, eax
  loc_0042A707: fnclex
  loc_0042A709: jge 0042A71Dh
  loc_0042A70B: push 000002B4h
  loc_0042A710: push 0040E0ECh
  loc_0042A715: push esi
  loc_0042A716: push eax
  loc_0042A717: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A71D: mov var_4, 00000000h
  loc_0042A724: mov eax, Me
  loc_0042A727: push eax
  loc_0042A728: mov ecx, [eax]
  loc_0042A72A: call [ecx+00000008h]
  loc_0042A72D: mov eax, var_4
  loc_0042A730: mov ecx, var_14
  loc_0042A733: pop edi
  loc_0042A734: pop esi
  loc_0042A735: mov fs:[00000000h], ecx
  loc_0042A73C: pop ebx
  loc_0042A73D: mov esp, ebp
  loc_0042A73F: pop ebp
  loc_0042A740: retn 0004h
End Sub

Private Sub mnuChangeAlpha_Click() '429CA0
  loc_00429CA0: push ebp
  loc_00429CA1: mov ebp, esp
  loc_00429CA3: sub esp, 0000000Ch
  loc_00429CA6: push 00401926h ; __vbaExceptHandler
  loc_00429CAB: mov eax, fs:[00000000h]
  loc_00429CB1: push eax
  loc_00429CB2: mov fs:[00000000h], esp
  loc_00429CB9: sub esp, 00000030h
  loc_00429CBC: push ebx
  loc_00429CBD: push esi
  loc_00429CBE: push edi
  loc_00429CBF: mov var_C, esp
  loc_00429CC2: mov var_8, 004016A8h
  loc_00429CC9: mov eax, Me
  loc_00429CCC: mov ecx, eax
  loc_00429CCE: and ecx, 00000001h
  loc_00429CD1: mov var_4, ecx
  loc_00429CD4: and al, FEh
  loc_00429CD6: push eax
  loc_00429CD7: mov Me, eax
  loc_00429CDA: mov edx, [eax]
  loc_00429CDC: call [edx+00000004h]
  loc_00429CDF: mov eax, [00430178h]
  loc_00429CE4: test eax, eax
  loc_00429CE6: jnz 00429CF8h
  loc_00429CE8: push 00430178h
  loc_00429CED: push 004051D0h
  loc_00429CF2: call [004010D4h] ; __vbaNew2
  loc_00429CF8: sub esp, 00000010h
  loc_00429CFB: mov ecx, 0000000Ah
  loc_00429D00: mov ebx, esp
  loc_00429D02: mov var_24, ecx
  loc_00429D05: mov eax, 80020004h
  loc_00429D0A: sub esp, 00000010h
  loc_00429D0D: mov [ebx], ecx
  loc_00429D0F: mov ecx, var_30
  loc_00429D12: mov edx, eax
  loc_00429D14: mov esi, [00430178h]
  loc_00429D1A: mov [ebx+00000004h], ecx
  loc_00429D1D: mov ecx, esp
  loc_00429D1F: mov edi, [esi]
  loc_00429D21: push esi
  loc_00429D22: mov [ebx+00000008h], eax
  loc_00429D25: mov eax, var_28
  loc_00429D28: mov [ebx+0000000Ch], eax
  loc_00429D2B: mov eax, var_24
  loc_00429D2E: mov [ecx], eax
  loc_00429D30: mov eax, var_20
  loc_00429D33: mov [ecx+00000004h], eax
  loc_00429D36: mov [ecx+00000008h], edx
  loc_00429D39: mov edx, var_18
  loc_00429D3C: mov [ecx+0000000Ch], edx
  loc_00429D3F: call [edi+000002B0h]
  loc_00429D45: test eax, eax
  loc_00429D47: fnclex
  loc_00429D49: jge 00429D5Dh
  loc_00429D4B: push 000002B0h
  loc_00429D50: push 004100A4h
  loc_00429D55: push esi
  loc_00429D56: push eax
  loc_00429D57: call [00401030h] ; __vbaHresultCheckObj
  loc_00429D5D: mov var_4, 00000000h
  loc_00429D64: mov eax, Me
  loc_00429D67: push eax
  loc_00429D68: mov ecx, [eax]
  loc_00429D6A: call [ecx+00000008h]
  loc_00429D6D: mov eax, var_4
  loc_00429D70: mov ecx, var_14
  loc_00429D73: pop edi
  loc_00429D74: pop esi
  loc_00429D75: mov fs:[00000000h], ecx
  loc_00429D7C: pop ebx
  loc_00429D7D: mov esp, ebp
  loc_00429D7F: pop ebp
  loc_00429D80: retn 0004h
End Sub

Private Sub mnuChangeUnits_Click() '42A060
  loc_0042A060: push ebp
  loc_0042A061: mov ebp, esp
  loc_0042A063: sub esp, 0000000Ch
  loc_0042A066: push 00401926h ; __vbaExceptHandler
  loc_0042A06B: mov eax, fs:[00000000h]
  loc_0042A071: push eax
  loc_0042A072: mov fs:[00000000h], esp
  loc_0042A079: sub esp, 00000030h
  loc_0042A07C: push ebx
  loc_0042A07D: push esi
  loc_0042A07E: push edi
  loc_0042A07F: mov var_C, esp
  loc_0042A082: mov var_8, 004016C8h
  loc_0042A089: mov eax, Me
  loc_0042A08C: mov ecx, eax
  loc_0042A08E: and ecx, 00000001h
  loc_0042A091: mov var_4, ecx
  loc_0042A094: and al, FEh
  loc_0042A096: push eax
  loc_0042A097: mov Me, eax
  loc_0042A09A: mov edx, [eax]
  loc_0042A09C: call [edx+00000004h]
  loc_0042A09F: mov eax, [004301A0h]
  loc_0042A0A4: test eax, eax
  loc_0042A0A6: jnz 0042A0B8h
  loc_0042A0A8: push 004301A0h
  loc_0042A0AD: push 004033C0h
  loc_0042A0B2: call [004010D4h] ; __vbaNew2
  loc_0042A0B8: sub esp, 00000010h
  loc_0042A0BB: mov ecx, 0000000Ah
  loc_0042A0C0: mov ebx, esp
  loc_0042A0C2: mov var_24, ecx
  loc_0042A0C5: mov eax, 80020004h
  loc_0042A0CA: sub esp, 00000010h
  loc_0042A0CD: mov [ebx], ecx
  loc_0042A0CF: mov ecx, var_30
  loc_0042A0D2: mov edx, eax
  loc_0042A0D4: mov esi, [004301A0h]
  loc_0042A0DA: mov [ebx+00000004h], ecx
  loc_0042A0DD: mov ecx, esp
  loc_0042A0DF: mov edi, [esi]
  loc_0042A0E1: push esi
  loc_0042A0E2: mov [ebx+00000008h], eax
  loc_0042A0E5: mov eax, var_28
  loc_0042A0E8: mov [ebx+0000000Ch], eax
  loc_0042A0EB: mov eax, var_24
  loc_0042A0EE: mov [ecx], eax
  loc_0042A0F0: mov eax, var_20
  loc_0042A0F3: mov [ecx+00000004h], eax
  loc_0042A0F6: mov [ecx+00000008h], edx
  loc_0042A0F9: mov edx, var_18
  loc_0042A0FC: mov [ecx+0000000Ch], edx
  loc_0042A0FF: call [edi+000002B0h]
  loc_0042A105: test eax, eax
  loc_0042A107: fnclex
  loc_0042A109: jge 0042A11Dh
  loc_0042A10B: push 000002B0h
  loc_0042A110: push 00410454h
  loc_0042A115: push esi
  loc_0042A116: push eax
  loc_0042A117: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A11D: mov var_4, 00000000h
  loc_0042A124: mov eax, Me
  loc_0042A127: push eax
  loc_0042A128: mov ecx, [eax]
  loc_0042A12A: call [ecx+00000008h]
  loc_0042A12D: mov eax, var_4
  loc_0042A130: mov ecx, var_14
  loc_0042A133: pop edi
  loc_0042A134: pop esi
  loc_0042A135: mov fs:[00000000h], ecx
  loc_0042A13C: pop ebx
  loc_0042A13D: mov esp, ebp
  loc_0042A13F: pop ebp
  loc_0042A140: retn 0004h
End Sub

Private Sub lblTestCon_Click() '429AD0
  loc_00429AD0: push ebp
  loc_00429AD1: mov ebp, esp
  loc_00429AD3: sub esp, 0000000Ch
  loc_00429AD6: push 00401926h ; __vbaExceptHandler
  loc_00429ADB: mov eax, fs:[00000000h]
  loc_00429AE1: push eax
  loc_00429AE2: mov fs:[00000000h], esp
  loc_00429AE9: sub esp, 00000008h
  loc_00429AEC: push ebx
  loc_00429AED: push esi
  loc_00429AEE: push edi
  loc_00429AEF: mov var_C, esp
  loc_00429AF2: mov var_8, 00401698h
  loc_00429AF9: mov esi, Me
  loc_00429AFC: mov eax, esi
  loc_00429AFE: and eax, 00000001h
  loc_00429B01: mov var_4, eax
  loc_00429B04: and esi, FFFFFFFEh
  loc_00429B07: push esi
  loc_00429B08: mov Me, esi
  loc_00429B0B: mov ecx, [esi]
  loc_00429B0D: call [ecx+00000004h]
  loc_00429B10: mov ax, [00430068h]
  loc_00429B16: add ax, 0001h
  loc_00429B1A: jo 00429B60h
  loc_00429B1C: cmp ax, 0005h
  loc_00429B20: mov [00430068h], ax
  loc_00429B26: jnz 00429B31h
  loc_00429B28: mov [00430068h], 0001h
  loc_00429B31: mov edx, [esi]
  loc_00429B33: push esi
  loc_00429B34: call [edx+00000700h]
  loc_00429B3A: mov var_4, 00000000h
  loc_00429B41: mov eax, Me
  loc_00429B44: push eax
  loc_00429B45: mov ecx, [eax]
  loc_00429B47: call [ecx+00000008h]
  loc_00429B4A: mov eax, var_4
  loc_00429B4D: mov ecx, var_14
  loc_00429B50: pop edi
  loc_00429B51: pop esi
  loc_00429B52: mov fs:[00000000h], ecx
  loc_00429B59: pop ebx
  loc_00429B5A: mov esp, ebp
  loc_00429B5C: pop ebp
  loc_00429B5D: retn 0004h
End Sub

Private Sub mnuIntro_Click() '42A3C0
  loc_0042A3C0: push ebp
  loc_0042A3C1: mov ebp, esp
  loc_0042A3C3: sub esp, 0000000Ch
  loc_0042A3C6: push 00401926h ; __vbaExceptHandler
  loc_0042A3CB: mov eax, fs:[00000000h]
  loc_0042A3D1: push eax
  loc_0042A3D2: mov fs:[00000000h], esp
  loc_0042A3D9: sub esp, 00000030h
  loc_0042A3DC: push ebx
  loc_0042A3DD: push esi
  loc_0042A3DE: push edi
  loc_0042A3DF: mov var_C, esp
  loc_0042A3E2: mov var_8, 004016F0h
  loc_0042A3E9: mov eax, Me
  loc_0042A3EC: mov ecx, eax
  loc_0042A3EE: and ecx, 00000001h
  loc_0042A3F1: mov var_4, ecx
  loc_0042A3F4: and al, FEh
  loc_0042A3F6: push eax
  loc_0042A3F7: mov Me, eax
  loc_0042A3FA: mov edx, [eax]
  loc_0042A3FC: call [edx+00000004h]
  loc_0042A3FF: mov eax, [00430088h]
  loc_0042A404: test eax, eax
  loc_0042A406: jnz 0042A418h
  loc_0042A408: push 00430088h
  loc_0042A40D: push 004058C0h
  loc_0042A412: call [004010D4h] ; __vbaNew2
  loc_0042A418: sub esp, 00000010h
  loc_0042A41B: mov ecx, 0000000Ah
  loc_0042A420: mov ebx, esp
  loc_0042A422: mov var_24, ecx
  loc_0042A425: mov eax, 80020004h
  loc_0042A42A: sub esp, 00000010h
  loc_0042A42D: mov [ebx], ecx
  loc_0042A42F: mov ecx, var_30
  loc_0042A432: mov edx, eax
  loc_0042A434: mov esi, [00430088h]
  loc_0042A43A: mov [ebx+00000004h], ecx
  loc_0042A43D: mov ecx, esp
  loc_0042A43F: mov edi, [esi]
  loc_0042A441: push esi
  loc_0042A442: mov [ebx+00000008h], eax
  loc_0042A445: mov eax, var_28
  loc_0042A448: mov [ebx+0000000Ch], eax
  loc_0042A44B: mov eax, var_24
  loc_0042A44E: mov [ecx], eax
  loc_0042A450: mov eax, var_20
  loc_0042A453: mov [ecx+00000004h], eax
  loc_0042A456: mov [ecx+00000008h], edx
  loc_0042A459: mov edx, var_18
  loc_0042A45C: mov [ecx+0000000Ch], edx
  loc_0042A45F: call [edi+000002B0h]
  loc_0042A465: test eax, eax
  loc_0042A467: fnclex
  loc_0042A469: jge 0042A47Dh
  loc_0042A46B: push 000002B0h
  loc_0042A470: push 0040DB64h
  loc_0042A475: push esi
  loc_0042A476: push eax
  loc_0042A477: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A47D: mov eax, [004300C4h]
  loc_0042A482: test eax, eax
  loc_0042A484: jnz 0042A496h
  loc_0042A486: push 004300C4h
  loc_0042A48B: push 00409228h
  loc_0042A490: call [004010D4h] ; __vbaNew2
  loc_0042A496: mov esi, [004300C4h]
  loc_0042A49C: push esi
  loc_0042A49D: mov eax, [esi]
  loc_0042A49F: call [eax+000002B4h]
  loc_0042A4A5: test eax, eax
  loc_0042A4A7: fnclex
  loc_0042A4A9: jge 0042A4BDh
  loc_0042A4AB: push 000002B4h
  loc_0042A4B0: push 0040E0ECh
  loc_0042A4B5: push esi
  loc_0042A4B6: push eax
  loc_0042A4B7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042A4BD: mov var_4, 00000000h
  loc_0042A4C4: mov eax, Me
  loc_0042A4C7: push eax
  loc_0042A4C8: mov ecx, [eax]
  loc_0042A4CA: call [ecx+00000008h]
  loc_0042A4CD: mov eax, var_4
  loc_0042A4D0: mov ecx, var_14
  loc_0042A4D3: pop edi
  loc_0042A4D4: pop esi
  loc_0042A4D5: mov fs:[00000000h], ecx
  loc_0042A4DC: pop ebx
  loc_0042A4DD: mov esp, ebp
  loc_0042A4DF: pop ebp
  loc_0042A4E0: retn 0004h
End Sub

Private Sub Proc_14_20_427D50
  loc_00427D50: push ebp
  loc_00427D51: mov ebp, esp
  loc_00427D53: sub esp, 00000008h
  loc_00427D56: push 00401926h ; __vbaExceptHandler
  loc_00427D5B: mov eax, fs:[00000000h]
  loc_00427D61: push eax
  loc_00427D62: mov fs:[00000000h], esp
  loc_00427D69: sub esp, 00000064h
  loc_00427D6C: push ebx
  loc_00427D6D: push esi
  loc_00427D6E: push edi
  loc_00427D6F: mov var_8, esp
  loc_00427D72: mov var_4, 00401670h
  loc_00427D79: mov ebx, Me
  loc_00427D7C: xor edi, edi
  loc_00427D7E: push edi
  loc_00427D7F: push 00000004h
  loc_00427D81: mov eax, [ebx]
  loc_00427D83: push ebx
  loc_00427D84: mov var_14, edi
  loc_00427D87: mov var_18, edi
  loc_00427D8A: mov var_1C, edi
  loc_00427D8D: mov var_20, edi
  loc_00427D90: mov var_24, edi
  loc_00427D93: mov var_28, edi
  loc_00427D96: mov var_2C, edi
  loc_00427D99: mov var_30, edi
  loc_00427D9C: mov var_34, edi
  loc_00427D9F: mov var_38, edi
  loc_00427DA2: mov var_3C, edi
  loc_00427DA5: mov var_40, edi
  loc_00427DA8: mov var_44, edi
  loc_00427DAB: mov var_54, edi
  loc_00427DAE: call [eax+000003A0h]
  loc_00427DB4: lea ecx, var_44
  loc_00427DB7: push eax
  loc_00427DB8: push ecx
  loc_00427DB9: call [0040103Ch] ; __vbaObjSet
  loc_00427DBF: lea edx, var_54
  loc_00427DC2: push eax
  loc_00427DC3: push edx
  loc_00427DC4: call [00401088h] ; __vbaLateIdCallLd
  loc_00427DCA: add esp, 00000010h
  loc_00427DCD: push eax
  loc_00427DCE: call [004010C4h] ; __vbaI2Var
  loc_00427DD4: mov si, ax
  loc_00427DD7: lea ecx, var_44
  loc_00427DDA: dec si
  loc_00427DDC: neg si
  loc_00427DDF: sbb esi, esi
  loc_00427DE1: inc esi
  loc_00427DE2: neg esi
  loc_00427DE4: call [00401114h] ; __vbaFreeObj
  loc_00427DEA: lea ecx, var_54
  loc_00427DED: call [00401010h] ; __vbaFreeVar
  loc_00427DF3: cmp si, di
  loc_00427DF6: jz 00427E3Dh
  loc_00427DF8: mov eax, [ebx]
  loc_00427DFA: push ebx
  loc_00427DFB: call [eax+00000378h]
  loc_00427E01: lea ecx, var_44
  loc_00427E04: push eax
  loc_00427E05: push ecx
  loc_00427E06: call [0040103Ch] ; __vbaObjSet
  loc_00427E0C: mov esi, eax
  loc_00427E0E: push FFFFFFFFh
  loc_00427E10: push esi
  loc_00427E11: mov edx, [esi]
  loc_00427E13: call [edx+00000074h]
  loc_00427E16: cmp eax, edi
  loc_00427E18: fnclex
  loc_00427E1A: jge 00427E2Bh
  loc_00427E1C: push 00000074h
  loc_00427E1E: push 0041148Ch
  loc_00427E23: push esi
  loc_00427E24: push eax
  loc_00427E25: call [00401030h] ; __vbaHresultCheckObj
  loc_00427E2B: lea ecx, var_44
  loc_00427E2E: call [00401114h] ; __vbaFreeObj
  loc_00427E34: mov eax, [ebx]
  loc_00427E36: push ebx
  loc_00427E37: call [eax+00000708h]
  loc_00427E3D: mov ecx, [ebx]
  loc_00427E3F: push edi
  loc_00427E40: push 00000004h
  loc_00427E42: push ebx
  loc_00427E43: call [ecx+000003A0h]
  loc_00427E49: lea edx, var_44
  loc_00427E4C: push eax
  loc_00427E4D: push edx
  loc_00427E4E: call [0040103Ch] ; __vbaObjSet
  loc_00427E54: push eax
  loc_00427E55: lea eax, var_54
  loc_00427E58: push eax
  loc_00427E59: call [00401088h] ; __vbaLateIdCallLd
  loc_00427E5F: add esp, 00000010h
  loc_00427E62: push eax
  loc_00427E63: call [004010C4h] ; __vbaI2Var
  loc_00427E69: mov si, ax
  loc_00427E6C: lea ecx, var_44
  loc_00427E6F: neg si
  loc_00427E72: sbb esi, esi
  loc_00427E74: inc esi
  loc_00427E75: neg esi
  loc_00427E77: call [00401114h] ; __vbaFreeObj
  loc_00427E7D: lea ecx, var_54
  loc_00427E80: call [00401010h] ; __vbaFreeVar
  loc_00427E86: cmp si, di
  loc_00427E89: jz 00428561h
  loc_00427E8F: mov ecx, [ebx]
  loc_00427E91: push ebx
  loc_00427E92: call [ecx+00000378h]
  loc_00427E98: lea edx, var_44
  loc_00427E9B: push eax
  loc_00427E9C: push edx
  loc_00427E9D: call [0040103Ch] ; __vbaObjSet
  loc_00427EA3: mov esi, eax
  loc_00427EA5: push edi
  loc_00427EA6: push esi
  loc_00427EA7: mov eax, [esi]
  loc_00427EA9: call [eax+00000074h]
  loc_00427EAC: cmp eax, edi
  loc_00427EAE: fnclex
  loc_00427EB0: jge 00427EC1h
  loc_00427EB2: push 00000074h
  loc_00427EB4: push 0041148Ch
  loc_00427EB9: push esi
  loc_00427EBA: push eax
  loc_00427EBB: call [00401030h] ; __vbaHresultCheckObj
  loc_00427EC1: lea ecx, var_44
  loc_00427EC4: call [00401114h] ; __vbaFreeObj
  loc_00427ECA: fld real4 ptr [00430064h]
  loc_00427ED0: fmul st0, real4 ptr [00430028h]
  loc_00427ED6: push 00410694h ; "( "
  loc_00427EDB: fadd st0, real4 ptr [00430060h]
  loc_00427EE1: fstp real4 ptr [ebx+0000003Ch]
  loc_00427EE4: fnstsw ax
  loc_00427EE6: test al, 0Dh
  loc_00427EE8: jnz 004285CDh
  loc_00427EEE: fld real4 ptr [00430064h]
  loc_00427EF4: fmul st0, real4 ptr [00430028h]
  loc_00427EFA: fsubr st0, real4 ptr [00430060h]
  loc_00427F00: fstp real4 ptr [ebx+00000040h]
  loc_00427F03: fnstsw ax
  loc_00427F05: test al, 0Dh
  loc_00427F07: jnz 004285CDh
  loc_00427F0D: mov ecx, [00430060h]
  loc_00427F13: push ecx
  loc_00427F14: call [0040107Ch] ; __vbaStrR4
  loc_00427F1A: mov esi, [004010FCh] ; __vbaStrMove
  loc_00427F20: mov edx, eax
  loc_00427F22: lea ecx, var_14
  loc_00427F25: call __vbaStrMove
  loc_00427F27: mov edi, [0040102Ch] ; __vbaStrCat
  loc_00427F2D: push eax
  loc_00427F2E: call edi
  loc_00427F30: mov edx, eax
  loc_00427F32: lea ecx, var_18
  loc_00427F35: call __vbaStrMove
  loc_00427F37: push eax
  loc_00427F38: push 004106A0h ; " )"
  loc_00427F3D: call edi
  loc_00427F3F: mov edx, eax
  loc_00427F41: lea ecx, var_1C
  loc_00427F44: call __vbaStrMove
  loc_00427F46: push eax
  loc_00427F47: push 004106ACh ; " ± "
  loc_00427F4C: call edi
  loc_00427F4E: mov edx, eax
  loc_00427F50: lea ecx, var_20
  loc_00427F53: call __vbaStrMove
  loc_00427F55: push eax
  loc_00427F56: push 00410694h ; "( "
  loc_00427F5B: call edi
  loc_00427F5D: mov edx, eax
  loc_00427F5F: lea ecx, var_24
  loc_00427F62: call __vbaStrMove
  loc_00427F64: mov edx, [00430028h]
  loc_00427F6A: push eax
  loc_00427F6B: push edx
  loc_00427F6C: call [0040107Ch] ; __vbaStrR4
  loc_00427F72: mov edx, eax
  loc_00427F74: lea ecx, var_28
  loc_00427F77: call __vbaStrMove
  loc_00427F79: push eax
  loc_00427F7A: call edi
  loc_00427F7C: mov edx, eax
  loc_00427F7E: lea ecx, var_2C
  loc_00427F81: call __vbaStrMove
  loc_00427F83: push eax
  loc_00427F84: push 004106B8h ; " ) * ( "
  loc_00427F89: call edi
  loc_00427F8B: mov edx, eax
  loc_00427F8D: lea ecx, var_30
  loc_00427F90: call __vbaStrMove
  loc_00427F92: push eax
  loc_00427F93: mov eax, [00430064h]
  loc_00427F98: push eax
  loc_00427F99: call [0040107Ch] ; __vbaStrR4
  loc_00427F9F: mov edx, eax
  loc_00427FA1: lea ecx, var_34
  loc_00427FA4: call __vbaStrMove
  loc_00427FA6: push eax
  loc_00427FA7: call edi
  loc_00427FA9: mov edx, eax
  loc_00427FAB: lea ecx, var_38
  loc_00427FAE: call __vbaStrMove
  loc_00427FB0: push eax
  loc_00427FB1: push 004106A0h ; " )"
  loc_00427FB6: call edi
  loc_00427FB8: mov edx, eax
  loc_00427FBA: lea ecx, var_3C
  loc_00427FBD: call __vbaStrMove
  loc_00427FBF: mov edx, eax
  loc_00427FC1: lea ecx, [ebx+00000048h]
  loc_00427FC4: call [004010E0h] ; __vbaStrCopy
  loc_00427FCA: lea ecx, var_3C
  loc_00427FCD: lea edx, var_38
  loc_00427FD0: push ecx
  loc_00427FD1: lea eax, var_34
  loc_00427FD4: push edx
  loc_00427FD5: lea ecx, var_30
  loc_00427FD8: push eax
  loc_00427FD9: lea edx, var_2C
  loc_00427FDC: push ecx
  loc_00427FDD: lea eax, var_28
  loc_00427FE0: push edx
  loc_00427FE1: lea ecx, var_24
  loc_00427FE4: push eax
  loc_00427FE5: lea edx, var_20
  loc_00427FE8: push ecx
  loc_00427FE9: lea eax, var_1C
  loc_00427FEC: push edx
  loc_00427FED: lea ecx, var_18
  loc_00427FF0: push eax
  loc_00427FF1: lea edx, var_14
  loc_00427FF4: push ecx
  loc_00427FF5: push edx
  loc_00427FF6: push 0000000Bh
  loc_00427FF8: call [004010E4h] ; __vbaFreeStrList
  loc_00427FFE: mov eax, [ebx+00000048h]
  loc_00428001: add esp, 00000030h
  loc_00428004: push eax
  loc_00428005: push 004106CCh ; " = "
  loc_0042800A: call edi
  loc_0042800C: mov edx, eax
  loc_0042800E: lea ecx, var_14
  loc_00428011: call __vbaStrMove
  loc_00428013: mov ecx, [ebx+00000040h]
  loc_00428016: push eax
  loc_00428017: push ecx
  loc_00428018: call [0040107Ch] ; __vbaStrR4
  loc_0042801E: mov edx, eax
  loc_00428020: lea ecx, var_18
  loc_00428023: call __vbaStrMove
  loc_00428025: push eax
  loc_00428026: call edi
  loc_00428028: mov edx, eax
  loc_0042802A: lea ecx, var_1C
  loc_0042802D: call __vbaStrMove
  loc_0042802F: push eax
  loc_00428030: push 004106D8h ; " to "
  loc_00428035: call edi
  loc_00428037: mov edx, eax
  loc_00428039: lea ecx, var_20
  loc_0042803C: call __vbaStrMove
  loc_0042803E: mov edx, [ebx+0000003Ch]
  loc_00428041: push eax
  loc_00428042: push edx
  loc_00428043: call [0040107Ch] ; __vbaStrR4
  loc_00428049: mov edx, eax
  loc_0042804B: lea ecx, var_24
  loc_0042804E: call __vbaStrMove
  loc_00428050: push eax
  loc_00428051: call edi
  loc_00428053: mov edx, eax
  loc_00428055: lea ecx, var_28
  loc_00428058: call __vbaStrMove
  loc_0042805A: mov edx, eax
  loc_0042805C: lea ecx, [ebx+00000048h]
  loc_0042805F: call [004010E0h] ; __vbaStrCopy
  loc_00428065: lea eax, var_28
  loc_00428068: lea ecx, var_24
  loc_0042806B: push eax
  loc_0042806C: lea edx, var_20
  loc_0042806F: push ecx
  loc_00428070: lea eax, var_1C
  loc_00428073: push edx
  loc_00428074: push eax
  loc_00428075: lea ecx, var_18
  loc_00428078: lea edx, var_14
  loc_0042807B: push ecx
  loc_0042807C: push edx
  loc_0042807D: push 00000006h
  loc_0042807F: call [004010E4h] ; __vbaFreeStrList
  loc_00428085: mov eax, [ebx]
  loc_00428087: add esp, 0000001Ch
  loc_0042808A: push ebx
  loc_0042808B: call [eax+00000360h]
  loc_00428091: lea ecx, var_44
  loc_00428094: push eax
  loc_00428095: push ecx
  loc_00428096: call [0040103Ch] ; __vbaObjSet
  loc_0042809C: mov ecx, [ebx+00000048h]
  loc_0042809F: mov edx, [eax]
  loc_004280A1: push ecx
  loc_004280A2: push eax
  loc_004280A3: mov var_58, eax
  loc_004280A6: call [edx+00000054h]
  loc_004280A9: test eax, eax
  loc_004280AB: fnclex
  loc_004280AD: jge 004280C1h
  loc_004280AF: mov edx, var_58
  loc_004280B2: push 00000054h
  loc_004280B4: push 0040E390h
  loc_004280B9: push edx
  loc_004280BA: push eax
  loc_004280BB: call [00401030h] ; __vbaHresultCheckObj
  loc_004280C1: lea ecx, var_44
  loc_004280C4: call [00401114h] ; __vbaFreeObj
  loc_004280CA: fld real4 ptr [00430020h]
  loc_004280D0: fsubr st0, real8 ptr [004011F8h]
  loc_004280D6: lea ecx, [ebx+00000044h]
  loc_004280D9: push 004106E8h ; "With "
  loc_004280DE: mov var_68, ecx
  loc_004280E1: fmul st0, real8 ptr [00401520h]
  loc_004280E7: fnstsw ax
  loc_004280E9: test al, 0Dh
  loc_004280EB: jnz 004285CDh
  loc_004280F1: fstp real4 ptr var_64
  loc_004280F4: mov eax, var_64
  loc_004280F7: push eax
  loc_004280F8: mov [ebx+00000038h], eax
  loc_004280FB: call [0040107Ch] ; __vbaStrR4
  loc_00428101: mov edx, eax
  loc_00428103: lea ecx, var_14
  loc_00428106: call __vbaStrMove
  loc_00428108: push eax
  loc_00428109: call edi
  loc_0042810B: mov edx, eax
  loc_0042810D: lea ecx, var_18
  loc_00428110: call __vbaStrMove
  loc_00428112: push eax
  loc_00428113: push 00411878h ; Chr(37) & " confidence we can say that for each additional "
  loc_00428118: call edi
  loc_0042811A: mov edx, eax
  loc_0042811C: lea ecx, var_1C
  loc_0042811F: call __vbaStrMove
  loc_00428121: mov edx, [0043001Ch]
  loc_00428127: push eax
  loc_00428128: push edx
  loc_00428129: call edi
  loc_0042812B: mov edx, eax
  loc_0042812D: lea ecx, var_20
  loc_00428130: call __vbaStrMove
  loc_00428132: push eax
  loc_00428133: push 00410E74h ; " increase in "
  loc_00428138: call edi
  loc_0042813A: mov edx, eax
  loc_0042813C: lea ecx, var_24
  loc_0042813F: call __vbaStrMove
  loc_00428141: push eax
  loc_00428142: mov eax, [00430018h]
  loc_00428147: push eax
  loc_00428148: call edi
  loc_0042814A: mov edx, eax
  loc_0042814C: lea ecx, var_28
  loc_0042814F: call __vbaStrMove
  loc_00428151: mov ecx, var_68
  loc_00428154: mov edx, eax
  loc_00428156: call [004010E0h] ; __vbaStrCopy
  loc_0042815C: lea ecx, var_28
  loc_0042815F: lea edx, var_24
  loc_00428162: push ecx
  loc_00428163: lea eax, var_20
  loc_00428166: push edx
  loc_00428167: lea ecx, var_1C
  loc_0042816A: push eax
  loc_0042816B: lea edx, var_18
  loc_0042816E: push ecx
  loc_0042816F: lea eax, var_14
  loc_00428172: push edx
  loc_00428173: push eax
  loc_00428174: push 00000006h
  loc_00428176: call [004010E4h] ; __vbaFreeStrList
  loc_0042817C: fld real4 ptr [ebx+0000003Ch]
  loc_0042817F: fcomp real4 ptr [004011E8h]
  loc_00428185: add esp, 0000001Ch
  loc_00428188: fnstsw ax
  loc_0042818A: test ah, 41h
  loc_0042818D: jz 00428196h
  loc_0042818F: mov ecx, 00000001h
  loc_00428194: jmp 00428198h
  loc_00428196: xor ecx, ecx
  loc_00428198: fld real4 ptr [ebx+00000040h]
  loc_0042819B: fcomp real4 ptr [004011E8h]
  loc_004281A1: fnstsw ax
  loc_004281A3: test ah, 41h
  loc_004281A6: jz 004281AFh
  loc_004281A8: mov eax, 00000001h
  loc_004281AD: jmp 004281B1h
  loc_004281AF: xor eax, eax
  loc_004281B1: or ecx, eax
  loc_004281B3: jnz 004282B3h
  loc_004281B9: mov ecx, var_68
  loc_004281BC: mov edx, [ecx]
  loc_004281BE: push edx
  loc_004281BF: push 00410760h ; ", the mean value of "
  loc_004281C4: call edi
  loc_004281C6: mov edx, eax
  loc_004281C8: lea ecx, var_14
  loc_004281CB: call __vbaStrMove
  loc_004281CD: push eax
  loc_004281CE: mov eax, [00430010h]
  loc_004281D3: push eax
  loc_004281D4: call edi
  loc_004281D6: mov edx, eax
  loc_004281D8: lea ecx, var_18
  loc_004281DB: call __vbaStrMove
  loc_004281DD: push eax
  loc_004281DE: push 004118E0h ; " increases from "
  loc_004281E3: call edi
  loc_004281E5: mov edx, eax
  loc_004281E7: lea ecx, var_1C
  loc_004281EA: call __vbaStrMove
  loc_004281EC: mov ecx, [ebx+00000040h]
  loc_004281EF: push eax
  loc_004281F0: push ecx
  loc_004281F1: call [0040107Ch] ; __vbaStrR4
  loc_004281F7: mov edx, eax
  loc_004281F9: lea ecx, var_20
  loc_004281FC: call __vbaStrMove
  loc_004281FE: push eax
  loc_004281FF: call edi
  loc_00428201: mov edx, eax
  loc_00428203: lea ecx, var_24
  loc_00428206: call __vbaStrMove
  loc_00428208: push eax
  loc_00428209: push 0040F028h
  loc_0042820E: call edi
  loc_00428210: mov edx, eax
  loc_00428212: lea ecx, var_28
  loc_00428215: call __vbaStrMove
  loc_00428217: mov edx, [00430014h]
  loc_0042821D: push eax
  loc_0042821E: push edx
  loc_0042821F: call edi
  loc_00428221: mov edx, eax
  loc_00428223: lea ecx, var_2C
  loc_00428226: call __vbaStrMove
  loc_00428228: push eax
  loc_00428229: push 004106D8h ; " to "
  loc_0042822E: call edi
  loc_00428230: mov edx, eax
  loc_00428232: lea ecx, var_30
  loc_00428235: call __vbaStrMove
  loc_00428237: push eax
  loc_00428238: mov eax, [ebx+0000003Ch]
  loc_0042823B: push eax
  loc_0042823C: call [0040107Ch] ; __vbaStrR4
  loc_00428242: mov edx, eax
  loc_00428244: lea ecx, var_34
  loc_00428247: call __vbaStrMove
  loc_00428249: push eax
  loc_0042824A: call edi
  loc_0042824C: mov edx, eax
  loc_0042824E: lea ecx, var_38
  loc_00428251: call __vbaStrMove
  loc_00428253: push eax
  loc_00428254: push 0040F028h
  loc_00428259: call edi
  loc_0042825B: mov edx, eax
  loc_0042825D: lea ecx, var_3C
  loc_00428260: call __vbaStrMove
  loc_00428262: mov ecx, [00430014h]
  loc_00428268: push eax
  loc_00428269: push ecx
  loc_0042826A: call edi
  loc_0042826C: mov edx, eax
  loc_0042826E: lea ecx, var_40
  loc_00428271: call __vbaStrMove
  loc_00428273: mov ecx, var_68
  loc_00428276: mov edx, eax
  loc_00428278: call [004010E0h] ; __vbaStrCopy
  loc_0042827E: lea edx, var_40
  loc_00428281: lea eax, var_3C
  loc_00428284: push edx
  loc_00428285: lea ecx, var_38
  loc_00428288: push eax
  loc_00428289: lea edx, var_34
  loc_0042828C: push ecx
  loc_0042828D: lea eax, var_30
  loc_00428290: push edx
  loc_00428291: lea ecx, var_2C
  loc_00428294: push eax
  loc_00428295: lea edx, var_28
  loc_00428298: push ecx
  loc_00428299: lea eax, var_24
  loc_0042829C: push edx
  loc_0042829D: lea ecx, var_20
  loc_004282A0: push eax
  loc_004282A1: lea edx, var_1C
  loc_004282A4: push ecx
  loc_004282A5: lea eax, var_18
  loc_004282A8: push edx
  loc_004282A9: lea ecx, var_14
  loc_004282AC: push eax
  loc_004282AD: push ecx
  loc_004282AE: jmp 00428516h
  loc_004282B3: fld real4 ptr [ebx+0000003Ch]
  loc_004282B6: fcomp real4 ptr [004011E8h]
  loc_004282BC: fnstsw ax
  loc_004282BE: test ah, 41h
  loc_004282C1: jz 004282CAh
  loc_004282C3: mov ecx, 00000001h
  loc_004282C8: jmp 004282CCh
  loc_004282CA: xor ecx, ecx
  loc_004282CC: fld real4 ptr [ebx+00000040h]
  loc_004282CF: fcomp real4 ptr [004011E8h]
  loc_004282D5: fnstsw ax
  loc_004282D7: test ah, 01h
  loc_004282DA: jnz 004282E3h
  loc_004282DC: mov eax, 00000001h
  loc_004282E1: jmp 004282E5h
  loc_004282E3: xor eax, eax
  loc_004282E5: mov edx, var_68
  loc_004282E8: or ecx, eax
  loc_004282EA: mov eax, [edx]
  loc_004282EC: push eax
  loc_004282ED: jnz 004283FDh
  loc_004282F3: push 00411908h ; ", the change in the mean value of "
  loc_004282F8: call edi
  loc_004282FA: mov edx, eax
  loc_004282FC: lea ecx, var_14
  loc_004282FF: call __vbaStrMove
  loc_00428301: mov ecx, [00430010h]
  loc_00428307: push eax
  loc_00428308: push ecx
  loc_00428309: call edi
  loc_0042830B: mov edx, eax
  loc_0042830D: lea ecx, var_18
  loc_00428310: call __vbaStrMove
  loc_00428312: push eax
  loc_00428313: push 00411954h ; " ranges from a decrease of "
  loc_00428318: call edi
  loc_0042831A: mov edx, eax
  loc_0042831C: lea ecx, var_1C
  loc_0042831F: call __vbaStrMove
  loc_00428321: fld real4 ptr [ebx+00000040h]
  loc_00428324: fabs
  loc_00428326: push eax
  loc_00428327: push ecx
  loc_00428328: fnstsw ax
  loc_0042832A: test al, 0Dh
  loc_0042832C: jnz 004285CDh
  loc_00428332: fstp real4 ptr var_6C
  loc_00428335: fld real4 ptr var_6C
  loc_00428338: fstp real4 ptr [esp]
  loc_0042833B: call [0040107Ch] ; __vbaStrR4
  loc_00428341: mov edx, eax
  loc_00428343: lea ecx, var_20
  loc_00428346: call __vbaStrMove
  loc_00428348: push eax
  loc_00428349: call edi
  loc_0042834B: mov edx, eax
  loc_0042834D: lea ecx, var_24
  loc_00428350: call __vbaStrMove
  loc_00428352: push eax
  loc_00428353: push 0040F028h
  loc_00428358: call edi
  loc_0042835A: mov edx, eax
  loc_0042835C: lea ecx, var_28
  loc_0042835F: call __vbaStrMove
  loc_00428361: mov edx, [00430014h]
  loc_00428367: push eax
  loc_00428368: push edx
  loc_00428369: call edi
  loc_0042836B: mov edx, eax
  loc_0042836D: lea ecx, var_2C
  loc_00428370: call __vbaStrMove
  loc_00428372: push eax
  loc_00428373: push 00411990h ; " to an increase "
  loc_00428378: call edi
  loc_0042837A: mov edx, eax
  loc_0042837C: lea ecx, var_30
  loc_0042837F: call __vbaStrMove
  loc_00428381: push eax
  loc_00428382: mov eax, [ebx+0000003Ch]
  loc_00428385: push eax
  loc_00428386: call [0040107Ch] ; __vbaStrR4
  loc_0042838C: mov edx, eax
  loc_0042838E: lea ecx, var_34
  loc_00428391: call __vbaStrMove
  loc_00428393: push eax
  loc_00428394: call edi
  loc_00428396: mov edx, eax
  loc_00428398: lea ecx, var_38
  loc_0042839B: call __vbaStrMove
  loc_0042839D: push eax
  loc_0042839E: push 0040F028h
  loc_004283A3: call edi
  loc_004283A5: mov edx, eax
  loc_004283A7: lea ecx, var_3C
  loc_004283AA: call __vbaStrMove
  loc_004283AC: mov ecx, [00430014h]
  loc_004283B2: push eax
  loc_004283B3: push ecx
  loc_004283B4: call edi
  loc_004283B6: mov edx, eax
  loc_004283B8: lea ecx, var_40
  loc_004283BB: call __vbaStrMove
  loc_004283BD: mov ecx, var_68
  loc_004283C0: mov edx, eax
  loc_004283C2: call [004010E0h] ; __vbaStrCopy
  loc_004283C8: lea edx, var_40
  loc_004283CB: lea eax, var_3C
  loc_004283CE: push edx
  loc_004283CF: lea ecx, var_38
  loc_004283D2: push eax
  loc_004283D3: lea edx, var_34
  loc_004283D6: push ecx
  loc_004283D7: lea eax, var_30
  loc_004283DA: push edx
  loc_004283DB: lea ecx, var_2C
  loc_004283DE: push eax
  loc_004283DF: lea edx, var_28
  loc_004283E2: push ecx
  loc_004283E3: lea eax, var_24
  loc_004283E6: push edx
  loc_004283E7: lea ecx, var_20
  loc_004283EA: push eax
  loc_004283EB: lea edx, var_1C
  loc_004283EE: push ecx
  loc_004283EF: lea eax, var_18
  loc_004283F2: push edx
  loc_004283F3: lea ecx, var_14
  loc_004283F6: push eax
  loc_004283F7: push ecx
  loc_004283F8: jmp 00428516h
  loc_004283FD: push 00410760h ; ", the mean value of "
  loc_00428402: call edi
  loc_00428404: mov edx, eax
  loc_00428406: lea ecx, var_14
  loc_00428409: call __vbaStrMove
  loc_0042840B: mov ecx, [00430010h]
  loc_00428411: push eax
  loc_00428412: push ecx
  loc_00428413: call edi
  loc_00428415: mov edx, eax
  loc_00428417: lea ecx, var_18
  loc_0042841A: call __vbaStrMove
  loc_0042841C: push eax
  loc_0042841D: push 004119B8h ; " decreases from "
  loc_00428422: call edi
  loc_00428424: mov edx, eax
  loc_00428426: lea ecx, var_1C
  loc_00428429: call __vbaStrMove
  loc_0042842B: fld real4 ptr [ebx+00000040h]
  loc_0042842E: fabs
  loc_00428430: push eax
  loc_00428431: push ecx
  loc_00428432: fnstsw ax
  loc_00428434: test al, 0Dh
  loc_00428436: jnz 004285CDh
  loc_0042843C: fstp real4 ptr var_70
  loc_0042843F: fld real4 ptr var_70
  loc_00428442: fstp real4 ptr [esp]
  loc_00428445: call [0040107Ch] ; __vbaStrR4
  loc_0042844B: mov edx, eax
  loc_0042844D: lea ecx, var_20
  loc_00428450: call __vbaStrMove
  loc_00428452: push eax
  loc_00428453: call edi
  loc_00428455: mov edx, eax
  loc_00428457: lea ecx, var_24
  loc_0042845A: call __vbaStrMove
  loc_0042845C: push eax
  loc_0042845D: push 0040F028h
  loc_00428462: call edi
  loc_00428464: mov edx, eax
  loc_00428466: lea ecx, var_28
  loc_00428469: call __vbaStrMove
  loc_0042846B: mov edx, [00430014h]
  loc_00428471: push eax
  loc_00428472: push edx
  loc_00428473: call edi
  loc_00428475: mov edx, eax
  loc_00428477: lea ecx, var_2C
  loc_0042847A: call __vbaStrMove
  loc_0042847C: push eax
  loc_0042847D: push 004106D8h ; " to "
  loc_00428482: call edi
  loc_00428484: mov edx, eax
  loc_00428486: lea ecx, var_30
  loc_00428489: call __vbaStrMove
  loc_0042848B: fld real4 ptr [ebx+0000003Ch]
  loc_0042848E: fabs
  loc_00428490: push eax
  loc_00428491: push ecx
  loc_00428492: fnstsw ax
  loc_00428494: test al, 0Dh
  loc_00428496: jnz 004285CDh
  loc_0042849C: fstp real4 ptr var_74
  loc_0042849F: fld real4 ptr var_74
  loc_004284A2: fstp real4 ptr [esp]
  loc_004284A5: call [0040107Ch] ; __vbaStrR4
  loc_004284AB: mov edx, eax
  loc_004284AD: lea ecx, var_34
  loc_004284B0: call __vbaStrMove
  loc_004284B2: push eax
  loc_004284B3: call edi
  loc_004284B5: mov edx, eax
  loc_004284B7: lea ecx, var_38
  loc_004284BA: call __vbaStrMove
  loc_004284BC: push eax
  loc_004284BD: push 0040F028h
  loc_004284C2: call edi
  loc_004284C4: mov edx, eax
  loc_004284C6: lea ecx, var_3C
  loc_004284C9: call __vbaStrMove
  loc_004284CB: push eax
  loc_004284CC: mov eax, [00430014h]
  loc_004284D1: push eax
  loc_004284D2: call edi
  loc_004284D4: mov edx, eax
  loc_004284D6: lea ecx, var_40
  loc_004284D9: call __vbaStrMove
  loc_004284DB: mov ecx, var_68
  loc_004284DE: mov edx, eax
  loc_004284E0: call [004010E0h] ; __vbaStrCopy
  loc_004284E6: lea ecx, var_40
  loc_004284E9: lea edx, var_3C
  loc_004284EC: push ecx
  loc_004284ED: lea eax, var_38
  loc_004284F0: push edx
  loc_004284F1: lea ecx, var_34
  loc_004284F4: push eax
  loc_004284F5: lea edx, var_30
  loc_004284F8: push ecx
  loc_004284F9: lea eax, var_2C
  loc_004284FC: push edx
  loc_004284FD: lea ecx, var_28
  loc_00428500: push eax
  loc_00428501: lea edx, var_24
  loc_00428504: push ecx
  loc_00428505: lea eax, var_20
  loc_00428508: push edx
  loc_00428509: lea ecx, var_1C
  loc_0042850C: push eax
  loc_0042850D: lea edx, var_18
  loc_00428510: push ecx
  loc_00428511: lea eax, var_14
  loc_00428514: push edx
  loc_00428515: push eax
  loc_00428516: push 0000000Ch
  loc_00428518: call [004010E4h] ; __vbaFreeStrList
  loc_0042851E: mov ecx, [ebx]
  loc_00428520: add esp, 00000034h
  loc_00428523: push ebx
  loc_00428524: call [ecx+00000358h]
  loc_0042852A: lea edx, var_44
  loc_0042852D: push eax
  loc_0042852E: push edx
  loc_0042852F: call [0040103Ch] ; __vbaObjSet
  loc_00428535: mov ecx, var_68
  loc_00428538: mov esi, eax
  loc_0042853A: mov edx, [ecx]
  loc_0042853C: mov eax, [esi]
  loc_0042853E: push edx
  loc_0042853F: push esi
  loc_00428540: call [eax+00000054h]
  loc_00428543: test eax, eax
  loc_00428545: fnclex
  loc_00428547: jge 00428558h
  loc_00428549: push 00000054h
  loc_0042854B: push 0040E390h
  loc_00428550: push esi
  loc_00428551: push eax
  loc_00428552: call [00401030h] ; __vbaHresultCheckObj
  loc_00428558: lea ecx, var_44
  loc_0042855B: call [00401114h] ; __vbaFreeObj
  loc_00428561: fwait
  loc_00428562: push 004285B8h
  loc_00428567: jmp 004285B7h
  loc_00428569: lea eax, var_40
  loc_0042856C: lea ecx, var_3C
  loc_0042856F: push eax
  loc_00428570: lea edx, var_38
  loc_00428573: push ecx
  loc_00428574: lea eax, var_34
  loc_00428577: push edx
  loc_00428578: lea ecx, var_30
  loc_0042857B: push eax
  loc_0042857C: lea edx, var_2C
  loc_0042857F: push ecx
  loc_00428580: lea eax, var_28
  loc_00428583: push edx
  loc_00428584: lea ecx, var_24
  loc_00428587: push eax
  loc_00428588: lea edx, var_20
  loc_0042858B: push ecx
  loc_0042858C: lea eax, var_1C
  loc_0042858F: push edx
  loc_00428590: lea ecx, var_18
  loc_00428593: push eax
  loc_00428594: lea edx, var_14
  loc_00428597: push ecx
  loc_00428598: push edx
  loc_00428599: push 0000000Ch
  loc_0042859B: call [004010E4h] ; __vbaFreeStrList
  loc_004285A1: add esp, 00000034h
  loc_004285A4: lea ecx, var_44
  loc_004285A7: call [00401114h] ; __vbaFreeObj
  loc_004285AD: lea ecx, var_54
  loc_004285B0: call [00401010h] ; __vbaFreeVar
  loc_004285B6: ret
  loc_004285B7: ret
  loc_004285B8: mov ecx, var_10
  loc_004285BB: pop edi
  loc_004285BC: pop esi
  loc_004285BD: xor eax, eax
  loc_004285BF: mov fs:[00000000h], ecx
  loc_004285C6: pop ebx
  loc_004285C7: mov esp, ebp
  loc_004285C9: pop ebp
  loc_004285CA: retn 0004h
End Sub

Private Function Proc_14_21_4285E0
  loc_004285E0: xor eax, eax
  loc_004285E2: retn 0004h
End Sub

Private Sub Proc_14_22_4285F0
  loc_004285F0: push ebp
  loc_004285F1: mov ebp, esp
  loc_004285F3: sub esp, 00000008h
  loc_004285F6: push 00401926h ; __vbaExceptHandler
  loc_004285FB: mov eax, fs:[00000000h]
  loc_00428601: push eax
  loc_00428602: mov fs:[00000000h], esp
  loc_00428609: sub esp, 000000D4h
  loc_0042860F: push ebx
  loc_00428610: push esi
  loc_00428611: push edi
  loc_00428612: mov var_8, esp
  loc_00428615: mov var_4, 00401688h
  loc_0042861C: mov edi, [004010E0h] ; __vbaStrCopy
  loc_00428622: xor esi, esi
  loc_00428624: mov edx, 004119E0h ; "Ho: b1 "
  loc_00428629: mov ecx, 00430040h
  loc_0042862E: mov var_14, esi
  loc_00428631: mov var_18, esi
  loc_00428634: mov var_1C, esi
  loc_00428637: mov var_20, esi
  loc_0042863A: mov var_24, esi
  loc_0042863D: mov var_28, esi
  loc_00428640: mov var_2C, esi
  loc_00428643: mov var_30, esi
  loc_00428646: mov var_34, esi
  loc_00428649: mov var_44, esi
  loc_0042864C: mov var_54, esi
  loc_0042864F: mov var_64, esi
  loc_00428652: mov var_74, esi
  loc_00428655: mov var_84, esi
  loc_0042865B: mov var_B8, esi
  loc_00428661: call edi
  loc_00428663: mov edx, 004119F4h ; "Ha: b1 "
  loc_00428668: mov ecx, 00430044h
  loc_0042866D: call edi
  loc_0042866F: mov eax, Me
  loc_00428672: push eax
  loc_00428673: mov ecx, [eax]
  loc_00428675: call [ecx+00000304h]
  loc_0042867B: lea edx, var_34
  loc_0042867E: push eax
  loc_0042867F: push edx
  loc_00428680: call [0040103Ch] ; __vbaObjSet
  loc_00428686: mov edi, eax
  loc_00428688: lea ecx, var_B8
  loc_0042868E: push ecx
  loc_0042868F: push edi
  loc_00428690: mov eax, [edi]
  loc_00428692: call [eax+000000E0h]
  loc_00428698: cmp eax, esi
  loc_0042869A: fnclex
  loc_0042869C: jge 004286B0h
  loc_0042869E: push 000000E0h
  loc_004286A3: push 00411A04h
  loc_004286A8: push edi
  loc_004286A9: push eax
  loc_004286AA: call [00401030h] ; __vbaHresultCheckObj
  loc_004286B0: xor edx, edx
  loc_004286B2: cmp var_B8, FFFFFFh
  loc_004286BA: lea ecx, var_34
  loc_004286BD: setz dl
  loc_004286C0: neg edx
  loc_004286C2: mov edi, edx
  loc_004286C4: call [00401114h] ; __vbaFreeObj
  loc_004286CA: cmp di, si
  loc_004286CD: jz 00428764h
  loc_004286D3: mov eax, [0043005Ch]
  loc_004286D8: mov ebx, [0040107Ch] ; __vbaStrR4
  loc_004286DE: push 00411A18h ; "³ "
  loc_004286E3: push eax
  loc_004286E4: call ebx
  loc_004286E6: mov esi, [004010FCh] ; __vbaStrMove
  loc_004286EC: mov edx, eax
  loc_004286EE: lea ecx, var_1C
  loc_004286F1: call __vbaStrMove
  loc_004286F3: mov edi, [0040102Ch] ; __vbaStrCat
  loc_004286F9: push eax
  loc_004286FA: call edi
  loc_004286FC: mov edx, eax
  loc_004286FE: mov ecx, 00430054h
  loc_00428703: call __vbaStrMove
  loc_00428705: lea ecx, var_1C
  loc_00428708: call [00401110h] ; __vbaFreeStr
  loc_0042870E: mov ecx, [0043005Ch]
  loc_00428714: push 00411A24h ; "< "
  loc_00428719: push ecx
  loc_0042871A: call ebx
  loc_0042871C: mov edx, eax
  loc_0042871E: lea ecx, var_1C
  loc_00428721: call __vbaStrMove
  loc_00428723: push eax
  loc_00428724: call edi
  loc_00428726: mov edx, eax
  loc_00428728: mov ecx, 00430058h
  loc_0042872D: call __vbaStrMove
  loc_0042872F: lea ecx, var_1C
  loc_00428732: call [00401110h] ; __vbaFreeStr
  loc_00428738: mov edx, [00430024h]
  loc_0042873E: push 00411A30h ; "t* < -"
  loc_00428743: push edx
  loc_00428744: call ebx
  loc_00428746: mov edx, eax
  loc_00428748: lea ecx, var_1C
  loc_0042874B: call __vbaStrMove
  loc_0042874D: push eax
  loc_0042874E: call edi
  loc_00428750: mov edx, eax
  loc_00428752: mov ecx, 0043004Ch
  loc_00428757: call __vbaStrMove
  loc_00428759: lea ecx, var_1C
  loc_0042875C: call [00401110h] ; __vbaFreeStr
  loc_00428762: jmp 00428776h
  loc_00428764: mov esi, [004010FCh] ; __vbaStrMove
  loc_0042876A: mov edi, [0040102Ch] ; __vbaStrCat
  loc_00428770: mov ebx, [0040107Ch] ; __vbaStrR4
  loc_00428776: mov eax, Me
  loc_00428779: push eax
  loc_0042877A: mov ecx, [eax]
  loc_0042877C: call [ecx+000002FCh]
  loc_00428782: lea edx, var_34
  loc_00428785: push eax
  loc_00428786: push edx
  loc_00428787: call [0040103Ch] ; __vbaObjSet
  loc_0042878D: mov ecx, [eax]
  loc_0042878F: lea edx, var_B8
  loc_00428795: push edx
  loc_00428796: push eax
  loc_00428797: mov var_BC, eax
  loc_0042879D: call [ecx+000000E0h]
  loc_004287A3: test eax, eax
  loc_004287A5: fnclex
  loc_004287A7: jge 004287C1h
  loc_004287A9: mov ecx, var_BC
  loc_004287AF: push 000000E0h
  loc_004287B4: push 00411A04h
  loc_004287B9: push ecx
  loc_004287BA: push eax
  loc_004287BB: call [00401030h] ; __vbaHresultCheckObj
  loc_004287C1: xor edx, edx
  loc_004287C3: cmp var_B8, FFFFFFh
  loc_004287CB: lea ecx, var_34
  loc_004287CE: setz dl
  loc_004287D1: neg edx
  loc_004287D3: mov var_C4, edx
  loc_004287D9: call [00401114h] ; __vbaFreeObj
  loc_004287DF: cmp var_C4, 0000h
  loc_004287E7: jz 00428866h
  loc_004287E9: mov eax, [0043005Ch]
  loc_004287EE: push 004114A0h ; "£ "
  loc_004287F3: push eax
  loc_004287F4: call ebx
  loc_004287F6: mov edx, eax
  loc_004287F8: lea ecx, var_1C
  loc_004287FB: call __vbaStrMove
  loc_004287FD: push eax
  loc_004287FE: call edi
  loc_00428800: mov edx, eax
  loc_00428802: mov ecx, 00430054h
  loc_00428807: call __vbaStrMove
  loc_00428809: lea ecx, var_1C
  loc_0042880C: call [00401110h] ; __vbaFreeStr
  loc_00428812: mov ecx, [0043005Ch]
  loc_00428818: push 00411868h ; "> "
  loc_0042881D: push ecx
  loc_0042881E: call ebx
  loc_00428820: mov edx, eax
  loc_00428822: lea ecx, var_1C
  loc_00428825: call __vbaStrMove
  loc_00428827: push eax
  loc_00428828: call edi
  loc_0042882A: mov edx, eax
  loc_0042882C: mov ecx, 00430058h
  loc_00428831: call __vbaStrMove
  loc_00428833: lea ecx, var_1C
  loc_00428836: call [00401110h] ; __vbaFreeStr
  loc_0042883C: mov edx, [00430024h]
  loc_00428842: push 00411660h ; "t* > "
  loc_00428847: push edx
  loc_00428848: call ebx
  loc_0042884A: mov edx, eax
  loc_0042884C: lea ecx, var_1C
  loc_0042884F: call __vbaStrMove
  loc_00428851: push eax
  loc_00428852: call edi
  loc_00428854: mov edx, eax
  loc_00428856: mov ecx, 0043004Ch
  loc_0042885B: call __vbaStrMove
  loc_0042885D: lea ecx, var_1C
  loc_00428860: call [00401110h] ; __vbaFreeStr
  loc_00428866: mov eax, Me
  loc_00428869: push eax
  loc_0042886A: mov ecx, [eax]
  loc_0042886C: call [ecx+00000300h]
  loc_00428872: lea edx, var_34
  loc_00428875: push eax
  loc_00428876: push edx
  loc_00428877: call [0040103Ch] ; __vbaObjSet
  loc_0042887D: mov ecx, [eax]
  loc_0042887F: lea edx, var_B8
  loc_00428885: push edx
  loc_00428886: push eax
  loc_00428887: mov var_BC, eax
  loc_0042888D: call [ecx+000000E0h]
  loc_00428893: test eax, eax
  loc_00428895: fnclex
  loc_00428897: jge 004288B1h
  loc_00428899: mov ecx, var_BC
  loc_0042889F: push 000000E0h
  loc_004288A4: push 00411A04h
  loc_004288A9: push ecx
  loc_004288AA: push eax
  loc_004288AB: call [00401030h] ; __vbaHresultCheckObj
  loc_004288B1: xor edx, edx
  loc_004288B3: cmp var_B8, FFFFFFh
  loc_004288BB: lea ecx, var_34
  loc_004288BE: setz dl
  loc_004288C1: neg edx
  loc_004288C3: mov var_C4, edx
  loc_004288C9: call [00401114h] ; __vbaFreeObj
  loc_004288CF: cmp var_C4, 0000h
  loc_004288D7: jz 00428995h
  loc_004288DD: mov eax, [0043005Ch]
  loc_004288E2: push 0041183Ch ; "= "
  loc_004288E7: push eax
  loc_004288E8: call ebx
  loc_004288EA: mov edx, eax
  loc_004288EC: lea ecx, var_1C
  loc_004288EF: call __vbaStrMove
  loc_004288F1: push eax
  loc_004288F2: call edi
  loc_004288F4: mov edx, eax
  loc_004288F6: mov ecx, 00430054h
  loc_004288FB: call __vbaStrMove
  loc_004288FD: lea ecx, var_1C
  loc_00428900: call [00401110h] ; __vbaFreeStr
  loc_00428906: mov ecx, [0043005Ch]
  loc_0042890C: push 00411848h ; "¹ "
  loc_00428911: push ecx
  loc_00428912: call ebx
  loc_00428914: mov edx, eax
  loc_00428916: lea ecx, var_1C
  loc_00428919: call __vbaStrMove
  loc_0042891B: push eax
  loc_0042891C: call edi
  loc_0042891E: mov edx, eax
  loc_00428920: mov ecx, 00430058h
  loc_00428925: call __vbaStrMove
  loc_00428927: lea ecx, var_1C
  loc_0042892A: call [00401110h] ; __vbaFreeStr
  loc_00428930: mov edx, [00430028h]
  loc_00428936: push 00411854h ; " t* > "
  loc_0042893B: push edx
  loc_0042893C: call ebx
  loc_0042893E: mov edx, eax
  loc_00428940: lea ecx, var_1C
  loc_00428943: call __vbaStrMove
  loc_00428945: push eax
  loc_00428946: call edi
  loc_00428948: mov edx, eax
  loc_0042894A: lea ecx, var_20
  loc_0042894D: call __vbaStrMove
  loc_0042894F: push eax
  loc_00428950: push 00411470h ; " or if t* < -"
  loc_00428955: call edi
  loc_00428957: mov edx, eax
  loc_00428959: lea ecx, var_24
  loc_0042895C: call __vbaStrMove
  loc_0042895E: push eax
  loc_0042895F: mov eax, [00430028h]
  loc_00428964: push eax
  loc_00428965: call ebx
  loc_00428967: mov edx, eax
  loc_00428969: lea ecx, var_28
  loc_0042896C: call __vbaStrMove
  loc_0042896E: push eax
  loc_0042896F: call edi
  loc_00428971: mov edx, eax
  loc_00428973: mov ecx, 0043004Ch
  loc_00428978: call __vbaStrMove
  loc_0042897A: lea ecx, var_28
  loc_0042897D: lea edx, var_24
  loc_00428980: push ecx
  loc_00428981: lea eax, var_20
  loc_00428984: push edx
  loc_00428985: lea ecx, var_1C
  loc_00428988: push eax
  loc_00428989: push ecx
  loc_0042898A: push 00000004h
  loc_0042898C: call [004010E4h] ; __vbaFreeStrList
  loc_00428992: add esp, 00000014h
  loc_00428995: mov eax, Me
  loc_00428998: push eax
  loc_00428999: mov edx, [eax]
  loc_0042899B: call [edx+00000350h]
  loc_004289A1: push eax
  loc_004289A2: lea eax, var_34
  loc_004289A5: push eax
  loc_004289A6: call [0040103Ch] ; __vbaObjSet
  loc_004289AC: mov ecx, [00430040h]
  loc_004289B2: mov edx, [00430054h]
  loc_004289B8: mov ebx, [eax]
  loc_004289BA: push ecx
  loc_004289BB: push edx
  loc_004289BC: mov var_BC, eax
  loc_004289C2: call edi
  loc_004289C4: mov edx, eax
  loc_004289C6: lea ecx, var_1C
  loc_004289C9: call __vbaStrMove
  loc_004289CB: mov var_D0, ebx
  loc_004289D1: mov ebx, var_BC
  loc_004289D7: push eax
  loc_004289D8: mov eax, var_D0
  loc_004289DE: push ebx
  loc_004289DF: call [eax+00000054h]
  loc_004289E2: test eax, eax
  loc_004289E4: fnclex
  loc_004289E6: jge 004289F7h
  loc_004289E8: push 00000054h
  loc_004289EA: push 0040E390h
  loc_004289EF: push ebx
  loc_004289F0: push eax
  loc_004289F1: call [00401030h] ; __vbaHresultCheckObj
  loc_004289F7: lea ecx, var_1C
  loc_004289FA: call [00401110h] ; __vbaFreeStr
  loc_00428A00: lea ecx, var_34
  loc_00428A03: call [00401114h] ; __vbaFreeObj
  loc_00428A09: mov eax, Me
  loc_00428A0C: push eax
  loc_00428A0D: mov ecx, [eax]
  loc_00428A0F: call [ecx+00000348h]
  loc_00428A15: lea edx, var_34
  loc_00428A18: push eax
  loc_00428A19: push edx
  loc_00428A1A: call [0040103Ch] ; __vbaObjSet
  loc_00428A20: mov ebx, [eax]
  loc_00428A22: mov ecx, [00430058h]
  loc_00428A28: mov var_BC, eax
  loc_00428A2E: mov eax, [00430044h]
  loc_00428A33: push eax
  loc_00428A34: push ecx
  loc_00428A35: call edi
  loc_00428A37: mov edx, eax
  loc_00428A39: lea ecx, var_1C
  loc_00428A3C: call __vbaStrMove
  loc_00428A3E: mov edx, ebx
  loc_00428A40: mov ebx, var_BC
  loc_00428A46: push eax
  loc_00428A47: push ebx
  loc_00428A48: call [edx+00000054h]
  loc_00428A4B: test eax, eax
  loc_00428A4D: fnclex
  loc_00428A4F: jge 00428A60h
  loc_00428A51: push 00000054h
  loc_00428A53: push 0040E390h
  loc_00428A58: push ebx
  loc_00428A59: push eax
  loc_00428A5A: call [00401030h] ; __vbaHresultCheckObj
  loc_00428A60: lea ecx, var_1C
  loc_00428A63: call [00401110h] ; __vbaFreeStr
  loc_00428A69: lea ecx, var_34
  loc_00428A6C: call [00401114h] ; __vbaFreeObj
  loc_00428A72: mov ebx, Me
  loc_00428A75: push ebx
  loc_00428A76: mov eax, [ebx]
  loc_00428A78: call [eax+00000314h]
  loc_00428A7E: lea ecx, var_34
  loc_00428A81: push eax
  loc_00428A82: push ecx
  loc_00428A83: call [0040103Ch] ; __vbaObjSet
  loc_00428A89: mov ecx, [0043004Ch]
  loc_00428A8F: mov edx, [eax]
  loc_00428A91: push ecx
  loc_00428A92: push eax
  loc_00428A93: mov var_BC, eax
  loc_00428A99: call [edx+00000054h]
  loc_00428A9C: test eax, eax
  loc_00428A9E: fnclex
  loc_00428AA0: jge 00428AB7h
  loc_00428AA2: mov edx, var_BC
  loc_00428AA8: push 00000054h
  loc_00428AAA: push 0040E390h
  loc_00428AAF: push edx
  loc_00428AB0: push eax
  loc_00428AB1: call [00401030h] ; __vbaHresultCheckObj
  loc_00428AB7: lea ecx, var_34
  loc_00428ABA: call [00401114h] ; __vbaFreeObj
  loc_00428AC0: fld real4 ptr [00430064h]
  loc_00428AC6: fcomp real4 ptr [004011E8h]
  loc_00428ACC: fnstsw ax
  loc_00428ACE: test ah, 40h
  loc_00428AD1: jz 00428C49h
  loc_00428AD7: mov ecx, 80020004h
  loc_00428ADC: mov eax, 0000000Ah
  loc_00428AE1: mov var_6C, ecx
  loc_00428AE4: mov var_5C, ecx
  loc_00428AE7: mov var_4C, ecx
  loc_00428AEA: lea edx, var_84
  loc_00428AF0: lea ecx, var_44
  loc_00428AF3: mov var_74, eax
  loc_00428AF6: mov var_64, eax
  loc_00428AF9: mov var_54, eax
  loc_00428AFC: mov var_7C, 00411A48h ; "Test Statistic can not be calculated for a standard error of zero."
  loc_00428B03: mov var_84, 00000008h
  loc_00428B0D: call [004010F4h] ; __vbaVarDup
  loc_00428B13: lea eax, var_74
  loc_00428B16: lea ecx, var_64
  loc_00428B19: push eax
  loc_00428B1A: lea edx, var_54
  loc_00428B1D: push ecx
  loc_00428B1E: push edx
  loc_00428B1F: lea eax, var_44
  loc_00428B22: push 00000000h
  loc_00428B24: push eax
  loc_00428B25: call [00401038h] ; rtcMsgBox
  loc_00428B2B: lea ecx, var_74
  loc_00428B2E: lea edx, var_64
  loc_00428B31: push ecx
  loc_00428B32: lea eax, var_54
  loc_00428B35: push edx
  loc_00428B36: lea ecx, var_44
  loc_00428B39: push eax
  loc_00428B3A: push ecx
  loc_00428B3B: push 00000004h
  loc_00428B3D: call [00401018h] ; __vbaFreeVarList
  loc_00428B43: mov edx, [ebx]
  loc_00428B45: add esp, 00000014h
  loc_00428B48: push ebx
  loc_00428B49: call [edx+0000032Ch]
  loc_00428B4F: mov edi, [0040103Ch] ; __vbaObjSet
  loc_00428B55: push eax
  loc_00428B56: lea eax, var_34
  loc_00428B59: push eax
  loc_00428B5A: call edi
  loc_00428B5C: mov esi, eax
  loc_00428B5E: push 0040F040h
  loc_00428B63: push esi
  loc_00428B64: mov ecx, [esi]
  loc_00428B66: call [ecx+00000054h]
  loc_00428B69: test eax, eax
  loc_00428B6B: fnclex
  loc_00428B6D: jge 00428B7Eh
  loc_00428B6F: push 00000054h
  loc_00428B71: push 0040E390h
  loc_00428B76: push esi
  loc_00428B77: push eax
  loc_00428B78: call [00401030h] ; __vbaHresultCheckObj
  loc_00428B7E: lea ecx, var_34
  loc_00428B81: call [00401114h] ; __vbaFreeObj
  loc_00428B87: mov edx, [ebx]
  loc_00428B89: push ebx
  loc_00428B8A: call [edx+00000320h]
  loc_00428B90: push eax
  loc_00428B91: lea eax, var_34
  loc_00428B94: push eax
  loc_00428B95: call edi
  loc_00428B97: mov esi, eax
  loc_00428B99: push 0040F040h
  loc_00428B9E: push esi
  loc_00428B9F: mov ecx, [esi]
  loc_00428BA1: call [ecx+00000054h]
  loc_00428BA4: test eax, eax
  loc_00428BA6: fnclex
  loc_00428BA8: jge 00428BB9h
  loc_00428BAA: push 00000054h
  loc_00428BAC: push 0040E390h
  loc_00428BB1: push esi
  loc_00428BB2: push eax
  loc_00428BB3: call [00401030h] ; __vbaHresultCheckObj
  loc_00428BB9: lea ecx, var_34
  loc_00428BBC: call [00401114h] ; __vbaFreeObj
  loc_00428BC2: mov edx, [ebx]
  loc_00428BC4: push ebx
  loc_00428BC5: call [edx+0000031Ch]
  loc_00428BCB: push eax
  loc_00428BCC: lea eax, var_34
  loc_00428BCF: push eax
  loc_00428BD0: call edi
  loc_00428BD2: mov esi, eax
  loc_00428BD4: push 0040F040h
  loc_00428BD9: push esi
  loc_00428BDA: mov ecx, [esi]
  loc_00428BDC: call [ecx+00000054h]
  loc_00428BDF: test eax, eax
  loc_00428BE1: fnclex
  loc_00428BE3: jge 00428BF4h
  loc_00428BE5: push 00000054h
  loc_00428BE7: push 0040E390h
  loc_00428BEC: push esi
  loc_00428BED: push eax
  loc_00428BEE: call [00401030h] ; __vbaHresultCheckObj
  loc_00428BF4: lea ecx, var_34
  loc_00428BF7: call [00401114h] ; __vbaFreeObj
  loc_00428BFD: mov edx, [ebx]
  loc_00428BFF: push ebx
  loc_00428C00: call [edx+0000030Ch]
  loc_00428C06: push eax
  loc_00428C07: lea eax, var_34
  loc_00428C0A: push eax
  loc_00428C0B: call edi
  loc_00428C0D: mov esi, eax
  loc_00428C0F: push 0040F040h
  loc_00428C14: push esi
  loc_00428C15: mov ecx, [esi]
  loc_00428C17: call [ecx+00000054h]
  loc_00428C1A: test eax, eax
  loc_00428C1C: fnclex
  loc_00428C1E: jge 00428C2Fh
  loc_00428C20: push 00000054h
  loc_00428C22: push 0040E390h
  loc_00428C27: push esi
  loc_00428C28: push eax
  loc_00428C29: call [00401030h] ; __vbaHresultCheckObj
  loc_00428C2F: lea ecx, var_34
  loc_00428C32: call [00401114h] ; __vbaFreeObj
  loc_00428C38: call [004010BCh] ; rtcBeep
  loc_00428C3E: fwait
  loc_00428C3F: push 00429A99h
  loc_00428C44: jmp 00429A98h
  loc_00428C49: fld real4 ptr [00430060h]
  loc_00428C4F: fsub st0, real4 ptr [0043005Ch]
  loc_00428C55: push 00000002h
  loc_00428C57: lea ecx, var_44
  loc_00428C5A: lea edx, var_18
  loc_00428C5D: mov var_84, 00004004h
  loc_00428C67: cmp [00430000h], 00000000h
  loc_00428C6E: jnz 00428C78h
  loc_00428C70: fdiv st0, real4 ptr [00430064h]
  loc_00428C76: jmp 00428C83h
  loc_00428C78: push [00430064h]
  loc_00428C7E: call 00401938h ; _adj_fdiv_m32
  loc_00428C83: mov var_7C, edx
  loc_00428C86: fstp real4 ptr var_18
  loc_00428C89: fnstsw ax
  loc_00428C8B: test al, 0Dh
  loc_00428C8D: jnz 00429AC0h
  loc_00428C93: lea eax, var_84
  loc_00428C99: push eax
  loc_00428C9A: push ecx
  loc_00428C9B: call [004010ACh] ; rtcRound
  loc_00428CA1: lea edx, var_44
  loc_00428CA4: push edx
  loc_00428CA5: call [00401084h] ; __vbaR4Var
  loc_00428CAB: fstp real4 ptr var_18
  loc_00428CAE: lea ecx, var_44
  loc_00428CB1: call [00401010h] ; __vbaFreeVar
  loc_00428CB7: fld real4 ptr [0043005Ch]
  loc_00428CBD: fcomp real4 ptr [004011E8h]
  loc_00428CC3: fnstsw ax
  loc_00428CC5: test ah, 01h
  loc_00428CC8: jnz 00428D88h
  loc_00428CCE: mov eax, [ebx]
  loc_00428CD0: push ebx
  loc_00428CD1: call [eax+0000032Ch]
  loc_00428CD7: lea ecx, var_34
  loc_00428CDA: push eax
  loc_00428CDB: push ecx
  loc_00428CDC: call [0040103Ch] ; __vbaObjSet
  loc_00428CE2: mov edx, [00430060h]
  loc_00428CE8: mov ebx, [eax]
  loc_00428CEA: push 00410694h ; "( "
  loc_00428CEF: push edx
  loc_00428CF0: mov var_BC, eax
  loc_00428CF6: call [0040107Ch] ; __vbaStrR4
  loc_00428CFC: mov edx, eax
  loc_00428CFE: lea ecx, var_1C
  loc_00428D01: call __vbaStrMove
  loc_00428D03: push eax
  loc_00428D04: call edi
  loc_00428D06: mov edx, eax
  loc_00428D08: lea ecx, var_20
  loc_00428D0B: call __vbaStrMove
  loc_00428D0D: push eax
  loc_00428D0E: push 00411AD4h ; " - "
  loc_00428D13: call edi
  loc_00428D15: mov edx, eax
  loc_00428D17: lea ecx, var_24
  loc_00428D1A: call __vbaStrMove
  loc_00428D1C: push eax
  loc_00428D1D: mov eax, [0043005Ch]
  loc_00428D22: push eax
  loc_00428D23: call [0040107Ch] ; __vbaStrR4
  loc_00428D29: mov edx, eax
  loc_00428D2B: lea ecx, var_28
  loc_00428D2E: call __vbaStrMove
  loc_00428D30: push eax
  loc_00428D31: call edi
  loc_00428D33: mov edx, eax
  loc_00428D35: lea ecx, var_2C
  loc_00428D38: call __vbaStrMove
  loc_00428D3A: push eax
  loc_00428D3B: push 004106A0h ; " )"
  loc_00428D40: call edi
  loc_00428D42: mov edx, eax
  loc_00428D44: lea ecx, var_30
  loc_00428D47: call __vbaStrMove
  loc_00428D49: mov ecx, ebx
  loc_00428D4B: mov ebx, var_BC
  loc_00428D51: push eax
  loc_00428D52: push ebx
  loc_00428D53: call [ecx+00000054h]
  loc_00428D56: test eax, eax
  loc_00428D58: fnclex
  loc_00428D5A: jge 00428D6Bh
  loc_00428D5C: push 00000054h
  loc_00428D5E: push 0040E390h
  loc_00428D63: push ebx
  loc_00428D64: push eax
  loc_00428D65: call [00401030h] ; __vbaHresultCheckObj
  loc_00428D6B: lea edx, var_30
  loc_00428D6E: lea eax, var_2C
  loc_00428D71: push edx
  loc_00428D72: lea ecx, var_28
  loc_00428D75: push eax
  loc_00428D76: lea edx, var_24
  loc_00428D79: push ecx
  loc_00428D7A: lea eax, var_20
  loc_00428D7D: push edx
  loc_00428D7E: lea ecx, var_1C
  loc_00428D81: push eax
  loc_00428D82: push ecx
  loc_00428D83: jmp 00428E48h
  loc_00428D88: mov edx, [ebx]
  loc_00428D8A: push ebx
  loc_00428D8B: call [edx+0000032Ch]
  loc_00428D91: push eax
  loc_00428D92: lea eax, var_34
  loc_00428D95: push eax
  loc_00428D96: call [0040103Ch] ; __vbaObjSet
  loc_00428D9C: mov ecx, [00430060h]
  loc_00428DA2: mov ebx, [eax]
  loc_00428DA4: push 00411AE0h ; "[ "
  loc_00428DA9: push ecx
  loc_00428DAA: mov var_BC, eax
  loc_00428DB0: call [0040107Ch] ; __vbaStrR4
  loc_00428DB6: mov edx, eax
  loc_00428DB8: lea ecx, var_1C
  loc_00428DBB: call __vbaStrMove
  loc_00428DBD: push eax
  loc_00428DBE: call edi
  loc_00428DC0: mov edx, eax
  loc_00428DC2: lea ecx, var_20
  loc_00428DC5: call __vbaStrMove
  loc_00428DC7: push eax
  loc_00428DC8: push 00411AECh ; " - ( "
  loc_00428DCD: call edi
  loc_00428DCF: mov edx, eax
  loc_00428DD1: lea ecx, var_24
  loc_00428DD4: call __vbaStrMove
  loc_00428DD6: mov edx, [0043005Ch]
  loc_00428DDC: push eax
  loc_00428DDD: push edx
  loc_00428DDE: call [0040107Ch] ; __vbaStrR4
  loc_00428DE4: mov edx, eax
  loc_00428DE6: lea ecx, var_28
  loc_00428DE9: call __vbaStrMove
  loc_00428DEB: push eax
  loc_00428DEC: call edi
  loc_00428DEE: mov edx, eax
  loc_00428DF0: lea ecx, var_2C
  loc_00428DF3: call __vbaStrMove
  loc_00428DF5: push eax
  loc_00428DF6: push 00411AFCh ; " ) ]"
  loc_00428DFB: call edi
  loc_00428DFD: mov edx, eax
  loc_00428DFF: lea ecx, var_30
  loc_00428E02: call __vbaStrMove
  loc_00428E04: mov var_DC, ebx
  loc_00428E0A: mov ebx, var_BC
  loc_00428E10: push eax
  loc_00428E11: mov eax, var_DC
  loc_00428E17: push ebx
  loc_00428E18: call [eax+00000054h]
  loc_00428E1B: test eax, eax
  loc_00428E1D: fnclex
  loc_00428E1F: jge 00428E30h
  loc_00428E21: push 00000054h
  loc_00428E23: push 0040E390h
  loc_00428E28: push ebx
  loc_00428E29: push eax
  loc_00428E2A: call [00401030h] ; __vbaHresultCheckObj
  loc_00428E30: lea ecx, var_30
  loc_00428E33: lea edx, var_2C
  loc_00428E36: push ecx
  loc_00428E37: lea eax, var_28
  loc_00428E3A: push edx
  loc_00428E3B: lea ecx, var_24
  loc_00428E3E: push eax
  loc_00428E3F: lea edx, var_20
  loc_00428E42: push ecx
  loc_00428E43: lea eax, var_1C
  loc_00428E46: push edx
  loc_00428E47: push eax
  loc_00428E48: push 00000006h
  loc_00428E4A: call [004010E4h] ; __vbaFreeStrList
  loc_00428E50: add esp, 0000001Ch
  loc_00428E53: lea ecx, var_34
  loc_00428E56: call [00401114h] ; __vbaFreeObj
  loc_00428E5C: mov eax, Me
  loc_00428E5F: push eax
  loc_00428E60: mov ecx, [eax]
  loc_00428E62: call [ecx+00000320h]
  loc_00428E68: lea edx, var_34
  loc_00428E6B: push eax
  loc_00428E6C: push edx
  loc_00428E6D: call [0040103Ch] ; __vbaObjSet
  loc_00428E73: mov ebx, [eax]
  loc_00428E75: mov var_BC, eax
  loc_00428E7B: mov eax, [00430064h]
  loc_00428E80: push 00410694h ; "( "
  loc_00428E85: push eax
  loc_00428E86: call [0040107Ch] ; __vbaStrR4
  loc_00428E8C: mov edx, eax
  loc_00428E8E: lea ecx, var_1C
  loc_00428E91: call __vbaStrMove
  loc_00428E93: push eax
  loc_00428E94: call edi
  loc_00428E96: mov edx, eax
  loc_00428E98: lea ecx, var_20
  loc_00428E9B: call __vbaStrMove
  loc_00428E9D: push eax
  loc_00428E9E: push 004106A0h ; " )"
  loc_00428EA3: call edi
  loc_00428EA5: mov edx, eax
  loc_00428EA7: lea ecx, var_24
  loc_00428EAA: call __vbaStrMove
  loc_00428EAC: mov ecx, ebx
  loc_00428EAE: mov ebx, var_BC
  loc_00428EB4: push eax
  loc_00428EB5: push ebx
  loc_00428EB6: call [ecx+00000054h]
  loc_00428EB9: test eax, eax
  loc_00428EBB: fnclex
  loc_00428EBD: jge 00428ECEh
  loc_00428EBF: push 00000054h
  loc_00428EC1: push 0040E390h
  loc_00428EC6: push ebx
  loc_00428EC7: push eax
  loc_00428EC8: call [00401030h] ; __vbaHresultCheckObj
  loc_00428ECE: mov ebx, [004010E4h] ; __vbaFreeStrList
  loc_00428ED4: lea edx, var_24
  loc_00428ED7: lea eax, var_20
  loc_00428EDA: push edx
  loc_00428EDB: lea ecx, var_1C
  loc_00428EDE: push eax
  loc_00428EDF: push ecx
  loc_00428EE0: push 00000003h
  loc_00428EE2: call ebx
  loc_00428EE4: add esp, 00000010h
  loc_00428EE7: lea ecx, var_34
  loc_00428EEA: call [00401114h] ; __vbaFreeObj
  loc_00428EF0: mov edx, [00430048h]
  loc_00428EF6: push edx
  loc_00428EF7: push 004106CCh ; " = "
  loc_00428EFC: call edi
  loc_00428EFE: mov edx, eax
  loc_00428F00: lea ecx, var_1C
  loc_00428F03: call __vbaStrMove
  loc_00428F05: push eax
  loc_00428F06: mov eax, var_18
  loc_00428F09: push eax
  loc_00428F0A: call [0040107Ch] ; __vbaStrR4
  loc_00428F10: mov edx, eax
  loc_00428F12: lea ecx, var_20
  loc_00428F15: call __vbaStrMove
  loc_00428F17: push eax
  loc_00428F18: call edi
  loc_00428F1A: mov edx, eax
  loc_00428F1C: mov ecx, 00430048h
  loc_00428F21: call __vbaStrMove
  loc_00428F23: lea ecx, var_20
  loc_00428F26: lea edx, var_1C
  loc_00428F29: push ecx
  loc_00428F2A: push edx
  loc_00428F2B: push 00000002h
  loc_00428F2D: call ebx
  loc_00428F2F: mov eax, Me
  loc_00428F32: add esp, 0000000Ch
  loc_00428F35: mov ecx, [eax]
  loc_00428F37: push eax
  loc_00428F38: call [ecx+0000031Ch]
  loc_00428F3E: lea edx, var_34
  loc_00428F41: push eax
  loc_00428F42: push edx
  loc_00428F43: call [0040103Ch] ; __vbaObjSet
  loc_00428F49: mov ebx, [eax]
  loc_00428F4B: mov var_BC, eax
  loc_00428F51: mov eax, var_18
  loc_00428F54: push 004106CCh ; " = "
  loc_00428F59: push eax
  loc_00428F5A: call [0040107Ch] ; __vbaStrR4
  loc_00428F60: mov edx, eax
  loc_00428F62: lea ecx, var_1C
  loc_00428F65: call __vbaStrMove
  loc_00428F67: push eax
  loc_00428F68: call edi
  loc_00428F6A: mov edx, eax
  loc_00428F6C: lea ecx, var_20
  loc_00428F6F: call __vbaStrMove
  loc_00428F71: mov ecx, ebx
  loc_00428F73: mov ebx, var_BC
  loc_00428F79: push eax
  loc_00428F7A: push ebx
  loc_00428F7B: call [ecx+00000054h]
  loc_00428F7E: test eax, eax
  loc_00428F80: fnclex
  loc_00428F82: jge 00428F93h
  loc_00428F84: push 00000054h
  loc_00428F86: push 0040E390h
  loc_00428F8B: push ebx
  loc_00428F8C: push eax
  loc_00428F8D: call [00401030h] ; __vbaHresultCheckObj
  loc_00428F93: lea edx, var_20
  loc_00428F96: lea eax, var_1C
  loc_00428F99: push edx
  loc_00428F9A: push eax
  loc_00428F9B: push 00000002h
  loc_00428F9D: call [004010E4h] ; __vbaFreeStrList
  loc_00428FA3: add esp, 0000000Ch
  loc_00428FA6: lea ecx, var_34
  loc_00428FA9: call [00401114h] ; __vbaFreeObj
  loc_00428FAF: mov ecx, Me
  loc_00428FB2: mov edx, [00430020h]
  loc_00428FB8: push 00411B0Ch ; "At alpha = "
  loc_00428FBD: push edx
  loc_00428FBE: lea ebx, [ecx+00000034h]
  loc_00428FC1: call [0040107Ch] ; __vbaStrR4
  loc_00428FC7: mov edx, eax
  loc_00428FC9: lea ecx, var_1C
  loc_00428FCC: call __vbaStrMove
  loc_00428FCE: push eax
  loc_00428FCF: call edi
  loc_00428FD1: mov edx, eax
  loc_00428FD3: lea ecx, var_20
  loc_00428FD6: call __vbaStrMove
  loc_00428FD8: push eax
  loc_00428FD9: push 00411B28h ; " we can "
  loc_00428FDE: call edi
  loc_00428FE0: mov edx, eax
  loc_00428FE2: lea ecx, var_24
  loc_00428FE5: call __vbaStrMove
  loc_00428FE7: mov edx, eax
  loc_00428FE9: mov ecx, ebx
  loc_00428FEB: call [004010E0h] ; __vbaStrCopy
  loc_00428FF1: lea eax, var_24
  loc_00428FF4: lea ecx, var_20
  loc_00428FF7: push eax
  loc_00428FF8: lea edx, var_1C
  loc_00428FFB: push ecx
  loc_00428FFC: push edx
  loc_00428FFD: push 00000003h
  loc_00428FFF: call [004010E4h] ; __vbaFreeStrList
  loc_00429005: mov eax, Me
  loc_00429008: add esp, 00000010h
  loc_0042900B: mov ecx, [eax]
  loc_0042900D: push eax
  loc_0042900E: call [ecx+000002FCh]
  loc_00429014: lea edx, var_34
  loc_00429017: push eax
  loc_00429018: push edx
  loc_00429019: call [0040103Ch] ; __vbaObjSet
  loc_0042901F: mov ecx, [eax]
  loc_00429021: lea edx, var_B8
  loc_00429027: push edx
  loc_00429028: push eax
  loc_00429029: mov var_BC, eax
  loc_0042902F: call [ecx+000000E0h]
  loc_00429035: test eax, eax
  loc_00429037: fnclex
  loc_00429039: jge 00429053h
  loc_0042903B: mov ecx, var_BC
  loc_00429041: push 000000E0h
  loc_00429046: push 00411A04h
  loc_0042904B: push ecx
  loc_0042904C: push eax
  loc_0042904D: call [00401030h] ; __vbaHresultCheckObj
  loc_00429053: xor edx, edx
  loc_00429055: cmp var_B8, FFFFFFh
  loc_0042905D: lea ecx, var_34
  loc_00429060: setz dl
  loc_00429063: neg edx
  loc_00429065: mov var_C4, edx
  loc_0042906B: call [00401114h] ; __vbaFreeObj
  loc_00429071: cmp var_C4, 0000h
  loc_00429079: jz 00429099h
  loc_0042907B: fld real4 ptr var_18
  loc_0042907E: fcomp real4 ptr [00430024h]
  loc_00429084: mov var_14, FFFFFFFFh
  loc_0042908B: fnstsw ax
  loc_0042908D: test ah, 41h
  loc_00429090: jz 00429099h
  loc_00429092: mov var_14, 00000000h
  loc_00429099: mov eax, Me
  loc_0042909C: push eax
  loc_0042909D: mov ecx, [eax]
  loc_0042909F: call [ecx+00000304h]
  loc_004290A5: lea edx, var_34
  loc_004290A8: push eax
  loc_004290A9: push edx
  loc_004290AA: call [0040103Ch] ; __vbaObjSet
  loc_004290B0: mov ecx, [eax]
  loc_004290B2: lea edx, var_B8
  loc_004290B8: push edx
  loc_004290B9: push eax
  loc_004290BA: mov var_BC, eax
  loc_004290C0: call [ecx+000000E0h]
  loc_004290C6: test eax, eax
  loc_004290C8: fnclex
  loc_004290CA: jge 004290E4h
  loc_004290CC: mov ecx, var_BC
  loc_004290D2: push 000000E0h
  loc_004290D7: push 00411A04h
  loc_004290DC: push ecx
  loc_004290DD: push eax
  loc_004290DE: call [00401030h] ; __vbaHresultCheckObj
  loc_004290E4: xor edx, edx
  loc_004290E6: cmp var_B8, FFFFFFh
  loc_004290EE: lea ecx, var_34
  loc_004290F1: setz dl
  loc_004290F4: neg edx
  loc_004290F6: mov var_C4, edx
  loc_004290FC: call [00401114h] ; __vbaFreeObj
  loc_00429102: cmp var_C4, 0000h
  loc_0042910A: jz 00429140h
  loc_0042910C: fld real4 ptr [00430024h]
  loc_00429112: fmul st0, real4 ptr [00401680h]
  loc_00429118: fnstsw ax
  loc_0042911A: test al, 0Dh
  loc_0042911C: jnz 00429AC0h
  loc_00429122: call [0040104Ch] ; __vbaFpR4
  loc_00429128: fcomp real4 ptr var_18
  loc_0042912B: mov var_14, FFFFFFFFh
  loc_00429132: fnstsw ax
  loc_00429134: test ah, 41h
  loc_00429137: jz 00429140h
  loc_00429139: mov var_14, 00000000h
  loc_00429140: mov eax, Me
  loc_00429143: push eax
  loc_00429144: mov ecx, [eax]
  loc_00429146: call [ecx+00000300h]
  loc_0042914C: lea edx, var_34
  loc_0042914F: push eax
  loc_00429150: push edx
  loc_00429151: call [0040103Ch] ; __vbaObjSet
  loc_00429157: mov ecx, [eax]
  loc_00429159: lea edx, var_B8
  loc_0042915F: push edx
  loc_00429160: push eax
  loc_00429161: mov var_BC, eax
  loc_00429167: call [ecx+000000E0h]
  loc_0042916D: test eax, eax
  loc_0042916F: fnclex
  loc_00429171: jge 0042918Bh
  loc_00429173: mov ecx, var_BC
  loc_00429179: push 000000E0h
  loc_0042917E: push 00411A04h
  loc_00429183: push ecx
  loc_00429184: push eax
  loc_00429185: call [00401030h] ; __vbaHresultCheckObj
  loc_0042918B: xor edx, edx
  loc_0042918D: cmp var_B8, FFFFFFh
  loc_00429195: lea ecx, var_34
  loc_00429198: setz dl
  loc_0042919B: neg edx
  loc_0042919D: mov var_C4, edx
  loc_004291A3: call [00401114h] ; __vbaFreeObj
  loc_004291A9: cmp var_C4, 0000h
  loc_004291B1: jz 004291F2h
  loc_004291B3: fld real4 ptr [00430028h]
  loc_004291B9: fmul st0, real4 ptr [00401680h]
  loc_004291BF: fnstsw ax
  loc_004291C1: test al, 0Dh
  loc_004291C3: jnz 00429AC0h
  loc_004291C9: call [0040104Ch] ; __vbaFpR4
  loc_004291CF: fcomp real4 ptr var_18
  loc_004291D2: fnstsw ax
  loc_004291D4: test ah, 41h
  loc_004291D7: jz 004291EDh
  loc_004291D9: fld real4 ptr var_18
  loc_004291DC: fcomp real4 ptr [00430028h]
  loc_004291E2: fnstsw ax
  loc_004291E4: test ah, 41h
  loc_004291E7: jz 004291EDh
  loc_004291E9: xor eax, eax
  loc_004291EB: jmp 004291F5h
  loc_004291ED: or eax, FFFFFFFFh
  loc_004291F0: jmp 004291F5h
  loc_004291F2: mov eax, var_14
  loc_004291F5: cmp ax, FFFFFFh
  loc_004291F9: jnz 00429205h
  loc_004291FB: mov eax, [ebx]
  loc_004291FD: push eax
  loc_004291FE: push 00411B40h ; "say that "
  loc_00429203: jmp 0042920Dh
  loc_00429205: mov ecx, [ebx]
  loc_00429207: push ecx
  loc_00429208: push 00411B58h ; "not say that "
  loc_0042920D: call edi
  loc_0042920F: mov edx, eax
  loc_00429211: lea ecx, var_1C
  loc_00429214: call __vbaStrMove
  loc_00429216: mov edx, eax
  loc_00429218: mov ecx, ebx
  loc_0042921A: call [004010E0h] ; __vbaStrCopy
  loc_00429220: lea ecx, var_1C
  loc_00429223: call [00401110h] ; __vbaFreeStr
  loc_00429229: fld real4 ptr [0043005Ch]
  loc_0042922F: fcomp real4 ptr [004011E8h]
  loc_00429235: fnstsw ax
  loc_00429237: test ah, 40h
  loc_0042923A: jnz 004295FBh
  loc_00429240: mov edx, [ebx]
  loc_00429242: push edx
  loc_00429243: push 00411B78h ; "for each one "
  loc_00429248: call edi
  loc_0042924A: mov edx, eax
  loc_0042924C: lea ecx, var_1C
  loc_0042924F: call __vbaStrMove
  loc_00429251: push eax
  loc_00429252: mov eax, [0043001Ch]
  loc_00429257: push eax
  loc_00429258: call edi
  loc_0042925A: mov edx, eax
  loc_0042925C: lea ecx, var_20
  loc_0042925F: call __vbaStrMove
  loc_00429261: push eax
  loc_00429262: push 00410E74h ; " increase in "
  loc_00429267: call edi
  loc_00429269: mov edx, eax
  loc_0042926B: lea ecx, var_24
  loc_0042926E: call __vbaStrMove
  loc_00429270: mov ecx, [00430018h]
  loc_00429276: push eax
  loc_00429277: push ecx
  loc_00429278: call edi
  loc_0042927A: mov edx, eax
  loc_0042927C: lea ecx, var_28
  loc_0042927F: call __vbaStrMove
  loc_00429281: mov edx, eax
  loc_00429283: mov ecx, ebx
  loc_00429285: call [004010E0h] ; __vbaStrCopy
  loc_0042928B: lea edx, var_28
  loc_0042928E: lea eax, var_24
  loc_00429291: push edx
  loc_00429292: lea ecx, var_20
  loc_00429295: push eax
  loc_00429296: lea edx, var_1C
  loc_00429299: push ecx
  loc_0042929A: push edx
  loc_0042929B: push 00000004h
  loc_0042929D: call [004010E4h] ; __vbaFreeStrList
  loc_004292A3: fld real4 ptr [0043005Ch]
  loc_004292A9: fcomp real4 ptr [004011E8h]
  loc_004292AF: add esp, 00000014h
  loc_004292B2: fnstsw ax
  loc_004292B4: test ah, 41h
  loc_004292B7: jnz 004292EFh
  loc_004292B9: mov eax, [ebx]
  loc_004292BB: push eax
  loc_004292BC: push 00411B98h ; " the increase in the mean "
  loc_004292C1: call edi
  loc_004292C3: mov edx, eax
  loc_004292C5: lea ecx, var_1C
  loc_004292C8: call __vbaStrMove
  loc_004292CA: mov ecx, [00430010h]
  loc_004292D0: push eax
  loc_004292D1: push ecx
  loc_004292D2: call edi
  loc_004292D4: mov edx, eax
  loc_004292D6: lea ecx, var_20
  loc_004292D9: call __vbaStrMove
  loc_004292DB: mov edx, eax
  loc_004292DD: mov ecx, ebx
  loc_004292DF: call [004010E0h] ; __vbaStrCopy
  loc_004292E5: lea edx, var_20
  loc_004292E8: lea eax, var_1C
  loc_004292EB: push edx
  loc_004292EC: push eax
  loc_004292ED: jmp 00429323h
  loc_004292EF: mov ecx, [ebx]
  loc_004292F1: push ecx
  loc_004292F2: push 00411BD4h ; " the decrease in the mean "
  loc_004292F7: call edi
  loc_004292F9: mov edx, eax
  loc_004292FB: lea ecx, var_1C
  loc_004292FE: call __vbaStrMove
  loc_00429300: mov edx, [00430010h]
  loc_00429306: push eax
  loc_00429307: push edx
  loc_00429308: call edi
  loc_0042930A: mov edx, eax
  loc_0042930C: lea ecx, var_20
  loc_0042930F: call __vbaStrMove
  loc_00429311: mov edx, eax
  loc_00429313: mov ecx, ebx
  loc_00429315: call [004010E0h] ; __vbaStrCopy
  loc_0042931B: lea eax, var_20
  loc_0042931E: lea ecx, var_1C
  loc_00429321: push eax
  loc_00429322: push ecx
  loc_00429323: push 00000002h
  loc_00429325: call [004010E4h] ; __vbaFreeStrList
  loc_0042932B: mov eax, Me
  loc_0042932E: add esp, 0000000Ch
  loc_00429331: mov edx, [eax]
  loc_00429333: push eax
  loc_00429334: call [edx+00000304h]
  loc_0042933A: push eax
  loc_0042933B: lea eax, var_34
  loc_0042933E: push eax
  loc_0042933F: call [0040103Ch] ; __vbaObjSet
  loc_00429345: mov ecx, [eax]
  loc_00429347: lea edx, var_B8
  loc_0042934D: push edx
  loc_0042934E: push eax
  loc_0042934F: mov var_BC, eax
  loc_00429355: call [ecx+000000E0h]
  loc_0042935B: test eax, eax
  loc_0042935D: fnclex
  loc_0042935F: jge 00429379h
  loc_00429361: mov ecx, var_BC
  loc_00429367: push 000000E0h
  loc_0042936C: push 00411A04h
  loc_00429371: push ecx
  loc_00429372: push eax
  loc_00429373: call [00401030h] ; __vbaHresultCheckObj
  loc_00429379: xor edx, edx
  loc_0042937B: cmp var_B8, FFFFFFh
  loc_00429383: lea ecx, var_34
  loc_00429386: setz dl
  loc_00429389: neg edx
  loc_0042938B: mov var_C4, edx
  loc_00429391: call [00401114h] ; __vbaFreeObj
  loc_00429397: cmp var_C4, 0000h
  loc_0042939F: jz 0042941Ah
  loc_004293A1: mov eax, [ebx]
  loc_004293A3: push eax
  loc_004293A4: push 00411670h ; " is less than "
  loc_004293A9: call edi
  loc_004293AB: mov edx, eax
  loc_004293AD: lea ecx, var_1C
  loc_004293B0: call __vbaStrMove
  loc_004293B2: mov ecx, [0043005Ch]
  loc_004293B8: push eax
  loc_004293B9: push ecx
  loc_004293BA: call [0040107Ch] ; __vbaStrR4
  loc_004293C0: mov edx, eax
  loc_004293C2: lea ecx, var_20
  loc_004293C5: call __vbaStrMove
  loc_004293C7: push eax
  loc_004293C8: call edi
  loc_004293CA: mov edx, eax
  loc_004293CC: lea ecx, var_24
  loc_004293CF: call __vbaStrMove
  loc_004293D1: push eax
  loc_004293D2: push 0040F028h
  loc_004293D7: call edi
  loc_004293D9: mov edx, eax
  loc_004293DB: lea ecx, var_28
  loc_004293DE: call __vbaStrMove
  loc_004293E0: mov edx, [00430014h]
  loc_004293E6: push eax
  loc_004293E7: push edx
  loc_004293E8: call edi
  loc_004293EA: mov edx, eax
  loc_004293EC: lea ecx, var_2C
  loc_004293EF: call __vbaStrMove
  loc_004293F1: mov edx, eax
  loc_004293F3: mov ecx, ebx
  loc_004293F5: call [004010E0h] ; __vbaStrCopy
  loc_004293FB: lea eax, var_2C
  loc_004293FE: lea ecx, var_28
  loc_00429401: push eax
  loc_00429402: lea edx, var_24
  loc_00429405: push ecx
  loc_00429406: lea eax, var_20
  loc_00429409: push edx
  loc_0042940A: lea ecx, var_1C
  loc_0042940D: push eax
  loc_0042940E: push ecx
  loc_0042940F: push 00000005h
  loc_00429411: call [004010E4h] ; __vbaFreeStrList
  loc_00429417: add esp, 00000018h
  loc_0042941A: mov eax, Me
  loc_0042941D: push eax
  loc_0042941E: mov edx, [eax]
  loc_00429420: call [edx+000002FCh]
  loc_00429426: push eax
  loc_00429427: lea eax, var_34
  loc_0042942A: push eax
  loc_0042942B: call [0040103Ch] ; __vbaObjSet
  loc_00429431: mov ecx, [eax]
  loc_00429433: lea edx, var_B8
  loc_00429439: push edx
  loc_0042943A: push eax
  loc_0042943B: mov var_BC, eax
  loc_00429441: call [ecx+000000E0h]
  loc_00429447: test eax, eax
  loc_00429449: fnclex
  loc_0042944B: jge 00429465h
  loc_0042944D: mov ecx, var_BC
  loc_00429453: push 000000E0h
  loc_00429458: push 00411A04h
  loc_0042945D: push ecx
  loc_0042945E: push eax
  loc_0042945F: call [00401030h] ; __vbaHresultCheckObj
  loc_00429465: xor edx, edx
  loc_00429467: cmp var_B8, FFFFFFh
  loc_0042946F: lea ecx, var_34
  loc_00429472: setz dl
  loc_00429475: neg edx
  loc_00429477: mov var_C4, edx
  loc_0042947D: call [00401114h] ; __vbaFreeObj
  loc_00429483: cmp var_C4, 0000h
  loc_0042948B: jz 00429506h
  loc_0042948D: mov eax, [ebx]
  loc_0042948F: push eax
  loc_00429490: push 004110D4h ; " is more than "
  loc_00429495: call edi
  loc_00429497: mov edx, eax
  loc_00429499: lea ecx, var_1C
  loc_0042949C: call __vbaStrMove
  loc_0042949E: mov ecx, [0043005Ch]
  loc_004294A4: push eax
  loc_004294A5: push ecx
  loc_004294A6: call [0040107Ch] ; __vbaStrR4
  loc_004294AC: mov edx, eax
  loc_004294AE: lea ecx, var_20
  loc_004294B1: call __vbaStrMove
  loc_004294B3: push eax
  loc_004294B4: call edi
  loc_004294B6: mov edx, eax
  loc_004294B8: lea ecx, var_24
  loc_004294BB: call __vbaStrMove
  loc_004294BD: push eax
  loc_004294BE: push 0040F028h
  loc_004294C3: call edi
  loc_004294C5: mov edx, eax
  loc_004294C7: lea ecx, var_28
  loc_004294CA: call __vbaStrMove
  loc_004294CC: mov edx, [00430014h]
  loc_004294D2: push eax
  loc_004294D3: push edx
  loc_004294D4: call edi
  loc_004294D6: mov edx, eax
  loc_004294D8: lea ecx, var_2C
  loc_004294DB: call __vbaStrMove
  loc_004294DD: mov edx, eax
  loc_004294DF: mov ecx, ebx
  loc_004294E1: call [004010E0h] ; __vbaStrCopy
  loc_004294E7: lea eax, var_2C
  loc_004294EA: lea ecx, var_28
  loc_004294ED: push eax
  loc_004294EE: lea edx, var_24
  loc_004294F1: push ecx
  loc_004294F2: lea eax, var_20
  loc_004294F5: push edx
  loc_004294F6: lea ecx, var_1C
  loc_004294F9: push eax
  loc_004294FA: push ecx
  loc_004294FB: push 00000005h
  loc_004294FD: call [004010E4h] ; __vbaFreeStrList
  loc_00429503: add esp, 00000018h
  loc_00429506: mov eax, Me
  loc_00429509: push eax
  loc_0042950A: mov edx, [eax]
  loc_0042950C: call [edx+00000300h]
  loc_00429512: push eax
  loc_00429513: lea eax, var_34
  loc_00429516: push eax
  loc_00429517: call [0040103Ch] ; __vbaObjSet
  loc_0042951D: mov ecx, [eax]
  loc_0042951F: lea edx, var_B8
  loc_00429525: push edx
  loc_00429526: push eax
  loc_00429527: mov var_BC, eax
  loc_0042952D: call [ecx+000000E0h]
  loc_00429533: test eax, eax
  loc_00429535: fnclex
  loc_00429537: jge 00429551h
  loc_00429539: mov ecx, var_BC
  loc_0042953F: push 000000E0h
  loc_00429544: push 00411A04h
  loc_00429549: push ecx
  loc_0042954A: push eax
  loc_0042954B: call [00401030h] ; __vbaHresultCheckObj
  loc_00429551: xor edx, edx
  loc_00429553: cmp var_B8, FFFFFFh
  loc_0042955B: lea ecx, var_34
  loc_0042955E: setz dl
  loc_00429561: neg edx
  loc_00429563: mov var_C4, edx
  loc_00429569: call [00401114h] ; __vbaFreeObj
  loc_0042956F: cmp var_C4, 0000h
  loc_00429577: jz 00429A08h
  loc_0042957D: mov eax, [ebx]
  loc_0042957F: push eax
  loc_00429580: push 004112A0h ; " differs from "
  loc_00429585: call edi
  loc_00429587: mov edx, eax
  loc_00429589: lea ecx, var_1C
  loc_0042958C: call __vbaStrMove
  loc_0042958E: mov ecx, [0043005Ch]
  loc_00429594: push eax
  loc_00429595: push ecx
  loc_00429596: call [0040107Ch] ; __vbaStrR4
  loc_0042959C: mov edx, eax
  loc_0042959E: lea ecx, var_20
  loc_004295A1: call __vbaStrMove
  loc_004295A3: push eax
  loc_004295A4: call edi
  loc_004295A6: mov edx, eax
  loc_004295A8: lea ecx, var_24
  loc_004295AB: call __vbaStrMove
  loc_004295AD: push eax
  loc_004295AE: push 0040F028h
  loc_004295B3: call edi
  loc_004295B5: mov edx, eax
  loc_004295B7: lea ecx, var_28
  loc_004295BA: call __vbaStrMove
  loc_004295BC: mov edx, [00430014h]
  loc_004295C2: push eax
  loc_004295C3: push edx
  loc_004295C4: call edi
  loc_004295C6: mov edx, eax
  loc_004295C8: lea ecx, var_2C
  loc_004295CB: call __vbaStrMove
  loc_004295CD: mov edx, eax
  loc_004295CF: mov ecx, ebx
  loc_004295D1: call [004010E0h] ; __vbaStrCopy
  loc_004295D7: lea eax, var_2C
  loc_004295DA: lea ecx, var_28
  loc_004295DD: push eax
  loc_004295DE: lea edx, var_24
  loc_004295E1: push ecx
  loc_004295E2: lea eax, var_20
  loc_004295E5: push edx
  loc_004295E6: lea ecx, var_1C
  loc_004295E9: push eax
  loc_004295EA: push ecx
  loc_004295EB: push 00000005h
  loc_004295ED: call [004010E4h] ; __vbaFreeStrList
  loc_004295F3: add esp, 00000018h
  loc_004295F6: jmp 00429A08h
  loc_004295FB: mov eax, Me
  loc_004295FE: push eax
  loc_004295FF: mov edx, [eax]
  loc_00429601: call [edx+00000304h]
  loc_00429607: push eax
  loc_00429608: lea eax, var_34
  loc_0042960B: push eax
  loc_0042960C: call [0040103Ch] ; __vbaObjSet
  loc_00429612: mov ecx, [eax]
  loc_00429614: lea edx, var_B8
  loc_0042961A: push edx
  loc_0042961B: push eax
  loc_0042961C: mov var_BC, eax
  loc_00429622: call [ecx+000000E0h]
  loc_00429628: test eax, eax
  loc_0042962A: fnclex
  loc_0042962C: jge 00429646h
  loc_0042962E: mov ecx, var_BC
  loc_00429634: push 000000E0h
  loc_00429639: push 00411A04h
  loc_0042963E: push ecx
  loc_0042963F: push eax
  loc_00429640: call [00401030h] ; __vbaHresultCheckObj
  loc_00429646: xor edx, edx
  loc_00429648: cmp var_B8, FFFFFFh
  loc_00429650: lea ecx, var_34
  loc_00429653: setz dl
  loc_00429656: neg edx
  loc_00429658: mov var_C4, edx
  loc_0042965E: call [00401114h] ; __vbaFreeObj
  loc_00429664: cmp var_C4, 0000h
  loc_0042966C: jz 00429700h
  loc_00429672: mov eax, [ebx]
  loc_00429674: push eax
  loc_00429675: push 00411C10h ; " there is a decrease in the mean "
  loc_0042967A: call edi
  loc_0042967C: mov edx, eax
  loc_0042967E: lea ecx, var_1C
  loc_00429681: call __vbaStrMove
  loc_00429683: mov ecx, [00430010h]
  loc_00429689: push eax
  loc_0042968A: push ecx
  loc_0042968B: call edi
  loc_0042968D: mov edx, eax
  loc_0042968F: lea ecx, var_20
  loc_00429692: call __vbaStrMove
  loc_00429694: push eax
  loc_00429695: push 00411C58h ; " for each one "
  loc_0042969A: call edi
  loc_0042969C: mov edx, eax
  loc_0042969E: lea ecx, var_24
  loc_004296A1: call __vbaStrMove
  loc_004296A3: mov edx, [0043001Ch]
  loc_004296A9: push eax
  loc_004296AA: push edx
  loc_004296AB: call edi
  loc_004296AD: mov edx, eax
  loc_004296AF: lea ecx, var_28
  loc_004296B2: call __vbaStrMove
  loc_004296B4: push eax
  loc_004296B5: push 00410E74h ; " increase in "
  loc_004296BA: call edi
  loc_004296BC: mov edx, eax
  loc_004296BE: lea ecx, var_2C
  loc_004296C1: call __vbaStrMove
  loc_004296C3: push eax
  loc_004296C4: mov eax, [00430018h]
  loc_004296C9: push eax
  loc_004296CA: call edi
  loc_004296CC: mov edx, eax
  loc_004296CE: lea ecx, var_30
  loc_004296D1: call __vbaStrMove
  loc_004296D3: mov edx, eax
  loc_004296D5: mov ecx, ebx
  loc_004296D7: call [004010E0h] ; __vbaStrCopy
  loc_004296DD: lea ecx, var_30
  loc_004296E0: lea edx, var_2C
  loc_004296E3: push ecx
  loc_004296E4: lea eax, var_28
  loc_004296E7: push edx
  loc_004296E8: lea ecx, var_24
  loc_004296EB: push eax
  loc_004296EC: lea edx, var_20
  loc_004296EF: push ecx
  loc_004296F0: lea eax, var_1C
  loc_004296F3: push edx
  loc_004296F4: push eax
  loc_004296F5: push 00000006h
  loc_004296F7: call [004010E4h] ; __vbaFreeStrList
  loc_004296FD: add esp, 0000001Ch
  loc_00429700: mov eax, Me
  loc_00429703: push eax
  loc_00429704: mov ecx, [eax]
  loc_00429706: call [ecx+000002FCh]
  loc_0042970C: lea edx, var_34
  loc_0042970F: push eax
  loc_00429710: push edx
  loc_00429711: call [0040103Ch] ; __vbaObjSet
  loc_00429717: mov ecx, [eax]
  loc_00429719: lea edx, var_B8
  loc_0042971F: push edx
  loc_00429720: push eax
  loc_00429721: mov var_BC, eax
  loc_00429727: call [ecx+000000E0h]
  loc_0042972D: test eax, eax
  loc_0042972F: fnclex
  loc_00429731: jge 0042974Bh
  loc_00429733: mov ecx, var_BC
  loc_00429739: push 000000E0h
  loc_0042973E: push 00411A04h
  loc_00429743: push ecx
  loc_00429744: push eax
  loc_00429745: call [00401030h] ; __vbaHresultCheckObj
  loc_0042974B: xor edx, edx
  loc_0042974D: cmp var_B8, FFFFFFh
  loc_00429755: lea ecx, var_34
  loc_00429758: setz dl
  loc_0042975B: neg edx
  loc_0042975D: mov var_C4, edx
  loc_00429763: call [00401114h] ; __vbaFreeObj
  loc_00429769: cmp var_C4, 0000h
  loc_00429771: jz 00429805h
  loc_00429777: mov eax, [ebx]
  loc_00429779: push eax
  loc_0042977A: push 00411C7Ch ; " there is a increase in the mean "
  loc_0042977F: call edi
  loc_00429781: mov edx, eax
  loc_00429783: lea ecx, var_1C
  loc_00429786: call __vbaStrMove
  loc_00429788: mov ecx, [00430010h]
  loc_0042978E: push eax
  loc_0042978F: push ecx
  loc_00429790: call edi
  loc_00429792: mov edx, eax
  loc_00429794: lea ecx, var_20
  loc_00429797: call __vbaStrMove
  loc_00429799: push eax
  loc_0042979A: push 00411C58h ; " for each one "
  loc_0042979F: call edi
  loc_004297A1: mov edx, eax
  loc_004297A3: lea ecx, var_24
  loc_004297A6: call __vbaStrMove
  loc_004297A8: mov edx, [0043001Ch]
  loc_004297AE: push eax
  loc_004297AF: push edx
  loc_004297B0: call edi
  loc_004297B2: mov edx, eax
  loc_004297B4: lea ecx, var_28
  loc_004297B7: call __vbaStrMove
  loc_004297B9: push eax
  loc_004297BA: push 00410E74h ; " increase in "
  loc_004297BF: call edi
  loc_004297C1: mov edx, eax
  loc_004297C3: lea ecx, var_2C
  loc_004297C6: call __vbaStrMove
  loc_004297C8: push eax
  loc_004297C9: mov eax, [00430018h]
  loc_004297CE: push eax
  loc_004297CF: call edi
  loc_004297D1: mov edx, eax
  loc_004297D3: lea ecx, var_30
  loc_004297D6: call __vbaStrMove
  loc_004297D8: mov edx, eax
  loc_004297DA: mov ecx, ebx
  loc_004297DC: call [004010E0h] ; __vbaStrCopy
  loc_004297E2: lea ecx, var_30
  loc_004297E5: lea edx, var_2C
  loc_004297E8: push ecx
  loc_004297E9: lea eax, var_28
  loc_004297EC: push edx
  loc_004297ED: lea ecx, var_24
  loc_004297F0: push eax
  loc_004297F1: lea edx, var_20
  loc_004297F4: push ecx
  loc_004297F5: lea eax, var_1C
  loc_004297F8: push edx
  loc_004297F9: push eax
  loc_004297FA: push 00000006h
  loc_004297FC: call [004010E4h] ; __vbaFreeStrList
  loc_00429802: add esp, 0000001Ch
  loc_00429805: mov eax, Me
  loc_00429808: push eax
  loc_00429809: mov ecx, [eax]
  loc_0042980B: call [ecx+00000300h]
  loc_00429811: lea edx, var_34
  loc_00429814: push eax
  loc_00429815: push edx
  loc_00429816: call [0040103Ch] ; __vbaObjSet
  loc_0042981C: mov ecx, [eax]
  loc_0042981E: lea edx, var_B8
  loc_00429824: push edx
  loc_00429825: push eax
  loc_00429826: mov var_BC, eax
  loc_0042982C: call [ecx+000000E0h]
  loc_00429832: test eax, eax
  loc_00429834: fnclex
  loc_00429836: jge 00429850h
  loc_00429838: mov ecx, var_BC
  loc_0042983E: push 000000E0h
  loc_00429843: push 00411A04h
  loc_00429848: push ecx
  loc_00429849: push eax
  loc_0042984A: call [00401030h] ; __vbaHresultCheckObj
  loc_00429850: xor edx, edx
  loc_00429852: cmp var_B8, FFFFFFh
  loc_0042985A: lea ecx, var_34
  loc_0042985D: setz dl
  loc_00429860: neg edx
  loc_00429862: mov var_C4, edx
  loc_00429868: call [00401114h] ; __vbaFreeObj
  loc_0042986E: cmp var_C4, 0000h
  loc_00429876: jz 00429A08h
  loc_0042987C: movsx eax, [00430068h]
  loc_00429883: dec eax
  loc_00429884: cmp eax, 00000003h
  loc_00429887: ja 004299CBh
  loc_0042988D: jmp [eax*4+00429AB0h]
  loc_00429894: mov eax, [ebx]
  loc_00429896: push eax
  loc_00429897: push 00411CC4h ; " changes in "
  loc_0042989C: call edi
  loc_0042989E: mov edx, eax
  loc_004298A0: lea ecx, var_1C
  loc_004298A3: call __vbaStrMove
  loc_004298A5: mov ecx, [00430018h]
  loc_004298AB: push eax
  loc_004298AC: push ecx
  loc_004298AD: call edi
  loc_004298AF: mov edx, eax
  loc_004298B1: lea ecx, var_20
  loc_004298B4: call __vbaStrMove
  loc_004298B6: push eax
  loc_004298B7: push 00411CE4h ; " is associated with changes in the mean of "
  loc_004298BC: call edi
  loc_004298BE: mov edx, eax
  loc_004298C0: lea ecx, var_24
  loc_004298C3: call __vbaStrMove
  loc_004298C5: mov edx, [00430010h]
  loc_004298CB: push eax
  loc_004298CC: push edx
  loc_004298CD: call edi
  loc_004298CF: mov edx, eax
  loc_004298D1: lea ecx, var_28
  loc_004298D4: call __vbaStrMove
  loc_004298D6: mov edx, eax
  loc_004298D8: mov ecx, ebx
  loc_004298DA: call [004010E0h] ; __vbaStrCopy
  loc_004298E0: lea eax, var_28
  loc_004298E3: lea ecx, var_24
  loc_004298E6: push eax
  loc_004298E7: lea edx, var_20
  loc_004298EA: push ecx
  loc_004298EB: lea eax, var_1C
  loc_004298EE: push edx
  loc_004298EF: push eax
  loc_004298F0: jmp 004299C0h
  loc_004298F5: mov ecx, [ebx]
  loc_004298F7: mov edx, [00430018h]
  loc_004298FD: push ecx
  loc_004298FE: push edx
  loc_004298FF: call edi
  loc_00429901: mov edx, eax
  loc_00429903: lea ecx, var_1C
  loc_00429906: call __vbaStrMove
  loc_00429908: push eax
  loc_00429909: push 00411D40h ; " helps predict "
  loc_0042990E: jmp 00429929h
  loc_00429910: mov ecx, [ebx]
  loc_00429912: mov edx, [00430018h]
  loc_00429918: push ecx
  loc_00429919: push edx
  loc_0042991A: call edi
  loc_0042991C: mov edx, eax
  loc_0042991E: lea ecx, var_1C
  loc_00429921: call __vbaStrMove
  loc_00429923: push eax
  loc_00429924: push 00411D64h ; " helps estimate the mean of "
  loc_00429929: call edi
  loc_0042992B: mov edx, eax
  loc_0042992D: lea ecx, var_20
  loc_00429930: call __vbaStrMove
  loc_00429932: push eax
  loc_00429933: mov eax, [00430010h]
  loc_00429938: push eax
  loc_00429939: call edi
  loc_0042993B: mov edx, eax
  loc_0042993D: lea ecx, var_24
  loc_00429940: call __vbaStrMove
  loc_00429942: mov edx, eax
  loc_00429944: mov ecx, ebx
  loc_00429946: call [004010E0h] ; __vbaStrCopy
  loc_0042994C: lea ecx, var_24
  loc_0042994F: lea edx, var_20
  loc_00429952: push ecx
  loc_00429953: lea eax, var_1C
  loc_00429956: push edx
  loc_00429957: push eax
  loc_00429958: push 00000003h
  loc_0042995A: call [004010E4h] ; __vbaFreeStrList
  loc_00429960: add esp, 00000010h
  loc_00429963: jmp 004299CBh
  loc_00429965: mov ecx, [ebx]
  loc_00429967: push ecx
  loc_00429968: push 00411DA4h ; " variation in "
  loc_0042996D: call edi
  loc_0042996F: mov edx, eax
  loc_00429971: lea ecx, var_1C
  loc_00429974: call __vbaStrMove
  loc_00429976: mov edx, [00430018h]
  loc_0042997C: push eax
  loc_0042997D: push edx
  loc_0042997E: call edi
  loc_00429980: mov edx, eax
  loc_00429982: lea ecx, var_20
  loc_00429985: call __vbaStrMove
  loc_00429987: push eax
  loc_00429988: push 00411DC8h ; " (though the model) is associated with variation in "
  loc_0042998D: call edi
  loc_0042998F: mov edx, eax
  loc_00429991: lea ecx, var_24
  loc_00429994: call __vbaStrMove
  loc_00429996: push eax
  loc_00429997: mov eax, [00430010h]
  loc_0042999C: push eax
  loc_0042999D: call edi
  loc_0042999F: mov edx, eax
  loc_004299A1: lea ecx, var_28
  loc_004299A4: call __vbaStrMove
  loc_004299A6: mov edx, eax
  loc_004299A8: mov ecx, ebx
  loc_004299AA: call [004010E0h] ; __vbaStrCopy
  loc_004299B0: lea ecx, var_28
  loc_004299B3: lea edx, var_24
  loc_004299B6: push ecx
  loc_004299B7: lea eax, var_20
  loc_004299BA: push edx
  loc_004299BB: lea ecx, var_1C
  loc_004299BE: push eax
  loc_004299BF: push ecx
  loc_004299C0: push 00000004h
  loc_004299C2: call [004010E4h] ; __vbaFreeStrList
  loc_004299C8: add esp, 00000014h
  loc_004299CB: mov edx, [ebx]
  loc_004299CD: push edx
  loc_004299CE: push 0040F720h ; vbCrLf
  loc_004299D3: call edi
  loc_004299D5: mov edx, eax
  loc_004299D7: lea ecx, var_1C
  loc_004299DA: call __vbaStrMove
  loc_004299DC: push eax
  loc_004299DD: push 00411E38h ; "(Click this sentence to see other possible conclusoins)"
  loc_004299E2: call edi
  loc_004299E4: mov edx, eax
  loc_004299E6: lea ecx, var_20
  loc_004299E9: call __vbaStrMove
  loc_004299EB: mov edx, eax
  loc_004299ED: mov ecx, ebx
  loc_004299EF: call [004010E0h] ; __vbaStrCopy
  loc_004299F5: lea eax, var_20
  loc_004299F8: lea ecx, var_1C
  loc_004299FB: push eax
  loc_004299FC: push ecx
  loc_004299FD: push 00000002h
  loc_004299FF: call [004010E4h] ; __vbaFreeStrList
  loc_00429A05: add esp, 0000000Ch
  loc_00429A08: mov eax, Me
  loc_00429A0B: push eax
  loc_00429A0C: mov edx, [eax]
  loc_00429A0E: call [edx+0000030Ch]
  loc_00429A14: push eax
  loc_00429A15: lea eax, var_34
  loc_00429A18: push eax
  loc_00429A19: call [0040103Ch] ; __vbaObjSet
  loc_00429A1F: mov edx, [ebx]
  loc_00429A21: mov esi, eax
  loc_00429A23: push edx
  loc_00429A24: push esi
  loc_00429A25: mov ecx, [esi]
  loc_00429A27: call [ecx+00000054h]
  loc_00429A2A: test eax, eax
  loc_00429A2C: fnclex
  loc_00429A2E: jge 00429A3Fh
  loc_00429A30: push 00000054h
  loc_00429A32: push 0040E390h
  loc_00429A37: push esi
  loc_00429A38: push eax
  loc_00429A39: call [00401030h] ; __vbaHresultCheckObj
  loc_00429A3F: lea ecx, var_34
  loc_00429A42: call [00401114h] ; __vbaFreeObj
  loc_00429A48: fwait
  loc_00429A49: push 00429A99h
  loc_00429A4E: jmp 00429A98h
  loc_00429A50: lea eax, var_30
  loc_00429A53: lea ecx, var_2C
  loc_00429A56: push eax
  loc_00429A57: lea edx, var_28
  loc_00429A5A: push ecx
  loc_00429A5B: lea eax, var_24
  loc_00429A5E: push edx
  loc_00429A5F: lea ecx, var_20
  loc_00429A62: push eax
  loc_00429A63: lea edx, var_1C
  loc_00429A66: push ecx
  loc_00429A67: push edx
  loc_00429A68: push 00000006h
  loc_00429A6A: call [004010E4h] ; __vbaFreeStrList
  loc_00429A70: add esp, 0000001Ch
  loc_00429A73: lea ecx, var_34
  loc_00429A76: call [00401114h] ; __vbaFreeObj
  loc_00429A7C: lea eax, var_74
  loc_00429A7F: lea ecx, var_64
  loc_00429A82: push eax
  loc_00429A83: lea edx, var_54
  loc_00429A86: push ecx
  loc_00429A87: lea eax, var_44
  loc_00429A8A: push edx
  loc_00429A8B: push eax
  loc_00429A8C: push 00000004h
  loc_00429A8E: call [00401018h] ; __vbaFreeVarList
  loc_00429A94: add esp, 00000014h
  loc_00429A97: ret
  loc_00429A98: ret
  loc_00429A99: mov ecx, var_10
  loc_00429A9C: pop edi
  loc_00429A9D: pop esi
  loc_00429A9E: xor eax, eax
  loc_00429AA0: mov fs:[00000000h], ecx
  loc_00429AA7: pop ebx
  loc_00429AA8: mov esp, ebp
  loc_00429AAA: pop ebp
  loc_00429AAB: retn 0004h
End Sub
