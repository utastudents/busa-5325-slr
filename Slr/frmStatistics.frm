VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "TABCTL32.OCX"
Begin VB.Form frmStatistics
  Caption = "Statistics / Point Estimates"
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 720
  ClientWidth = 8385
  ClientHeight = 4755
  StartUpPosition = 2 'CenterScreen
  Begin TabDlg.SSTab SSTab1
    Left = 120
    Top = 240
    Width = 11655
    Height = 7815
    TabIndex = 0
    OleObjectBlob = "frmStatistics.frx":0000
    Begin VB.CommandButton cmdEstPred
      Caption = "Prediction Interval"
      Index = 1
      Left = -67680
      Top = 7080
      Width = 3495
      Height = 435
      TabIndex = 27
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "Click here to go to prediction intervals."
    End
    Begin VB.CommandButton cmdEstPred
      Caption = "Estimation Interval"
      Index = 0
      Left = -67680
      Top = 7200
      Width = 3735
      Height = 435
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
      ToolTipText = "Click here to go to estimation intervals"
    End
    Begin VB.CommandButton cmdGotoSlope
      Caption = "Inferences on Slope"
      Left = -67560
      Top = 7200
      Width = 3495
      Height = 435
      TabIndex = 25
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "Click to go to inferences on the slope."
    End
    Begin VB.Frame Frame12
      Caption = "Residual"
      Left = -74400
      Top = 5160
      Width = 10695
      Height = 2295
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
      Begin VB.Label lblResidual
        Left = 120
        Top = 360
        Width = 10455
        Height = 1695
        TabIndex = 24
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
    Begin VB.Frame Frame11
      Caption = "Error:"
      Left = -74400
      Top = 2520
      Width = 10695
      Height = 2415
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
      Begin VB.Label lblError
        Left = 120
        Top = 360
        Width = 9975
        Height = 1815
        TabIndex = 22
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
      Caption = "Cofficient of Determination, r-squared"
      Left = -74520
      Top = 5160
      Width = 10815
      Height = 2055
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
      Begin VB.Label lblR2
        Left = 120
        Top = 360
        Width = 10575
        Height = 1575
        TabIndex = 20
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
    Begin VB.Frame fraS
      Caption = "Point Estimate: This is also called the mean-squared error (MSE)"
      Left = 360
      Top = 5040
      Width = 10935
      Height = 2415
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
      Begin VB.Label lblS
        Left = 240
        Top = 720
        Width = 10455
        Height = 1575
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
    End
    Begin VB.Frame fraSigma
      Caption = "Parameter: The values are assumed the same  for any value of X."
      Left = 360
      Top = 2520
      Width = 10815
      Height = 2415
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
      Begin VB.Label lblSigma
        Left = 240
        Top = 720
        Width = 10335
        Height = 1575
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
      End
    End
    Begin VB.Frame Frame7
      Caption = "Pearson's Correlation Coefficient: r"
      Left = -74520
      Top = 2760
      Width = 10815
      Height = 2055
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
      Begin VB.Label lblR
        Left = 120
        Top = 360
        Width = 10575
        Height = 1575
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
    End
    Begin VB.Frame fraYhat
      Caption = "Point Estimate: These value are also a straight line function of X."
      Left = -74760
      Top = 4920
      Width = 11055
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
      Begin VB.Label lblYhat
        Left = 360
        Top = 600
        Width = 10575
        Height = 1455
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
    Begin VB.Frame fraMuY
      Caption = "Parameter: The values fall on a straight line function of X."
      Left = -74760
      Top = 2760
      Width = 11055
      Height = 2055
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
      Begin VB.Label lblMuY
        Left = 120
        Top = 360
        Width = 10815
        Height = 1455
        TabIndex = 10
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
    Begin VB.Frame fraYhatnew
      Caption = "Point Estimate: Unlike the Y values, the estimates are on a (L.S.) line."
      Left = -74640
      Top = 4680
      Width = 11055
      Height = 2175
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
      Begin VB.Label lblYhatnew
        Left = 120
        Top = 480
        Width = 10815
        Height = 1575
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
      End
    End
    Begin VB.Frame fraY
      Caption = "Parameter: These values are normally distributed for any value of X."
      Left = -74640
      Top = 2640
      Width = 11055
      Height = 1935
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
      Begin VB.Label lblY
        Left = 240
        Top = 480
        Width = 10695
        Height = 1335
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
      Caption = "Point Estimate"
      Left = -74640
      Top = 4680
      Width = 10695
      Height = 2295
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
      Begin VB.Label lblBetaHat
        Left = 120
        Top = 480
        Width = 10455
        Height = 1575
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
      Caption = "Parameter"
      Left = -74640
      Top = 2160
      Width = 10695
      Height = 2295
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
      Begin VB.Label lblbeta
        Left = 120
        Top = 360
        Width = 10335
        Height = 1695
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
  Begin VB.Menu mnuInstructions
    Caption = "&Instructions"
  End
  Begin VB.Menu mnuChange
    Caption = "&Change"
    Begin VB.Menu mnuChangeNames
      Caption = "&Variable Names"
    End
    Begin VB.Menu mnuChangeXg
      Caption = "The Given Value of &X: Xg"
    End
    Begin VB.Menu mnuChangeLSLine
      Caption = "Slope and Intercept Estimates"
    End
    Begin VB.Menu mnuChangeCorr
      Caption = "&R and/or  R-Squared"
    End
    Begin VB.Menu mnuChangeUnits
      Caption = "&Units"
    End
  End
  Begin VB.Menu mnuGoto
    Caption = "&Go To"
    Begin VB.Menu mnuIntro
      Caption = "Intro&duction"
    End
    Begin VB.Menu mnuQuestions
      Caption = "&Questions"
    End
    Begin VB.Menu mnuInferences
      Caption = "I&nferences"
      Begin VB.Menu mnuSlope
        Caption = "&Slope"
      End
      Begin VB.Menu mnuEstPred
        Caption = "Estimation and &Prediction"
      End
    End
    Begin VB.Menu mnuAssumptions
      Caption = "&Assumptions"
    End
  End
  Begin VB.Menu mnuExit
    Caption = "&Exit"
  End
End

Attribute VB_Name = "frmStatistics"


Private Sub cmdGotoSlope_Click() '42AF90
  loc_0042AF90: push ebp
  loc_0042AF91: mov ebp, esp
  loc_0042AF93: sub esp, 0000000Ch
  loc_0042AF96: push 00401926h ; __vbaExceptHandler
  loc_0042AF9B: mov eax, fs:[00000000h]
  loc_0042AFA1: push eax
  loc_0042AFA2: mov fs:[00000000h], esp
  loc_0042AFA9: sub esp, 00000030h
  loc_0042AFAC: push ebx
  loc_0042AFAD: push esi
  loc_0042AFAE: push edi
  loc_0042AFAF: mov var_C, esp
  loc_0042AFB2: mov var_8, 00401768h
  loc_0042AFB9: mov eax, Me
  loc_0042AFBC: mov ecx, eax
  loc_0042AFBE: and ecx, 00000001h
  loc_0042AFC1: mov var_4, ecx
  loc_0042AFC4: and al, FEh
  loc_0042AFC6: push eax
  loc_0042AFC7: mov Me, eax
  loc_0042AFCA: mov edx, [eax]
  loc_0042AFCC: call [edx+00000004h]
  loc_0042AFCF: mov eax, [004300C4h]
  loc_0042AFD4: test eax, eax
  loc_0042AFD6: jnz 0042AFE8h
  loc_0042AFD8: push 004300C4h
  loc_0042AFDD: push 00409228h
  loc_0042AFE2: call [004010D4h] ; __vbaNew2
  loc_0042AFE8: sub esp, 00000010h
  loc_0042AFEB: mov ecx, 0000000Ah
  loc_0042AFF0: mov ebx, esp
  loc_0042AFF2: mov var_24, ecx
  loc_0042AFF5: mov eax, 80020004h
  loc_0042AFFA: sub esp, 00000010h
  loc_0042AFFD: mov [ebx], ecx
  loc_0042AFFF: mov ecx, var_30
  loc_0042B002: mov edx, eax
  loc_0042B004: mov esi, [004300C4h]
  loc_0042B00A: mov [ebx+00000004h], ecx
  loc_0042B00D: mov ecx, esp
  loc_0042B00F: mov edi, [esi]
  loc_0042B011: push esi
  loc_0042B012: mov [ebx+00000008h], eax
  loc_0042B015: mov eax, var_28
  loc_0042B018: mov [ebx+0000000Ch], eax
  loc_0042B01B: mov eax, var_24
  loc_0042B01E: mov [ecx], eax
  loc_0042B020: mov eax, var_20
  loc_0042B023: mov [ecx+00000004h], eax
  loc_0042B026: mov [ecx+00000008h], edx
  loc_0042B029: mov edx, var_18
  loc_0042B02C: mov [ecx+0000000Ch], edx
  loc_0042B02F: call [edi+000002B0h]
  loc_0042B035: test eax, eax
  loc_0042B037: fnclex
  loc_0042B039: jge 0042B04Dh
  loc_0042B03B: push 000002B0h
  loc_0042B040: push 0040E0ECh
  loc_0042B045: push esi
  loc_0042B046: push eax
  loc_0042B047: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B04D: mov eax, [004300ECh]
  loc_0042B052: test eax, eax
  loc_0042B054: jnz 0042B066h
  loc_0042B056: push 004300ECh
  loc_0042B05B: push 0040A624h
  loc_0042B060: call [004010D4h] ; __vbaNew2
  loc_0042B066: mov esi, [004300ECh]
  loc_0042B06C: push esi
  loc_0042B06D: mov eax, [esi]
  loc_0042B06F: call [eax+000002B4h]
  loc_0042B075: test eax, eax
  loc_0042B077: fnclex
  loc_0042B079: jge 0042B08Dh
  loc_0042B07B: push 000002B4h
  loc_0042B080: push 0040ECECh
  loc_0042B085: push esi
  loc_0042B086: push eax
  loc_0042B087: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B08D: mov var_4, 00000000h
  loc_0042B094: mov eax, Me
  loc_0042B097: push eax
  loc_0042B098: mov ecx, [eax]
  loc_0042B09A: call [ecx+00000008h]
  loc_0042B09D: mov eax, var_4
  loc_0042B0A0: mov ecx, var_14
  loc_0042B0A3: pop edi
  loc_0042B0A4: pop esi
  loc_0042B0A5: mov fs:[00000000h], ecx
  loc_0042B0AC: pop ebx
  loc_0042B0AD: mov esp, ebp
  loc_0042B0AF: pop ebp
  loc_0042B0B0: retn 0004h
End Sub

Private Sub mnuChangeCorr_Click() '42CB70
  loc_0042CB70: push ebp
  loc_0042CB71: mov ebp, esp
  loc_0042CB73: sub esp, 0000000Ch
  loc_0042CB76: push 00401926h ; __vbaExceptHandler
  loc_0042CB7B: mov eax, fs:[00000000h]
  loc_0042CB81: push eax
  loc_0042CB82: mov fs:[00000000h], esp
  loc_0042CB89: sub esp, 00000030h
  loc_0042CB8C: push ebx
  loc_0042CB8D: push esi
  loc_0042CB8E: push edi
  loc_0042CB8F: mov var_C, esp
  loc_0042CB92: mov var_8, 004017A0h
  loc_0042CB99: mov eax, Me
  loc_0042CB9C: mov ecx, eax
  loc_0042CB9E: and ecx, 00000001h
  loc_0042CBA1: mov var_4, ecx
  loc_0042CBA4: and al, FEh
  loc_0042CBA6: push eax
  loc_0042CBA7: mov Me, eax
  loc_0042CBAA: mov edx, [eax]
  loc_0042CBAC: call [edx+00000004h]
  loc_0042CBAF: mov eax, [00430114h]
  loc_0042CBB4: test eax, eax
  loc_0042CBB6: jnz 0042CBC8h
  loc_0042CBB8: push 00430114h
  loc_0042CBBD: push 00404514h
  loc_0042CBC2: call [004010D4h] ; __vbaNew2
  loc_0042CBC8: sub esp, 00000010h
  loc_0042CBCB: mov ecx, 0000000Ah
  loc_0042CBD0: mov ebx, esp
  loc_0042CBD2: mov var_24, ecx
  loc_0042CBD5: mov eax, 80020004h
  loc_0042CBDA: sub esp, 00000010h
  loc_0042CBDD: mov [ebx], ecx
  loc_0042CBDF: mov ecx, var_30
  loc_0042CBE2: mov edx, eax
  loc_0042CBE4: mov esi, [00430114h]
  loc_0042CBEA: mov [ebx+00000004h], ecx
  loc_0042CBED: mov ecx, esp
  loc_0042CBEF: mov edi, [esi]
  loc_0042CBF1: push esi
  loc_0042CBF2: mov [ebx+00000008h], eax
  loc_0042CBF5: mov eax, var_28
  loc_0042CBF8: mov [ebx+0000000Ch], eax
  loc_0042CBFB: mov eax, var_24
  loc_0042CBFE: mov [ecx], eax
  loc_0042CC00: mov eax, var_20
  loc_0042CC03: mov [ecx+00000004h], eax
  loc_0042CC06: mov [ecx+00000008h], edx
  loc_0042CC09: mov edx, var_18
  loc_0042CC0C: mov [ecx+0000000Ch], edx
  loc_0042CC0F: call [edi+000002B0h]
  loc_0042CC15: test eax, eax
  loc_0042CC17: fnclex
  loc_0042CC19: jge 0042CC2Dh
  loc_0042CC1B: push 000002B0h
  loc_0042CC20: push 0040F348h
  loc_0042CC25: push esi
  loc_0042CC26: push eax
  loc_0042CC27: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CC2D: mov var_4, 00000000h
  loc_0042CC34: mov eax, Me
  loc_0042CC37: push eax
  loc_0042CC38: mov ecx, [eax]
  loc_0042CC3A: call [ecx+00000008h]
  loc_0042CC3D: mov eax, var_4
  loc_0042CC40: mov ecx, var_14
  loc_0042CC43: pop edi
  loc_0042CC44: pop esi
  loc_0042CC45: mov fs:[00000000h], ecx
  loc_0042CC4C: pop ebx
  loc_0042CC4D: mov esp, ebp
  loc_0042CC4F: pop ebp
  loc_0042CC50: retn 0004h
End Sub

Private Sub cmdEstPred_Click(Index As Integer) '42ADF0
  loc_0042ADF0: push ebp
  loc_0042ADF1: mov ebp, esp
  loc_0042ADF3: sub esp, 0000000Ch
  loc_0042ADF6: push 00401926h ; __vbaExceptHandler
  loc_0042ADFB: mov eax, fs:[00000000h]
  loc_0042AE01: push eax
  loc_0042AE02: mov fs:[00000000h], esp
  loc_0042AE09: sub esp, 00000034h
  loc_0042AE0C: push ebx
  loc_0042AE0D: push esi
  loc_0042AE0E: push edi
  loc_0042AE0F: mov var_C, esp
  loc_0042AE12: mov var_8, 00401758h
  loc_0042AE19: mov eax, Me
  loc_0042AE1C: mov ecx, eax
  loc_0042AE1E: and ecx, 00000001h
  loc_0042AE21: mov var_4, ecx
  loc_0042AE24: and al, FEh
  loc_0042AE26: push eax
  loc_0042AE27: mov Me, eax
  loc_0042AE2A: mov edx, [eax]
  loc_0042AE2C: call [edx+00000004h]
  loc_0042AE2F: mov eax, [004300B0h]
  loc_0042AE34: mov var_18, 00000000h
  loc_0042AE3B: test eax, eax
  loc_0042AE3D: jnz 0042AE4Fh
  loc_0042AE3F: push 004300B0h
  loc_0042AE44: push 00407528h
  loc_0042AE49: call [004010D4h] ; __vbaNew2
  loc_0042AE4F: sub esp, 00000010h
  loc_0042AE52: mov ecx, 0000000Ah
  loc_0042AE57: mov ebx, esp
  loc_0042AE59: mov var_28, ecx
  loc_0042AE5C: mov eax, 80020004h
  loc_0042AE61: sub esp, 00000010h
  loc_0042AE64: mov [ebx], ecx
  loc_0042AE66: mov ecx, var_34
  loc_0042AE69: mov edx, eax
  loc_0042AE6B: mov esi, [004300B0h]
  loc_0042AE71: mov [ebx+00000004h], ecx
  loc_0042AE74: mov ecx, esp
  loc_0042AE76: mov edi, [esi]
  loc_0042AE78: push esi
  loc_0042AE79: mov [ebx+00000008h], eax
  loc_0042AE7C: mov eax, var_2C
  loc_0042AE7F: mov [ebx+0000000Ch], eax
  loc_0042AE82: mov eax, var_28
  loc_0042AE85: mov ebx, var_24
  loc_0042AE88: mov [ecx], eax
  loc_0042AE8A: mov [ecx+00000004h], ebx
  loc_0042AE8D: mov [ecx+00000008h], edx
  loc_0042AE90: mov edx, var_1C
  loc_0042AE93: mov [ecx+0000000Ch], edx
  loc_0042AE96: call [edi+000002B0h]
  loc_0042AE9C: test eax, eax
  loc_0042AE9E: fnclex
  loc_0042AEA0: jge 0042AEB4h
  loc_0042AEA2: push 000002B0h
  loc_0042AEA7: push 0040DE98h
  loc_0042AEAC: push esi
  loc_0042AEAD: push eax
  loc_0042AEAE: call [00401030h] ; __vbaHresultCheckObj
  loc_0042AEB4: mov eax, [004300ECh]
  loc_0042AEB9: test eax, eax
  loc_0042AEBB: jnz 0042AECDh
  loc_0042AEBD: push 004300ECh
  loc_0042AEC2: push 0040A624h
  loc_0042AEC7: call [004010D4h] ; __vbaNew2
  loc_0042AECD: mov esi, [004300ECh]
  loc_0042AED3: push esi
  loc_0042AED4: mov eax, [esi]
  loc_0042AED6: call [eax+000002B4h]
  loc_0042AEDC: test eax, eax
  loc_0042AEDE: fnclex
  loc_0042AEE0: jge 0042AEF4h
  loc_0042AEE2: push 000002B4h
  loc_0042AEE7: push 0040ECECh
  loc_0042AEEC: push esi
  loc_0042AEED: push eax
  loc_0042AEEE: call [00401030h] ; __vbaHresultCheckObj
  loc_0042AEF4: mov eax, [004300B0h]
  loc_0042AEF9: mov esi, Index
  loc_0042AEFC: test eax, eax
  loc_0042AEFE: mov edi, 00004002h
  loc_0042AF03: jnz 0042AF1Ah
  loc_0042AF05: push 004300B0h
  loc_0042AF0A: push 00407528h
  loc_0042AF0F: call [004010D4h] ; __vbaNew2
  loc_0042AF15: mov eax, [004300B0h]
  loc_0042AF1A: sub esp, 00000010h
  loc_0042AF1D: mov edx, var_1C
  loc_0042AF20: mov ecx, esp
  loc_0042AF22: push 00000004h
  loc_0042AF24: push eax
  loc_0042AF25: mov [ecx], edi
  loc_0042AF27: mov [ecx+00000004h], ebx
  loc_0042AF2A: mov [ecx+00000008h], esi
  loc_0042AF2D: mov [ecx+0000000Ch], edx
  loc_0042AF30: mov ecx, [eax]
  loc_0042AF32: call [ecx+0000035Ch]
  loc_0042AF38: lea edx, var_18
  loc_0042AF3B: push eax
  loc_0042AF3C: push edx
  loc_0042AF3D: call [0040103Ch] ; __vbaObjSet
  loc_0042AF43: push eax
  loc_0042AF44: call [00401104h] ; __vbaLateIdSt
  loc_0042AF4A: lea ecx, var_18
  loc_0042AF4D: call [00401114h] ; __vbaFreeObj
  loc_0042AF53: mov var_4, 00000000h
  loc_0042AF5A: push 0042AF6Ch
  loc_0042AF5F: jmp 0042AF6Bh
  loc_0042AF61: lea ecx, var_18
  loc_0042AF64: call [00401114h] ; __vbaFreeObj
  loc_0042AF6A: ret
  loc_0042AF6B: ret
  loc_0042AF6C: mov eax, Me
  loc_0042AF6F: push eax
  loc_0042AF70: mov ecx, [eax]
  loc_0042AF72: call [ecx+00000008h]
  loc_0042AF75: mov eax, var_4
  loc_0042AF78: mov ecx, var_14
  loc_0042AF7B: pop edi
  loc_0042AF7C: pop esi
  loc_0042AF7D: mov fs:[00000000h], ecx
  loc_0042AF84: pop ebx
  loc_0042AF85: mov esp, ebp
  loc_0042AF87: pop ebp
  loc_0042AF88: retn 0008h
End Sub

Private Sub mnuChangeLSLine_Click() '42CC60
  loc_0042CC60: push ebp
  loc_0042CC61: mov ebp, esp
  loc_0042CC63: sub esp, 0000000Ch
  loc_0042CC66: push 00401926h ; __vbaExceptHandler
  loc_0042CC6B: mov eax, fs:[00000000h]
  loc_0042CC71: push eax
  loc_0042CC72: mov fs:[00000000h], esp
  loc_0042CC79: sub esp, 00000030h
  loc_0042CC7C: push ebx
  loc_0042CC7D: push esi
  loc_0042CC7E: push edi
  loc_0042CC7F: mov var_C, esp
  loc_0042CC82: mov var_8, 004017A8h
  loc_0042CC89: mov eax, Me
  loc_0042CC8C: mov ecx, eax
  loc_0042CC8E: and ecx, 00000001h
  loc_0042CC91: mov var_4, ecx
  loc_0042CC94: and al, FEh
  loc_0042CC96: push eax
  loc_0042CC97: mov Me, eax
  loc_0042CC9A: mov edx, [eax]
  loc_0042CC9C: call [edx+00000004h]
  loc_0042CC9F: mov eax, [00430150h]
  loc_0042CCA4: test eax, eax
  loc_0042CCA6: jnz 0042CCB8h
  loc_0042CCA8: push 00430150h
  loc_0042CCAD: push 00403F48h
  loc_0042CCB2: call [004010D4h] ; __vbaNew2
  loc_0042CCB8: sub esp, 00000010h
  loc_0042CCBB: mov ecx, 0000000Ah
  loc_0042CCC0: mov ebx, esp
  loc_0042CCC2: mov var_24, ecx
  loc_0042CCC5: mov eax, 80020004h
  loc_0042CCCA: sub esp, 00000010h
  loc_0042CCCD: mov [ebx], ecx
  loc_0042CCCF: mov ecx, var_30
  loc_0042CCD2: mov edx, eax
  loc_0042CCD4: mov esi, [00430150h]
  loc_0042CCDA: mov [ebx+00000004h], ecx
  loc_0042CCDD: mov ecx, esp
  loc_0042CCDF: mov edi, [esi]
  loc_0042CCE1: push esi
  loc_0042CCE2: mov [ebx+00000008h], eax
  loc_0042CCE5: mov eax, var_28
  loc_0042CCE8: mov [ebx+0000000Ch], eax
  loc_0042CCEB: mov eax, var_24
  loc_0042CCEE: mov [ecx], eax
  loc_0042CCF0: mov eax, var_20
  loc_0042CCF3: mov [ecx+00000004h], eax
  loc_0042CCF6: mov [ecx+00000008h], edx
  loc_0042CCF9: mov edx, var_18
  loc_0042CCFC: mov [ecx+0000000Ch], edx
  loc_0042CCFF: call [edi+000002B0h]
  loc_0042CD05: test eax, eax
  loc_0042CD07: fnclex
  loc_0042CD09: jge 0042CD1Dh
  loc_0042CD0B: push 000002B0h
  loc_0042CD10: push 0040FD4Ch
  loc_0042CD15: push esi
  loc_0042CD16: push eax
  loc_0042CD17: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CD1D: mov var_4, 00000000h
  loc_0042CD24: mov eax, Me
  loc_0042CD27: push eax
  loc_0042CD28: mov ecx, [eax]
  loc_0042CD2A: call [ecx+00000008h]
  loc_0042CD2D: mov eax, var_4
  loc_0042CD30: mov ecx, var_14
  loc_0042CD33: pop edi
  loc_0042CD34: pop esi
  loc_0042CD35: mov fs:[00000000h], ecx
  loc_0042CD3C: pop ebx
  loc_0042CD3D: mov esp, ebp
  loc_0042CD3F: pop ebp
  loc_0042CD40: retn 0004h
End Sub

Private Sub mnuInstructions_Click() '42D790
  loc_0042D790: push ebp
  loc_0042D791: mov ebp, esp
  loc_0042D793: sub esp, 0000000Ch
  loc_0042D796: push 00401926h ; __vbaExceptHandler
  loc_0042D79B: mov eax, fs:[00000000h]
  loc_0042D7A1: push eax
  loc_0042D7A2: mov fs:[00000000h], esp
  loc_0042D7A9: sub esp, 00000088h
  loc_0042D7AF: push ebx
  loc_0042D7B0: push esi
  loc_0042D7B1: push edi
  loc_0042D7B2: mov var_C, esp
  loc_0042D7B5: mov var_8, 00401800h
  loc_0042D7BC: mov eax, Me
  loc_0042D7BF: mov ecx, eax
  loc_0042D7C1: and ecx, 00000001h
  loc_0042D7C4: mov var_4, ecx
  loc_0042D7C7: and al, FEh
  loc_0042D7C9: push eax
  loc_0042D7CA: mov Me, eax
  loc_0042D7CD: mov edx, [eax]
  loc_0042D7CF: call [edx+00000004h]
  loc_0042D7D2: mov edi, [004010F4h] ; __vbaVarDup
  loc_0042D7D8: mov ecx, 80020004h
  loc_0042D7DD: xor esi, esi
  loc_0042D7DF: mov var_4C, ecx
  loc_0042D7E2: mov eax, 0000000Ah
  loc_0042D7E7: mov var_3C, ecx
  loc_0042D7EA: mov ebx, 00000008h
  loc_0042D7EF: mov var_44, esi
  loc_0042D7F2: mov var_54, esi
  loc_0042D7F5: mov var_74, esi
  loc_0042D7F8: lea edx, var_74
  loc_0042D7FB: lea ecx, var_34
  loc_0042D7FE: mov var_24, esi
  loc_0042D801: mov var_34, esi
  loc_0042D804: mov var_64, esi
  loc_0042D807: mov var_54, eax
  loc_0042D80A: mov var_44, eax
  loc_0042D80D: mov var_6C, 0040E0D0h ; "Instructions"
  loc_0042D814: mov var_74, ebx
  loc_0042D817: call edi
  loc_0042D819: lea edx, var_64
  loc_0042D81C: lea ecx, var_24
  loc_0042D81F: mov var_5C, 0040E034h ; "Hold mouse over statements to see the statement in context of the example."
  loc_0042D826: mov var_64, ebx
  loc_0042D829: call edi
  loc_0042D82B: lea eax, var_54
  loc_0042D82E: lea ecx, var_44
  loc_0042D831: push eax
  loc_0042D832: lea edx, var_34
  loc_0042D835: push ecx
  loc_0042D836: push edx
  loc_0042D837: lea eax, var_24
  loc_0042D83A: push esi
  loc_0042D83B: push eax
  loc_0042D83C: call [00401038h] ; rtcMsgBox
  loc_0042D842: lea ecx, var_54
  loc_0042D845: lea edx, var_44
  loc_0042D848: push ecx
  loc_0042D849: lea eax, var_34
  loc_0042D84C: push edx
  loc_0042D84D: lea ecx, var_24
  loc_0042D850: push eax
  loc_0042D851: push ecx
  loc_0042D852: push 00000004h
  loc_0042D854: call [00401018h] ; __vbaFreeVarList
  loc_0042D85A: add esp, 00000014h
  loc_0042D85D: mov var_4, esi
  loc_0042D860: push 0042D884h
  loc_0042D865: jmp 0042D883h
  loc_0042D867: lea edx, var_54
  loc_0042D86A: lea eax, var_44
  loc_0042D86D: push edx
  loc_0042D86E: lea ecx, var_34
  loc_0042D871: push eax
  loc_0042D872: lea edx, var_24
  loc_0042D875: push ecx
  loc_0042D876: push edx
  loc_0042D877: push 00000004h
  loc_0042D879: call [00401018h] ; __vbaFreeVarList
  loc_0042D87F: add esp, 00000014h
  loc_0042D882: ret
  loc_0042D883: ret
  loc_0042D884: mov eax, Me
  loc_0042D887: push eax
  loc_0042D888: mov ecx, [eax]
  loc_0042D88A: call [ecx+00000008h]
  loc_0042D88D: mov eax, var_4
  loc_0042D890: mov ecx, var_14
  loc_0042D893: pop edi
  loc_0042D894: pop esi
  loc_0042D895: mov fs:[00000000h], ecx
  loc_0042D89C: pop ebx
  loc_0042D89D: mov esp, ebp
  loc_0042D89F: pop ebp
  loc_0042D8A0: retn 0004h
End Sub

Private Sub mnuExit_Click() '42D1F0
  loc_0042D1F0: push ebp
  loc_0042D1F1: mov ebp, esp
  loc_0042D1F3: sub esp, 0000000Ch
  loc_0042D1F6: push 00401926h ; __vbaExceptHandler
  loc_0042D1FB: mov eax, fs:[00000000h]
  loc_0042D201: push eax
  loc_0042D202: mov fs:[00000000h], esp
  loc_0042D209: sub esp, 00000018h
  loc_0042D20C: push ebx
  loc_0042D20D: push esi
  loc_0042D20E: push edi
  loc_0042D20F: mov var_C, esp
  loc_0042D212: mov var_8, 004017D0h
  loc_0042D219: mov edi, Me
  loc_0042D21C: mov eax, edi
  loc_0042D21E: and eax, 00000001h
  loc_0042D221: mov var_4, eax
  loc_0042D224: and edi, FFFFFFFEh
  loc_0042D227: push edi
  loc_0042D228: mov Me, edi
  loc_0042D22B: mov ecx, [edi]
  loc_0042D22D: call [ecx+00000004h]
  loc_0042D230: mov eax, [00430A24h]
  loc_0042D235: xor ebx, ebx
  loc_0042D237: cmp eax, ebx
  loc_0042D239: mov var_18, ebx
  loc_0042D23C: jnz 0042D24Eh
  loc_0042D23E: push 00430A24h
  loc_0042D243: push 0040EC7Ch
  loc_0042D248: call [004010D4h] ; __vbaNew2
  loc_0042D24E: mov esi, [00430A24h]
  loc_0042D254: lea eax, var_18
  loc_0042D257: push edi
  loc_0042D258: push eax
  loc_0042D259: mov edx, [esi]
  loc_0042D25B: mov var_2C, edx
  loc_0042D25E: call [00401044h] ; __vbaObjSetAddref
  loc_0042D264: mov ecx, var_2C
  loc_0042D267: push eax
  loc_0042D268: push esi
  loc_0042D269: call [ecx+00000010h]
  loc_0042D26C: cmp eax, ebx
  loc_0042D26E: fnclex
  loc_0042D270: jge 0042D281h
  loc_0042D272: push 00000010h
  loc_0042D274: push 0040EC6Ch
  loc_0042D279: push esi
  loc_0042D27A: push eax
  loc_0042D27B: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D281: lea ecx, var_18
  loc_0042D284: call [00401114h] ; __vbaFreeObj
  loc_0042D28A: mov var_4, ebx
  loc_0042D28D: push 0042D29Fh
  loc_0042D292: jmp 0042D29Eh
  loc_0042D294: lea ecx, var_18
  loc_0042D297: call [00401114h] ; __vbaFreeObj
  loc_0042D29D: ret
  loc_0042D29E: ret
  loc_0042D29F: mov eax, Me
  loc_0042D2A2: push eax
  loc_0042D2A3: mov edx, [eax]
  loc_0042D2A5: call [edx+00000008h]
  loc_0042D2A8: mov eax, var_4
  loc_0042D2AB: mov ecx, var_14
  loc_0042D2AE: pop edi
  loc_0042D2AF: pop esi
  loc_0042D2B0: mov fs:[00000000h], ecx
  loc_0042D2B7: pop ebx
  loc_0042D2B8: mov esp, ebp
  loc_0042D2BA: pop ebp
  loc_0042D2BB: retn 0004h
End Sub

Private Sub mnuEstPred_Click() '42D0C0
  loc_0042D0C0: push ebp
  loc_0042D0C1: mov ebp, esp
  loc_0042D0C3: sub esp, 0000000Ch
  loc_0042D0C6: push 00401926h ; __vbaExceptHandler
  loc_0042D0CB: mov eax, fs:[00000000h]
  loc_0042D0D1: push eax
  loc_0042D0D2: mov fs:[00000000h], esp
  loc_0042D0D9: sub esp, 00000030h
  loc_0042D0DC: push ebx
  loc_0042D0DD: push esi
  loc_0042D0DE: push edi
  loc_0042D0DF: mov var_C, esp
  loc_0042D0E2: mov var_8, 004017C8h
  loc_0042D0E9: mov eax, Me
  loc_0042D0EC: mov ecx, eax
  loc_0042D0EE: and ecx, 00000001h
  loc_0042D0F1: mov var_4, ecx
  loc_0042D0F4: and al, FEh
  loc_0042D0F6: push eax
  loc_0042D0F7: mov Me, eax
  loc_0042D0FA: mov edx, [eax]
  loc_0042D0FC: call [edx+00000004h]
  loc_0042D0FF: mov eax, [004300B0h]
  loc_0042D104: test eax, eax
  loc_0042D106: jnz 0042D118h
  loc_0042D108: push 004300B0h
  loc_0042D10D: push 00407528h
  loc_0042D112: call [004010D4h] ; __vbaNew2
  loc_0042D118: sub esp, 00000010h
  loc_0042D11B: mov ecx, 0000000Ah
  loc_0042D120: mov ebx, esp
  loc_0042D122: mov var_24, ecx
  loc_0042D125: mov eax, 80020004h
  loc_0042D12A: sub esp, 00000010h
  loc_0042D12D: mov [ebx], ecx
  loc_0042D12F: mov ecx, var_30
  loc_0042D132: mov edx, eax
  loc_0042D134: mov esi, [004300B0h]
  loc_0042D13A: mov [ebx+00000004h], ecx
  loc_0042D13D: mov ecx, esp
  loc_0042D13F: mov edi, [esi]
  loc_0042D141: push esi
  loc_0042D142: mov [ebx+00000008h], eax
  loc_0042D145: mov eax, var_28
  loc_0042D148: mov [ebx+0000000Ch], eax
  loc_0042D14B: mov eax, var_24
  loc_0042D14E: mov [ecx], eax
  loc_0042D150: mov eax, var_20
  loc_0042D153: mov [ecx+00000004h], eax
  loc_0042D156: mov [ecx+00000008h], edx
  loc_0042D159: mov edx, var_18
  loc_0042D15C: mov [ecx+0000000Ch], edx
  loc_0042D15F: call [edi+000002B0h]
  loc_0042D165: test eax, eax
  loc_0042D167: fnclex
  loc_0042D169: jge 0042D17Dh
  loc_0042D16B: push 000002B0h
  loc_0042D170: push 0040DE98h
  loc_0042D175: push esi
  loc_0042D176: push eax
  loc_0042D177: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D17D: mov eax, [004300ECh]
  loc_0042D182: test eax, eax
  loc_0042D184: jnz 0042D196h
  loc_0042D186: push 004300ECh
  loc_0042D18B: push 0040A624h
  loc_0042D190: call [004010D4h] ; __vbaNew2
  loc_0042D196: mov esi, [004300ECh]
  loc_0042D19C: push esi
  loc_0042D19D: mov eax, [esi]
  loc_0042D19F: call [eax+000002B4h]
  loc_0042D1A5: test eax, eax
  loc_0042D1A7: fnclex
  loc_0042D1A9: jge 0042D1BDh
  loc_0042D1AB: push 000002B4h
  loc_0042D1B0: push 0040ECECh
  loc_0042D1B5: push esi
  loc_0042D1B6: push eax
  loc_0042D1B7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D1BD: mov var_4, 00000000h
  loc_0042D1C4: mov eax, Me
  loc_0042D1C7: push eax
  loc_0042D1C8: mov ecx, [eax]
  loc_0042D1CA: call [ecx+00000008h]
  loc_0042D1CD: mov eax, var_4
  loc_0042D1D0: mov ecx, var_14
  loc_0042D1D3: pop edi
  loc_0042D1D4: pop esi
  loc_0042D1D5: mov fs:[00000000h], ecx
  loc_0042D1DC: pop ebx
  loc_0042D1DD: mov esp, ebp
  loc_0042D1DF: pop ebp
  loc_0042D1E0: retn 0004h
End Sub

Private Sub mnuQuestions_Click() '42D660
  loc_0042D660: push ebp
  loc_0042D661: mov ebp, esp
  loc_0042D663: sub esp, 0000000Ch
  loc_0042D666: push 00401926h ; __vbaExceptHandler
  loc_0042D66B: mov eax, fs:[00000000h]
  loc_0042D671: push eax
  loc_0042D672: mov fs:[00000000h], esp
  loc_0042D679: sub esp, 00000030h
  loc_0042D67C: push ebx
  loc_0042D67D: push esi
  loc_0042D67E: push edi
  loc_0042D67F: mov var_C, esp
  loc_0042D682: mov var_8, 004017F8h
  loc_0042D689: mov eax, Me
  loc_0042D68C: mov ecx, eax
  loc_0042D68E: and ecx, 00000001h
  loc_0042D691: mov var_4, ecx
  loc_0042D694: and al, FEh
  loc_0042D696: push eax
  loc_0042D697: mov Me, eax
  loc_0042D69A: mov edx, [eax]
  loc_0042D69C: call [edx+00000004h]
  loc_0042D69F: mov eax, [00430164h]
  loc_0042D6A4: test eax, eax
  loc_0042D6A6: jnz 0042D6B8h
  loc_0042D6A8: push 00430164h
  loc_0042D6AD: push 00408108h
  loc_0042D6B2: call [004010D4h] ; __vbaNew2
  loc_0042D6B8: sub esp, 00000010h
  loc_0042D6BB: mov ecx, 0000000Ah
  loc_0042D6C0: mov ebx, esp
  loc_0042D6C2: mov var_24, ecx
  loc_0042D6C5: mov eax, 80020004h
  loc_0042D6CA: sub esp, 00000010h
  loc_0042D6CD: mov [ebx], ecx
  loc_0042D6CF: mov ecx, var_30
  loc_0042D6D2: mov edx, eax
  loc_0042D6D4: mov esi, [00430164h]
  loc_0042D6DA: mov [ebx+00000004h], ecx
  loc_0042D6DD: mov ecx, esp
  loc_0042D6DF: mov edi, [esi]
  loc_0042D6E1: push esi
  loc_0042D6E2: mov [ebx+00000008h], eax
  loc_0042D6E5: mov eax, var_28
  loc_0042D6E8: mov [ebx+0000000Ch], eax
  loc_0042D6EB: mov eax, var_24
  loc_0042D6EE: mov [ecx], eax
  loc_0042D6F0: mov eax, var_20
  loc_0042D6F3: mov [ecx+00000004h], eax
  loc_0042D6F6: mov [ecx+00000008h], edx
  loc_0042D6F9: mov edx, var_18
  loc_0042D6FC: mov [ecx+0000000Ch], edx
  loc_0042D6FF: call [edi+000002B0h]
  loc_0042D705: test eax, eax
  loc_0042D707: fnclex
  loc_0042D709: jge 0042D71Dh
  loc_0042D70B: push 000002B0h
  loc_0042D710: push 0040FF58h
  loc_0042D715: push esi
  loc_0042D716: push eax
  loc_0042D717: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D71D: mov eax, [004300ECh]
  loc_0042D722: test eax, eax
  loc_0042D724: jnz 0042D736h
  loc_0042D726: push 004300ECh
  loc_0042D72B: push 0040A624h
  loc_0042D730: call [004010D4h] ; __vbaNew2
  loc_0042D736: mov esi, [004300ECh]
  loc_0042D73C: push esi
  loc_0042D73D: mov eax, [esi]
  loc_0042D73F: call [eax+000002B4h]
  loc_0042D745: test eax, eax
  loc_0042D747: fnclex
  loc_0042D749: jge 0042D75Dh
  loc_0042D74B: push 000002B4h
  loc_0042D750: push 0040ECECh
  loc_0042D755: push esi
  loc_0042D756: push eax
  loc_0042D757: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D75D: mov var_4, 00000000h
  loc_0042D764: mov eax, Me
  loc_0042D767: push eax
  loc_0042D768: mov ecx, [eax]
  loc_0042D76A: call [ecx+00000008h]
  loc_0042D76D: mov eax, var_4
  loc_0042D770: mov ecx, var_14
  loc_0042D773: pop edi
  loc_0042D774: pop esi
  loc_0042D775: mov fs:[00000000h], ecx
  loc_0042D77C: pop ebx
  loc_0042D77D: mov esp, ebp
  loc_0042D77F: pop ebp
  loc_0042D780: retn 0004h
End Sub

Private Sub mnuSlope_Click() '42D8B0
  loc_0042D8B0: push ebp
  loc_0042D8B1: mov ebp, esp
  loc_0042D8B3: sub esp, 0000000Ch
  loc_0042D8B6: push 00401926h ; __vbaExceptHandler
  loc_0042D8BB: mov eax, fs:[00000000h]
  loc_0042D8C1: push eax
  loc_0042D8C2: mov fs:[00000000h], esp
  loc_0042D8C9: sub esp, 00000030h
  loc_0042D8CC: push ebx
  loc_0042D8CD: push esi
  loc_0042D8CE: push edi
  loc_0042D8CF: mov var_C, esp
  loc_0042D8D2: mov var_8, 00401810h
  loc_0042D8D9: mov eax, Me
  loc_0042D8DC: mov ecx, eax
  loc_0042D8DE: and ecx, 00000001h
  loc_0042D8E1: mov var_4, ecx
  loc_0042D8E4: and al, FEh
  loc_0042D8E6: push eax
  loc_0042D8E7: mov Me, eax
  loc_0042D8EA: mov edx, [eax]
  loc_0042D8EC: call [edx+00000004h]
  loc_0042D8EF: mov eax, [004300C4h]
  loc_0042D8F4: test eax, eax
  loc_0042D8F6: jnz 0042D908h
  loc_0042D8F8: push 004300C4h
  loc_0042D8FD: push 00409228h
  loc_0042D902: call [004010D4h] ; __vbaNew2
  loc_0042D908: sub esp, 00000010h
  loc_0042D90B: mov ecx, 0000000Ah
  loc_0042D910: mov ebx, esp
  loc_0042D912: mov var_24, ecx
  loc_0042D915: mov eax, 80020004h
  loc_0042D91A: sub esp, 00000010h
  loc_0042D91D: mov [ebx], ecx
  loc_0042D91F: mov ecx, var_30
  loc_0042D922: mov edx, eax
  loc_0042D924: mov esi, [004300C4h]
  loc_0042D92A: mov [ebx+00000004h], ecx
  loc_0042D92D: mov ecx, esp
  loc_0042D92F: mov edi, [esi]
  loc_0042D931: push esi
  loc_0042D932: mov [ebx+00000008h], eax
  loc_0042D935: mov eax, var_28
  loc_0042D938: mov [ebx+0000000Ch], eax
  loc_0042D93B: mov eax, var_24
  loc_0042D93E: mov [ecx], eax
  loc_0042D940: mov eax, var_20
  loc_0042D943: mov [ecx+00000004h], eax
  loc_0042D946: mov [ecx+00000008h], edx
  loc_0042D949: mov edx, var_18
  loc_0042D94C: mov [ecx+0000000Ch], edx
  loc_0042D94F: call [edi+000002B0h]
  loc_0042D955: test eax, eax
  loc_0042D957: fnclex
  loc_0042D959: jge 0042D96Dh
  loc_0042D95B: push 000002B0h
  loc_0042D960: push 0040E0ECh
  loc_0042D965: push esi
  loc_0042D966: push eax
  loc_0042D967: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D96D: mov eax, [004300ECh]
  loc_0042D972: test eax, eax
  loc_0042D974: jnz 0042D986h
  loc_0042D976: push 004300ECh
  loc_0042D97B: push 0040A624h
  loc_0042D980: call [004010D4h] ; __vbaNew2
  loc_0042D986: mov esi, [004300ECh]
  loc_0042D98C: push esi
  loc_0042D98D: mov eax, [esi]
  loc_0042D98F: call [eax+000002B4h]
  loc_0042D995: test eax, eax
  loc_0042D997: fnclex
  loc_0042D999: jge 0042D9ADh
  loc_0042D99B: push 000002B4h
  loc_0042D9A0: push 0040ECECh
  loc_0042D9A5: push esi
  loc_0042D9A6: push eax
  loc_0042D9A7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D9AD: mov var_4, 00000000h
  loc_0042D9B4: mov eax, Me
  loc_0042D9B7: push eax
  loc_0042D9B8: mov ecx, [eax]
  loc_0042D9BA: call [ecx+00000008h]
  loc_0042D9BD: mov eax, var_4
  loc_0042D9C0: mov ecx, var_14
  loc_0042D9C3: pop edi
  loc_0042D9C4: pop esi
  loc_0042D9C5: mov fs:[00000000h], ecx
  loc_0042D9CC: pop ebx
  loc_0042D9CD: mov esp, ebp
  loc_0042D9CF: pop ebp
  loc_0042D9D0: retn 0004h
End Sub

Private Sub mnuChangeUnits_Click() '42CEE0
  loc_0042CEE0: push ebp
  loc_0042CEE1: mov ebp, esp
  loc_0042CEE3: sub esp, 0000000Ch
  loc_0042CEE6: push 00401926h ; __vbaExceptHandler
  loc_0042CEEB: mov eax, fs:[00000000h]
  loc_0042CEF1: push eax
  loc_0042CEF2: mov fs:[00000000h], esp
  loc_0042CEF9: sub esp, 00000030h
  loc_0042CEFC: push ebx
  loc_0042CEFD: push esi
  loc_0042CEFE: push edi
  loc_0042CEFF: mov var_C, esp
  loc_0042CF02: mov var_8, 004017B8h
  loc_0042CF09: mov eax, Me
  loc_0042CF0C: mov ecx, eax
  loc_0042CF0E: and ecx, 00000001h
  loc_0042CF11: mov var_4, ecx
  loc_0042CF14: and al, FEh
  loc_0042CF16: push eax
  loc_0042CF17: mov Me, eax
  loc_0042CF1A: mov edx, [eax]
  loc_0042CF1C: call [edx+00000004h]
  loc_0042CF1F: mov eax, [004301A0h]
  loc_0042CF24: test eax, eax
  loc_0042CF26: jnz 0042CF38h
  loc_0042CF28: push 004301A0h
  loc_0042CF2D: push 004033C0h
  loc_0042CF32: call [004010D4h] ; __vbaNew2
  loc_0042CF38: sub esp, 00000010h
  loc_0042CF3B: mov ecx, 0000000Ah
  loc_0042CF40: mov ebx, esp
  loc_0042CF42: mov var_24, ecx
  loc_0042CF45: mov eax, 80020004h
  loc_0042CF4A: sub esp, 00000010h
  loc_0042CF4D: mov [ebx], ecx
  loc_0042CF4F: mov ecx, var_30
  loc_0042CF52: mov edx, eax
  loc_0042CF54: mov esi, [004301A0h]
  loc_0042CF5A: mov [ebx+00000004h], ecx
  loc_0042CF5D: mov ecx, esp
  loc_0042CF5F: mov edi, [esi]
  loc_0042CF61: push esi
  loc_0042CF62: mov [ebx+00000008h], eax
  loc_0042CF65: mov eax, var_28
  loc_0042CF68: mov [ebx+0000000Ch], eax
  loc_0042CF6B: mov eax, var_24
  loc_0042CF6E: mov [ecx], eax
  loc_0042CF70: mov eax, var_20
  loc_0042CF73: mov [ecx+00000004h], eax
  loc_0042CF76: mov [ecx+00000008h], edx
  loc_0042CF79: mov edx, var_18
  loc_0042CF7C: mov [ecx+0000000Ch], edx
  loc_0042CF7F: call [edi+000002B0h]
  loc_0042CF85: test eax, eax
  loc_0042CF87: fnclex
  loc_0042CF89: jge 0042CF9Dh
  loc_0042CF8B: push 000002B0h
  loc_0042CF90: push 00410454h
  loc_0042CF95: push esi
  loc_0042CF96: push eax
  loc_0042CF97: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CF9D: mov var_4, 00000000h
  loc_0042CFA4: mov eax, Me
  loc_0042CFA7: push eax
  loc_0042CFA8: mov ecx, [eax]
  loc_0042CFAA: call [ecx+00000008h]
  loc_0042CFAD: mov eax, var_4
  loc_0042CFB0: mov ecx, var_14
  loc_0042CFB3: pop edi
  loc_0042CFB4: pop esi
  loc_0042CFB5: mov fs:[00000000h], ecx
  loc_0042CFBC: pop ebx
  loc_0042CFBD: mov esp, ebp
  loc_0042CFBF: pop ebp
  loc_0042CFC0: retn 0004h
End Sub

Private Sub mnuIntro_Click() '42D530
  loc_0042D530: push ebp
  loc_0042D531: mov ebp, esp
  loc_0042D533: sub esp, 0000000Ch
  loc_0042D536: push 00401926h ; __vbaExceptHandler
  loc_0042D53B: mov eax, fs:[00000000h]
  loc_0042D541: push eax
  loc_0042D542: mov fs:[00000000h], esp
  loc_0042D549: sub esp, 00000030h
  loc_0042D54C: push ebx
  loc_0042D54D: push esi
  loc_0042D54E: push edi
  loc_0042D54F: mov var_C, esp
  loc_0042D552: mov var_8, 004017F0h
  loc_0042D559: mov eax, Me
  loc_0042D55C: mov ecx, eax
  loc_0042D55E: and ecx, 00000001h
  loc_0042D561: mov var_4, ecx
  loc_0042D564: and al, FEh
  loc_0042D566: push eax
  loc_0042D567: mov Me, eax
  loc_0042D56A: mov edx, [eax]
  loc_0042D56C: call [edx+00000004h]
  loc_0042D56F: mov eax, [00430088h]
  loc_0042D574: test eax, eax
  loc_0042D576: jnz 0042D588h
  loc_0042D578: push 00430088h
  loc_0042D57D: push 004058C0h
  loc_0042D582: call [004010D4h] ; __vbaNew2
  loc_0042D588: sub esp, 00000010h
  loc_0042D58B: mov ecx, 0000000Ah
  loc_0042D590: mov ebx, esp
  loc_0042D592: mov var_24, ecx
  loc_0042D595: mov eax, 80020004h
  loc_0042D59A: sub esp, 00000010h
  loc_0042D59D: mov [ebx], ecx
  loc_0042D59F: mov ecx, var_30
  loc_0042D5A2: mov edx, eax
  loc_0042D5A4: mov esi, [00430088h]
  loc_0042D5AA: mov [ebx+00000004h], ecx
  loc_0042D5AD: mov ecx, esp
  loc_0042D5AF: mov edi, [esi]
  loc_0042D5B1: push esi
  loc_0042D5B2: mov [ebx+00000008h], eax
  loc_0042D5B5: mov eax, var_28
  loc_0042D5B8: mov [ebx+0000000Ch], eax
  loc_0042D5BB: mov eax, var_24
  loc_0042D5BE: mov [ecx], eax
  loc_0042D5C0: mov eax, var_20
  loc_0042D5C3: mov [ecx+00000004h], eax
  loc_0042D5C6: mov [ecx+00000008h], edx
  loc_0042D5C9: mov edx, var_18
  loc_0042D5CC: mov [ecx+0000000Ch], edx
  loc_0042D5CF: call [edi+000002B0h]
  loc_0042D5D5: test eax, eax
  loc_0042D5D7: fnclex
  loc_0042D5D9: jge 0042D5EDh
  loc_0042D5DB: push 000002B0h
  loc_0042D5E0: push 0040DB64h
  loc_0042D5E5: push esi
  loc_0042D5E6: push eax
  loc_0042D5E7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D5ED: mov eax, [004300ECh]
  loc_0042D5F2: test eax, eax
  loc_0042D5F4: jnz 0042D606h
  loc_0042D5F6: push 004300ECh
  loc_0042D5FB: push 0040A624h
  loc_0042D600: call [004010D4h] ; __vbaNew2
  loc_0042D606: mov esi, [004300ECh]
  loc_0042D60C: push esi
  loc_0042D60D: mov eax, [esi]
  loc_0042D60F: call [eax+000002B4h]
  loc_0042D615: test eax, eax
  loc_0042D617: fnclex
  loc_0042D619: jge 0042D62Dh
  loc_0042D61B: push 000002B4h
  loc_0042D620: push 0040ECECh
  loc_0042D625: push esi
  loc_0042D626: push eax
  loc_0042D627: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D62D: mov var_4, 00000000h
  loc_0042D634: mov eax, Me
  loc_0042D637: push eax
  loc_0042D638: mov ecx, [eax]
  loc_0042D63A: call [ecx+00000008h]
  loc_0042D63D: mov eax, var_4
  loc_0042D640: mov ecx, var_14
  loc_0042D643: pop edi
  loc_0042D644: pop esi
  loc_0042D645: mov fs:[00000000h], ecx
  loc_0042D64C: pop ebx
  loc_0042D64D: mov esp, ebp
  loc_0042D64F: pop ebp
  loc_0042D650: retn 0004h
End Sub

Private Sub Form_Load() '42B130
  loc_0042B130: push ebp
  loc_0042B131: mov ebp, esp
  loc_0042B133: sub esp, 0000000Ch
  loc_0042B136: push 00401926h ; __vbaExceptHandler
  loc_0042B13B: mov eax, fs:[00000000h]
  loc_0042B141: push eax
  loc_0042B142: mov fs:[00000000h], esp
  loc_0042B149: sub esp, 00000008h
  loc_0042B14C: push ebx
  loc_0042B14D: push esi
  loc_0042B14E: push edi
  loc_0042B14F: mov var_C, esp
  loc_0042B152: mov var_8, 00401778h
  loc_0042B159: mov esi, Me
  loc_0042B15C: mov eax, esi
  loc_0042B15E: and eax, 00000001h
  loc_0042B161: mov var_4, eax
  loc_0042B164: and esi, FFFFFFFEh
  loc_0042B167: push esi
  loc_0042B168: mov Me, esi
  loc_0042B16B: mov ecx, [esi]
  loc_0042B16D: call [ecx+00000004h]
  loc_0042B170: mov edx, [esi]
  loc_0042B172: push esi
  loc_0042B173: call [edx+00000708h]
  loc_0042B179: mov var_4, 00000000h
  loc_0042B180: mov eax, Me
  loc_0042B183: push eax
  loc_0042B184: mov ecx, [eax]
  loc_0042B186: call [ecx+00000008h]
  loc_0042B189: mov eax, var_4
  loc_0042B18C: mov ecx, var_14
  loc_0042B18F: pop edi
  loc_0042B190: pop esi
  loc_0042B191: mov fs:[00000000h], ecx
  loc_0042B198: pop ebx
  loc_0042B199: mov esp, ebp
  loc_0042B19B: pop ebp
  loc_0042B19C: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '42D2C0
  loc_0042D2C0: push ebp
  loc_0042D2C1: mov ebp, esp
  loc_0042D2C3: sub esp, 0000000Ch
  loc_0042D2C6: push 00401926h ; __vbaExceptHandler
  loc_0042D2CB: mov eax, fs:[00000000h]
  loc_0042D2D1: push eax
  loc_0042D2D2: mov fs:[00000000h], esp
  loc_0042D2D9: sub esp, 0000009Ch
  loc_0042D2DF: push ebx
  loc_0042D2E0: push esi
  loc_0042D2E1: push edi
  loc_0042D2E2: mov var_C, esp
  loc_0042D2E5: mov var_8, 004017E0h
  loc_0042D2EC: mov edi, Me
  loc_0042D2EF: mov eax, edi
  loc_0042D2F1: and eax, 00000001h
  loc_0042D2F4: mov var_4, eax
  loc_0042D2F7: and edi, FFFFFFFEh
  loc_0042D2FA: push edi
  loc_0042D2FB: mov Me, edi
  loc_0042D2FE: mov ecx, [edi]
  loc_0042D300: call [ecx+00000004h]
  loc_0042D303: mov ebx, [004010F4h] ; __vbaVarDup
  loc_0042D309: mov ecx, 80020004h
  loc_0042D30E: xor esi, esi
  loc_0042D310: mov var_54, ecx
  loc_0042D313: mov eax, 0000000Ah
  loc_0042D318: mov var_44, ecx
  loc_0042D31B: mov var_4C, esi
  loc_0042D31E: mov var_5C, esi
  loc_0042D321: mov var_7C, esi
  loc_0042D324: lea edx, var_7C
  loc_0042D327: lea ecx, var_3C
  loc_0042D32A: mov var_1C, esi
  loc_0042D32D: mov var_2C, esi
  loc_0042D330: mov var_3C, esi
  loc_0042D333: mov var_6C, esi
  loc_0042D336: mov var_5C, eax
  loc_0042D339: mov var_4C, eax
  loc_0042D33C: mov var_74, 0040ECD4h ; "Exit Check"
  loc_0042D343: mov var_7C, 00000008h
  loc_0042D34A: call ebx
  loc_0042D34C: lea edx, var_6C
  loc_0042D34F: lea ecx, var_2C
  loc_0042D352: mov var_64, 0040EC90h ; "Are you sure you want to exit?"
  loc_0042D359: mov var_6C, 00000008h
  loc_0042D360: call ebx
  loc_0042D362: lea edx, var_5C
  loc_0042D365: lea eax, var_4C
  loc_0042D368: push edx
  loc_0042D369: lea ecx, var_3C
  loc_0042D36C: push eax
  loc_0042D36D: push ecx
  loc_0042D36E: lea edx, var_2C
  loc_0042D371: push 00000004h
  loc_0042D373: push edx
  loc_0042D374: call [00401038h] ; rtcMsgBox
  loc_0042D37A: mov ecx, eax
  loc_0042D37C: call [00401078h] ; __vbaI2I4
  loc_0042D382: mov ebx, eax
  loc_0042D384: lea eax, var_5C
  loc_0042D387: lea ecx, var_4C
  loc_0042D38A: push eax
  loc_0042D38B: lea edx, var_3C
  loc_0042D38E: push ecx
  loc_0042D38F: lea eax, var_2C
  loc_0042D392: push edx
  loc_0042D393: push eax
  loc_0042D394: push 00000004h
  loc_0042D396: call [00401018h] ; __vbaFreeVarList
  loc_0042D39C: add esp, 00000014h
  loc_0042D39F: cmp bx, 0007h
  loc_0042D3A3: jnz 0042D3AFh
  loc_0042D3A5: mov ecx, Cancel
  loc_0042D3A8: mov [ecx], FFFFFFh
  loc_0042D3AD: jmp 0042D40Fh
  loc_0042D3AF: cmp [00430A24h], esi
  loc_0042D3B5: jnz 0042D3C7h
  loc_0042D3B7: push 00430A24h
  loc_0042D3BC: push 0040EC7Ch
  loc_0042D3C1: call [004010D4h] ; __vbaNew2
  loc_0042D3C7: mov ebx, [00430A24h]
  loc_0042D3CD: lea eax, var_1C
  loc_0042D3D0: push edi
  loc_0042D3D1: push eax
  loc_0042D3D2: mov edx, [ebx]
  loc_0042D3D4: mov var_B0, edx
  loc_0042D3DA: call [00401044h] ; __vbaObjSetAddref
  loc_0042D3E0: mov ecx, var_B0
  loc_0042D3E6: push eax
  loc_0042D3E7: push ebx
  loc_0042D3E8: call [ecx+00000010h]
  loc_0042D3EB: cmp eax, esi
  loc_0042D3ED: fnclex
  loc_0042D3EF: jge 0042D400h
  loc_0042D3F1: push 00000010h
  loc_0042D3F3: push 0040EC6Ch
  loc_0042D3F8: push ebx
  loc_0042D3F9: push eax
  loc_0042D3FA: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D400: lea ecx, var_1C
  loc_0042D403: call [00401114h] ; __vbaFreeObj
  loc_0042D409: call [0040101Ch] ; __vbaEnd
  loc_0042D40F: mov var_4, esi
  loc_0042D412: push 0042D43Fh
  loc_0042D417: jmp 0042D43Eh
  loc_0042D419: lea ecx, var_1C
  loc_0042D41C: call [00401114h] ; __vbaFreeObj
  loc_0042D422: lea edx, var_5C
  loc_0042D425: lea eax, var_4C
  loc_0042D428: push edx
  loc_0042D429: lea ecx, var_3C
  loc_0042D42C: push eax
  loc_0042D42D: lea edx, var_2C
  loc_0042D430: push ecx
  loc_0042D431: push edx
  loc_0042D432: push 00000004h
  loc_0042D434: call [00401018h] ; __vbaFreeVarList
  loc_0042D43A: add esp, 00000014h
  loc_0042D43D: ret
  loc_0042D43E: ret
  loc_0042D43F: mov eax, Me
  loc_0042D442: push eax
  loc_0042D443: mov ecx, [eax]
  loc_0042D445: call [ecx+00000008h]
  loc_0042D448: mov eax, var_4
  loc_0042D44B: mov ecx, var_14
  loc_0042D44E: pop edi
  loc_0042D44F: pop esi
  loc_0042D450: mov fs:[00000000h], ecx
  loc_0042D457: pop ebx
  loc_0042D458: mov esp, ebp
  loc_0042D45A: pop ebp
  loc_0042D45B: retn 000Ch
End Sub

Private Sub Form_Activate() '42B0C0
  loc_0042B0C0: push ebp
  loc_0042B0C1: mov ebp, esp
  loc_0042B0C3: sub esp, 0000000Ch
  loc_0042B0C6: push 00401926h ; __vbaExceptHandler
  loc_0042B0CB: mov eax, fs:[00000000h]
  loc_0042B0D1: push eax
  loc_0042B0D2: mov fs:[00000000h], esp
  loc_0042B0D9: sub esp, 00000008h
  loc_0042B0DC: push ebx
  loc_0042B0DD: push esi
  loc_0042B0DE: push edi
  loc_0042B0DF: mov var_C, esp
  loc_0042B0E2: mov var_8, 00401770h
  loc_0042B0E9: mov esi, Me
  loc_0042B0EC: mov eax, esi
  loc_0042B0EE: and eax, 00000001h
  loc_0042B0F1: mov var_4, eax
  loc_0042B0F4: and esi, FFFFFFFEh
  loc_0042B0F7: push esi
  loc_0042B0F8: mov Me, esi
  loc_0042B0FB: mov ecx, [esi]
  loc_0042B0FD: call [ecx+00000004h]
  loc_0042B100: mov edx, [esi]
  loc_0042B102: push esi
  loc_0042B103: call [edx+00000708h]
  loc_0042B109: mov var_4, 00000000h
  loc_0042B110: mov eax, Me
  loc_0042B113: push eax
  loc_0042B114: mov ecx, [eax]
  loc_0042B116: call [ecx+00000008h]
  loc_0042B119: mov eax, var_4
  loc_0042B11C: mov ecx, var_14
  loc_0042B11F: pop edi
  loc_0042B120: pop esi
  loc_0042B121: mov fs:[00000000h], ecx
  loc_0042B128: pop ebx
  loc_0042B129: mov esp, ebp
  loc_0042B12B: pop ebp
  loc_0042B12C: retn 0004h
End Sub

Private Sub mnuChangeXg_Click() '42CFD0
  loc_0042CFD0: push ebp
  loc_0042CFD1: mov ebp, esp
  loc_0042CFD3: sub esp, 0000000Ch
  loc_0042CFD6: push 00401926h ; __vbaExceptHandler
  loc_0042CFDB: mov eax, fs:[00000000h]
  loc_0042CFE1: push eax
  loc_0042CFE2: mov fs:[00000000h], esp
  loc_0042CFE9: sub esp, 00000030h
  loc_0042CFEC: push ebx
  loc_0042CFED: push esi
  loc_0042CFEE: push edi
  loc_0042CFEF: mov var_C, esp
  loc_0042CFF2: mov var_8, 004017C0h
  loc_0042CFF9: mov eax, Me
  loc_0042CFFC: mov ecx, eax
  loc_0042CFFE: and ecx, 00000001h
  loc_0042D001: mov var_4, ecx
  loc_0042D004: and al, FEh
  loc_0042D006: push eax
  loc_0042D007: mov Me, eax
  loc_0042D00A: mov edx, [eax]
  loc_0042D00C: call [edx+00000004h]
  loc_0042D00F: mov eax, [004301B4h]
  loc_0042D014: test eax, eax
  loc_0042D016: jnz 0042D028h
  loc_0042D018: push 004301B4h
  loc_0042D01D: push 00402480h
  loc_0042D022: call [004010D4h] ; __vbaNew2
  loc_0042D028: sub esp, 00000010h
  loc_0042D02B: mov ecx, 0000000Ah
  loc_0042D030: mov ebx, esp
  loc_0042D032: mov var_24, ecx
  loc_0042D035: mov eax, 80020004h
  loc_0042D03A: sub esp, 00000010h
  loc_0042D03D: mov [ebx], ecx
  loc_0042D03F: mov ecx, var_30
  loc_0042D042: mov edx, eax
  loc_0042D044: mov esi, [004301B4h]
  loc_0042D04A: mov [ebx+00000004h], ecx
  loc_0042D04D: mov ecx, esp
  loc_0042D04F: mov edi, [esi]
  loc_0042D051: push esi
  loc_0042D052: mov [ebx+00000008h], eax
  loc_0042D055: mov eax, var_28
  loc_0042D058: mov [ebx+0000000Ch], eax
  loc_0042D05B: mov eax, var_24
  loc_0042D05E: mov [ecx], eax
  loc_0042D060: mov eax, var_20
  loc_0042D063: mov [ecx+00000004h], eax
  loc_0042D066: mov [ecx+00000008h], edx
  loc_0042D069: mov edx, var_18
  loc_0042D06C: mov [ecx+0000000Ch], edx
  loc_0042D06F: call [edi+000002B0h]
  loc_0042D075: test eax, eax
  loc_0042D077: fnclex
  loc_0042D079: jge 0042D08Dh
  loc_0042D07B: push 000002B0h
  loc_0042D080: push 00410574h
  loc_0042D085: push esi
  loc_0042D086: push eax
  loc_0042D087: call [00401030h] ; __vbaHresultCheckObj
  loc_0042D08D: mov var_4, 00000000h
  loc_0042D094: mov eax, Me
  loc_0042D097: push eax
  loc_0042D098: mov ecx, [eax]
  loc_0042D09A: call [ecx+00000008h]
  loc_0042D09D: mov eax, var_4
  loc_0042D0A0: mov ecx, var_14
  loc_0042D0A3: pop edi
  loc_0042D0A4: pop esi
  loc_0042D0A5: mov fs:[00000000h], ecx
  loc_0042D0AC: pop ebx
  loc_0042D0AD: mov esp, ebp
  loc_0042D0AF: pop ebp
  loc_0042D0B0: retn 0004h
End Sub

Private Sub mnuChangeNames_Click() '42CD50
  loc_0042CD50: push ebp
  loc_0042CD51: mov ebp, esp
  loc_0042CD53: sub esp, 0000000Ch
  loc_0042CD56: push 00401926h ; __vbaExceptHandler
  loc_0042CD5B: mov eax, fs:[00000000h]
  loc_0042CD61: push eax
  loc_0042CD62: mov fs:[00000000h], esp
  loc_0042CD69: sub esp, 00000030h
  loc_0042CD6C: push ebx
  loc_0042CD6D: push esi
  loc_0042CD6E: push edi
  loc_0042CD6F: mov var_C, esp
  loc_0042CD72: mov var_8, 004017B0h
  loc_0042CD79: mov eax, Me
  loc_0042CD7C: mov ecx, eax
  loc_0042CD7E: and ecx, 00000001h
  loc_0042CD81: mov var_4, ecx
  loc_0042CD84: and al, FEh
  loc_0042CD86: push eax
  loc_0042CD87: mov Me, eax
  loc_0042CD8A: mov edx, [eax]
  loc_0042CD8C: call [edx+00000004h]
  loc_0042CD8F: mov eax, [004300D8h]
  loc_0042CD94: test eax, eax
  loc_0042CD96: jnz 0042CDA8h
  loc_0042CD98: push 004300D8h
  loc_0042CD9D: push 00402E04h
  loc_0042CDA2: call [004010D4h] ; __vbaNew2
  loc_0042CDA8: sub esp, 00000010h
  loc_0042CDAB: mov ecx, 0000000Ah
  loc_0042CDB0: mov ebx, esp
  loc_0042CDB2: mov var_24, ecx
  loc_0042CDB5: mov eax, 80020004h
  loc_0042CDBA: sub esp, 00000010h
  loc_0042CDBD: mov [ebx], ecx
  loc_0042CDBF: mov ecx, var_30
  loc_0042CDC2: mov edx, eax
  loc_0042CDC4: mov esi, [004300D8h]
  loc_0042CDCA: mov [ebx+00000004h], ecx
  loc_0042CDCD: mov ecx, esp
  loc_0042CDCF: mov edi, [esi]
  loc_0042CDD1: push esi
  loc_0042CDD2: mov [ebx+00000008h], eax
  loc_0042CDD5: mov eax, var_28
  loc_0042CDD8: mov [ebx+0000000Ch], eax
  loc_0042CDDB: mov eax, var_24
  loc_0042CDDE: mov [ecx], eax
  loc_0042CDE0: mov eax, var_20
  loc_0042CDE3: mov [ecx+00000004h], eax
  loc_0042CDE6: mov [ecx+00000008h], edx
  loc_0042CDE9: mov edx, var_18
  loc_0042CDEC: mov [ecx+0000000Ch], edx
  loc_0042CDEF: call [edi+000002B0h]
  loc_0042CDF5: test eax, eax
  loc_0042CDF7: fnclex
  loc_0042CDF9: jge 0042CE0Dh
  loc_0042CDFB: push 000002B0h
  loc_0042CE00: push 0040E260h
  loc_0042CE05: push esi
  loc_0042CE06: push eax
  loc_0042CE07: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CE0D: mov var_4, 00000000h
  loc_0042CE14: mov eax, Me
  loc_0042CE17: push eax
  loc_0042CE18: mov ecx, [eax]
  loc_0042CE1A: call [ecx+00000008h]
  loc_0042CE1D: mov eax, var_4
  loc_0042CE20: mov ecx, var_14
  loc_0042CE23: pop edi
  loc_0042CE24: pop esi
  loc_0042CE25: mov fs:[00000000h], ecx
  loc_0042CE2C: pop ebx
  loc_0042CE2D: mov esp, ebp
  loc_0042CE2F: pop ebp
  loc_0042CE30: retn 0004h
End Sub

Private Sub mnuAssumptions_Click() '42CA40
  loc_0042CA40: push ebp
  loc_0042CA41: mov ebp, esp
  loc_0042CA43: sub esp, 0000000Ch
  loc_0042CA46: push 00401926h ; __vbaExceptHandler
  loc_0042CA4B: mov eax, fs:[00000000h]
  loc_0042CA51: push eax
  loc_0042CA52: mov fs:[00000000h], esp
  loc_0042CA59: sub esp, 00000030h
  loc_0042CA5C: push ebx
  loc_0042CA5D: push esi
  loc_0042CA5E: push edi
  loc_0042CA5F: mov var_C, esp
  loc_0042CA62: mov var_8, 00401798h
  loc_0042CA69: mov eax, Me
  loc_0042CA6C: mov ecx, eax
  loc_0042CA6E: and ecx, 00000001h
  loc_0042CA71: mov var_4, ecx
  loc_0042CA74: and al, FEh
  loc_0042CA76: push eax
  loc_0042CA77: mov Me, eax
  loc_0042CA7A: mov edx, [eax]
  loc_0042CA7C: call [edx+00000004h]
  loc_0042CA7F: mov eax, [0043009Ch]
  loc_0042CA84: test eax, eax
  loc_0042CA86: jnz 0042CA98h
  loc_0042CA88: push 0043009Ch
  loc_0042CA8D: push 00405FC0h
  loc_0042CA92: call [004010D4h] ; __vbaNew2
  loc_0042CA98: sub esp, 00000010h
  loc_0042CA9B: mov ecx, 0000000Ah
  loc_0042CAA0: mov ebx, esp
  loc_0042CAA2: mov var_24, ecx
  loc_0042CAA5: mov eax, 80020004h
  loc_0042CAAA: sub esp, 00000010h
  loc_0042CAAD: mov [ebx], ecx
  loc_0042CAAF: mov ecx, var_30
  loc_0042CAB2: mov edx, eax
  loc_0042CAB4: mov esi, [0043009Ch]
  loc_0042CABA: mov [ebx+00000004h], ecx
  loc_0042CABD: mov ecx, esp
  loc_0042CABF: mov edi, [esi]
  loc_0042CAC1: push esi
  loc_0042CAC2: mov [ebx+00000008h], eax
  loc_0042CAC5: mov eax, var_28
  loc_0042CAC8: mov [ebx+0000000Ch], eax
  loc_0042CACB: mov eax, var_24
  loc_0042CACE: mov [ecx], eax
  loc_0042CAD0: mov eax, var_20
  loc_0042CAD3: mov [ecx+00000004h], eax
  loc_0042CAD6: mov [ecx+00000008h], edx
  loc_0042CAD9: mov edx, var_18
  loc_0042CADC: mov [ecx+0000000Ch], edx
  loc_0042CADF: call [edi+000002B0h]
  loc_0042CAE5: test eax, eax
  loc_0042CAE7: fnclex
  loc_0042CAE9: jge 0042CAFDh
  loc_0042CAEB: push 000002B0h
  loc_0042CAF0: push 0040DDE0h
  loc_0042CAF5: push esi
  loc_0042CAF6: push eax
  loc_0042CAF7: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CAFD: mov eax, [004300ECh]
  loc_0042CB02: test eax, eax
  loc_0042CB04: jnz 0042CB16h
  loc_0042CB06: push 004300ECh
  loc_0042CB0B: push 0040A624h
  loc_0042CB10: call [004010D4h] ; __vbaNew2
  loc_0042CB16: mov esi, [004300ECh]
  loc_0042CB1C: push esi
  loc_0042CB1D: mov eax, [esi]
  loc_0042CB1F: call [eax+000002B4h]
  loc_0042CB25: test eax, eax
  loc_0042CB27: fnclex
  loc_0042CB29: jge 0042CB3Dh
  loc_0042CB2B: push 000002B4h
  loc_0042CB30: push 0040ECECh
  loc_0042CB35: push esi
  loc_0042CB36: push eax
  loc_0042CB37: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CB3D: mov var_4, 00000000h
  loc_0042CB44: mov eax, Me
  loc_0042CB47: push eax
  loc_0042CB48: mov ecx, [eax]
  loc_0042CB4A: call [ecx+00000008h]
  loc_0042CB4D: mov eax, var_4
  loc_0042CB50: mov ecx, var_14
  loc_0042CB53: pop edi
  loc_0042CB54: pop esi
  loc_0042CB55: mov fs:[00000000h], ecx
  loc_0042CB5C: pop ebx
  loc_0042CB5D: mov esp, ebp
  loc_0042CB5F: pop ebp
  loc_0042CB60: retn 0004h
End Sub

Private Sub Proc_15_17_42B1A0
  loc_0042B1A0: push ebp
  loc_0042B1A1: mov ebp, esp
  loc_0042B1A3: sub esp, 00000008h
  loc_0042B1A6: push 00401926h ; __vbaExceptHandler
  loc_0042B1AB: mov eax, fs:[00000000h]
  loc_0042B1B1: push eax
  loc_0042B1B2: mov fs:[00000000h], esp
  loc_0042B1B9: sub esp, 0000017Ch
  loc_0042B1BF: push ebx
  loc_0042B1C0: push esi
  loc_0042B1C1: push edi
  loc_0042B1C2: mov var_8, esp
  loc_0042B1C5: mov var_4, 00401788h
  loc_0042B1CC: mov esi, Me
  loc_0042B1CF: xor ebx, ebx
  loc_0042B1D1: push esi
  loc_0042B1D2: mov var_14, ebx
  loc_0042B1D5: mov eax, [esi]
  loc_0042B1D7: mov var_18, ebx
  loc_0042B1DA: mov var_1C, ebx
  loc_0042B1DD: mov var_20, ebx
  loc_0042B1E0: mov var_24, ebx
  loc_0042B1E3: mov var_28, ebx
  loc_0042B1E6: mov var_2C, ebx
  loc_0042B1E9: mov var_30, ebx
  loc_0042B1EC: mov var_34, ebx
  loc_0042B1EF: mov var_38, ebx
  loc_0042B1F2: mov var_3C, ebx
  loc_0042B1F5: mov var_40, ebx
  loc_0042B1F8: mov var_50, ebx
  loc_0042B1FB: mov var_60, ebx
  loc_0042B1FE: mov var_70, ebx
  loc_0042B201: mov var_80, ebx
  loc_0042B204: mov var_90, ebx
  loc_0042B20A: mov var_A0, ebx
  loc_0042B210: mov var_B0, ebx
  loc_0042B216: mov var_C0, ebx
  loc_0042B21C: mov var_D0, ebx
  loc_0042B222: mov var_E0, ebx
  loc_0042B228: mov var_F0, ebx
  loc_0042B22E: mov var_100, ebx
  loc_0042B234: mov var_110, ebx
  loc_0042B23A: mov var_120, ebx
  loc_0042B240: mov var_130, ebx
  loc_0042B246: call [eax+00000360h]
  loc_0042B24C: mov edi, [0040103Ch] ; __vbaObjSet
  loc_0042B252: lea ecx, var_40
  loc_0042B255: push eax
  loc_0042B256: push ecx
  loc_0042B257: call edi
  loc_0042B259: mov edx, [eax]
  loc_0042B25B: push 00411EECh ; "The change in the mean of Y given a one unit change in X."
  loc_0042B260: push eax
  loc_0042B261: mov var_134, eax
  loc_0042B267: call [edx+00000054h]
  loc_0042B26A: cmp eax, ebx
  loc_0042B26C: fnclex
  loc_0042B26E: jge 0042B285h
  loc_0042B270: mov ecx, var_134
  loc_0042B276: push 00000054h
  loc_0042B278: push 0040E390h
  loc_0042B27D: push ecx
  loc_0042B27E: push eax
  loc_0042B27F: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B285: mov ebx, [00401114h] ; __vbaFreeObj
  loc_0042B28B: lea ecx, var_40
  loc_0042B28E: call ebx
  loc_0042B290: mov edx, [esi]
  loc_0042B292: push esi
  loc_0042B293: call [edx+00000358h]
  loc_0042B299: push eax
  loc_0042B29A: lea eax, var_40
  loc_0042B29D: push eax
  loc_0042B29E: call edi
  loc_0042B2A0: mov ecx, [eax]
  loc_0042B2A2: push 00411F64h ; "The estimated change in the mean of Y given a one unit change in X."
  loc_0042B2A7: push eax
  loc_0042B2A8: mov var_134, eax
  loc_0042B2AE: call [ecx+00000054h]
  loc_0042B2B1: test eax, eax
  loc_0042B2B3: fnclex
  loc_0042B2B5: jge 0042B2CCh
  loc_0042B2B7: mov edx, var_134
  loc_0042B2BD: push 00000054h
  loc_0042B2BF: push 0040E390h
  loc_0042B2C4: push edx
  loc_0042B2C5: push eax
  loc_0042B2C6: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B2CC: lea ecx, var_40
  loc_0042B2CF: call ebx
  loc_0042B2D1: mov eax, [esi]
  loc_0042B2D3: push esi
  loc_0042B2D4: call [eax+00000340h]
  loc_0042B2DA: lea ecx, var_40
  loc_0042B2DD: push eax
  loc_0042B2DE: push ecx
  loc_0042B2DF: call edi
  loc_0042B2E1: mov edx, [eax]
  loc_0042B2E3: push 00411FF0h ; "The expected value of Y given X is equal to Xg."
  loc_0042B2E8: push eax
  loc_0042B2E9: mov var_134, eax
  loc_0042B2EF: call [edx+00000054h]
  loc_0042B2F2: test eax, eax
  loc_0042B2F4: fnclex
  loc_0042B2F6: jge 0042B30Dh
  loc_0042B2F8: mov ecx, var_134
  loc_0042B2FE: push 00000054h
  loc_0042B300: push 0040E390h
  loc_0042B305: push ecx
  loc_0042B306: push eax
  loc_0042B307: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B30D: lea ecx, var_40
  loc_0042B310: call ebx
  loc_0042B312: mov edx, [esi]
  loc_0042B314: push esi
  loc_0042B315: call [edx+00000338h]
  loc_0042B31B: push eax
  loc_0042B31C: lea eax, var_40
  loc_0042B31F: push eax
  loc_0042B320: call edi
  loc_0042B322: mov ecx, [eax]
  loc_0042B324: push 00412054h ; "The estimate of the expected value of Y when X equals Xg."
  loc_0042B329: push eax
  loc_0042B32A: mov var_134, eax
  loc_0042B330: call [ecx+00000054h]
  loc_0042B333: test eax, eax
  loc_0042B335: fnclex
  loc_0042B337: jge 0042B34Eh
  loc_0042B339: mov edx, var_134
  loc_0042B33F: push 00000054h
  loc_0042B341: push 0040E390h
  loc_0042B346: push edx
  loc_0042B347: push eax
  loc_0042B348: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B34E: lea ecx, var_40
  loc_0042B351: call ebx
  loc_0042B353: mov eax, [esi]
  loc_0042B355: push esi
  loc_0042B356: call [eax+00000350h]
  loc_0042B35C: lea ecx, var_40
  loc_0042B35F: push eax
  loc_0042B360: push ecx
  loc_0042B361: call edi
  loc_0042B363: mov edx, [eax]
  loc_0042B365: push 004120CCh ; "The value of Y when X takes on the value of Xg."
  loc_0042B36A: push eax
  loc_0042B36B: mov var_134, eax
  loc_0042B371: call [edx+00000054h]
  loc_0042B374: test eax, eax
  loc_0042B376: fnclex
  loc_0042B378: jge 0042B38Fh
  loc_0042B37A: mov ecx, var_134
  loc_0042B380: push 00000054h
  loc_0042B382: push 0040E390h
  loc_0042B387: push ecx
  loc_0042B388: push eax
  loc_0042B389: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B38F: lea ecx, var_40
  loc_0042B392: call ebx
  loc_0042B394: mov edx, [esi]
  loc_0042B396: push esi
  loc_0042B397: call [edx+00000348h]
  loc_0042B39D: push eax
  loc_0042B39E: lea eax, var_40
  loc_0042B3A1: push eax
  loc_0042B3A2: call edi
  loc_0042B3A4: mov ecx, [eax]
  loc_0042B3A6: push 00412130h ; "The estimate of the value of Y when X = Xg."
  loc_0042B3AB: push eax
  loc_0042B3AC: mov var_134, eax
  loc_0042B3B2: call [ecx+00000054h]
  loc_0042B3B5: test eax, eax
  loc_0042B3B7: fnclex
  loc_0042B3B9: jge 0042B3D0h
  loc_0042B3BB: mov edx, var_134
  loc_0042B3C1: push 00000054h
  loc_0042B3C3: push 0040E390h
  loc_0042B3C8: push edx
  loc_0042B3C9: push eax
  loc_0042B3CA: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B3D0: lea ecx, var_40
  loc_0042B3D3: call ebx
  loc_0042B3D5: mov eax, [esi]
  loc_0042B3D7: push esi
  loc_0042B3D8: call [eax+00000328h]
  loc_0042B3DE: lea ecx, var_40
  loc_0042B3E1: push eax
  loc_0042B3E2: push ecx
  loc_0042B3E3: call edi
  loc_0042B3E5: mov edx, [eax]
  loc_0042B3E7: push 0041218Ch ; "The variance of Y given X = Xg."
  loc_0042B3EC: push eax
  loc_0042B3ED: mov var_134, eax
  loc_0042B3F3: call [edx+00000054h]
  loc_0042B3F6: test eax, eax
  loc_0042B3F8: fnclex
  loc_0042B3FA: jge 0042B411h
  loc_0042B3FC: mov ecx, var_134
  loc_0042B402: push 00000054h
  loc_0042B404: push 0040E390h
  loc_0042B409: push ecx
  loc_0042B40A: push eax
  loc_0042B40B: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B411: lea ecx, var_40
  loc_0042B414: call ebx
  loc_0042B416: mov edx, [esi]
  loc_0042B418: push esi
  loc_0042B419: call [edx+00000320h]
  loc_0042B41F: push eax
  loc_0042B420: lea eax, var_40
  loc_0042B423: push eax
  loc_0042B424: call edi
  loc_0042B426: mov ecx, [eax]
  loc_0042B428: push 004121D0h ; "The estimate of the variance of Y given X = Xg."
  loc_0042B42D: push eax
  loc_0042B42E: mov var_134, eax
  loc_0042B434: call [ecx+00000054h]
  loc_0042B437: test eax, eax
  loc_0042B439: fnclex
  loc_0042B43B: jge 0042B452h
  loc_0042B43D: mov edx, var_134
  loc_0042B443: push 00000054h
  loc_0042B445: push 0040E390h
  loc_0042B44A: push edx
  loc_0042B44B: push eax
  loc_0042B44C: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B452: lea ecx, var_40
  loc_0042B455: call ebx
  loc_0042B457: mov eax, [esi]
  loc_0042B459: push esi
  loc_0042B45A: call [eax+00000330h]
  loc_0042B460: lea ecx, var_40
  loc_0042B463: push eax
  loc_0042B464: push ecx
  loc_0042B465: call edi
  loc_0042B467: mov edx, [eax]
  loc_0042B469: push 00412234h ; "There is an r standard deviation change in the estimate of the mean of Y for each one standard deviation increase in X."
  loc_0042B46E: push eax
  loc_0042B46F: mov var_134, eax
  loc_0042B475: call [edx+00000054h]
  loc_0042B478: test eax, eax
  loc_0042B47A: fnclex
  loc_0042B47C: jge 0042B493h
  loc_0042B47E: mov ecx, var_134
  loc_0042B484: push 00000054h
  loc_0042B486: push 0040E390h
  loc_0042B48B: push ecx
  loc_0042B48C: push eax
  loc_0042B48D: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B493: lea ecx, var_40
  loc_0042B496: call ebx
  loc_0042B498: mov edx, [esi]
  loc_0042B49A: push esi
  loc_0042B49B: call [edx+00000318h]
  loc_0042B4A1: push eax
  loc_0042B4A2: lea eax, var_40
  loc_0042B4A5: push eax
  loc_0042B4A6: call edi
  loc_0042B4A8: mov ecx, [eax]
  loc_0042B4AA: push 004123BCh ; "r-squared percent of the sample variation in Y is associated with changes in X through the model."
  loc_0042B4AF: push eax
  loc_0042B4B0: mov var_134, eax
  loc_0042B4B6: call [ecx+00000054h]
  loc_0042B4B9: test eax, eax
  loc_0042B4BB: fnclex
  loc_0042B4BD: jge 0042B4D4h
  loc_0042B4BF: mov edx, var_134
  loc_0042B4C5: push 00000054h
  loc_0042B4C7: push 0040E390h
  loc_0042B4CC: push edx
  loc_0042B4CD: push eax
  loc_0042B4CE: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B4D4: lea ecx, var_40
  loc_0042B4D7: call ebx
  loc_0042B4D9: mov eax, [esi]
  loc_0042B4DB: push esi
  loc_0042B4DC: call [eax+00000310h]
  loc_0042B4E2: lea ecx, var_40
  loc_0042B4E5: push eax
  loc_0042B4E6: push ecx
  loc_0042B4E7: call edi
  loc_0042B4E9: mov edx, [eax]
  loc_0042B4EB: push 00412484h ; " The difference between Y and the mean of Y when X = Xg."
  loc_0042B4F0: push eax
  loc_0042B4F1: mov var_134, eax
  loc_0042B4F7: call [edx+00000054h]
  loc_0042B4FA: test eax, eax
  loc_0042B4FC: fnclex
  loc_0042B4FE: jge 0042B515h
  loc_0042B500: mov ecx, var_134
  loc_0042B506: push 00000054h
  loc_0042B508: push 0040E390h
  loc_0042B50D: push ecx
  loc_0042B50E: push eax
  loc_0042B50F: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B515: lea ecx, var_40
  loc_0042B518: call ebx
  loc_0042B51A: mov edx, [esi]
  loc_0042B51C: push esi
  loc_0042B51D: call [edx+00000308h]
  loc_0042B523: push eax
  loc_0042B524: lea eax, var_40
  loc_0042B527: push eax
  loc_0042B528: call edi
  loc_0042B52A: mov ecx, [eax]
  loc_0042B52C: push 004124FCh ; "The differnce between Y and the estimated mean of Y when X equals Xg."
  loc_0042B531: push eax
  loc_0042B532: mov var_134, eax
  loc_0042B538: call [ecx+00000054h]
  loc_0042B53B: test eax, eax
  loc_0042B53D: fnclex
  loc_0042B53F: jge 0042B556h
  loc_0042B541: mov edx, var_134
  loc_0042B547: push 00000054h
  loc_0042B549: push 0040E390h
  loc_0042B54E: push edx
  loc_0042B54F: push eax
  loc_0042B550: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B556: lea ecx, var_40
  loc_0042B559: call ebx
  loc_0042B55B: mov eax, [esi]
  loc_0042B55D: push esi
  loc_0042B55E: call [eax+00000360h]
  loc_0042B564: lea ecx, var_40
  loc_0042B567: push eax
  loc_0042B568: push ecx
  loc_0042B569: call edi
  loc_0042B56B: mov edx, [00430010h]
  loc_0042B571: mov edi, [0040102Ch] ; __vbaStrCat
  loc_0042B577: mov ebx, [eax]
  loc_0042B579: push 00412328h ; "The change in the mean of "
  loc_0042B57E: push edx
  loc_0042B57F: mov var_134, eax
  loc_0042B585: call edi
  loc_0042B587: mov esi, [004010FCh] ; __vbaStrMove
  loc_0042B58D: mov edx, eax
  loc_0042B58F: lea ecx, var_14
  loc_0042B592: call __vbaStrMove
  loc_0042B594: push eax
  loc_0042B595: push 00412364h ; " given a one "
  loc_0042B59A: call edi
  loc_0042B59C: mov edx, eax
  loc_0042B59E: lea ecx, var_18
  loc_0042B5A1: call __vbaStrMove
  loc_0042B5A3: push eax
  loc_0042B5A4: mov eax, [0043001Ch]
  loc_0042B5A9: push eax
  loc_0042B5AA: call edi
  loc_0042B5AC: mov edx, eax
  loc_0042B5AE: lea ecx, var_1C
  loc_0042B5B1: call __vbaStrMove
  loc_0042B5B3: push eax
  loc_0042B5B4: push 00410E74h ; " increase in "
  loc_0042B5B9: call edi
  loc_0042B5BB: mov edx, eax
  loc_0042B5BD: lea ecx, var_20
  loc_0042B5C0: call __vbaStrMove
  loc_0042B5C2: mov ecx, [00430018h]
  loc_0042B5C8: push eax
  loc_0042B5C9: push ecx
  loc_0042B5CA: call edi
  loc_0042B5CC: mov edx, eax
  loc_0042B5CE: lea ecx, var_24
  loc_0042B5D1: call __vbaStrMove
  loc_0042B5D3: push eax
  loc_0042B5D4: push 0040DD3Ch ; "."
  loc_0042B5D9: call edi
  loc_0042B5DB: mov edx, eax
  loc_0042B5DD: lea ecx, var_28
  loc_0042B5E0: call __vbaStrMove
  loc_0042B5E2: mov edx, ebx
  loc_0042B5E4: mov ebx, var_134
  loc_0042B5EA: push eax
  loc_0042B5EB: push ebx
  loc_0042B5EC: call [edx+0000019Ch]
  loc_0042B5F2: test eax, eax
  loc_0042B5F4: fnclex
  loc_0042B5F6: jge 0042B60Ah
  loc_0042B5F8: push 0000019Ch
  loc_0042B5FD: push 0040E390h
  loc_0042B602: push ebx
  loc_0042B603: push eax
  loc_0042B604: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B60A: lea eax, var_28
  loc_0042B60D: lea ecx, var_24
  loc_0042B610: push eax
  loc_0042B611: lea edx, var_20
  loc_0042B614: push ecx
  loc_0042B615: lea eax, var_1C
  loc_0042B618: push edx
  loc_0042B619: lea ecx, var_18
  loc_0042B61C: push eax
  loc_0042B61D: lea edx, var_14
  loc_0042B620: push ecx
  loc_0042B621: push edx
  loc_0042B622: push 00000006h
  loc_0042B624: call [004010E4h] ; __vbaFreeStrList
  loc_0042B62A: add esp, 0000001Ch
  loc_0042B62D: lea ecx, var_40
  loc_0042B630: call [00401114h] ; __vbaFreeObj
  loc_0042B636: fld real4 ptr [00430060h]
  loc_0042B63C: fcomp real4 ptr [004011E8h]
  loc_0042B642: fnstsw ax
  loc_0042B644: test ah, 01h
  loc_0042B647: mov eax, Me
  loc_0042B64A: jnz 0042B73Bh
  loc_0042B650: mov ecx, [eax]
  loc_0042B652: push eax
  loc_0042B653: call [ecx+00000358h]
  loc_0042B659: lea edx, var_40
  loc_0042B65C: push eax
  loc_0042B65D: push edx
  loc_0042B65E: call [0040103Ch] ; __vbaObjSet
  loc_0042B664: mov ebx, [eax]
  loc_0042B666: mov var_134, eax
  loc_0042B66C: mov eax, [00430060h]
  loc_0042B671: push 00412384h ; "There is an estimated "
  loc_0042B676: push eax
  loc_0042B677: call [0040107Ch] ; __vbaStrR4
  loc_0042B67D: mov edx, eax
  loc_0042B67F: lea ecx, var_14
  loc_0042B682: call __vbaStrMove
  loc_0042B684: push eax
  loc_0042B685: call edi
  loc_0042B687: mov edx, eax
  loc_0042B689: lea ecx, var_18
  loc_0042B68C: call __vbaStrMove
  loc_0042B68E: push eax
  loc_0042B68F: push 0040F028h
  loc_0042B694: call edi
  loc_0042B696: mov edx, eax
  loc_0042B698: lea ecx, var_1C
  loc_0042B69B: call __vbaStrMove
  loc_0042B69D: mov ecx, [00430014h]
  loc_0042B6A3: push eax
  loc_0042B6A4: push ecx
  loc_0042B6A5: call edi
  loc_0042B6A7: mov edx, eax
  loc_0042B6A9: lea ecx, var_20
  loc_0042B6AC: call __vbaStrMove
  loc_0042B6AE: push eax
  loc_0042B6AF: push 00412598h ; " increase in the mean of "
  loc_0042B6B4: call edi
  loc_0042B6B6: mov edx, eax
  loc_0042B6B8: lea ecx, var_24
  loc_0042B6BB: call __vbaStrMove
  loc_0042B6BD: mov edx, [00430010h]
  loc_0042B6C3: push eax
  loc_0042B6C4: push edx
  loc_0042B6C5: call edi
  loc_0042B6C7: mov edx, eax
  loc_0042B6C9: lea ecx, var_28
  loc_0042B6CC: call __vbaStrMove
  loc_0042B6CE: push eax
  loc_0042B6CF: push 00412364h ; " given a one "
  loc_0042B6D4: call edi
  loc_0042B6D6: mov edx, eax
  loc_0042B6D8: lea ecx, var_2C
  loc_0042B6DB: call __vbaStrMove
  loc_0042B6DD: push eax
  loc_0042B6DE: mov eax, [0043001Ch]
  loc_0042B6E3: push eax
  loc_0042B6E4: call edi
  loc_0042B6E6: mov edx, eax
  loc_0042B6E8: lea ecx, var_30
  loc_0042B6EB: call __vbaStrMove
  loc_0042B6ED: push eax
  loc_0042B6EE: push 00410E74h ; " increase in "
  loc_0042B6F3: call edi
  loc_0042B6F5: mov edx, eax
  loc_0042B6F7: lea ecx, var_34
  loc_0042B6FA: call __vbaStrMove
  loc_0042B6FC: mov ecx, [00430018h]
  loc_0042B702: push eax
  loc_0042B703: push ecx
  loc_0042B704: call edi
  loc_0042B706: mov edx, eax
  loc_0042B708: lea ecx, var_38
  loc_0042B70B: call __vbaStrMove
  loc_0042B70D: push eax
  loc_0042B70E: push 0040DD3Ch ; "."
  loc_0042B713: call edi
  loc_0042B715: mov edx, eax
  loc_0042B717: lea ecx, var_3C
  loc_0042B71A: call __vbaStrMove
  loc_0042B71C: push eax
  loc_0042B71D: mov edx, ebx
  loc_0042B71F: mov ebx, var_134
  loc_0042B725: push ebx
  loc_0042B726: call [edx+0000019Ch]
  loc_0042B72C: test eax, eax
  loc_0042B72E: fnclex
  loc_0042B730: jge 0042B84Bh
  loc_0042B736: jmp 0042B839h
  loc_0042B73B: mov edx, [eax]
  loc_0042B73D: push eax
  loc_0042B73E: call [edx+00000358h]
  loc_0042B744: push eax
  loc_0042B745: lea eax, var_40
  loc_0042B748: push eax
  loc_0042B749: call [0040103Ch] ; __vbaObjSet
  loc_0042B74F: fld real4 ptr [00430060h]
  loc_0042B755: mov ebx, [eax]
  loc_0042B757: mov var_134, eax
  loc_0042B75D: fabs
  loc_0042B75F: fnstsw ax
  loc_0042B761: test al, 0Dh
  loc_0042B763: jnz 0042C960h
  loc_0042B769: fstp real4 ptr var_148
  loc_0042B76F: fld real4 ptr var_148
  loc_0042B775: push 00412384h ; "There is an estimated "
  loc_0042B77A: push ecx
  loc_0042B77B: fstp real4 ptr [esp]
  loc_0042B77E: call [0040107Ch] ; __vbaStrR4
  loc_0042B784: mov edx, eax
  loc_0042B786: lea ecx, var_14
  loc_0042B789: call __vbaStrMove
  loc_0042B78B: push eax
  loc_0042B78C: call edi
  loc_0042B78E: mov edx, eax
  loc_0042B790: lea ecx, var_18
  loc_0042B793: call __vbaStrMove
  loc_0042B795: push eax
  loc_0042B796: push 0040F028h
  loc_0042B79B: call edi
  loc_0042B79D: mov edx, eax
  loc_0042B79F: lea ecx, var_1C
  loc_0042B7A2: call __vbaStrMove
  loc_0042B7A4: mov ecx, [00430014h]
  loc_0042B7AA: push eax
  loc_0042B7AB: push ecx
  loc_0042B7AC: call edi
  loc_0042B7AE: mov edx, eax
  loc_0042B7B0: lea ecx, var_20
  loc_0042B7B3: call __vbaStrMove
  loc_0042B7B5: push eax
  loc_0042B7B6: push 004125D0h ; " decrease in the mean of "
  loc_0042B7BB: call edi
  loc_0042B7BD: mov edx, eax
  loc_0042B7BF: lea ecx, var_24
  loc_0042B7C2: call __vbaStrMove
  loc_0042B7C4: mov edx, [00430010h]
  loc_0042B7CA: push eax
  loc_0042B7CB: push edx
  loc_0042B7CC: call edi
  loc_0042B7CE: mov edx, eax
  loc_0042B7D0: lea ecx, var_28
  loc_0042B7D3: call __vbaStrMove
  loc_0042B7D5: push eax
  loc_0042B7D6: push 00412364h ; " given a one "
  loc_0042B7DB: call edi
  loc_0042B7DD: mov edx, eax
  loc_0042B7DF: lea ecx, var_2C
  loc_0042B7E2: call __vbaStrMove
  loc_0042B7E4: push eax
  loc_0042B7E5: mov eax, [0043001Ch]
  loc_0042B7EA: push eax
  loc_0042B7EB: call edi
  loc_0042B7ED: mov edx, eax
  loc_0042B7EF: lea ecx, var_30
  loc_0042B7F2: call __vbaStrMove
  loc_0042B7F4: push eax
  loc_0042B7F5: push 00410E74h ; " increase in "
  loc_0042B7FA: call edi
  loc_0042B7FC: mov edx, eax
  loc_0042B7FE: lea ecx, var_34
  loc_0042B801: call __vbaStrMove
  loc_0042B803: mov ecx, [00430018h]
  loc_0042B809: push eax
  loc_0042B80A: push ecx
  loc_0042B80B: call edi
  loc_0042B80D: mov edx, eax
  loc_0042B80F: lea ecx, var_38
  loc_0042B812: call __vbaStrMove
  loc_0042B814: push eax
  loc_0042B815: push 0040DD3Ch ; "."
  loc_0042B81A: call edi
  loc_0042B81C: mov edx, eax
  loc_0042B81E: lea ecx, var_3C
  loc_0042B821: call __vbaStrMove
  loc_0042B823: mov edx, ebx
  loc_0042B825: mov ebx, var_134
  loc_0042B82B: push eax
  loc_0042B82C: push ebx
  loc_0042B82D: call [edx+0000019Ch]
  loc_0042B833: test eax, eax
  loc_0042B835: fnclex
  loc_0042B837: jge 0042B84Bh
  loc_0042B839: push 0000019Ch
  loc_0042B83E: push 0040E390h
  loc_0042B843: push ebx
  loc_0042B844: push eax
  loc_0042B845: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B84B: lea eax, var_3C
  loc_0042B84E: lea ecx, var_38
  loc_0042B851: push eax
  loc_0042B852: lea edx, var_34
  loc_0042B855: push ecx
  loc_0042B856: lea eax, var_30
  loc_0042B859: push edx
  loc_0042B85A: lea ecx, var_2C
  loc_0042B85D: push eax
  loc_0042B85E: lea edx, var_28
  loc_0042B861: push ecx
  loc_0042B862: lea eax, var_24
  loc_0042B865: push edx
  loc_0042B866: lea ecx, var_20
  loc_0042B869: push eax
  loc_0042B86A: lea edx, var_1C
  loc_0042B86D: push ecx
  loc_0042B86E: lea eax, var_18
  loc_0042B871: push edx
  loc_0042B872: lea ecx, var_14
  loc_0042B875: push eax
  loc_0042B876: push ecx
  loc_0042B877: push 0000000Bh
  loc_0042B879: call [004010E4h] ; __vbaFreeStrList
  loc_0042B87F: add esp, 00000030h
  loc_0042B882: lea ecx, var_40
  loc_0042B885: call [00401114h] ; __vbaFreeObj
  loc_0042B88B: mov eax, Me
  loc_0042B88E: push eax
  loc_0042B88F: mov edx, [eax]
  loc_0042B891: call [edx+00000340h]
  loc_0042B897: push eax
  loc_0042B898: lea eax, var_40
  loc_0042B89B: push eax
  loc_0042B89C: call [0040103Ch] ; __vbaObjSet
  loc_0042B8A2: mov ecx, [00430010h]
  loc_0042B8A8: mov ebx, [eax]
  loc_0042B8AA: push 00412608h ; "The expected value of "
  loc_0042B8AF: push ecx
  loc_0042B8B0: mov var_134, eax
  loc_0042B8B6: call edi
  loc_0042B8B8: mov edx, eax
  loc_0042B8BA: lea ecx, var_14
  loc_0042B8BD: call __vbaStrMove
  loc_0042B8BF: push eax
  loc_0042B8C0: push 0041263Ch ; " given "
  loc_0042B8C5: call edi
  loc_0042B8C7: mov edx, eax
  loc_0042B8C9: lea ecx, var_18
  loc_0042B8CC: call __vbaStrMove
  loc_0042B8CE: mov edx, [00430018h]
  loc_0042B8D4: push eax
  loc_0042B8D5: push edx
  loc_0042B8D6: call edi
  loc_0042B8D8: mov edx, eax
  loc_0042B8DA: lea ecx, var_1C
  loc_0042B8DD: call __vbaStrMove
  loc_0042B8DF: push eax
  loc_0042B8E0: push 0040FF18h ; " is equal to "
  loc_0042B8E5: call edi
  loc_0042B8E7: mov edx, eax
  loc_0042B8E9: lea ecx, var_20
  loc_0042B8EC: call __vbaStrMove
  loc_0042B8EE: push eax
  loc_0042B8EF: mov eax, [0043002Ch]
  loc_0042B8F4: push eax
  loc_0042B8F5: call edi
  loc_0042B8F7: mov edx, eax
  loc_0042B8F9: lea ecx, var_24
  loc_0042B8FC: call __vbaStrMove
  loc_0042B8FE: push eax
  loc_0042B8FF: push 0040DD3Ch ; "."
  loc_0042B904: call edi
  loc_0042B906: mov edx, eax
  loc_0042B908: lea ecx, var_28
  loc_0042B90B: call __vbaStrMove
  loc_0042B90D: mov ecx, ebx
  loc_0042B90F: mov ebx, var_134
  loc_0042B915: push eax
  loc_0042B916: push ebx
  loc_0042B917: call [ecx+0000019Ch]
  loc_0042B91D: test eax, eax
  loc_0042B91F: fnclex
  loc_0042B921: jge 0042B935h
  loc_0042B923: push 0000019Ch
  loc_0042B928: push 0040E390h
  loc_0042B92D: push ebx
  loc_0042B92E: push eax
  loc_0042B92F: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B935: lea edx, var_28
  loc_0042B938: lea eax, var_24
  loc_0042B93B: push edx
  loc_0042B93C: lea ecx, var_20
  loc_0042B93F: push eax
  loc_0042B940: lea edx, var_1C
  loc_0042B943: push ecx
  loc_0042B944: lea eax, var_18
  loc_0042B947: push edx
  loc_0042B948: lea ecx, var_14
  loc_0042B94B: push eax
  loc_0042B94C: push ecx
  loc_0042B94D: push 00000006h
  loc_0042B94F: call [004010E4h] ; __vbaFreeStrList
  loc_0042B955: add esp, 0000001Ch
  loc_0042B958: lea ecx, var_40
  loc_0042B95B: call [00401114h] ; __vbaFreeObj
  loc_0042B961: mov eax, Me
  loc_0042B964: push eax
  loc_0042B965: mov edx, [eax]
  loc_0042B967: call [edx+0000033Ch]
  loc_0042B96D: push eax
  loc_0042B96E: lea eax, var_40
  loc_0042B971: push eax
  loc_0042B972: call [0040103Ch] ; __vbaObjSet
  loc_0042B978: mov ecx, [00430018h]
  loc_0042B97E: mov ebx, [eax]
  loc_0042B980: push 00412650h ; "The means are assumed to fall on a straight line function of "
  loc_0042B985: push ecx
  loc_0042B986: mov var_134, eax
  loc_0042B98C: call edi
  loc_0042B98E: mov edx, eax
  loc_0042B990: lea ecx, var_14
  loc_0042B993: call __vbaStrMove
  loc_0042B995: push eax
  loc_0042B996: push 0040DD3Ch ; "."
  loc_0042B99B: call edi
  loc_0042B99D: mov edx, eax
  loc_0042B99F: lea ecx, var_18
  loc_0042B9A2: call __vbaStrMove
  loc_0042B9A4: mov edx, ebx
  loc_0042B9A6: mov ebx, var_134
  loc_0042B9AC: push eax
  loc_0042B9AD: push ebx
  loc_0042B9AE: call [edx+0000014Ch]
  loc_0042B9B4: test eax, eax
  loc_0042B9B6: fnclex
  loc_0042B9B8: jge 0042B9CCh
  loc_0042B9BA: push 0000014Ch
  loc_0042B9BF: push 0040E728h
  loc_0042B9C4: push ebx
  loc_0042B9C5: push eax
  loc_0042B9C6: call [00401030h] ; __vbaHresultCheckObj
  loc_0042B9CC: lea eax, var_18
  loc_0042B9CF: lea ecx, var_14
  loc_0042B9D2: push eax
  loc_0042B9D3: push ecx
  loc_0042B9D4: push 00000002h
  loc_0042B9D6: call [004010E4h] ; __vbaFreeStrList
  loc_0042B9DC: add esp, 0000000Ch
  loc_0042B9DF: lea ecx, var_40
  loc_0042B9E2: call [00401114h] ; __vbaFreeObj
  loc_0042B9E8: mov eax, Me
  loc_0042B9EB: push eax
  loc_0042B9EC: mov edx, [eax]
  loc_0042B9EE: call [edx+00000338h]
  loc_0042B9F4: push eax
  loc_0042B9F5: lea eax, var_40
  loc_0042B9F8: push eax
  loc_0042B9F9: call [0040103Ch] ; __vbaObjSet
  loc_0042B9FF: mov ecx, [00430010h]
  loc_0042BA05: mov ebx, [eax]
  loc_0042BA07: push 004126D0h ; "The estimate of the expected value of "
  loc_0042BA0C: push ecx
  loc_0042BA0D: mov var_134, eax
  loc_0042BA13: call edi
  loc_0042BA15: mov edx, eax
  loc_0042BA17: lea ecx, var_14
  loc_0042BA1A: call __vbaStrMove
  loc_0042BA1C: push eax
  loc_0042BA1D: push 00410684h ; " is "
  loc_0042BA22: call edi
  loc_0042BA24: mov edx, eax
  loc_0042BA26: lea ecx, var_18
  loc_0042BA29: call __vbaStrMove
  loc_0042BA2B: mov edx, [00430038h]
  loc_0042BA31: push eax
  loc_0042BA32: push edx
  loc_0042BA33: call [0040107Ch] ; __vbaStrR4
  loc_0042BA39: mov edx, eax
  loc_0042BA3B: lea ecx, var_1C
  loc_0042BA3E: call __vbaStrMove
  loc_0042BA40: push eax
  loc_0042BA41: call edi
  loc_0042BA43: mov edx, eax
  loc_0042BA45: lea ecx, var_20
  loc_0042BA48: call __vbaStrMove
  loc_0042BA4A: push eax
  loc_0042BA4B: push 0040F028h
  loc_0042BA50: call edi
  loc_0042BA52: mov edx, eax
  loc_0042BA54: lea ecx, var_24
  loc_0042BA57: call __vbaStrMove
  loc_0042BA59: push eax
  loc_0042BA5A: mov eax, [00430014h]
  loc_0042BA5F: push eax
  loc_0042BA60: call edi
  loc_0042BA62: mov edx, eax
  loc_0042BA64: lea ecx, var_28
  loc_0042BA67: call __vbaStrMove
  loc_0042BA69: push eax
  loc_0042BA6A: push 0040FF04h ; " when "
  loc_0042BA6F: call edi
  loc_0042BA71: mov edx, eax
  loc_0042BA73: lea ecx, var_2C
  loc_0042BA76: call __vbaStrMove
  loc_0042BA78: mov ecx, [00430018h]
  loc_0042BA7E: push eax
  loc_0042BA7F: push ecx
  loc_0042BA80: call edi
  loc_0042BA82: mov edx, eax
  loc_0042BA84: lea ecx, var_30
  loc_0042BA87: call __vbaStrMove
  loc_0042BA89: push eax
  loc_0042BA8A: push 00410824h ; " equals "
  loc_0042BA8F: call edi
  loc_0042BA91: mov edx, eax
  loc_0042BA93: lea ecx, var_34
  loc_0042BA96: call __vbaStrMove
  loc_0042BA98: mov edx, [0043002Ch]
  loc_0042BA9E: push eax
  loc_0042BA9F: push edx
  loc_0042BAA0: call edi
  loc_0042BAA2: mov edx, eax
  loc_0042BAA4: lea ecx, var_38
  loc_0042BAA7: call __vbaStrMove
  loc_0042BAA9: push eax
  loc_0042BAAA: push 0040DD3Ch ; "."
  loc_0042BAAF: call edi
  loc_0042BAB1: mov edx, eax
  loc_0042BAB3: lea ecx, var_3C
  loc_0042BAB6: call __vbaStrMove
  loc_0042BAB8: mov var_158, ebx
  loc_0042BABE: mov ebx, var_134
  loc_0042BAC4: push eax
  loc_0042BAC5: mov eax, var_158
  loc_0042BACB: push ebx
  loc_0042BACC: call [eax+0000019Ch]
  loc_0042BAD2: test eax, eax
  loc_0042BAD4: fnclex
  loc_0042BAD6: jge 0042BAEAh
  loc_0042BAD8: push 0000019Ch
  loc_0042BADD: push 0040E390h
  loc_0042BAE2: push ebx
  loc_0042BAE3: push eax
  loc_0042BAE4: call [00401030h] ; __vbaHresultCheckObj
  loc_0042BAEA: lea ecx, var_3C
  loc_0042BAED: lea edx, var_38
  loc_0042BAF0: push ecx
  loc_0042BAF1: lea eax, var_34
  loc_0042BAF4: push edx
  loc_0042BAF5: lea ecx, var_30
  loc_0042BAF8: push eax
  loc_0042BAF9: lea edx, var_2C
  loc_0042BAFC: push ecx
  loc_0042BAFD: lea eax, var_28
  loc_0042BB00: push edx
  loc_0042BB01: lea ecx, var_24
  loc_0042BB04: push eax
  loc_0042BB05: lea edx, var_20
  loc_0042BB08: push ecx
  loc_0042BB09: lea eax, var_1C
  loc_0042BB0C: push edx
  loc_0042BB0D: lea ecx, var_18
  loc_0042BB10: push eax
  loc_0042BB11: lea edx, var_14
  loc_0042BB14: push ecx
  loc_0042BB15: push edx
  loc_0042BB16: push 0000000Bh
  loc_0042BB18: call [004010E4h] ; __vbaFreeStrList
  loc_0042BB1E: add esp, 00000030h
  loc_0042BB21: lea ecx, var_40
  loc_0042BB24: call [00401114h] ; __vbaFreeObj
  loc_0042BB2A: mov eax, Me
  loc_0042BB2D: push eax
  loc_0042BB2E: mov ecx, [eax]
  loc_0042BB30: call [ecx+00000334h]
  loc_0042BB36: lea edx, var_40
  loc_0042BB39: push eax
  loc_0042BB3A: push edx
  loc_0042BB3B: call [0040103Ch] ; __vbaObjSet
  loc_0042BB41: mov var_134, eax
  loc_0042BB47: lea eax, var_D0
  loc_0042BB4D: push 00000002h
  loc_0042BB4F: lea ecx, var_50
  loc_0042BB52: mov ebx, 00000008h
  loc_0042BB57: push eax
  loc_0042BB58: push ecx
  loc_0042BB59: mov var_D8, 00412724h ; "These estimates fall on the line: "
  loc_0042BB63: mov var_E0, ebx
  loc_0042BB69: mov var_C8, 00430074h
  loc_0042BB73: mov var_D0, 00004004h
  loc_0042BB7D: call [004010ACh] ; rtcRound
  loc_0042BB83: lea edx, var_100
  loc_0042BB89: push 00000002h
  loc_0042BB8B: lea eax, var_80
  loc_0042BB8E: push edx
  loc_0042BB8F: push eax
  loc_0042BB90: mov var_E8, 0041258Ch ; " + "
  loc_0042BB9A: mov var_F0, ebx
  loc_0042BBA0: mov var_F8, 00430060h
  loc_0042BBAA: mov var_100, 00004004h
  loc_0042BBB4: call [004010ACh] ; rtcRound
  loc_0042BBBA: mov ecx, [00430018h]
  loc_0042BBC0: mov edx, var_134
  loc_0042BBC6: mov var_108, 00412770h ; " * Xg (the given value of "
  loc_0042BBD0: mov var_110, ebx
  loc_0042BBD6: mov var_118, ecx
  loc_0042BBDC: mov var_120, ebx
  loc_0042BBE2: mov var_128, 004127ACh ; ")."
  loc_0042BBEC: mov var_130, ebx
  loc_0042BBF2: mov ebx, [edx]
  loc_0042BBF4: lea eax, var_E0
  loc_0042BBFA: lea ecx, var_50
  loc_0042BBFD: push eax
  loc_0042BBFE: lea edx, var_60
  loc_0042BC01: push ecx
  loc_0042BC02: push edx
  loc_0042BC03: call [004010C0h] ; __vbaVarCat
  loc_0042BC09: push eax
  loc_0042BC0A: lea eax, var_F0
  loc_0042BC10: lea ecx, var_70
  loc_0042BC13: push eax
  loc_0042BC14: push ecx
  loc_0042BC15: call [004010C0h] ; __vbaVarCat
  loc_0042BC1B: push eax
  loc_0042BC1C: lea edx, var_80
  loc_0042BC1F: lea eax, var_90
  loc_0042BC25: push edx
  loc_0042BC26: push eax
  loc_0042BC27: call [004010C0h] ; __vbaVarCat
  loc_0042BC2D: lea ecx, var_110
  loc_0042BC33: push eax
  loc_0042BC34: lea edx, var_A0
  loc_0042BC3A: push ecx
  loc_0042BC3B: push edx
  loc_0042BC3C: call [004010C0h] ; __vbaVarCat
  loc_0042BC42: push eax
  loc_0042BC43: lea eax, var_120
  loc_0042BC49: lea ecx, var_B0
  loc_0042BC4F: push eax
  loc_0042BC50: push ecx
  loc_0042BC51: call [004010C0h] ; __vbaVarCat
  loc_0042BC57: push eax
  loc_0042BC58: lea edx, var_130
  loc_0042BC5E: lea eax, var_C0
  loc_0042BC64: push edx
  loc_0042BC65: push eax
  loc_0042BC66: call [004010C0h] ; __vbaVarCat
  loc_0042BC6C: lea ecx, var_14
  loc_0042BC6F: push eax
  loc_0042BC70: push ecx
  loc_0042BC71: call [004010B8h] ; __vbaStrVarVal
  loc_0042BC77: mov edx, ebx
  loc_0042BC79: mov ebx, var_134
  loc_0042BC7F: push eax
  loc_0042BC80: push ebx
  loc_0042BC81: call [edx+0000014Ch]
  loc_0042BC87: test eax, eax
  loc_0042BC89: fnclex
  loc_0042BC8B: jge 0042BC9Fh
  loc_0042BC8D: push 0000014Ch
  loc_0042BC92: push 0040E728h
  loc_0042BC97: push ebx
  loc_0042BC98: push eax
  loc_0042BC99: call [00401030h] ; __vbaHresultCheckObj
  loc_0042BC9F: lea ecx, var_14
  loc_0042BCA2: call [00401110h] ; __vbaFreeStr
  loc_0042BCA8: lea ecx, var_40
  loc_0042BCAB: call [00401114h] ; __vbaFreeObj
  loc_0042BCB1: lea eax, var_C0
  loc_0042BCB7: lea ecx, var_B0
  loc_0042BCBD: push eax
  loc_0042BCBE: lea edx, var_A0
  loc_0042BCC4: push ecx
  loc_0042BCC5: lea eax, var_90
  loc_0042BCCB: push edx
  loc_0042BCCC: lea ecx, var_80
  loc_0042BCCF: push eax
  loc_0042BCD0: lea edx, var_70
  loc_0042BCD3: push ecx
  loc_0042BCD4: lea eax, var_60
  loc_0042BCD7: push edx
  loc_0042BCD8: lea ecx, var_50
  loc_0042BCDB: push eax
  loc_0042BCDC: push ecx
  loc_0042BCDD: push 00000008h
  loc_0042BCDF: call [00401018h] ; __vbaFreeVarList
  loc_0042BCE5: mov eax, Me
  loc_0042BCE8: add esp, 00000024h
  loc_0042BCEB: mov edx, [eax]
  loc_0042BCED: push eax
  loc_0042BCEE: call [edx+00000350h]
  loc_0042BCF4: push eax
  loc_0042BCF5: lea eax, var_40
  loc_0042BCF8: push eax
  loc_0042BCF9: call [0040103Ch] ; __vbaObjSet
  loc_0042BCFF: mov ecx, [00430010h]
  loc_0042BD05: mov ebx, [eax]
  loc_0042BD07: push 004127B8h ; "The value of "
  loc_0042BD0C: push ecx
  loc_0042BD0D: mov var_134, eax
  loc_0042BD13: call edi
  loc_0042BD15: mov edx, eax
  loc_0042BD17: lea ecx, var_14
  loc_0042BD1A: call __vbaStrMove
  loc_0042BD1C: push eax
  loc_0042BD1D: push 0040FF04h ; " when "
  loc_0042BD22: call edi
  loc_0042BD24: mov edx, eax
  loc_0042BD26: lea ecx, var_18
  loc_0042BD29: call __vbaStrMove
  loc_0042BD2B: mov edx, [00430018h]
  loc_0042BD31: push eax
  loc_0042BD32: push edx
  loc_0042BD33: call edi
  loc_0042BD35: mov edx, eax
  loc_0042BD37: lea ecx, var_1C
  loc_0042BD3A: call __vbaStrMove
  loc_0042BD3C: push eax
  loc_0042BD3D: push 004127D8h ; " takes on the value of "
  loc_0042BD42: call edi
  loc_0042BD44: mov edx, eax
  loc_0042BD46: lea ecx, var_20
  loc_0042BD49: call __vbaStrMove
  loc_0042BD4B: push eax
  loc_0042BD4C: mov eax, [0043002Ch]
  loc_0042BD51: push eax
  loc_0042BD52: call edi
  loc_0042BD54: mov edx, eax
  loc_0042BD56: lea ecx, var_24
  loc_0042BD59: call __vbaStrMove
  loc_0042BD5B: push eax
  loc_0042BD5C: push 0040DD3Ch ; "."
  loc_0042BD61: call edi
  loc_0042BD63: mov edx, eax
  loc_0042BD65: lea ecx, var_28
  loc_0042BD68: call __vbaStrMove
  loc_0042BD6A: mov ecx, ebx
  loc_0042BD6C: mov ebx, var_134
  loc_0042BD72: push eax
  loc_0042BD73: push ebx
  loc_0042BD74: call [ecx+0000019Ch]
  loc_0042BD7A: test eax, eax
  loc_0042BD7C: fnclex
  loc_0042BD7E: jge 0042BD92h
  loc_0042BD80: push 0000019Ch
  loc_0042BD85: push 0040E390h
  loc_0042BD8A: push ebx
  loc_0042BD8B: push eax
  loc_0042BD8C: call [00401030h] ; __vbaHresultCheckObj
  loc_0042BD92: lea edx, var_28
  loc_0042BD95: lea eax, var_24
  loc_0042BD98: push edx
  loc_0042BD99: lea ecx, var_20
  loc_0042BD9C: push eax
  loc_0042BD9D: lea edx, var_1C
  loc_0042BDA0: push ecx
  loc_0042BDA1: lea eax, var_18
  loc_0042BDA4: push edx
  loc_0042BDA5: lea ecx, var_14
  loc_0042BDA8: push eax
  loc_0042BDA9: push ecx
  loc_0042BDAA: push 00000006h
  loc_0042BDAC: call [004010E4h] ; __vbaFreeStrList
  loc_0042BDB2: add esp, 0000001Ch
  loc_0042BDB5: lea ecx, var_40
  loc_0042BDB8: call [00401114h] ; __vbaFreeObj
  loc_0042BDBE: mov eax, Me
  loc_0042BDC1: push eax
  loc_0042BDC2: mov edx, [eax]
  loc_0042BDC4: call [edx+0000034Ch]
  loc_0042BDCA: push eax
  loc_0042BDCB: lea eax, var_40
  loc_0042BDCE: push eax
  loc_0042BDCF: call [0040103Ch] ; __vbaObjSet
  loc_0042BDD5: mov ecx, [00430010h]
  loc_0042BDDB: mov ebx, [eax]
  loc_0042BDDD: push 0041280Ch ; "These values are assumed to follow a normal distribution around the mean of "
  loc_0042BDE2: push ecx
  loc_0042BDE3: mov var_134, eax
  loc_0042BDE9: call edi
  loc_0042BDEB: mov edx, eax
  loc_0042BDED: lea ecx, var_14
  loc_0042BDF0: call __vbaStrMove
  loc_0042BDF2: push eax
  loc_0042BDF3: push 0040DD3Ch ; "."
  loc_0042BDF8: call edi
  loc_0042BDFA: mov edx, eax
  loc_0042BDFC: lea ecx, var_18
  loc_0042BDFF: call __vbaStrMove
  loc_0042BE01: mov edx, ebx
  loc_0042BE03: mov ebx, var_134
  loc_0042BE09: push eax
  loc_0042BE0A: push ebx
  loc_0042BE0B: call [edx+0000014Ch]
  loc_0042BE11: test eax, eax
  loc_0042BE13: fnclex
  loc_0042BE15: jge 0042BE29h
  loc_0042BE17: push 0000014Ch
  loc_0042BE1C: push 0040E728h
  loc_0042BE21: push ebx
  loc_0042BE22: push eax
  loc_0042BE23: call [00401030h] ; __vbaHresultCheckObj
  loc_0042BE29: lea eax, var_18
  loc_0042BE2C: lea ecx, var_14
  loc_0042BE2F: push eax
  loc_0042BE30: push ecx
  loc_0042BE31: push 00000002h
  loc_0042BE33: call [004010E4h] ; __vbaFreeStrList
  loc_0042BE39: add esp, 0000000Ch
  loc_0042BE3C: lea ecx, var_40
  loc_0042BE3F: call [00401114h] ; __vbaFreeObj
  loc_0042BE45: mov eax, Me
  loc_0042BE48: push eax
  loc_0042BE49: mov edx, [eax]
  loc_0042BE4B: call [edx+00000348h]
  loc_0042BE51: push eax
  loc_0042BE52: lea eax, var_40
  loc_0042BE55: push eax
  loc_0042BE56: call [0040103Ch] ; __vbaObjSet
  loc_0042BE5C: mov ecx, [00430010h]
  loc_0042BE62: mov ebx, [eax]
  loc_0042BE64: push 004128ACh ; "The prediction of the actual value of "
  loc_0042BE69: push ecx
  loc_0042BE6A: mov var_134, eax
  loc_0042BE70: call edi
  loc_0042BE72: mov edx, eax
  loc_0042BE74: lea ecx, var_14
  loc_0042BE77: call __vbaStrMove
  loc_0042BE79: push eax
  loc_0042BE7A: push 00410684h ; " is "
  loc_0042BE7F: call edi
  loc_0042BE81: mov edx, eax
  loc_0042BE83: lea ecx, var_18
  loc_0042BE86: call __vbaStrMove
  loc_0042BE88: mov edx, [00430038h]
  loc_0042BE8E: push eax
  loc_0042BE8F: push edx
  loc_0042BE90: call [0040107Ch] ; __vbaStrR4
  loc_0042BE96: mov edx, eax
  loc_0042BE98: lea ecx, var_1C
  loc_0042BE9B: call __vbaStrMove
  loc_0042BE9D: push eax
  loc_0042BE9E: call edi
  loc_0042BEA0: mov edx, eax
  loc_0042BEA2: lea ecx, var_20
  loc_0042BEA5: call __vbaStrMove
  loc_0042BEA7: push eax
  loc_0042BEA8: push 0040F028h
  loc_0042BEAD: call edi
  loc_0042BEAF: mov edx, eax
  loc_0042BEB1: lea ecx, var_24
  loc_0042BEB4: call __vbaStrMove
  loc_0042BEB6: push eax
  loc_0042BEB7: mov eax, [00430014h]
  loc_0042BEBC: push eax
  loc_0042BEBD: call edi
  loc_0042BEBF: mov edx, eax
  loc_0042BEC1: lea ecx, var_28
  loc_0042BEC4: call __vbaStrMove
  loc_0042BEC6: push eax
  loc_0042BEC7: push 0040FF04h ; " when "
  loc_0042BECC: call edi
  loc_0042BECE: mov edx, eax
  loc_0042BED0: lea ecx, var_2C
  loc_0042BED3: call __vbaStrMove
  loc_0042BED5: mov ecx, [00430018h]
  loc_0042BEDB: push eax
  loc_0042BEDC: push ecx
  loc_0042BEDD: call edi
  loc_0042BEDF: mov edx, eax
  loc_0042BEE1: lea ecx, var_30
  loc_0042BEE4: call __vbaStrMove
  loc_0042BEE6: push eax
  loc_0042BEE7: push 00410824h ; " equals "
  loc_0042BEEC: call edi
  loc_0042BEEE: mov edx, eax
  loc_0042BEF0: lea ecx, var_34
  loc_0042BEF3: call __vbaStrMove
  loc_0042BEF5: mov edx, [0043002Ch]
  loc_0042BEFB: push eax
  loc_0042BEFC: push edx
  loc_0042BEFD: call edi
  loc_0042BEFF: mov edx, eax
  loc_0042BF01: lea ecx, var_38
  loc_0042BF04: call __vbaStrMove
  loc_0042BF06: push eax
  loc_0042BF07: push 0040DD3Ch ; "."
  loc_0042BF0C: call edi
  loc_0042BF0E: mov edx, eax
  loc_0042BF10: lea ecx, var_3C
  loc_0042BF13: call __vbaStrMove
  loc_0042BF15: mov var_168, ebx
  loc_0042BF1B: mov ebx, var_134
  loc_0042BF21: push eax
  loc_0042BF22: mov eax, var_168
  loc_0042BF28: push ebx
  loc_0042BF29: call [eax+0000019Ch]
  loc_0042BF2F: test eax, eax
  loc_0042BF31: fnclex
  loc_0042BF33: jge 0042BF47h
  loc_0042BF35: push 0000019Ch
  loc_0042BF3A: push 0040E390h
  loc_0042BF3F: push ebx
  loc_0042BF40: push eax
  loc_0042BF41: call [00401030h] ; __vbaHresultCheckObj
  loc_0042BF47: lea ecx, var_3C
  loc_0042BF4A: lea edx, var_38
  loc_0042BF4D: push ecx
  loc_0042BF4E: lea eax, var_34
  loc_0042BF51: push edx
  loc_0042BF52: lea ecx, var_30
  loc_0042BF55: push eax
  loc_0042BF56: lea edx, var_2C
  loc_0042BF59: push ecx
  loc_0042BF5A: lea eax, var_28
  loc_0042BF5D: push edx
  loc_0042BF5E: lea ecx, var_24
  loc_0042BF61: push eax
  loc_0042BF62: lea edx, var_20
  loc_0042BF65: push ecx
  loc_0042BF66: lea eax, var_1C
  loc_0042BF69: push edx
  loc_0042BF6A: lea ecx, var_18
  loc_0042BF6D: push eax
  loc_0042BF6E: lea edx, var_14
  loc_0042BF71: push ecx
  loc_0042BF72: push edx
  loc_0042BF73: push 0000000Bh
  loc_0042BF75: call [004010E4h] ; __vbaFreeStrList
  loc_0042BF7B: add esp, 00000030h
  loc_0042BF7E: lea ecx, var_40
  loc_0042BF81: call [00401114h] ; __vbaFreeObj
  loc_0042BF87: mov eax, Me
  loc_0042BF8A: push eax
  loc_0042BF8B: mov ecx, [eax]
  loc_0042BF8D: call [ecx+00000344h]
  loc_0042BF93: lea edx, var_40
  loc_0042BF96: push eax
  loc_0042BF97: push edx
  loc_0042BF98: call [0040103Ch] ; __vbaObjSet
  loc_0042BF9E: mov var_134, eax
  loc_0042BFA4: lea eax, var_D0
  loc_0042BFAA: push 00000002h
  loc_0042BFAC: lea ecx, var_50
  loc_0042BFAF: mov ebx, 00000008h
  loc_0042BFB4: push eax
  loc_0042BFB5: push ecx
  loc_0042BFB6: mov var_D8, 00412724h ; "These estimates fall on the line: "
  loc_0042BFC0: mov var_E0, ebx
  loc_0042BFC6: mov var_C8, 00430074h
  loc_0042BFD0: mov var_D0, 00004004h
  loc_0042BFDA: call [004010ACh] ; rtcRound
  loc_0042BFE0: lea edx, var_100
  loc_0042BFE6: push 00000002h
  loc_0042BFE8: lea eax, var_80
  loc_0042BFEB: push edx
  loc_0042BFEC: push eax
  loc_0042BFED: mov var_E8, 0041258Ch ; " + "
  loc_0042BFF7: mov var_F0, ebx
  loc_0042BFFD: mov var_F8, 00430060h
  loc_0042C007: mov var_100, 00004004h
  loc_0042C011: call [004010ACh] ; rtcRound
  loc_0042C017: mov ecx, [00430018h]
  loc_0042C01D: mov edx, var_134
  loc_0042C023: mov var_108, 00412770h ; " * Xg (the given value of "
  loc_0042C02D: mov var_110, ebx
  loc_0042C033: mov var_118, ecx
  loc_0042C039: mov var_120, ebx
  loc_0042C03F: mov var_128, 004127ACh ; ")."
  loc_0042C049: mov var_130, ebx
  loc_0042C04F: mov ebx, [edx]
  loc_0042C051: lea eax, var_E0
  loc_0042C057: lea ecx, var_50
  loc_0042C05A: push eax
  loc_0042C05B: lea edx, var_60
  loc_0042C05E: push ecx
  loc_0042C05F: push edx
  loc_0042C060: call [004010C0h] ; __vbaVarCat
  loc_0042C066: push eax
  loc_0042C067: lea eax, var_F0
  loc_0042C06D: lea ecx, var_70
  loc_0042C070: push eax
  loc_0042C071: push ecx
  loc_0042C072: call [004010C0h] ; __vbaVarCat
  loc_0042C078: push eax
  loc_0042C079: lea edx, var_80
  loc_0042C07C: lea eax, var_90
  loc_0042C082: push edx
  loc_0042C083: push eax
  loc_0042C084: call [004010C0h] ; __vbaVarCat
  loc_0042C08A: lea ecx, var_110
  loc_0042C090: push eax
  loc_0042C091: lea edx, var_A0
  loc_0042C097: push ecx
  loc_0042C098: push edx
  loc_0042C099: call [004010C0h] ; __vbaVarCat
  loc_0042C09F: push eax
  loc_0042C0A0: lea eax, var_120
  loc_0042C0A6: lea ecx, var_B0
  loc_0042C0AC: push eax
  loc_0042C0AD: push ecx
  loc_0042C0AE: call [004010C0h] ; __vbaVarCat
  loc_0042C0B4: push eax
  loc_0042C0B5: lea edx, var_130
  loc_0042C0BB: lea eax, var_C0
  loc_0042C0C1: push edx
  loc_0042C0C2: push eax
  loc_0042C0C3: call [004010C0h] ; __vbaVarCat
  loc_0042C0C9: lea ecx, var_14
  loc_0042C0CC: push eax
  loc_0042C0CD: push ecx
  loc_0042C0CE: call [004010B8h] ; __vbaStrVarVal
  loc_0042C0D4: mov edx, ebx
  loc_0042C0D6: mov ebx, var_134
  loc_0042C0DC: push eax
  loc_0042C0DD: push ebx
  loc_0042C0DE: call [edx+0000014Ch]
  loc_0042C0E4: test eax, eax
  loc_0042C0E6: fnclex
  loc_0042C0E8: jge 0042C0FCh
  loc_0042C0EA: push 0000014Ch
  loc_0042C0EF: push 0040E728h
  loc_0042C0F4: push ebx
  loc_0042C0F5: push eax
  loc_0042C0F6: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C0FC: lea ecx, var_14
  loc_0042C0FF: call [00401110h] ; __vbaFreeStr
  loc_0042C105: lea ecx, var_40
  loc_0042C108: call [00401114h] ; __vbaFreeObj
  loc_0042C10E: lea eax, var_C0
  loc_0042C114: lea ecx, var_B0
  loc_0042C11A: push eax
  loc_0042C11B: lea edx, var_A0
  loc_0042C121: push ecx
  loc_0042C122: lea eax, var_90
  loc_0042C128: push edx
  loc_0042C129: lea ecx, var_80
  loc_0042C12C: push eax
  loc_0042C12D: lea edx, var_70
  loc_0042C130: push ecx
  loc_0042C131: lea eax, var_60
  loc_0042C134: push edx
  loc_0042C135: lea ecx, var_50
  loc_0042C138: push eax
  loc_0042C139: push ecx
  loc_0042C13A: push 00000008h
  loc_0042C13C: call [00401018h] ; __vbaFreeVarList
  loc_0042C142: mov eax, Me
  loc_0042C145: add esp, 00000024h
  loc_0042C148: mov edx, [eax]
  loc_0042C14A: push eax
  loc_0042C14B: call [edx+00000328h]
  loc_0042C151: push eax
  loc_0042C152: lea eax, var_40
  loc_0042C155: push eax
  loc_0042C156: call [0040103Ch] ; __vbaObjSet
  loc_0042C15C: mov ecx, [00430010h]
  loc_0042C162: mov ebx, [eax]
  loc_0042C164: push 00412900h ; "The variance of "
  loc_0042C169: push ecx
  loc_0042C16A: mov var_134, eax
  loc_0042C170: call edi
  loc_0042C172: mov edx, eax
  loc_0042C174: lea ecx, var_14
  loc_0042C177: call __vbaStrMove
  loc_0042C179: push eax
  loc_0042C17A: push 0041263Ch ; " given "
  loc_0042C17F: call edi
  loc_0042C181: mov edx, eax
  loc_0042C183: lea ecx, var_18
  loc_0042C186: call __vbaStrMove
  loc_0042C188: mov edx, [00430018h]
  loc_0042C18E: push eax
  loc_0042C18F: push edx
  loc_0042C190: call edi
  loc_0042C192: mov edx, eax
  loc_0042C194: lea ecx, var_1C
  loc_0042C197: call __vbaStrMove
  loc_0042C199: push eax
  loc_0042C19A: push 004106CCh ; " = "
  loc_0042C19F: call edi
  loc_0042C1A1: mov edx, eax
  loc_0042C1A3: lea ecx, var_20
  loc_0042C1A6: call __vbaStrMove
  loc_0042C1A8: push eax
  loc_0042C1A9: mov eax, [0043002Ch]
  loc_0042C1AE: push eax
  loc_0042C1AF: call edi
  loc_0042C1B1: mov edx, eax
  loc_0042C1B3: lea ecx, var_24
  loc_0042C1B6: call __vbaStrMove
  loc_0042C1B8: push eax
  loc_0042C1B9: push 0040DD3Ch ; "."
  loc_0042C1BE: call edi
  loc_0042C1C0: mov edx, eax
  loc_0042C1C2: lea ecx, var_28
  loc_0042C1C5: call __vbaStrMove
  loc_0042C1C7: mov ecx, ebx
  loc_0042C1C9: mov ebx, var_134
  loc_0042C1CF: push eax
  loc_0042C1D0: push ebx
  loc_0042C1D1: call [ecx+0000019Ch]
  loc_0042C1D7: test eax, eax
  loc_0042C1D9: fnclex
  loc_0042C1DB: jge 0042C1EFh
  loc_0042C1DD: push 0000019Ch
  loc_0042C1E2: push 0040E390h
  loc_0042C1E7: push ebx
  loc_0042C1E8: push eax
  loc_0042C1E9: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C1EF: lea edx, var_28
  loc_0042C1F2: lea eax, var_24
  loc_0042C1F5: push edx
  loc_0042C1F6: lea ecx, var_20
  loc_0042C1F9: push eax
  loc_0042C1FA: lea edx, var_1C
  loc_0042C1FD: push ecx
  loc_0042C1FE: lea eax, var_18
  loc_0042C201: push edx
  loc_0042C202: lea ecx, var_14
  loc_0042C205: push eax
  loc_0042C206: push ecx
  loc_0042C207: push 00000006h
  loc_0042C209: call [004010E4h] ; __vbaFreeStrList
  loc_0042C20F: add esp, 0000001Ch
  loc_0042C212: lea ecx, var_40
  loc_0042C215: call [00401114h] ; __vbaFreeObj
  loc_0042C21B: mov eax, Me
  loc_0042C21E: push eax
  loc_0042C21F: mov edx, [eax]
  loc_0042C221: call [edx+00000324h]
  loc_0042C227: push eax
  loc_0042C228: lea eax, var_40
  loc_0042C22B: push eax
  loc_0042C22C: call [0040103Ch] ; __vbaObjSet
  loc_0042C232: mov ecx, [00430010h]
  loc_0042C238: mov ebx, [eax]
  loc_0042C23A: push 00412900h ; "The variance of "
  loc_0042C23F: push ecx
  loc_0042C240: mov var_134, eax
  loc_0042C246: call edi
  loc_0042C248: mov edx, eax
  loc_0042C24A: lea ecx, var_14
  loc_0042C24D: call __vbaStrMove
  loc_0042C24F: push eax
  loc_0042C250: push 00412928h ; " is assumed to be a constant for any value of "
  loc_0042C255: call edi
  loc_0042C257: mov edx, eax
  loc_0042C259: lea ecx, var_18
  loc_0042C25C: call __vbaStrMove
  loc_0042C25E: mov edx, [00430018h]
  loc_0042C264: push eax
  loc_0042C265: push edx
  loc_0042C266: call edi
  loc_0042C268: mov edx, eax
  loc_0042C26A: lea ecx, var_1C
  loc_0042C26D: call __vbaStrMove
  loc_0042C26F: push eax
  loc_0042C270: push 0040DD3Ch ; "."
  loc_0042C275: call edi
  loc_0042C277: mov edx, eax
  loc_0042C279: lea ecx, var_20
  loc_0042C27C: call __vbaStrMove
  loc_0042C27E: mov var_174, ebx
  loc_0042C284: mov ebx, var_134
  loc_0042C28A: push eax
  loc_0042C28B: mov eax, var_174
  loc_0042C291: push ebx
  loc_0042C292: call [eax+0000014Ch]
  loc_0042C298: test eax, eax
  loc_0042C29A: fnclex
  loc_0042C29C: jge 0042C2B0h
  loc_0042C29E: push 0000014Ch
  loc_0042C2A3: push 0040E728h
  loc_0042C2A8: push ebx
  loc_0042C2A9: push eax
  loc_0042C2AA: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C2B0: lea ecx, var_20
  loc_0042C2B3: lea edx, var_1C
  loc_0042C2B6: push ecx
  loc_0042C2B7: lea eax, var_18
  loc_0042C2BA: push edx
  loc_0042C2BB: lea ecx, var_14
  loc_0042C2BE: push eax
  loc_0042C2BF: push ecx
  loc_0042C2C0: push 00000004h
  loc_0042C2C2: call [004010E4h] ; __vbaFreeStrList
  loc_0042C2C8: add esp, 00000014h
  loc_0042C2CB: lea ecx, var_40
  loc_0042C2CE: call [00401114h] ; __vbaFreeObj
  loc_0042C2D4: mov eax, Me
  loc_0042C2D7: push eax
  loc_0042C2D8: mov edx, [eax]
  loc_0042C2DA: call [edx+00000320h]
  loc_0042C2E0: push eax
  loc_0042C2E1: lea eax, var_40
  loc_0042C2E4: push eax
  loc_0042C2E5: call [0040103Ch] ; __vbaObjSet
  loc_0042C2EB: mov ecx, [00430010h]
  loc_0042C2F1: mov ebx, [eax]
  loc_0042C2F3: push 0041298Ch ; "The estimate of the variance of "
  loc_0042C2F8: push ecx
  loc_0042C2F9: mov var_134, eax
  loc_0042C2FF: call edi
  loc_0042C301: mov edx, eax
  loc_0042C303: lea ecx, var_14
  loc_0042C306: call __vbaStrMove
  loc_0042C308: push eax
  loc_0042C309: push 0041263Ch ; " given "
  loc_0042C30E: call edi
  loc_0042C310: mov edx, eax
  loc_0042C312: lea ecx, var_18
  loc_0042C315: call __vbaStrMove
  loc_0042C317: mov edx, [00430018h]
  loc_0042C31D: push eax
  loc_0042C31E: push edx
  loc_0042C31F: call edi
  loc_0042C321: mov edx, eax
  loc_0042C323: lea ecx, var_1C
  loc_0042C326: call __vbaStrMove
  loc_0042C328: push eax
  loc_0042C329: push 004106CCh ; " = "
  loc_0042C32E: call edi
  loc_0042C330: mov edx, eax
  loc_0042C332: lea ecx, var_20
  loc_0042C335: call __vbaStrMove
  loc_0042C337: push eax
  loc_0042C338: mov eax, [0043002Ch]
  loc_0042C33D: push eax
  loc_0042C33E: call edi
  loc_0042C340: mov edx, eax
  loc_0042C342: lea ecx, var_24
  loc_0042C345: call __vbaStrMove
  loc_0042C347: push eax
  loc_0042C348: push 0040DD3Ch ; "."
  loc_0042C34D: call edi
  loc_0042C34F: mov edx, eax
  loc_0042C351: lea ecx, var_28
  loc_0042C354: call __vbaStrMove
  loc_0042C356: mov ecx, ebx
  loc_0042C358: mov ebx, var_134
  loc_0042C35E: push eax
  loc_0042C35F: push ebx
  loc_0042C360: call [ecx+0000019Ch]
  loc_0042C366: test eax, eax
  loc_0042C368: fnclex
  loc_0042C36A: jge 0042C37Eh
  loc_0042C36C: push 0000019Ch
  loc_0042C371: push 0040E390h
  loc_0042C376: push ebx
  loc_0042C377: push eax
  loc_0042C378: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C37E: lea edx, var_28
  loc_0042C381: lea eax, var_24
  loc_0042C384: push edx
  loc_0042C385: lea ecx, var_20
  loc_0042C388: push eax
  loc_0042C389: lea edx, var_1C
  loc_0042C38C: push ecx
  loc_0042C38D: lea eax, var_18
  loc_0042C390: push edx
  loc_0042C391: lea ecx, var_14
  loc_0042C394: push eax
  loc_0042C395: push ecx
  loc_0042C396: push 00000006h
  loc_0042C398: call [004010E4h] ; __vbaFreeStrList
  loc_0042C39E: add esp, 0000001Ch
  loc_0042C3A1: lea ecx, var_40
  loc_0042C3A4: call [00401114h] ; __vbaFreeObj
  loc_0042C3AA: fld real4 ptr [0043006Ch]
  loc_0042C3B0: fcomp real4 ptr [004011E8h]
  loc_0042C3B6: fnstsw ax
  loc_0042C3B8: test ah, 01h
  loc_0042C3BB: mov eax, Me
  loc_0042C3BE: jnz 0042C49Ah
  loc_0042C3C4: mov edx, [eax]
  loc_0042C3C6: push eax
  loc_0042C3C7: call [edx+00000330h]
  loc_0042C3CD: push eax
  loc_0042C3CE: lea eax, var_40
  loc_0042C3D1: push eax
  loc_0042C3D2: call [0040103Ch] ; __vbaObjSet
  loc_0042C3D8: mov ecx, [0043006Ch]
  loc_0042C3DE: mov ebx, [eax]
  loc_0042C3E0: push 004129D4h ; "There is a "
  loc_0042C3E5: push ecx
  loc_0042C3E6: mov var_134, eax
  loc_0042C3EC: call [0040107Ch] ; __vbaStrR4
  loc_0042C3F2: mov edx, eax
  loc_0042C3F4: lea ecx, var_14
  loc_0042C3F7: call __vbaStrMove
  loc_0042C3F9: push eax
  loc_0042C3FA: call edi
  loc_0042C3FC: mov edx, eax
  loc_0042C3FE: lea ecx, var_18
  loc_0042C401: call __vbaStrMove
  loc_0042C403: push eax
  loc_0042C404: push 004129F0h ; " standard deviation increase in the estimate of the mean of "
  loc_0042C409: call edi
  loc_0042C40B: mov edx, eax
  loc_0042C40D: lea ecx, var_1C
  loc_0042C410: call __vbaStrMove
  loc_0042C412: mov edx, [00430010h]
  loc_0042C418: push eax
  loc_0042C419: push edx
  loc_0042C41A: call edi
  loc_0042C41C: mov edx, eax
  loc_0042C41E: lea ecx, var_20
  loc_0042C421: call __vbaStrMove
  loc_0042C423: push eax
  loc_0042C424: push 00412A70h ; " for each one standard deviation increase in "
  loc_0042C429: call edi
  loc_0042C42B: mov edx, eax
  loc_0042C42D: lea ecx, var_24
  loc_0042C430: call __vbaStrMove
  loc_0042C432: push eax
  loc_0042C433: mov eax, [00430018h]
  loc_0042C438: push eax
  loc_0042C439: call edi
  loc_0042C43B: mov edx, eax
  loc_0042C43D: lea ecx, var_28
  loc_0042C440: call __vbaStrMove
  loc_0042C442: push eax
  loc_0042C443: push 0040DD3Ch ; "."
  loc_0042C448: call edi
  loc_0042C44A: mov edx, eax
  loc_0042C44C: lea ecx, var_2C
  loc_0042C44F: call __vbaStrMove
  loc_0042C451: mov ecx, ebx
  loc_0042C453: mov ebx, var_134
  loc_0042C459: push eax
  loc_0042C45A: push ebx
  loc_0042C45B: call [ecx+0000019Ch]
  loc_0042C461: test eax, eax
  loc_0042C463: fnclex
  loc_0042C465: jge 0042C479h
  loc_0042C467: push 0000019Ch
  loc_0042C46C: push 0040E390h
  loc_0042C471: push ebx
  loc_0042C472: push eax
  loc_0042C473: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C479: lea edx, var_2C
  loc_0042C47C: lea eax, var_28
  loc_0042C47F: push edx
  loc_0042C480: lea ecx, var_24
  loc_0042C483: push eax
  loc_0042C484: lea edx, var_20
  loc_0042C487: push ecx
  loc_0042C488: lea eax, var_1C
  loc_0042C48B: push edx
  loc_0042C48C: lea ecx, var_18
  loc_0042C48F: push eax
  loc_0042C490: lea edx, var_14
  loc_0042C493: push ecx
  loc_0042C494: push edx
  loc_0042C495: jmp 0042C586h
  loc_0042C49A: mov ecx, [eax]
  loc_0042C49C: push eax
  loc_0042C49D: call [ecx+00000330h]
  loc_0042C4A3: lea edx, var_40
  loc_0042C4A6: push eax
  loc_0042C4A7: push edx
  loc_0042C4A8: call [0040103Ch] ; __vbaObjSet
  loc_0042C4AE: fld real4 ptr [0043006Ch]
  loc_0042C4B4: mov ebx, [eax]
  loc_0042C4B6: mov var_134, eax
  loc_0042C4BC: fabs
  loc_0042C4BE: fnstsw ax
  loc_0042C4C0: test al, 0Dh
  loc_0042C4C2: jnz 0042C960h
  loc_0042C4C8: fstp real4 ptr var_180
  loc_0042C4CE: fld real4 ptr var_180
  loc_0042C4D4: push 004129D4h ; "There is a "
  loc_0042C4D9: push ecx
  loc_0042C4DA: fstp real4 ptr [esp]
  loc_0042C4DD: call [0040107Ch] ; __vbaStrR4
  loc_0042C4E3: mov edx, eax
  loc_0042C4E5: lea ecx, var_14
  loc_0042C4E8: call __vbaStrMove
  loc_0042C4EA: push eax
  loc_0042C4EB: call edi
  loc_0042C4ED: mov edx, eax
  loc_0042C4EF: lea ecx, var_18
  loc_0042C4F2: call __vbaStrMove
  loc_0042C4F4: push eax
  loc_0042C4F5: push 00412AE0h ; " standard deviation decrease in the estimate of the mean of "
  loc_0042C4FA: call edi
  loc_0042C4FC: mov edx, eax
  loc_0042C4FE: lea ecx, var_1C
  loc_0042C501: call __vbaStrMove
  loc_0042C503: push eax
  loc_0042C504: mov eax, [00430010h]
  loc_0042C509: push eax
  loc_0042C50A: call edi
  loc_0042C50C: mov edx, eax
  loc_0042C50E: lea ecx, var_20
  loc_0042C511: call __vbaStrMove
  loc_0042C513: push eax
  loc_0042C514: push 00412A70h ; " for each one standard deviation increase in "
  loc_0042C519: call edi
  loc_0042C51B: mov edx, eax
  loc_0042C51D: lea ecx, var_24
  loc_0042C520: call __vbaStrMove
  loc_0042C522: mov ecx, [00430018h]
  loc_0042C528: push eax
  loc_0042C529: push ecx
  loc_0042C52A: call edi
  loc_0042C52C: mov edx, eax
  loc_0042C52E: lea ecx, var_28
  loc_0042C531: call __vbaStrMove
  loc_0042C533: push eax
  loc_0042C534: push 0040DD3Ch ; "."
  loc_0042C539: call edi
  loc_0042C53B: mov edx, eax
  loc_0042C53D: lea ecx, var_2C
  loc_0042C540: call __vbaStrMove
  loc_0042C542: mov edx, ebx
  loc_0042C544: mov ebx, var_134
  loc_0042C54A: push eax
  loc_0042C54B: push ebx
  loc_0042C54C: call [edx+0000019Ch]
  loc_0042C552: test eax, eax
  loc_0042C554: fnclex
  loc_0042C556: jge 0042C56Ah
  loc_0042C558: push 0000019Ch
  loc_0042C55D: push 0040E390h
  loc_0042C562: push ebx
  loc_0042C563: push eax
  loc_0042C564: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C56A: lea eax, var_2C
  loc_0042C56D: lea ecx, var_28
  loc_0042C570: push eax
  loc_0042C571: lea edx, var_24
  loc_0042C574: push ecx
  loc_0042C575: lea eax, var_20
  loc_0042C578: push edx
  loc_0042C579: lea ecx, var_1C
  loc_0042C57C: push eax
  loc_0042C57D: lea edx, var_18
  loc_0042C580: push ecx
  loc_0042C581: lea eax, var_14
  loc_0042C584: push edx
  loc_0042C585: push eax
  loc_0042C586: push 00000007h
  loc_0042C588: call [004010E4h] ; __vbaFreeStrList
  loc_0042C58E: add esp, 00000020h
  loc_0042C591: lea ecx, var_40
  loc_0042C594: call [00401114h] ; __vbaFreeObj
  loc_0042C59A: mov eax, Me
  loc_0042C59D: push eax
  loc_0042C59E: mov ecx, [eax]
  loc_0042C5A0: call [ecx+00000318h]
  loc_0042C5A6: lea edx, var_40
  loc_0042C5A9: push eax
  loc_0042C5AA: push edx
  loc_0042C5AB: call [0040103Ch] ; __vbaObjSet
  loc_0042C5B1: fld real4 ptr [00430070h]
  loc_0042C5B7: fmul st0, real4 ptr [00401780h]
  loc_0042C5BD: mov ebx, [eax]
  loc_0042C5BF: mov var_134, eax
  loc_0042C5C5: push ecx
  loc_0042C5C6: fnstsw ax
  loc_0042C5C8: test al, 0Dh
  loc_0042C5CA: jnz 0042C960h
  loc_0042C5D0: fstp real4 ptr [esp]
  loc_0042C5D3: call [0040107Ch] ; __vbaStrR4
  loc_0042C5D9: mov edx, eax
  loc_0042C5DB: lea ecx, var_14
  loc_0042C5DE: call __vbaStrMove
  loc_0042C5E0: push eax
  loc_0042C5E1: push 00412B60h ; " percent of the sample variation in "
  loc_0042C5E6: call edi
  loc_0042C5E8: mov edx, eax
  loc_0042C5EA: lea ecx, var_18
  loc_0042C5ED: call __vbaStrMove
  loc_0042C5EF: push eax
  loc_0042C5F0: mov eax, [00430010h]
  loc_0042C5F5: push eax
  loc_0042C5F6: call edi
  loc_0042C5F8: mov edx, eax
  loc_0042C5FA: lea ecx, var_1C
  loc_0042C5FD: call __vbaStrMove
  loc_0042C5FF: push eax
  loc_0042C600: push 00412BB0h ; " is associated with changes in "
  loc_0042C605: call edi
  loc_0042C607: mov edx, eax
  loc_0042C609: lea ecx, var_20
  loc_0042C60C: call __vbaStrMove
  loc_0042C60E: mov ecx, [00430018h]
  loc_0042C614: push eax
  loc_0042C615: push ecx
  loc_0042C616: call edi
  loc_0042C618: mov edx, eax
  loc_0042C61A: lea ecx, var_24
  loc_0042C61D: call __vbaStrMove
  loc_0042C61F: push eax
  loc_0042C620: push 00412BF4h ; " through the model."
  loc_0042C625: call edi
  loc_0042C627: mov edx, eax
  loc_0042C629: lea ecx, var_28
  loc_0042C62C: call __vbaStrMove
  loc_0042C62E: mov edx, ebx
  loc_0042C630: mov ebx, var_134
  loc_0042C636: push eax
  loc_0042C637: push ebx
  loc_0042C638: call [edx+0000019Ch]
  loc_0042C63E: test eax, eax
  loc_0042C640: fnclex
  loc_0042C642: jge 0042C656h
  loc_0042C644: push 0000019Ch
  loc_0042C649: push 0040E390h
  loc_0042C64E: push ebx
  loc_0042C64F: push eax
  loc_0042C650: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C656: lea eax, var_28
  loc_0042C659: lea ecx, var_24
  loc_0042C65C: push eax
  loc_0042C65D: lea edx, var_20
  loc_0042C660: push ecx
  loc_0042C661: lea eax, var_1C
  loc_0042C664: push edx
  loc_0042C665: lea ecx, var_18
  loc_0042C668: push eax
  loc_0042C669: lea edx, var_14
  loc_0042C66C: push ecx
  loc_0042C66D: push edx
  loc_0042C66E: push 00000006h
  loc_0042C670: call [004010E4h] ; __vbaFreeStrList
  loc_0042C676: add esp, 0000001Ch
  loc_0042C679: lea ecx, var_40
  loc_0042C67C: call [00401114h] ; __vbaFreeObj
  loc_0042C682: mov eax, Me
  loc_0042C685: push eax
  loc_0042C686: mov ecx, [eax]
  loc_0042C688: call [ecx+00000310h]
  loc_0042C68E: lea edx, var_40
  loc_0042C691: push eax
  loc_0042C692: push edx
  loc_0042C693: call [0040103Ch] ; __vbaObjSet
  loc_0042C699: mov ebx, [eax]
  loc_0042C69B: mov var_134, eax
  loc_0042C6A1: mov eax, [00430010h]
  loc_0042C6A6: push 00412C20h ; " The difference between "
  loc_0042C6AB: push eax
  loc_0042C6AC: call edi
  loc_0042C6AE: mov edx, eax
  loc_0042C6B0: lea ecx, var_14
  loc_0042C6B3: call __vbaStrMove
  loc_0042C6B5: push eax
  loc_0042C6B6: push 00412C58h ; " and the mean "
  loc_0042C6BB: call edi
  loc_0042C6BD: mov edx, eax
  loc_0042C6BF: lea ecx, var_18
  loc_0042C6C2: call __vbaStrMove
  loc_0042C6C4: mov ecx, [00430010h]
  loc_0042C6CA: push eax
  loc_0042C6CB: push ecx
  loc_0042C6CC: call edi
  loc_0042C6CE: mov edx, eax
  loc_0042C6D0: lea ecx, var_1C
  loc_0042C6D3: call __vbaStrMove
  loc_0042C6D5: push eax
  loc_0042C6D6: push 0040FF04h ; " when "
  loc_0042C6DB: call edi
  loc_0042C6DD: mov edx, eax
  loc_0042C6DF: lea ecx, var_20
  loc_0042C6E2: call __vbaStrMove
  loc_0042C6E4: mov edx, [00430018h]
  loc_0042C6EA: push eax
  loc_0042C6EB: push edx
  loc_0042C6EC: call edi
  loc_0042C6EE: mov edx, eax
  loc_0042C6F0: lea ecx, var_24
  loc_0042C6F3: call __vbaStrMove
  loc_0042C6F5: push eax
  loc_0042C6F6: push 004106CCh ; " = "
  loc_0042C6FB: call edi
  loc_0042C6FD: mov edx, eax
  loc_0042C6FF: lea ecx, var_28
  loc_0042C702: call __vbaStrMove
  loc_0042C704: push eax
  loc_0042C705: mov eax, [0043002Ch]
  loc_0042C70A: push eax
  loc_0042C70B: call edi
  loc_0042C70D: mov edx, eax
  loc_0042C70F: lea ecx, var_2C
  loc_0042C712: call __vbaStrMove
  loc_0042C714: push eax
  loc_0042C715: push 0040F028h
  loc_0042C71A: call edi
  loc_0042C71C: mov edx, eax
  loc_0042C71E: lea ecx, var_30
  loc_0042C721: call __vbaStrMove
  loc_0042C723: mov ecx, [0043001Ch]
  loc_0042C729: push eax
  loc_0042C72A: push ecx
  loc_0042C72B: call edi
  loc_0042C72D: mov edx, eax
  loc_0042C72F: lea ecx, var_34
  loc_0042C732: call __vbaStrMove
  loc_0042C734: push eax
  loc_0042C735: push 0040DD3Ch ; "."
  loc_0042C73A: call edi
  loc_0042C73C: mov edx, eax
  loc_0042C73E: lea ecx, var_38
  loc_0042C741: call __vbaStrMove
  loc_0042C743: mov edx, ebx
  loc_0042C745: mov ebx, var_134
  loc_0042C74B: push eax
  loc_0042C74C: push ebx
  loc_0042C74D: call [edx+0000019Ch]
  loc_0042C753: test eax, eax
  loc_0042C755: fnclex
  loc_0042C757: jge 0042C76Bh
  loc_0042C759: push 0000019Ch
  loc_0042C75E: push 0040E390h
  loc_0042C763: push ebx
  loc_0042C764: push eax
  loc_0042C765: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C76B: lea eax, var_38
  loc_0042C76E: lea ecx, var_34
  loc_0042C771: push eax
  loc_0042C772: lea edx, var_30
  loc_0042C775: push ecx
  loc_0042C776: lea eax, var_2C
  loc_0042C779: push edx
  loc_0042C77A: lea ecx, var_28
  loc_0042C77D: push eax
  loc_0042C77E: lea edx, var_24
  loc_0042C781: push ecx
  loc_0042C782: lea eax, var_20
  loc_0042C785: push edx
  loc_0042C786: lea ecx, var_1C
  loc_0042C789: push eax
  loc_0042C78A: lea edx, var_18
  loc_0042C78D: push ecx
  loc_0042C78E: lea eax, var_14
  loc_0042C791: push edx
  loc_0042C792: push eax
  loc_0042C793: push 0000000Ah
  loc_0042C795: call [004010E4h] ; __vbaFreeStrList
  loc_0042C79B: add esp, 0000002Ch
  loc_0042C79E: lea ecx, var_40
  loc_0042C7A1: call [00401114h] ; __vbaFreeObj
  loc_0042C7A7: mov eax, Me
  loc_0042C7AA: push eax
  loc_0042C7AB: mov ecx, [eax]
  loc_0042C7AD: call [ecx+00000308h]
  loc_0042C7B3: lea edx, var_40
  loc_0042C7B6: push eax
  loc_0042C7B7: push edx
  loc_0042C7B8: call [0040103Ch] ; __vbaObjSet
  loc_0042C7BE: mov ebx, [eax]
  loc_0042C7C0: mov var_134, eax
  loc_0042C7C6: mov eax, [00430010h]
  loc_0042C7CB: push 00412C20h ; " The difference between "
  loc_0042C7D0: push eax
  loc_0042C7D1: call edi
  loc_0042C7D3: mov edx, eax
  loc_0042C7D5: lea ecx, var_14
  loc_0042C7D8: call __vbaStrMove
  loc_0042C7DA: push eax
  loc_0042C7DB: push 00412C7Ch ; " and the estimated mean "
  loc_0042C7E0: call edi
  loc_0042C7E2: mov edx, eax
  loc_0042C7E4: lea ecx, var_18
  loc_0042C7E7: call __vbaStrMove
  loc_0042C7E9: mov ecx, [00430010h]
  loc_0042C7EF: push eax
  loc_0042C7F0: push ecx
  loc_0042C7F1: call edi
  loc_0042C7F3: mov edx, eax
  loc_0042C7F5: lea ecx, var_1C
  loc_0042C7F8: call __vbaStrMove
  loc_0042C7FA: push eax
  loc_0042C7FB: push 0040FF04h ; " when "
  loc_0042C800: call edi
  loc_0042C802: mov edx, eax
  loc_0042C804: lea ecx, var_20
  loc_0042C807: call __vbaStrMove
  loc_0042C809: mov edx, [00430018h]
  loc_0042C80F: push eax
  loc_0042C810: push edx
  loc_0042C811: call edi
  loc_0042C813: mov edx, eax
  loc_0042C815: lea ecx, var_24
  loc_0042C818: call __vbaStrMove
  loc_0042C81A: push eax
  loc_0042C81B: push 004106CCh ; " = "
  loc_0042C820: call edi
  loc_0042C822: mov edx, eax
  loc_0042C824: lea ecx, var_28
  loc_0042C827: call __vbaStrMove
  loc_0042C829: push eax
  loc_0042C82A: mov eax, [0043002Ch]
  loc_0042C82F: push eax
  loc_0042C830: call edi
  loc_0042C832: mov edx, eax
  loc_0042C834: lea ecx, var_2C
  loc_0042C837: call __vbaStrMove
  loc_0042C839: push eax
  loc_0042C83A: push 0040F028h
  loc_0042C83F: call edi
  loc_0042C841: mov edx, eax
  loc_0042C843: lea ecx, var_30
  loc_0042C846: call __vbaStrMove
  loc_0042C848: mov ecx, [0043001Ch]
  loc_0042C84E: push eax
  loc_0042C84F: push ecx
  loc_0042C850: call edi
  loc_0042C852: mov edx, eax
  loc_0042C854: lea ecx, var_34
  loc_0042C857: call __vbaStrMove
  loc_0042C859: push eax
  loc_0042C85A: push 0040DD3Ch ; "."
  loc_0042C85F: call edi
  loc_0042C861: mov edx, eax
  loc_0042C863: lea ecx, var_38
  loc_0042C866: call __vbaStrMove
  loc_0042C868: mov esi, var_134
  loc_0042C86E: push eax
  loc_0042C86F: push esi
  loc_0042C870: call [ebx+0000019Ch]
  loc_0042C876: test eax, eax
  loc_0042C878: fnclex
  loc_0042C87A: jge 0042C88Eh
  loc_0042C87C: push 0000019Ch
  loc_0042C881: push 0040E390h
  loc_0042C886: push esi
  loc_0042C887: push eax
  loc_0042C888: call [00401030h] ; __vbaHresultCheckObj
  loc_0042C88E: lea edx, var_38
  loc_0042C891: lea eax, var_34
  loc_0042C894: push edx
  loc_0042C895: lea ecx, var_30
  loc_0042C898: push eax
  loc_0042C899: lea edx, var_2C
  loc_0042C89C: push ecx
  loc_0042C89D: lea eax, var_28
  loc_0042C8A0: push edx
  loc_0042C8A1: lea ecx, var_24
  loc_0042C8A4: push eax
  loc_0042C8A5: lea edx, var_20
  loc_0042C8A8: push ecx
  loc_0042C8A9: lea eax, var_1C
  loc_0042C8AC: push edx
  loc_0042C8AD: lea ecx, var_18
  loc_0042C8B0: push eax
  loc_0042C8B1: lea edx, var_14
  loc_0042C8B4: push ecx
  loc_0042C8B5: push edx
  loc_0042C8B6: push 0000000Ah
  loc_0042C8B8: call [004010E4h] ; __vbaFreeStrList
  loc_0042C8BE: add esp, 0000002Ch
  loc_0042C8C1: lea ecx, var_40
  loc_0042C8C4: call [00401114h] ; __vbaFreeObj
  loc_0042C8CA: fwait
  loc_0042C8CB: push 0042C94Bh
  loc_0042C8D0: jmp 0042C94Ah
  loc_0042C8D2: lea eax, var_3C
  loc_0042C8D5: lea ecx, var_38
  loc_0042C8D8: push eax
  loc_0042C8D9: lea edx, var_34
  loc_0042C8DC: push ecx
  loc_0042C8DD: lea eax, var_30
  loc_0042C8E0: push edx
  loc_0042C8E1: lea ecx, var_2C
  loc_0042C8E4: push eax
  loc_0042C8E5: lea edx, var_28
  loc_0042C8E8: push ecx
  loc_0042C8E9: lea eax, var_24
  loc_0042C8EC: push edx
  loc_0042C8ED: lea ecx, var_20
  loc_0042C8F0: push eax
  loc_0042C8F1: lea edx, var_1C
  loc_0042C8F4: push ecx
  loc_0042C8F5: lea eax, var_18
  loc_0042C8F8: push edx
  loc_0042C8F9: lea ecx, var_14
  loc_0042C8FC: push eax
  loc_0042C8FD: push ecx
  loc_0042C8FE: push 0000000Bh
  loc_0042C900: call [004010E4h] ; __vbaFreeStrList
  loc_0042C906: add esp, 00000030h
  loc_0042C909: lea ecx, var_40
  loc_0042C90C: call [00401114h] ; __vbaFreeObj
  loc_0042C912: lea edx, var_C0
  loc_0042C918: lea eax, var_B0
  loc_0042C91E: push edx
  loc_0042C91F: lea ecx, var_A0
  loc_0042C925: push eax
  loc_0042C926: lea edx, var_90
  loc_0042C92C: push ecx
  loc_0042C92D: lea eax, var_80
  loc_0042C930: push edx
  loc_0042C931: lea ecx, var_70
  loc_0042C934: push eax
  loc_0042C935: lea edx, var_60
  loc_0042C938: push ecx
  loc_0042C939: lea eax, var_50
  loc_0042C93C: push edx
  loc_0042C93D: push eax
  loc_0042C93E: push 00000008h
  loc_0042C940: call [00401018h] ; __vbaFreeVarList
  loc_0042C946: add esp, 00000024h
  loc_0042C949: ret
  loc_0042C94A: ret
  loc_0042C94B: mov ecx, var_10
  loc_0042C94E: pop edi
  loc_0042C94F: pop esi
  loc_0042C950: xor eax, eax
  loc_0042C952: mov fs:[00000000h], ecx
  loc_0042C959: pop ebx
  loc_0042C95A: mov esp, ebp
  loc_0042C95C: pop ebp
  loc_0042C95D: retn 0004h
End Sub

Private Function Proc_15_18_42C970(arg_C) '42C970
  loc_0042C970: mov eax, [0043009Ch]
  loc_0042C975: sub esp, 00000010h
  loc_0042C978: test eax, eax
  loc_0042C97A: push ebx
  loc_0042C97B: push ebp
  loc_0042C97C: push esi
  loc_0042C97D: push edi
  loc_0042C97E: jnz 0042C990h
  loc_0042C980: push 0043009Ch
  loc_0042C985: push 00405FC0h
  loc_0042C98A: call [004010D4h] ; __vbaNew2
  loc_0042C990: sub esp, 00000010h
  loc_0042C993: mov ecx, 0000000Ah
  loc_0042C998: mov ebp, esp
  loc_0042C99A: mov edi, ecx
  loc_0042C99C: mov eax, 80020004h
  loc_0042C9A1: sub esp, 00000010h
  loc_0042C9A4: mov [ebp], ecx
  loc_0042C9A7: mov ecx, arg_2C
  loc_0042C9AB: mov esi, [0043009Ch]
  loc_0042C9B1: mov edx, eax
  loc_0042C9B3: mov Proc_15_18_42C970, ecx
  loc_0042C9B6: mov ecx, esp
  loc_0042C9B8: mov ebx, [esi]
  loc_0042C9BA: push esi
  loc_0042C9BB: mov Me, eax
  loc_0042C9BE: mov eax, arg_38
  loc_0042C9C2: mov arg_C, eax
  loc_0042C9C5: mov eax, arg_30
  loc_0042C9C9: mov [ecx], edi
  loc_0042C9CB: mov [ecx+00000004h], eax
  loc_0042C9CE: mov [ecx+00000008h], edx
  loc_0042C9D1: mov edx, arg_38
  loc_0042C9D5: mov [ecx+0000000Ch], edx
  loc_0042C9D8: call [ebx+000002B0h]
  loc_0042C9DE: mov edi, [00401030h] ; __vbaHresultCheckObj
  loc_0042C9E4: test eax, eax
  loc_0042C9E6: fnclex
  loc_0042C9E8: jge 0042C9F8h
  loc_0042C9EA: push 000002B0h
  loc_0042C9EF: push 0040DDE0h
  loc_0042C9F4: push esi
  loc_0042C9F5: push eax
  loc_0042C9F6: call edi
  loc_0042C9F8: mov eax, [004300ECh]
  loc_0042C9FD: test eax, eax
  loc_0042C9FF: jnz 0042CA11h
  loc_0042CA01: push 004300ECh
  loc_0042CA06: push 0040A624h
  loc_0042CA0B: call [004010D4h] ; __vbaNew2
  loc_0042CA11: mov esi, [004300ECh]
  loc_0042CA17: push esi
  loc_0042CA18: mov eax, [esi]
  loc_0042CA1A: call [eax+000002B4h]
  loc_0042CA20: test eax, eax
  loc_0042CA22: fnclex
  loc_0042CA24: jge 0042CA34h
  loc_0042CA26: push 000002B4h
  loc_0042CA2B: push 0040ECECh
  loc_0042CA30: push esi
  loc_0042CA31: push eax
  loc_0042CA32: call edi
  loc_0042CA34: pop edi
  loc_0042CA35: pop esi
  loc_0042CA36: pop ebp
  loc_0042CA37: xor eax, eax
  loc_0042CA39: pop ebx
  loc_0042CA3A: add esp, 00000010h
  loc_0042CA3D: retn 0004h
End Function

Private Function Proc_15_19_42CE40(arg_C) '42CE40
  loc_0042CE40: mov eax, [00430128h]
  loc_0042CE45: sub esp, 00000010h
  loc_0042CE48: test eax, eax
  loc_0042CE4A: push ebx
  loc_0042CE4B: push ebp
  loc_0042CE4C: push esi
  loc_0042CE4D: push edi
  loc_0042CE4E: jnz 0042CE60h
  loc_0042CE50: push 00430128h
  loc_0042CE55: push 00404AE0h
  loc_0042CE5A: call [004010D4h] ; __vbaNew2
  loc_0042CE60: sub esp, 00000010h
  loc_0042CE63: mov ecx, 0000000Ah
  loc_0042CE68: mov ebp, esp
  loc_0042CE6A: mov edi, ecx
  loc_0042CE6C: mov eax, 80020004h
  loc_0042CE71: sub esp, 00000010h
  loc_0042CE74: mov [ebp], ecx
  loc_0042CE77: mov ecx, arg_2C
  loc_0042CE7B: mov esi, [00430128h]
  loc_0042CE81: mov edx, eax
  loc_0042CE83: mov Proc_15_19_42CE40, ecx
  loc_0042CE86: mov ecx, esp
  loc_0042CE88: mov ebx, [esi]
  loc_0042CE8A: push esi
  loc_0042CE8B: mov Me, eax
  loc_0042CE8E: mov eax, arg_38
  loc_0042CE92: mov arg_C, eax
  loc_0042CE95: mov eax, arg_30
  loc_0042CE99: mov [ecx], edi
  loc_0042CE9B: mov [ecx+00000004h], eax
  loc_0042CE9E: mov [ecx+00000008h], edx
  loc_0042CEA1: mov edx, arg_38
  loc_0042CEA5: mov [ecx+0000000Ch], edx
  loc_0042CEA8: call [ebx+000002B0h]
  loc_0042CEAE: test eax, eax
  loc_0042CEB0: fnclex
  loc_0042CEB2: jge 0042CEC6h
  loc_0042CEB4: push 000002B0h
  loc_0042CEB9: push 0040F8B4h
  loc_0042CEBE: push esi
  loc_0042CEBF: push eax
  loc_0042CEC0: call [00401030h] ; __vbaHresultCheckObj
  loc_0042CEC6: pop edi
  loc_0042CEC7: pop esi
  loc_0042CEC8: pop ebp
  loc_0042CEC9: xor eax, eax
  loc_0042CECB: pop ebx
  loc_0042CECC: add esp, 00000010h
  loc_0042CECF: retn 0004h
End Function

Private Function Proc_15_20_42D460(arg_C) '42D460
  loc_0042D460: mov eax, [00430088h]
  loc_0042D465: sub esp, 00000010h
  loc_0042D468: test eax, eax
  loc_0042D46A: push ebx
  loc_0042D46B: push ebp
  loc_0042D46C: push esi
  loc_0042D46D: push edi
  loc_0042D46E: jnz 0042D480h
  loc_0042D470: push 00430088h
  loc_0042D475: push 004058C0h
  loc_0042D47A: call [004010D4h] ; __vbaNew2
  loc_0042D480: sub esp, 00000010h
  loc_0042D483: mov ecx, 0000000Ah
  loc_0042D488: mov ebp, esp
  loc_0042D48A: mov edi, ecx
  loc_0042D48C: mov eax, 80020004h
  loc_0042D491: sub esp, 00000010h
  loc_0042D494: mov [ebp], ecx
  loc_0042D497: mov ecx, arg_2C
  loc_0042D49B: mov esi, [00430088h]
  loc_0042D4A1: mov edx, eax
  loc_0042D4A3: mov Proc_15_20_42D460, ecx
  loc_0042D4A6: mov ecx, esp
  loc_0042D4A8: mov ebx, [esi]
  loc_0042D4AA: push esi
  loc_0042D4AB: mov Me, eax
  loc_0042D4AE: mov eax, arg_38
  loc_0042D4B2: mov arg_C, eax
  loc_0042D4B5: mov eax, arg_30
  loc_0042D4B9: mov [ecx], edi
  loc_0042D4BB: mov [ecx+00000004h], eax
  loc_0042D4BE: mov [ecx+00000008h], edx
  loc_0042D4C1: mov edx, arg_38
  loc_0042D4C5: mov [ecx+0000000Ch], edx
  loc_0042D4C8: call [ebx+000002B0h]
  loc_0042D4CE: mov edi, [00401030h] ; __vbaHresultCheckObj
  loc_0042D4D4: test eax, eax
  loc_0042D4D6: fnclex
  loc_0042D4D8: jge 0042D4E8h
  loc_0042D4DA: push 000002B0h
  loc_0042D4DF: push 0040DB64h
  loc_0042D4E4: push esi
  loc_0042D4E5: push eax
  loc_0042D4E6: call edi
  loc_0042D4E8: mov eax, [004300ECh]
  loc_0042D4ED: test eax, eax
  loc_0042D4EF: jnz 0042D501h
  loc_0042D4F1: push 004300ECh
  loc_0042D4F6: push 0040A624h
  loc_0042D4FB: call [004010D4h] ; __vbaNew2
  loc_0042D501: mov esi, [004300ECh]
  loc_0042D507: push esi
  loc_0042D508: mov eax, [esi]
  loc_0042D50A: call [eax+000002B4h]
  loc_0042D510: test eax, eax
  loc_0042D512: fnclex
  loc_0042D514: jge 0042D524h
  loc_0042D516: push 000002B4h
  loc_0042D51B: push 0040ECECh
  loc_0042D520: push esi
  loc_0042D521: push eax
  loc_0042D522: call edi
  loc_0042D524: pop edi
  loc_0042D525: pop esi
  loc_0042D526: pop ebp
  loc_0042D527: xor eax, eax
  loc_0042D529: pop ebx
  loc_0042D52A: add esp, 00000010h
  loc_0042D52D: retn 0004h
End Function
