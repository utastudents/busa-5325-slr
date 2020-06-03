VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "COMDLG32.OCX"
Begin VB.Form frmSLRSensativity
  Caption = "SLR - Sensitivity"
  WindowState = 2
  ScaleMode = 0
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmSLRSensativity.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 165
  ClientTop = 510
  ClientWidth = 11355
  ClientHeight = 8025
  ScaleLeft = 0
  ScaleTop = 0
  ScaleWidth = 100
  ScaleHeight = 100
  StartUpPosition = 2 'CenterScreen
  Begin VB.HScrollBar hsbXg
    Left = 8220
    Top = 4530
    Width = 2325
    Height = 225
    Visible = 0   'False
    TabIndex = 38
    SmallChange = 30
    LargeChange = 100
    Value = 2000
  End
  Begin VB.PictureBox Picture1
    BackColor = &H80000005&
    ForeColor = &H80000008&
    Left = 7200
    Top = 390
    Width = 4095
    Height = 3255
    TabIndex = 2
    ScaleMode = 1
    AutoRedraw = False
    FontTransparent = True
    Appearance = 0 'Flat
    Begin VB.Line linPredInt
      X1 = 2640
      Y1 = 660
      X2 = 2640
      Y2 = 1710
      Visible = 0   'False
      BorderStyle = 3 'Dot
    End
    Begin VB.Line linEstInt
      X1 = 1980
      Y1 = 1260
      X2 = 1980
      Y2 = 2040
      Visible = 0   'False
    End
    Begin VB.Shape shpXg
      Left = 1440
      Top = 2520
      Width = 100
      Height = 100
      Visible = 0   'False
      Shape = 3
      FillStyle = 0
    End
    Begin VB.Line LinLSLine
      X1 = 720
      Y1 = 2520
      X2 = 3480
      Y2 = 600
      Visible = 0   'False
    End
    Begin VB.Shape Shape1
      Index = 0
      Left = 720
      Top = 2280
      Width = 200
      Height = 200
      Visible = 0   'False
      Shape = 3
    End
  End
  Begin MSComDlg.CommonDialog dlgSave
    OleObjectBlob = "frmSLRSensativity.frx":030A
    Left = 10470
    Top = 7530
  End
  Begin TabDlg.SSTab SSTab1
    Left = 120
    Top = 240
    Width = 5715
    Height = 7185
    TabIndex = 0
    OleObjectBlob = "frmSLRSensativity.frx":036E
    Begin VB.PictureBox picResidPlot
      BackColor = &H80000009&
      Left = -74280
      Top = 1560
      Width = 4695
      Height = 4455
      TabIndex = 41
      ScaleMode = 0
      ScaleLeft = 0
      ScaleTop = 0
      ScaleWidth = 101.295
      ScaleHeight = 4455
      AutoRedraw = False
      FontTransparent = True
    End
    Begin MSFlexGridLib.MSFlexGrid flxDiagnostics
      Left = -74760
      Top = 1680
      Width = 5055
      Height = 4335
      TabIndex = 40
      OleObjectBlob = "frmSLRSensativity.frx":0D90
    End
    Begin VB.TextBox txtXg
      Left = -73920
      Top = 4080
      Width = 1350
      Height = 555
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
    Begin MSFlexGridLib.MSFlexGrid flxData
      Left = 360
      Top = 1560
      Width = 4875
      Height = 5025
      TabIndex = 1
      OleObjectBlob = "frmSLRSensativity.frx":0E7C
    End
    Begin VB.Label Label5
      Caption = "Yhat"
      Left = -72480
      Top = 6480
      Width = 975
      Height = 375
      TabIndex = 43
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
    Begin VB.Label Label4
      Caption = "R e s i d u a l"
      Left = -74880
      Top = 1800
      Width = 255
      Height = 3855
      TabIndex = 42
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
    Begin VB.Label lblSEEst
      Caption = "label1"
      Left = -74640
      Top = 6600
      Width = 1185
      Height = 495
      TabIndex = 37
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
    Begin VB.Label Label15
      Caption = "Est.S.E."
      Left = -74760
      Top = 6000
      Width = 1245
      Height = 465
      TabIndex = 36
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
    Begin VB.Label lblYhatg
      Caption = "label1"
      Left = -71040
      Top = 4080
      Width = 1035
      Height = 420
      TabIndex = 35
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
    Begin VB.Label Label14
      Caption = "Yhat ="
      Left = -72570
      Top = 4080
      Width = 1230
      Height = 345
      TabIndex = 34
      Alignment = 1 'Right Justify
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
    Begin VB.Label lblSEPred
      Caption = "label1"
      Left = -74760
      Top = 5160
      Width = 1170
      Height = 450
      TabIndex = 33
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
    Begin VB.Label label13
      Caption = "Pred S.E."
      Left = -74880
      Top = 4680
      Width = 1635
      Height = 660
      TabIndex = 32
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
    Begin VB.Label lblUpMean
      Caption = "label1"
      Left = -71040
      Top = 6600
      Width = 1305
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
    End
    Begin VB.Label lblLowMean
      Caption = "label1"
      Left = -72600
      Top = 6600
      Width = 1425
      Height = 495
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
    Begin VB.Label lblUpPred
      Caption = "label1"
      Left = -71040
      Top = 5280
      Width = 1095
      Height = 450
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
    End
    Begin VB.Label lblLowPred
      Caption = "label1"
      Left = -72600
      Top = 5280
      Width = 975
      Height = 420
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
    Begin VB.Label Label12
      Caption = "Upper"
      Left = -71040
      Top = 6000
      Width = 1140
      Height = 555
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
    Begin VB.Label Label11
      Caption = "Lower"
      Left = -72600
      Top = 6000
      Width = 1245
      Height = 525
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
    Begin VB.Line Line4
      X1 = -74760
      Y1 = 5880
      X2 = -69825
      Y2 = 5880
    End
    Begin VB.Label Label10
      Caption = "Upper"
      Left = -71040
      Top = 4680
      Width = 1095
      Height = 420
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
    Begin VB.Label Label9
      Caption = "Lower"
      Left = -72600
      Top = 4680
      Width = 1095
      Height = 660
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
    Begin VB.Label Label8
      Caption = " Xg = "
      Left = -74880
      Top = 4080
      Width = 870
      Height = 495
      TabIndex = 18
      Alignment = 1 'Right Justify
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
    Begin VB.Line Line3
      X1 = -74880
      Y1 = 3960
      X2 = -69990
      Y2 = 3960
    End
    Begin VB.Label lblRsquare
      Caption = "label1"
      Left = -70920
      Top = 3360
      Width = 1230
      Height = 495
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
    Begin VB.Label Label7
      Caption = "R-Sq. = "
      Left = -72240
      Top = 3360
      Width = 1230
      Height = 435
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
    Begin VB.Label lblMSE
      Caption = "label1"
      Left = -73680
      Top = 3360
      Width = 1260
      Height = 375
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
    End
    Begin VB.Label lblMSEc
      Caption = "MSE = "
      Left = -74880
      Top = 3360
      Width = 1230
      Height = 405
      TabIndex = 14
      Alignment = 1 'Right Justify
      BeginProperty Font
        Name = "MS Sans Serif"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "Mean Squared Error (the estimate of the variance of Y given any X)"
    End
    Begin VB.Line Line2
      X1 = -74880
      Y1 = 3240
      X2 = -69855
      Y2 = 3240
    End
    Begin VB.Label lblTTest
      Caption = "label1"
      Left = -70920
      Top = 2760
      Width = 870
      Height = 405
      TabIndex = 13
      AutoSize = -1  'True
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblSbeta0Hat
      Left = -72480
      Top = 2400
      Width = 1050
      Height = 285
      TabIndex = 12
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblSBetaHat
      Caption = "label1"
      Left = -72480
      Top = 2760
      Width = 870
      Height = 405
      TabIndex = 11
      AutoSize = -1  'True
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblBetaHat
      Caption = "label1"
      Left = -73800
      Top = 2760
      Width = 870
      Height = 405
      TabIndex = 10
      AutoSize = -1  'True
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblBeta0Hat
      Caption = "label1"
      Left = -73785
      Top = 2400
      Width = 870
      Height = 405
      TabIndex = 9
      AutoSize = -1  'True
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblBetaOne
      Caption = "b1"
      Left = -74760
      Top = 2760
      Width = 615
      Height = 420
      TabIndex = 8
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label lblBetaZero
      Caption = "b0"
      Left = -74760
      Top = 2400
      Width = 735
      Height = 600
      TabIndex = 7
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Line Line1
      X1 = -74760
      Y1 = 2280
      X2 = -69735
      Y2 = 2280
    End
    Begin VB.Label lblTestCaption
      Caption = "T-Test B1 ="
      Left = -71160
      Top = 1410
      Width = 1695
      Height = 915
      TabIndex = 6
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "T-Test"
    End
    Begin VB.Label Label3
      Caption = "S. E.  Est"
      Left = -72570
      Top = 1380
      Width = 765
      Height = 915
      TabIndex = 5
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "Estimated Standard Error of the Beta Estimate"
    End
    Begin VB.Label Label2
      Caption = "Beta Est."
      Left = -73785
      Top = 1410
      Width = 735
      Height = 900
      TabIndex = 4
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
      ToolTipText = "(Least Squares) Point Estimate of the Betas"
    End
    Begin VB.Label Label1
      Caption = "Beta"
      Left = -74760
      Top = 1815
      Width = 675
      Height = 405
      TabIndex = 3
      AutoSize = -1  'True
      BeginProperty Font
        Name = "Times New Roman"
        Size = 18
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
  End
  Begin VB.Label lblHelp
    Caption = "frmSLRSensativity.frx":0F68
    BackColor = &H80FFFF&
    Left = 6000
    Top = 5010
    Width = 5055
    Height = 2265
    Visible = 0   'False
    TabIndex = 39
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
  Begin VB.Label lblXmax
    Left = 10260
    Top = 3780
    Width = 1035
    Height = 405
    TabIndex = 31
    Alignment = 1 'Right Justify
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
  Begin VB.Label lblXmin
    Left = 7290
    Top = 3780
    Width = 1185
    Height = 435
    TabIndex = 30
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
  Begin VB.Label lblYmin
    Left = 5700
    Top = 3210
    Width = 1335
    Height = 525
    TabIndex = 29
    Alignment = 1 'Right Justify
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
  Begin VB.Label lblYmax
    Left = 5640
    Top = 510
    Width = 1365
    Height = 555
    TabIndex = 28
    Alignment = 1 'Right Justify
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
  Begin VB.Menu mnuFile
    Caption = "&File"
    Begin VB.Menu mnuFileOpen
      Caption = "&Open Data File (*.txt)"
    End
    Begin VB.Menu mnuFileSave
      Caption = "&Save Data"
    End
    Begin VB.Menu mnuExit
      Caption = "E&xit"
    End
  End
  Begin VB.Menu mnuGraphics
    Caption = "&View"
    Begin VB.Menu mnuShowLSline
      Caption = "&Least Squares Line"
    End
    Begin VB.Menu mnuShowXg
      Caption = "&Given Value of X"
    End
    Begin VB.Menu mnuShowEst
      Caption = "&Estimation Interval"
    End
    Begin VB.Menu mnuShowPred
      Caption = "&Prediction Interval"
    End
    Begin VB.Menu mnuShowSlider
      Caption = "&Slider for Different Xg Values"
    End
  End
  Begin VB.Menu mnuChange
    Caption = "&Edit"
    Begin VB.Menu mnuChangeTTable
      Caption = "Alpha and/or T-table value"
    End
    Begin VB.Menu mnuChangeHypValue
      Caption = "&Hypothesized Value of Beta"
    End
  End
  Begin VB.Menu mnuHelp
    Caption = "&Help"
    Begin VB.Menu mnuViewHelp
      Caption = "&Instructions On"
    End
    Begin VB.Menu mnuAbout
      Caption = "&About this program"
    End
  End
End

Attribute VB_Name = "frmSLRSensativity"


Private Sub flxData_UnknownEvent_9 '40E290
  loc_0040E290: push ebp
  loc_0040E291: mov ebp, esp
  loc_0040E293: sub esp, 0000000Ch
  loc_0040E296: push 00401546h ; __vbaExceptHandler
  loc_0040E29B: mov eax, fs:[00000000h]
  loc_0040E2A1: push eax
  loc_0040E2A2: mov fs:[00000000h], esp
  loc_0040E2A9: sub esp, 00000008h
  loc_0040E2AC: push ebx
  loc_0040E2AD: push esi
  loc_0040E2AE: push edi
  loc_0040E2AF: mov var_C, esp
  loc_0040E2B2: mov var_8, 004012C8h
  loc_0040E2B9: mov eax, Me
  loc_0040E2BC: mov ecx, eax
  loc_0040E2BE: and ecx, 00000001h
  loc_0040E2C1: mov var_4, ecx
  loc_0040E2C4: and al, FEh
  loc_0040E2C6: push eax
  loc_0040E2C7: mov Me, eax
  loc_0040E2CA: mov edx, [eax]
  loc_0040E2CC: call [edx+00000004h]
  loc_0040E2CF: mov var_4, 00000000h
  loc_0040E2D6: mov eax, Me
  loc_0040E2D9: push eax
  loc_0040E2DA: mov ecx, [eax]
  loc_0040E2DC: call [ecx+00000008h]
  loc_0040E2DF: mov eax, var_4
  loc_0040E2E2: mov ecx, var_14
  loc_0040E2E5: pop edi
  loc_0040E2E6: pop esi
  loc_0040E2E7: mov fs:[00000000h], ecx
  loc_0040E2EE: pop ebx
  loc_0040E2EF: mov esp, ebp
  loc_0040E2F1: pop ebp
  loc_0040E2F2: retn 0004h
End Sub

Private Sub hsbXg_Change() '40C480
  loc_0040C480: push ebp
  loc_0040C481: mov ebp, esp
  loc_0040C483: sub esp, 0000000Ch
  loc_0040C486: push 00401546h ; __vbaExceptHandler
  loc_0040C48B: mov eax, fs:[00000000h]
  loc_0040C491: push eax
  loc_0040C492: mov fs:[00000000h], esp
  loc_0040C499: sub esp, 00000050h
  loc_0040C49C: push ebx
  loc_0040C49D: push esi
  loc_0040C49E: push edi
  loc_0040C49F: mov var_C, esp
  loc_0040C4A2: mov var_8, 00401208h
  loc_0040C4A9: mov esi, Me
  loc_0040C4AC: mov eax, esi
  loc_0040C4AE: and eax, 00000001h
  loc_0040C4B1: mov var_4, eax
  loc_0040C4B4: and esi, FFFFFFFEh
  loc_0040C4B7: push esi
  loc_0040C4B8: mov Me, esi
  loc_0040C4BB: mov ecx, [esi]
  loc_0040C4BD: call [ecx+00000004h]
  loc_0040C4C0: mov edx, [esi]
  loc_0040C4C2: xor ebx, ebx
  loc_0040C4C4: push esi
  loc_0040C4C5: mov var_18, ebx
  loc_0040C4C8: mov var_28, ebx
  loc_0040C4CB: mov var_38, ebx
  loc_0040C4CE: mov var_3C, ebx
  loc_0040C4D1: mov var_40, ebx
  loc_0040C4D4: mov var_4C, ebx
  loc_0040C4D7: call [edx+000002FCh]
  loc_0040C4DD: push eax
  loc_0040C4DE: lea eax, var_4C
  loc_0040C4E1: push eax
  loc_0040C4E2: call [0040105Ch] ; __vbaObjSet
  loc_0040C4E8: mov eax, var_4C
  loc_0040C4EB: lea edx, var_3C
  loc_0040C4EE: push edx
  loc_0040C4EF: push eax
  loc_0040C4F0: mov ecx, [eax]
  loc_0040C4F2: call [ecx+000000B8h]
  loc_0040C4F8: cmp eax, ebx
  loc_0040C4FA: fnclex
  loc_0040C4FC: jge 0040C513h
  loc_0040C4FE: mov ecx, var_4C
  loc_0040C501: push 000000B8h
  loc_0040C506: push 00403814h
  loc_0040C50B: push ecx
  loc_0040C50C: push eax
  loc_0040C50D: call [00401040h] ; __vbaHresultCheckObj
  loc_0040C513: mov eax, var_4C
  loc_0040C516: lea ecx, var_40
  loc_0040C519: push ecx
  loc_0040C51A: push eax
  loc_0040C51B: mov edx, [eax]
  loc_0040C51D: call [edx+000000A0h]
  loc_0040C523: cmp eax, ebx
  loc_0040C525: fnclex
  loc_0040C527: jge 0040C53Eh
  loc_0040C529: mov edx, var_4C
  loc_0040C52C: push 000000A0h
  loc_0040C531: push 00403814h
  loc_0040C536: push edx
  loc_0040C537: push eax
  loc_0040C538: call [00401040h] ; __vbaHresultCheckObj
  loc_0040C53E: movsx eax, var_3C
  loc_0040C542: fld real4 ptr [esi+00000044h]
  loc_0040C545: fsub st0, real4 ptr [esi+00000040h]
  loc_0040C548: movsx ecx, var_40
  loc_0040C54C: mov var_58, eax
  loc_0040C54F: mov var_60, ecx
  loc_0040C552: fild real4 ptr var_58
  loc_0040C555: push ecx
  loc_0040C556: fstp real4 ptr var_5C
  loc_0040C559: fmul st0, real4 ptr var_5C
  loc_0040C55C: fild real4 ptr var_60
  loc_0040C55F: fstp real4 ptr var_64
  loc_0040C562: cmp [00415000h], 00000000h
  loc_0040C569: jnz 0040C570h
  loc_0040C56B: fdiv st0, real4 ptr var_64
  loc_0040C56E: jmp 0040C578h
  loc_0040C570: push var_64
  loc_0040C573: call 00401558h ; _adj_fdiv_m32
  loc_0040C578: fadd st0, real4 ptr [esi+00000040h]
  loc_0040C57B: fnstsw ax
  loc_0040C57D: test al, 0Dh
  loc_0040C57F: jnz 0040C6A6h
  loc_0040C585: fstp real4 ptr [esp]
  loc_0040C588: call [004010A8h] ; __vbaStrR4
  loc_0040C58E: mov edi, [00401164h] ; __vbaStrMove
  loc_0040C594: mov edx, eax
  loc_0040C596: mov ecx, 0041502Ch
  loc_0040C59B: call edi
  loc_0040C59D: lea edx, var_38
  loc_0040C5A0: push 00000002h
  loc_0040C5A2: lea eax, var_28
  loc_0040C5A5: push edx
  loc_0040C5A6: push eax
  loc_0040C5A7: mov var_30, 0041502Ch
  loc_0040C5AE: mov var_38, 00004008h
  loc_0040C5B5: call [004010E8h] ; rtcRound
  loc_0040C5BB: lea ecx, var_28
  loc_0040C5BE: push ecx
  loc_0040C5BF: call [00401018h] ; __vbaStrVarMove
  loc_0040C5C5: mov edx, eax
  loc_0040C5C7: mov ecx, 0041502Ch
  loc_0040C5CC: call edi
  loc_0040C5CE: lea ecx, var_28
  loc_0040C5D1: call [00401014h] ; __vbaFreeVar
  loc_0040C5D7: mov edx, [esi]
  loc_0040C5D9: push esi
  loc_0040C5DA: call [edx+0000031Ch]
  loc_0040C5E0: push eax
  loc_0040C5E1: lea eax, var_18
  loc_0040C5E4: push eax
  loc_0040C5E5: call [0040105Ch] ; __vbaObjSet
  loc_0040C5EB: mov edx, [0041502Ch]
  loc_0040C5F1: mov edi, eax
  loc_0040C5F3: push edx
  loc_0040C5F4: push edi
  loc_0040C5F5: mov ecx, [edi]
  loc_0040C5F7: call [ecx+000000A4h]
  loc_0040C5FD: cmp eax, ebx
  loc_0040C5FF: fnclex
  loc_0040C601: jge 0040C615h
  loc_0040C603: push 000000A4h
  loc_0040C608: push 00403848h
  loc_0040C60D: push edi
  loc_0040C60E: push eax
  loc_0040C60F: call [00401040h] ; __vbaHresultCheckObj
  loc_0040C615: lea ecx, var_18
  loc_0040C618: call [00401180h] ; __vbaFreeObj
  loc_0040C61E: mov eax, [00415028h]
  loc_0040C623: mov ecx, [0041502Ch]
  loc_0040C629: push eax
  loc_0040C62A: push ecx
  loc_0040C62B: call [004010A0h] ; __vbaR4Str
  loc_0040C631: push ecx
  loc_0040C632: lea edx, [esi+0000003Ch]
  loc_0040C635: fstp real4 ptr [esp]
  loc_0040C638: push 00000001h
  loc_0040C63A: push 00000002h
  loc_0040C63C: push edx
  loc_0040C63D: call 00409C80h
  loc_0040C642: mov eax, [esi]
  loc_0040C644: push esi
  loc_0040C645: call [eax+00000780h]
  loc_0040C64B: mov ecx, [esi]
  loc_0040C64D: push esi
  loc_0040C64E: call [ecx+0000078Ch]
  loc_0040C654: lea edx, var_4C
  loc_0040C657: push ebx
  loc_0040C658: push edx
  loc_0040C659: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040C65F: mov var_4, ebx
  loc_0040C662: fwait
  loc_0040C663: push 0040C687h
  loc_0040C668: jmp 0040C67Dh
  loc_0040C66A: lea ecx, var_18
  loc_0040C66D: call [00401180h] ; __vbaFreeObj
  loc_0040C673: lea ecx, var_28
  loc_0040C676: call [00401014h] ; __vbaFreeVar
  loc_0040C67C: ret
  loc_0040C67D: lea ecx, var_4C
  loc_0040C680: call [00401180h] ; __vbaFreeObj
  loc_0040C686: ret
  loc_0040C687: mov eax, Me
  loc_0040C68A: push eax
  loc_0040C68B: mov ecx, [eax]
  loc_0040C68D: call [ecx+00000008h]
  loc_0040C690: mov eax, var_4
  loc_0040C693: mov ecx, var_14
  loc_0040C696: pop edi
  loc_0040C697: pop esi
  loc_0040C698: mov fs:[00000000h], ecx
  loc_0040C69F: pop ebx
  loc_0040C6A0: mov esp, ebp
  loc_0040C6A2: pop ebp
  loc_0040C6A3: retn 0004h
End Sub

Private Sub SSTab1_UnknownEvent_10 '411760
  loc_00411760: push ebp
  loc_00411761: mov ebp, esp
  loc_00411763: sub esp, 0000000Ch
  loc_00411766: push 00401546h ; __vbaExceptHandler
  loc_0041176B: mov eax, fs:[00000000h]
  loc_00411771: push eax
  loc_00411772: mov fs:[00000000h], esp
  loc_00411779: sub esp, 00000020h
  loc_0041177C: push ebx
  loc_0041177D: push esi
  loc_0041177E: push edi
  loc_0041177F: mov var_C, esp
  loc_00411782: mov var_8, 00401408h
  loc_00411789: mov esi, Me
  loc_0041178C: mov eax, esi
  loc_0041178E: and eax, 00000001h
  loc_00411791: mov var_4, eax
  loc_00411794: and esi, FFFFFFFEh
  loc_00411797: push esi
  loc_00411798: mov Me, esi
  loc_0041179B: mov ecx, [esi]
  loc_0041179D: call [ecx+00000004h]
  loc_004117A0: mov edx, [esi]
  loc_004117A2: xor ebx, ebx
  loc_004117A4: push ebx
  loc_004117A5: push 00000004h
  loc_004117A7: push esi
  loc_004117A8: mov var_18, ebx
  loc_004117AB: mov var_28, ebx
  loc_004117AE: call [edx+00000408h]
  loc_004117B4: push eax
  loc_004117B5: lea eax, var_18
  loc_004117B8: push eax
  loc_004117B9: call [0040105Ch] ; __vbaObjSet
  loc_004117BF: lea ecx, var_28
  loc_004117C2: push eax
  loc_004117C3: push ecx
  loc_004117C4: call [004010B4h] ; __vbaLateIdCallLd
  loc_004117CA: add esp, 00000010h
  loc_004117CD: push eax
  loc_004117CE: call [00401104h] ; __vbaI2Var
  loc_004117D4: xor edx, edx
  loc_004117D6: cmp ax, 0003h
  loc_004117DA: setz dl
  loc_004117DD: neg edx
  loc_004117DF: lea ecx, var_18
  loc_004117E2: mov edi, edx
  loc_004117E4: call [00401180h] ; __vbaFreeObj
  loc_004117EA: lea ecx, var_28
  loc_004117ED: call [00401014h] ; __vbaFreeVar
  loc_004117F3: cmp di, bx
  loc_004117F6: jz 00411801h
  loc_004117F8: mov eax, [esi]
  loc_004117FA: push esi
  loc_004117FB: call [eax+00000718h]
  loc_00411801: mov var_4, ebx
  loc_00411804: push 0041181Fh
  loc_00411809: jmp 0041181Eh
  loc_0041180B: lea ecx, var_18
  loc_0041180E: call [00401180h] ; __vbaFreeObj
  loc_00411814: lea ecx, var_28
  loc_00411817: call [00401014h] ; __vbaFreeVar
  loc_0041181D: ret
  loc_0041181E: ret
  loc_0041181F: mov eax, Me
  loc_00411822: push eax
  loc_00411823: mov ecx, [eax]
  loc_00411825: call [ecx+00000008h]
  loc_00411828: mov eax, var_4
  loc_0041182B: mov ecx, var_14
  loc_0041182E: pop edi
  loc_0041182F: pop esi
  loc_00411830: mov fs:[00000000h], ecx
  loc_00411837: pop ebx
  loc_00411838: mov esp, ebp
  loc_0041183A: pop ebp
  loc_0041183B: retn 0008h
End Sub

Private Sub mnuFileOpen_Click() '40D3C0
  loc_0040D3C0: push ebp
  loc_0040D3C1: mov ebp, esp
  loc_0040D3C3: sub esp, 0000000Ch
  loc_0040D3C6: push 00401546h ; __vbaExceptHandler
  loc_0040D3CB: mov eax, fs:[00000000h]
  loc_0040D3D1: push eax
  loc_0040D3D2: mov fs:[00000000h], esp
  loc_0040D3D9: sub esp, 0000008Ch
  loc_0040D3DF: push ebx
  loc_0040D3E0: push esi
  loc_0040D3E1: push edi
  loc_0040D3E2: mov var_C, esp
  loc_0040D3E5: mov var_8, 00401260h
  loc_0040D3EC: mov esi, Me
  loc_0040D3EF: mov eax, esi
  loc_0040D3F1: and eax, 00000001h
  loc_0040D3F4: mov var_4, eax
  loc_0040D3F7: and esi, FFFFFFFEh
  loc_0040D3FA: push esi
  loc_0040D3FB: mov Me, esi
  loc_0040D3FE: mov ecx, [esi]
  loc_0040D400: call [ecx+00000004h]
  loc_0040D403: mov edx, [esi]
  loc_0040D405: lea eax, var_18
  loc_0040D408: xor edi, edi
  loc_0040D40A: push eax
  loc_0040D40B: push esi
  loc_0040D40C: mov var_18, edi
  loc_0040D40F: mov var_1C, edi
  loc_0040D412: mov var_20, edi
  loc_0040D415: mov var_24, edi
  loc_0040D418: mov var_34, edi
  loc_0040D41B: mov var_44, edi
  loc_0040D41E: mov var_54, edi
  loc_0040D421: mov var_64, edi
  loc_0040D424: call [edx+00000724h]
  loc_0040D42A: mov ecx, [esi]
  loc_0040D42C: push esi
  loc_0040D42D: call [ecx+0000072Ch]
  loc_0040D433: mov eax, var_18
  loc_0040D436: mov ebx, var_70
  loc_0040D439: mov [esi+00000038h], ax
  loc_0040D43D: add ax, 0001h
  loc_0040D441: jo 0040D61Ch
  loc_0040D447: sub esp, 00000010h
  loc_0040D44A: mov ecx, 00000003h
  loc_0040D44F: mov edx, esp
  loc_0040D451: movsx eax, ax
  loc_0040D454: mov [edx], ecx
  loc_0040D456: mov ecx, [esi]
  loc_0040D458: push 00000004h
  loc_0040D45A: push esi
  loc_0040D45B: mov [edx+00000004h], ebx
  loc_0040D45E: mov [edx+00000008h], eax
  loc_0040D461: mov eax, var_68
  loc_0040D464: mov [edx+0000000Ch], eax
  loc_0040D467: call [ecx+00000410h]
  loc_0040D46D: lea edx, var_24
  loc_0040D470: push eax
  loc_0040D471: push edx
  loc_0040D472: call [0040105Ch] ; __vbaObjSet
  loc_0040D478: push eax
  loc_0040D479: call [0040116Ch] ; __vbaLateIdSt
  loc_0040D47F: lea ecx, var_24
  loc_0040D482: call [00401180h] ; __vbaFreeObj
  loc_0040D488: mov ax, [esi+00000038h]
  loc_0040D48C: mov ecx, 00000003h
  loc_0040D491: add ax, 0001h
  loc_0040D495: jo 0040D61Ch
  loc_0040D49B: sub esp, 00000010h
  loc_0040D49E: mov edx, esp
  loc_0040D4A0: movsx eax, ax
  loc_0040D4A3: mov [edx], ecx
  loc_0040D4A5: mov ecx, [esi]
  loc_0040D4A7: push 00000004h
  loc_0040D4A9: push esi
  loc_0040D4AA: mov [edx+00000004h], ebx
  loc_0040D4AD: mov [edx+00000008h], eax
  loc_0040D4B0: mov eax, var_68
  loc_0040D4B3: mov [edx+0000000Ch], eax
  loc_0040D4B6: call [ecx+0000040Ch]
  loc_0040D4BC: lea edx, var_24
  loc_0040D4BF: push eax
  loc_0040D4C0: push edx
  loc_0040D4C1: call [0040105Ch] ; __vbaObjSet
  loc_0040D4C7: push eax
  loc_0040D4C8: call [0040116Ch] ; __vbaLateIdSt
  loc_0040D4CE: lea ecx, var_24
  loc_0040D4D1: call [00401180h] ; __vbaFreeObj
  loc_0040D4D7: mov eax, 0000000Ah
  loc_0040D4DC: mov ecx, 80020004h
  loc_0040D4E1: mov var_64, eax
  loc_0040D4E4: mov var_54, eax
  loc_0040D4E7: mov var_44, eax
  loc_0040D4EA: mov ax, [esi+00000038h]
  loc_0040D4EE: sub ax, 0002h
  loc_0040D4F2: push 00403AF0h ; "Using menu above, please change ttable value based on "
  loc_0040D4F7: jo 0040D61Ch
  loc_0040D4FD: push eax
  loc_0040D4FE: mov var_5C, ecx
  loc_0040D501: mov var_4C, ecx
  loc_0040D504: mov var_3C, ecx
  loc_0040D507: call [00401000h] ; __vbaStrI2
  loc_0040D50D: mov ebx, [00401164h] ; __vbaStrMove
  loc_0040D513: mov edx, eax
  loc_0040D515: lea ecx, var_1C
  loc_0040D518: call ebx
  loc_0040D51A: push eax
  loc_0040D51B: call [0040103Ch] ; __vbaStrCat
  loc_0040D521: mov edx, eax
  loc_0040D523: lea ecx, var_20
  loc_0040D526: call ebx
  loc_0040D528: push eax
  loc_0040D529: push 00403B64h ; " degrees of freedom"
  loc_0040D52E: call [0040103Ch] ; __vbaStrCat
  loc_0040D534: lea ecx, var_64
  loc_0040D537: mov var_2C, eax
  loc_0040D53A: lea edx, var_54
  loc_0040D53D: push ecx
  loc_0040D53E: lea eax, var_44
  loc_0040D541: push edx
  loc_0040D542: push eax
  loc_0040D543: lea ecx, var_34
  loc_0040D546: push edi
  loc_0040D547: push ecx
  loc_0040D548: mov var_34, 00000008h
  loc_0040D54F: call [00401060h] ; rtcMsgBox
  loc_0040D555: lea edx, var_20
  loc_0040D558: lea eax, var_1C
  loc_0040D55B: push edx
  loc_0040D55C: push eax
  loc_0040D55D: push 00000002h
  loc_0040D55F: call [00401130h] ; __vbaFreeStrList
  loc_0040D565: lea ecx, var_64
  loc_0040D568: lea edx, var_54
  loc_0040D56B: push ecx
  loc_0040D56C: lea eax, var_44
  loc_0040D56F: push edx
  loc_0040D570: lea ecx, var_34
  loc_0040D573: push eax
  loc_0040D574: push ecx
  loc_0040D575: push 00000004h
  loc_0040D577: call [00401024h] ; __vbaFreeVarList
  loc_0040D57D: mov edx, [esi]
  loc_0040D57F: add esp, 00000020h
  loc_0040D582: push esi
  loc_0040D583: call [edx+00000728h]
  loc_0040D589: mov eax, [esi]
  loc_0040D58B: push esi
  loc_0040D58C: call [eax+000006F8h]
  loc_0040D592: cmp eax, edi
  loc_0040D594: jge 0040D5A8h
  loc_0040D596: push 000006F8h
  loc_0040D59B: push 004030CCh
  loc_0040D5A0: push esi
  loc_0040D5A1: push eax
  loc_0040D5A2: call [00401040h] ; __vbaHresultCheckObj
  loc_0040D5A8: mov ecx, [esi]
  loc_0040D5AA: push esi
  loc_0040D5AB: call [ecx+00000730h]
  loc_0040D5B1: mov edx, [esi]
  loc_0040D5B3: push esi
  loc_0040D5B4: call [edx+0000073Ch]
  loc_0040D5BA: mov var_4, edi
  loc_0040D5BD: push 0040D5FDh
  loc_0040D5C2: jmp 0040D5FCh
  loc_0040D5C4: lea eax, var_20
  loc_0040D5C7: lea ecx, var_1C
  loc_0040D5CA: push eax
  loc_0040D5CB: push ecx
  loc_0040D5CC: push 00000002h
  loc_0040D5CE: call [00401130h] ; __vbaFreeStrList
  loc_0040D5D4: add esp, 0000000Ch
  loc_0040D5D7: lea ecx, var_24
  loc_0040D5DA: call [00401180h] ; __vbaFreeObj
  loc_0040D5E0: lea edx, var_64
  loc_0040D5E3: lea eax, var_54
  loc_0040D5E6: push edx
  loc_0040D5E7: lea ecx, var_44
  loc_0040D5EA: push eax
  loc_0040D5EB: lea edx, var_34
  loc_0040D5EE: push ecx
  loc_0040D5EF: push edx
  loc_0040D5F0: push 00000004h
  loc_0040D5F2: call [00401024h] ; __vbaFreeVarList
  loc_0040D5F8: add esp, 00000014h
  loc_0040D5FB: ret
  loc_0040D5FC: ret
  loc_0040D5FD: mov eax, Me
  loc_0040D600: push eax
  loc_0040D601: mov ecx, [eax]
  loc_0040D603: call [ecx+00000008h]
  loc_0040D606: mov eax, var_4
  loc_0040D609: mov ecx, var_14
  loc_0040D60C: pop edi
  loc_0040D60D: pop esi
  loc_0040D60E: mov fs:[00000000h], ecx
  loc_0040D615: pop ebx
  loc_0040D616: mov esp, ebp
  loc_0040D618: pop ebp
  loc_0040D619: retn 0004h
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '40EDD0
  loc_0040EDD0: push ebp
  loc_0040EDD1: mov ebp, esp
  loc_0040EDD3: sub esp, 0000000Ch
  loc_0040EDD6: push 00401546h ; __vbaExceptHandler
  loc_0040EDDB: mov eax, fs:[00000000h]
  loc_0040EDE1: push eax
  loc_0040EDE2: mov fs:[00000000h], esp
  loc_0040EDE9: sub esp, 00000038h
  loc_0040EDEC: push ebx
  loc_0040EDED: push esi
  loc_0040EDEE: push edi
  loc_0040EDEF: mov var_C, esp
  loc_0040EDF2: mov var_8, 00401340h
  loc_0040EDF9: mov esi, Me
  loc_0040EDFC: mov eax, esi
  loc_0040EDFE: and eax, 00000001h
  loc_0040EE01: mov var_4, eax
  loc_0040EE04: and esi, FFFFFFFEh
  loc_0040EE07: push esi
  loc_0040EE08: mov Me, esi
  loc_0040EE0B: mov ecx, [esi]
  loc_0040EE0D: call [ecx+00000004h]
  loc_0040EE10: mov edx, Button
  loc_0040EE13: xor eax, eax
  loc_0040EE15: mov var_1C, eax
  loc_0040EE18: mov var_2C, eax
  loc_0040EE1B: cmp [edx], 0001h
  loc_0040EE1F: jnz 0040F065h
  loc_0040EE25: mov edi, var_28
  loc_0040EE28: sub esp, 00000010h
  loc_0040EE2B: mov edx, esp
  loc_0040EE2D: mov ecx, 00000003h
  loc_0040EE32: mov ebx, var_20
  loc_0040EE35: push 0000000Bh
  loc_0040EE37: mov [edx], ecx
  loc_0040EE39: push esi
  loc_0040EE3A: mov [edx+00000004h], edi
  loc_0040EE3D: mov [edx+00000008h], eax
  loc_0040EE40: mov eax, [esi]
  loc_0040EE42: mov [edx+0000000Ch], ebx
  loc_0040EE45: call [eax+00000410h]
  loc_0040EE4B: lea ecx, var_1C
  loc_0040EE4E: push eax
  loc_0040EE4F: push ecx
  loc_0040EE50: call [0040105Ch] ; __vbaObjSet
  loc_0040EE56: push eax
  loc_0040EE57: call [0040116Ch] ; __vbaLateIdSt
  loc_0040EE5D: lea ecx, var_1C
  loc_0040EE60: call [00401180h] ; __vbaFreeObj
  loc_0040EE66: mov ax, [esi+00000038h]
  loc_0040EE6A: mov var_40, 00000001h
  loc_0040EE71: sub ax, 0001h
  loc_0040EE75: jo 0040F099h
  loc_0040EE7B: xor ecx, ecx
  loc_0040EE7D: mov var_44, eax
  loc_0040EE80: mov var_18, ecx
  loc_0040EE83: cmp cx, ax
  loc_0040EE86: jg 0040EF22h
  loc_0040EE8C: sub esp, 00000010h
  loc_0040EE8F: movsx eax, cx
  loc_0040EE92: mov edx, esp
  loc_0040EE94: mov ecx, 00000003h
  loc_0040EE99: push 0000000Ah
  loc_0040EE9B: push esi
  loc_0040EE9C: mov [edx], ecx
  loc_0040EE9E: mov [edx+00000004h], edi
  loc_0040EEA1: mov [edx+00000008h], eax
  loc_0040EEA4: mov eax, [esi]
  loc_0040EEA6: mov [edx+0000000Ch], ebx
  loc_0040EEA9: call [eax+00000410h]
  loc_0040EEAF: lea ecx, var_1C
  loc_0040EEB2: push eax
  loc_0040EEB3: push ecx
  loc_0040EEB4: call [0040105Ch] ; __vbaObjSet
  loc_0040EEBA: push eax
  loc_0040EEBB: call [0040116Ch] ; __vbaLateIdSt
  loc_0040EEC1: lea ecx, var_1C
  loc_0040EEC4: call [00401180h] ; __vbaFreeObj
  loc_0040EECA: sub esp, 00000010h
  loc_0040EECD: mov ecx, 00000003h
  loc_0040EED2: mov edx, esp
  loc_0040EED4: xor eax, eax
  loc_0040EED6: push 00000026h
  loc_0040EED8: push esi
  loc_0040EED9: mov [edx], ecx
  loc_0040EEDB: mov [edx+00000004h], edi
  loc_0040EEDE: mov [edx+00000008h], eax
  loc_0040EEE1: mov eax, [esi]
  loc_0040EEE3: mov [edx+0000000Ch], ebx
  loc_0040EEE6: call [eax+00000410h]
  loc_0040EEEC: lea ecx, var_1C
  loc_0040EEEF: push eax
  loc_0040EEF0: push ecx
  loc_0040EEF1: call [0040105Ch] ; __vbaObjSet
  loc_0040EEF7: push eax
  loc_0040EEF8: call [0040116Ch] ; __vbaLateIdSt
  loc_0040EEFE: lea ecx, var_1C
  loc_0040EF01: call [00401180h] ; __vbaFreeObj
  loc_0040EF07: mov dx, var_40
  loc_0040EF0B: mov eax, var_44
  loc_0040EF0E: add dx, var_18
  loc_0040EF12: jo 0040F099h
  loc_0040EF18: mov var_18, edx
  loc_0040EF1B: mov ecx, edx
  loc_0040EF1D: jmp 0040EE83h
  loc_0040EF22: movsx eax, [esi+00000034h]
  loc_0040EF26: sub esp, 00000010h
  loc_0040EF29: mov ecx, 00000003h
  loc_0040EF2E: mov edx, esp
  loc_0040EF30: push 0000000Ah
  loc_0040EF32: push esi
  loc_0040EF33: mov [edx], ecx
  loc_0040EF35: mov [edx+00000004h], edi
  loc_0040EF38: mov [edx+00000008h], eax
  loc_0040EF3B: mov eax, [esi]
  loc_0040EF3D: mov [edx+0000000Ch], ebx
  loc_0040EF40: call [eax+00000410h]
  loc_0040EF46: lea ecx, var_1C
  loc_0040EF49: push eax
  loc_0040EF4A: push ecx
  loc_0040EF4B: call [0040105Ch] ; __vbaObjSet
  loc_0040EF51: push eax
  loc_0040EF52: call [0040116Ch] ; __vbaLateIdSt
  loc_0040EF58: lea ecx, var_1C
  loc_0040EF5B: call [00401180h] ; __vbaFreeObj
  loc_0040EF61: sub esp, 00000010h
  loc_0040EF64: mov ecx, 00000003h
  loc_0040EF69: mov edx, esp
  loc_0040EF6B: mov eax, 000000FFh
  loc_0040EF70: push 00000026h
  loc_0040EF72: push esi
  loc_0040EF73: mov [edx], ecx
  loc_0040EF75: mov [edx+00000004h], edi
  loc_0040EF78: mov [edx+00000008h], eax
  loc_0040EF7B: mov eax, [esi]
  loc_0040EF7D: mov [edx+0000000Ch], ebx
  loc_0040EF80: call [eax+00000410h]
  loc_0040EF86: lea ecx, var_1C
  loc_0040EF89: push eax
  loc_0040EF8A: push ecx
  loc_0040EF8B: call [0040105Ch] ; __vbaObjSet
  loc_0040EF91: push eax
  loc_0040EF92: call [0040116Ch] ; __vbaLateIdSt
  loc_0040EF98: lea ecx, var_1C
  loc_0040EF9B: call [00401180h] ; __vbaFreeObj
  loc_0040EFA1: sub esp, 00000010h
  loc_0040EFA4: mov ecx, 00000003h
  loc_0040EFA9: mov edx, esp
  loc_0040EFAB: xor eax, eax
  loc_0040EFAD: push 0000000Bh
  loc_0040EFAF: push esi
  loc_0040EFB0: mov [edx], ecx
  loc_0040EFB2: mov [edx+00000004h], edi
  loc_0040EFB5: mov [edx+00000008h], eax
  loc_0040EFB8: mov eax, [esi]
  loc_0040EFBA: mov [edx+0000000Ch], ebx
  loc_0040EFBD: call [eax+0000040Ch]
  loc_0040EFC3: lea ecx, var_1C
  loc_0040EFC6: push eax
  loc_0040EFC7: push ecx
  loc_0040EFC8: call [0040105Ch] ; __vbaObjSet
  loc_0040EFCE: push eax
  loc_0040EFCF: call [0040116Ch] ; __vbaLateIdSt
  loc_0040EFD5: lea ecx, var_1C
  loc_0040EFD8: call [00401180h] ; __vbaFreeObj
  loc_0040EFDE: movsx eax, [esi+00000034h]
  loc_0040EFE2: sub esp, 00000010h
  loc_0040EFE5: mov ecx, 00000003h
  loc_0040EFEA: mov edx, esp
  loc_0040EFEC: push 0000000Ah
  loc_0040EFEE: push esi
  loc_0040EFEF: mov [edx], ecx
  loc_0040EFF1: mov [edx+00000004h], edi
  loc_0040EFF4: mov [edx+00000008h], eax
  loc_0040EFF7: mov eax, [esi]
  loc_0040EFF9: mov [edx+0000000Ch], ebx
  loc_0040EFFC: call [eax+0000040Ch]
  loc_0040F002: lea ecx, var_1C
  loc_0040F005: push eax
  loc_0040F006: push ecx
  loc_0040F007: call [0040105Ch] ; __vbaObjSet
  loc_0040F00D: push eax
  loc_0040F00E: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F014: lea ecx, var_1C
  loc_0040F017: call [00401180h] ; __vbaFreeObj
  loc_0040F01D: mov eax, 000000FFh
  loc_0040F022: sub esp, 00000010h
  loc_0040F025: mov ecx, 00000003h
  loc_0040F02A: mov edx, esp
  loc_0040F02C: push 00000026h
  loc_0040F02E: push esi
  loc_0040F02F: mov [edx], ecx
  loc_0040F031: mov [edx+00000004h], edi
  loc_0040F034: mov [edx+00000008h], eax
  loc_0040F037: mov eax, [esi]
  loc_0040F039: mov [edx+0000000Ch], ebx
  loc_0040F03C: call [eax+0000040Ch]
  loc_0040F042: lea ecx, var_1C
  loc_0040F045: push eax
  loc_0040F046: push ecx
  loc_0040F047: call [0040105Ch] ; __vbaObjSet
  loc_0040F04D: push eax
  loc_0040F04E: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F054: lea ecx, var_1C
  loc_0040F057: call [00401180h] ; __vbaFreeObj
  loc_0040F05D: mov [esi+00000036h], FFFFFFh
  loc_0040F063: xor eax, eax
  loc_0040F065: mov var_4, eax
  loc_0040F068: push 0040F07Ah
  loc_0040F06D: jmp 0040F079h
  loc_0040F06F: lea ecx, var_1C
  loc_0040F072: call [00401180h] ; __vbaFreeObj
  loc_0040F078: ret
  loc_0040F079: ret
  loc_0040F07A: mov eax, Me
  loc_0040F07D: push eax
  loc_0040F07E: mov edx, [eax]
  loc_0040F080: call [edx+00000008h]
  loc_0040F083: mov eax, var_4
  loc_0040F086: mov ecx, var_14
  loc_0040F089: pop edi
  loc_0040F08A: pop esi
  loc_0040F08B: mov fs:[00000000h], ecx
  loc_0040F092: pop ebx
  loc_0040F093: mov esp, ebp
  loc_0040F095: pop ebp
  loc_0040F096: retn 0014h
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '40F0A0
  loc_0040F0A0: push ebp
  loc_0040F0A1: mov ebp, esp
  loc_0040F0A3: sub esp, 0000000Ch
  loc_0040F0A6: push 00401546h ; __vbaExceptHandler
  loc_0040F0AB: mov eax, fs:[00000000h]
  loc_0040F0B1: push eax
  loc_0040F0B2: mov fs:[00000000h], esp
  loc_0040F0B9: sub esp, 000000B8h
  loc_0040F0BF: push ebx
  loc_0040F0C0: push esi
  loc_0040F0C1: push edi
  loc_0040F0C2: mov var_C, esp
  loc_0040F0C5: mov var_8, 00401358h
  loc_0040F0CC: mov esi, Me
  loc_0040F0CF: mov eax, esi
  loc_0040F0D1: and eax, 00000001h
  loc_0040F0D4: mov var_4, eax
  loc_0040F0D7: and esi, FFFFFFFEh
  loc_0040F0DA: push esi
  loc_0040F0DB: mov Me, esi
  loc_0040F0DE: mov ecx, [esi]
  loc_0040F0E0: call [ecx+00000004h]
  loc_0040F0E3: xor eax, eax
  loc_0040F0E5: cmp [esi+00000036h], ax
  loc_0040F0E9: mov var_18, eax
  loc_0040F0EC: mov var_1C, eax
  loc_0040F0EF: mov var_20, eax
  loc_0040F0F2: mov var_30, eax
  loc_0040F0F5: mov var_40, eax
  loc_0040F0F8: mov var_50, eax
  loc_0040F0FB: mov var_60, eax
  loc_0040F0FE: mov var_74, eax
  loc_0040F101: mov var_78, eax
  loc_0040F104: mov var_7C, eax
  loc_0040F107: mov var_80, eax
  loc_0040F10A: mov var_84, eax
  loc_0040F110: mov var_88, eax
  loc_0040F116: mov var_A4, eax
  loc_0040F11C: jnz 0040F70Bh
  loc_0040F122: mov dx, [esi+00000038h]
  loc_0040F126: mov edi, [0040105Ch] ; __vbaObjSet
  loc_0040F12C: mov ebx, [00401180h] ; __vbaFreeObj
  loc_0040F132: mov var_AC, dx
  loc_0040F139: mov [esi+0000008Ch], 0001h
  loc_0040F142: mov cx, var_AC
  loc_0040F149: cmp [esi+0000008Ch], cx
  loc_0040F150: jg 0040F717h
  loc_0040F156: mov edx, [esi]
  loc_0040F158: push esi
  loc_0040F159: call [edx+00000314h]
  loc_0040F15F: push eax
  loc_0040F160: lea eax, var_1C
  loc_0040F163: push eax
  loc_0040F164: call edi
  loc_0040F166: mov ecx, [eax]
  loc_0040F168: lea edx, var_20
  loc_0040F16B: push edx
  loc_0040F16C: mov dx, [esi+0000008Ch]
  loc_0040F173: push edx
  loc_0040F174: push eax
  loc_0040F175: mov var_8C, eax
  loc_0040F17B: call [ecx+00000040h]
  loc_0040F17E: test eax, eax
  loc_0040F180: fnclex
  loc_0040F182: jge 0040F199h
  loc_0040F184: mov ecx, var_8C
  loc_0040F18A: push 00000040h
  loc_0040F18C: push 00403BE8h
  loc_0040F191: push ecx
  loc_0040F192: push eax
  loc_0040F193: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F199: mov eax, var_20
  loc_0040F19C: lea edx, var_A4
  loc_0040F1A2: push eax
  loc_0040F1A3: push edx
  loc_0040F1A4: mov var_20, 00000000h
  loc_0040F1AB: call edi
  loc_0040F1AD: lea ecx, var_1C
  loc_0040F1B0: call ebx
  loc_0040F1B2: mov eax, var_A4
  loc_0040F1B8: lea edx, var_74
  loc_0040F1BB: push edx
  loc_0040F1BC: push eax
  loc_0040F1BD: mov ecx, [eax]
  loc_0040F1BF: call [ecx+00000068h]
  loc_0040F1C2: test eax, eax
  loc_0040F1C4: fnclex
  loc_0040F1C6: jge 0040F1DDh
  loc_0040F1C8: mov ecx, var_A4
  loc_0040F1CE: push 00000068h
  loc_0040F1D0: push 00403C28h
  loc_0040F1D5: push ecx
  loc_0040F1D6: push eax
  loc_0040F1D7: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F1DD: mov eax, var_A4
  loc_0040F1E3: lea ecx, var_78
  loc_0040F1E6: push ecx
  loc_0040F1E7: push eax
  loc_0040F1E8: mov edx, [eax]
  loc_0040F1EA: call [edx+00000068h]
  loc_0040F1ED: test eax, eax
  loc_0040F1EF: fnclex
  loc_0040F1F1: jge 0040F208h
  loc_0040F1F3: mov edx, var_A4
  loc_0040F1F9: push 00000068h
  loc_0040F1FB: push 00403C28h
  loc_0040F200: push edx
  loc_0040F201: push eax
  loc_0040F202: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F208: mov eax, var_A4
  loc_0040F20E: lea edx, var_7C
  loc_0040F211: push edx
  loc_0040F212: push eax
  loc_0040F213: mov ecx, [eax]
  loc_0040F215: call [ecx+00000078h]
  loc_0040F218: test eax, eax
  loc_0040F21A: fnclex
  loc_0040F21C: jge 0040F233h
  loc_0040F21E: mov ecx, var_A4
  loc_0040F224: push 00000078h
  loc_0040F226: push 00403C28h
  loc_0040F22B: push ecx
  loc_0040F22C: push eax
  loc_0040F22D: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F233: mov eax, var_A4
  loc_0040F239: lea ecx, var_80
  loc_0040F23C: push ecx
  loc_0040F23D: push eax
  loc_0040F23E: mov edx, [eax]
  loc_0040F240: call [edx+00000070h]
  loc_0040F243: test eax, eax
  loc_0040F245: fnclex
  loc_0040F247: jge 0040F25Eh
  loc_0040F249: mov edx, var_A4
  loc_0040F24F: push 00000070h
  loc_0040F251: push 00403C28h
  loc_0040F256: push edx
  loc_0040F257: push eax
  loc_0040F258: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F25E: mov eax, var_A4
  loc_0040F264: lea edx, var_84
  loc_0040F26A: push edx
  loc_0040F26B: push eax
  loc_0040F26C: mov ecx, [eax]
  loc_0040F26E: call [ecx+00000070h]
  loc_0040F271: test eax, eax
  loc_0040F273: fnclex
  loc_0040F275: jge 0040F28Ch
  loc_0040F277: mov ecx, var_A4
  loc_0040F27D: push 00000070h
  loc_0040F27F: push 00403C28h
  loc_0040F284: push ecx
  loc_0040F285: push eax
  loc_0040F286: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F28C: mov eax, var_A4
  loc_0040F292: lea ecx, var_88
  loc_0040F298: push ecx
  loc_0040F299: push eax
  loc_0040F29A: mov edx, [eax]
  loc_0040F29C: call [edx+00000080h]
  loc_0040F2A2: test eax, eax
  loc_0040F2A4: fnclex
  loc_0040F2A6: jge 0040F2C0h
  loc_0040F2A8: mov edx, var_A4
  loc_0040F2AE: push 00000080h
  loc_0040F2B3: push 00403C28h
  loc_0040F2B8: push edx
  loc_0040F2B9: push eax
  loc_0040F2BA: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F2C0: mov eax, X
  loc_0040F2C3: mov var_BC, 00000001h
  loc_0040F2CD: fld real4 ptr [eax]
  loc_0040F2CF: fcomp real4 ptr var_74
  loc_0040F2D2: fnstsw ax
  loc_0040F2D4: test ah, 41h
  loc_0040F2D7: jz 0040F2E3h
  loc_0040F2D9: mov var_BC, 00000000h
  loc_0040F2E3: fld real4 ptr var_7C
  loc_0040F2E6: fadd st0, real4 ptr var_78
  loc_0040F2E9: fnstsw ax
  loc_0040F2EB: test al, 0Dh
  loc_0040F2ED: jnz 0040FCD3h
  loc_0040F2F3: call [00401074h] ; __vbaFpR4
  loc_0040F2F9: mov ecx, X
  loc_0040F2FC: mov var_C0, 00000001h
  loc_0040F306: fcomp real4 ptr [ecx]
  loc_0040F308: fnstsw ax
  loc_0040F30A: test ah, 41h
  loc_0040F30D: jz 0040F319h
  loc_0040F30F: mov var_C0, 00000000h
  loc_0040F319: mov edx, Y
  loc_0040F31C: mov var_C4, 00000001h
  loc_0040F326: fld real4 ptr [edx]
  loc_0040F328: fcomp real4 ptr var_80
  loc_0040F32B: fnstsw ax
  loc_0040F32D: test ah, 41h
  loc_0040F330: jz 0040F33Ch
  loc_0040F332: mov var_C4, 00000000h
  loc_0040F33C: fld real4 ptr var_88
  loc_0040F342: fadd st0, real4 ptr var_84
  loc_0040F348: fnstsw ax
  loc_0040F34A: test al, 0Dh
  loc_0040F34C: jnz 0040FCD3h
  loc_0040F352: call [00401074h] ; __vbaFpR4
  loc_0040F358: mov eax, Y
  loc_0040F35B: fcomp real4 ptr [eax]
  loc_0040F35D: fnstsw ax
  loc_0040F35F: test ah, 41h
  loc_0040F362: jnz 0040F36Bh
  loc_0040F364: mov eax, 00000001h
  loc_0040F369: jmp 0040F36Dh
  loc_0040F36B: xor eax, eax
  loc_0040F36D: mov ecx, var_C4
  loc_0040F373: neg eax
  loc_0040F375: neg ecx
  loc_0040F377: and eax, ecx
  loc_0040F379: mov ecx, var_C0
  loc_0040F37F: neg ecx
  loc_0040F381: and eax, ecx
  loc_0040F383: mov ecx, var_BC
  loc_0040F389: neg ecx
  loc_0040F38B: and eax, ecx
  loc_0040F38D: test ax, ax
  loc_0040F390: jz 0040F544h
  loc_0040F396: sub esp, 00000010h
  loc_0040F399: mov ecx, 00000003h
  loc_0040F39E: mov edx, esp
  loc_0040F3A0: xor eax, eax
  loc_0040F3A2: push 0000000Bh
  loc_0040F3A4: push esi
  loc_0040F3A5: mov [edx], ecx
  loc_0040F3A7: mov ecx, var_5C
  loc_0040F3AA: mov [edx+00000004h], ecx
  loc_0040F3AD: mov ecx, [esi]
  loc_0040F3AF: mov [edx+00000008h], eax
  loc_0040F3B2: mov eax, var_54
  loc_0040F3B5: mov [edx+0000000Ch], eax
  loc_0040F3B8: call [ecx+00000410h]
  loc_0040F3BE: lea edx, var_1C
  loc_0040F3C1: push eax
  loc_0040F3C2: push edx
  loc_0040F3C3: call edi
  loc_0040F3C5: push eax
  loc_0040F3C6: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F3CC: lea ecx, var_1C
  loc_0040F3CF: call ebx
  loc_0040F3D1: movsx eax, [esi+0000008Ch]
  loc_0040F3D8: sub esp, 00000010h
  loc_0040F3DB: mov ecx, 00000003h
  loc_0040F3E0: mov edx, esp
  loc_0040F3E2: push 0000000Ah
  loc_0040F3E4: push esi
  loc_0040F3E5: mov [edx], ecx
  loc_0040F3E7: mov ecx, var_5C
  loc_0040F3EA: mov [edx+00000004h], ecx
  loc_0040F3ED: mov ecx, [esi]
  loc_0040F3EF: mov [edx+00000008h], eax
  loc_0040F3F2: mov eax, var_54
  loc_0040F3F5: mov [edx+0000000Ch], eax
  loc_0040F3F8: call [ecx+00000410h]
  loc_0040F3FE: lea edx, var_1C
  loc_0040F401: push eax
  loc_0040F402: push edx
  loc_0040F403: call edi
  loc_0040F405: push eax
  loc_0040F406: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F40C: lea ecx, var_1C
  loc_0040F40F: call ebx
  loc_0040F411: sub esp, 00000010h
  loc_0040F414: mov ecx, 00000003h
  loc_0040F419: mov edx, esp
  loc_0040F41B: mov eax, 000000FFh
  loc_0040F420: push 00000026h
  loc_0040F422: push esi
  loc_0040F423: mov [edx], ecx
  loc_0040F425: mov ecx, var_5C
  loc_0040F428: mov [edx+00000004h], ecx
  loc_0040F42B: mov ecx, [esi]
  loc_0040F42D: mov [edx+00000008h], eax
  loc_0040F430: mov eax, var_54
  loc_0040F433: mov [edx+0000000Ch], eax
  loc_0040F436: call [ecx+00000410h]
  loc_0040F43C: lea edx, var_1C
  loc_0040F43F: push eax
  loc_0040F440: push edx
  loc_0040F441: call edi
  loc_0040F443: push eax
  loc_0040F444: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F44A: lea ecx, var_1C
  loc_0040F44D: call ebx
  loc_0040F44F: sub esp, 00000010h
  loc_0040F452: mov ecx, 00000003h
  loc_0040F457: mov edx, esp
  loc_0040F459: xor eax, eax
  loc_0040F45B: push 0000000Bh
  loc_0040F45D: push esi
  loc_0040F45E: mov [edx], ecx
  loc_0040F460: mov ecx, var_5C
  loc_0040F463: mov [edx+00000004h], ecx
  loc_0040F466: mov ecx, [esi]
  loc_0040F468: mov [edx+00000008h], eax
  loc_0040F46B: mov eax, var_54
  loc_0040F46E: mov [edx+0000000Ch], eax
  loc_0040F471: call [ecx+0000040Ch]
  loc_0040F477: push eax
  loc_0040F478: lea edx, var_1C
  loc_0040F47B: push edx
  loc_0040F47C: call edi
  loc_0040F47E: push eax
  loc_0040F47F: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F485: lea ecx, var_1C
  loc_0040F488: call ebx
  loc_0040F48A: movsx eax, [esi+0000008Ch]
  loc_0040F491: sub esp, 00000010h
  loc_0040F494: mov ecx, 00000003h
  loc_0040F499: mov edx, esp
  loc_0040F49B: push 0000000Ah
  loc_0040F49D: push esi
  loc_0040F49E: mov [edx], ecx
  loc_0040F4A0: mov ecx, var_5C
  loc_0040F4A3: mov [edx+00000004h], ecx
  loc_0040F4A6: mov ecx, [esi]
  loc_0040F4A8: mov [edx+00000008h], eax
  loc_0040F4AB: mov eax, var_54
  loc_0040F4AE: mov [edx+0000000Ch], eax
  loc_0040F4B1: call [ecx+0000040Ch]
  loc_0040F4B7: lea edx, var_1C
  loc_0040F4BA: push eax
  loc_0040F4BB: push edx
  loc_0040F4BC: call edi
  loc_0040F4BE: push eax
  loc_0040F4BF: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F4C5: lea ecx, var_1C
  loc_0040F4C8: call ebx
  loc_0040F4CA: sub esp, 00000010h
  loc_0040F4CD: mov ecx, 00000003h
  loc_0040F4D2: mov edx, esp
  loc_0040F4D4: mov eax, 000000FFh
  loc_0040F4D9: push 00000026h
  loc_0040F4DB: push esi
  loc_0040F4DC: mov [edx], ecx
  loc_0040F4DE: mov ecx, var_5C
  loc_0040F4E1: mov [edx+00000004h], ecx
  loc_0040F4E4: mov ecx, [esi]
  loc_0040F4E6: mov [edx+00000008h], eax
  loc_0040F4E9: mov eax, var_54
  loc_0040F4EC: mov [edx+0000000Ch], eax
  loc_0040F4EF: call [ecx+0000040Ch]
  loc_0040F4F5: lea edx, var_1C
  loc_0040F4F8: push eax
  loc_0040F4F9: push edx
  loc_0040F4FA: call edi
  loc_0040F4FC: push eax
  loc_0040F4FD: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F503: lea ecx, var_1C
  loc_0040F506: call ebx
  loc_0040F508: mov eax, var_A4
  loc_0040F50E: push 000000FFh
  loc_0040F513: push eax
  loc_0040F514: mov ecx, [eax]
  loc_0040F516: call [ecx+00000064h]
  loc_0040F519: test eax, eax
  loc_0040F51B: fnclex
  loc_0040F51D: jge 0040F534h
  loc_0040F51F: mov edx, var_A4
  loc_0040F525: push 00000064h
  loc_0040F527: push 00403C28h
  loc_0040F52C: push edx
  loc_0040F52D: push eax
  loc_0040F52E: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F534: mov ax, [esi+0000008Ch]
  loc_0040F53B: mov [esi+00000034h], ax
  loc_0040F53F: jmp 0040F6D9h
  loc_0040F544: sub esp, 00000010h
  loc_0040F547: mov ecx, 00000003h
  loc_0040F54C: mov edx, esp
  loc_0040F54E: xor eax, eax
  loc_0040F550: push 0000000Bh
  loc_0040F552: push esi
  loc_0040F553: mov [edx], ecx
  loc_0040F555: mov ecx, var_5C
  loc_0040F558: mov [edx+00000004h], ecx
  loc_0040F55B: mov ecx, [esi]
  loc_0040F55D: mov [edx+00000008h], eax
  loc_0040F560: mov eax, var_54
  loc_0040F563: mov [edx+0000000Ch], eax
  loc_0040F566: call [ecx+00000410h]
  loc_0040F56C: lea edx, var_1C
  loc_0040F56F: push eax
  loc_0040F570: push edx
  loc_0040F571: call edi
  loc_0040F573: push eax
  loc_0040F574: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F57A: lea ecx, var_1C
  loc_0040F57D: call ebx
  loc_0040F57F: movsx eax, [esi+0000008Ch]
  loc_0040F586: sub esp, 00000010h
  loc_0040F589: mov ecx, 00000003h
  loc_0040F58E: mov edx, esp
  loc_0040F590: push 0000000Ah
  loc_0040F592: push esi
  loc_0040F593: mov [edx], ecx
  loc_0040F595: mov ecx, var_5C
  loc_0040F598: mov [edx+00000004h], ecx
  loc_0040F59B: mov ecx, [esi]
  loc_0040F59D: mov [edx+00000008h], eax
  loc_0040F5A0: mov eax, var_54
  loc_0040F5A3: mov [edx+0000000Ch], eax
  loc_0040F5A6: call [ecx+00000410h]
  loc_0040F5AC: lea edx, var_1C
  loc_0040F5AF: push eax
  loc_0040F5B0: push edx
  loc_0040F5B1: call edi
  loc_0040F5B3: push eax
  loc_0040F5B4: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F5BA: lea ecx, var_1C
  loc_0040F5BD: call ebx
  loc_0040F5BF: sub esp, 00000010h
  loc_0040F5C2: mov ecx, 00000003h
  loc_0040F5C7: mov edx, esp
  loc_0040F5C9: xor eax, eax
  loc_0040F5CB: push 00000026h
  loc_0040F5CD: push esi
  loc_0040F5CE: mov [edx], ecx
  loc_0040F5D0: mov ecx, var_5C
  loc_0040F5D3: mov [edx+00000004h], ecx
  loc_0040F5D6: mov ecx, [esi]
  loc_0040F5D8: mov [edx+00000008h], eax
  loc_0040F5DB: mov eax, var_54
  loc_0040F5DE: mov [edx+0000000Ch], eax
  loc_0040F5E1: call [ecx+00000410h]
  loc_0040F5E7: lea edx, var_1C
  loc_0040F5EA: push eax
  loc_0040F5EB: push edx
  loc_0040F5EC: call edi
  loc_0040F5EE: push eax
  loc_0040F5EF: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F5F5: lea ecx, var_1C
  loc_0040F5F8: call ebx
  loc_0040F5FA: sub esp, 00000010h
  loc_0040F5FD: mov ecx, 00000003h
  loc_0040F602: mov edx, esp
  loc_0040F604: xor eax, eax
  loc_0040F606: push 0000000Bh
  loc_0040F608: push esi
  loc_0040F609: mov [edx], ecx
  loc_0040F60B: mov ecx, var_5C
  loc_0040F60E: mov [edx+00000004h], ecx
  loc_0040F611: mov ecx, [esi]
  loc_0040F613: mov [edx+00000008h], eax
  loc_0040F616: mov eax, var_54
  loc_0040F619: mov [edx+0000000Ch], eax
  loc_0040F61C: call [ecx+0000040Ch]
  loc_0040F622: push eax
  loc_0040F623: lea edx, var_1C
  loc_0040F626: push edx
  loc_0040F627: call edi
  loc_0040F629: push eax
  loc_0040F62A: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F630: lea ecx, var_1C
  loc_0040F633: call ebx
  loc_0040F635: movsx eax, [esi+0000008Ch]
  loc_0040F63C: sub esp, 00000010h
  loc_0040F63F: mov ecx, 00000003h
  loc_0040F644: mov edx, esp
  loc_0040F646: push 0000000Ah
  loc_0040F648: push esi
  loc_0040F649: mov [edx], ecx
  loc_0040F64B: mov ecx, var_5C
  loc_0040F64E: mov [edx+00000004h], ecx
  loc_0040F651: mov ecx, [esi]
  loc_0040F653: mov [edx+00000008h], eax
  loc_0040F656: mov eax, var_54
  loc_0040F659: mov [edx+0000000Ch], eax
  loc_0040F65C: call [ecx+0000040Ch]
  loc_0040F662: lea edx, var_1C
  loc_0040F665: push eax
  loc_0040F666: push edx
  loc_0040F667: call edi
  loc_0040F669: push eax
  loc_0040F66A: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F670: lea ecx, var_1C
  loc_0040F673: call ebx
  loc_0040F675: sub esp, 00000010h
  loc_0040F678: mov ecx, 00000003h
  loc_0040F67D: mov edx, esp
  loc_0040F67F: xor eax, eax
  loc_0040F681: push 00000026h
  loc_0040F683: push esi
  loc_0040F684: mov [edx], ecx
  loc_0040F686: mov ecx, var_5C
  loc_0040F689: mov [edx+00000004h], ecx
  loc_0040F68C: mov ecx, [esi]
  loc_0040F68E: mov [edx+00000008h], eax
  loc_0040F691: mov eax, var_54
  loc_0040F694: mov [edx+0000000Ch], eax
  loc_0040F697: call [ecx+0000040Ch]
  loc_0040F69D: lea edx, var_1C
  loc_0040F6A0: push eax
  loc_0040F6A1: push edx
  loc_0040F6A2: call edi
  loc_0040F6A4: push eax
  loc_0040F6A5: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F6AB: lea ecx, var_1C
  loc_0040F6AE: call ebx
  loc_0040F6B0: mov eax, var_A4
  loc_0040F6B6: push 00000000h
  loc_0040F6B8: push eax
  loc_0040F6B9: mov ecx, [eax]
  loc_0040F6BB: call [ecx+00000064h]
  loc_0040F6BE: test eax, eax
  loc_0040F6C0: fnclex
  loc_0040F6C2: jge 0040F6D9h
  loc_0040F6C4: mov edx, var_A4
  loc_0040F6CA: push 00000064h
  loc_0040F6CC: push 00403C28h
  loc_0040F6D1: push edx
  loc_0040F6D2: push eax
  loc_0040F6D3: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F6D9: lea eax, var_A4
  loc_0040F6DF: push 00000000h
  loc_0040F6E1: push eax
  loc_0040F6E2: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040F6E8: mov cx, [esi+0000008Ch]
  loc_0040F6EF: mov eax, 00000001h
  loc_0040F6F4: add cx, ax
  loc_0040F6F7: jo 0040FCD8h
  loc_0040F6FD: mov [esi+0000008Ch], cx
  loc_0040F704: xor eax, eax
  loc_0040F706: jmp 0040F142h
  loc_0040F70B: mov edi, [0040105Ch] ; __vbaObjSet
  loc_0040F711: mov ebx, [00401180h] ; __vbaFreeObj
  loc_0040F717: cmp [esi+00000036h], FFFFFFh
  loc_0040F71C: jnz 0040FC6Bh
  loc_0040F722: mov edx, [esi]
  loc_0040F724: push esi
  loc_0040F725: call [edx+00000314h]
  loc_0040F72B: push eax
  loc_0040F72C: lea eax, var_1C
  loc_0040F72F: push eax
  loc_0040F730: call edi
  loc_0040F732: mov ecx, [eax]
  loc_0040F734: lea edx, var_20
  loc_0040F737: push edx
  loc_0040F738: mov dx, [esi+00000034h]
  loc_0040F73C: push edx
  loc_0040F73D: push eax
  loc_0040F73E: mov var_8C, eax
  loc_0040F744: call [ecx+00000040h]
  loc_0040F747: test eax, eax
  loc_0040F749: fnclex
  loc_0040F74B: jge 0040F762h
  loc_0040F74D: mov ecx, var_8C
  loc_0040F753: push 00000040h
  loc_0040F755: push 00403BE8h
  loc_0040F75A: push ecx
  loc_0040F75B: push eax
  loc_0040F75C: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F762: mov eax, Y
  loc_0040F765: mov ecx, var_20
  loc_0040F768: push ecx
  loc_0040F769: mov var_94, ecx
  loc_0040F76F: fld real4 ptr [eax]
  loc_0040F771: fsub st0, real4 ptr [00401350h]
  loc_0040F777: mov edx, [ecx]
  loc_0040F779: fnstsw ax
  loc_0040F77B: test al, 0Dh
  loc_0040F77D: jnz 0040FCD3h
  loc_0040F783: fstp real4 ptr [esp]
  loc_0040F786: push ecx
  loc_0040F787: call [edx+00000074h]
  loc_0040F78A: test eax, eax
  loc_0040F78C: fnclex
  loc_0040F78E: jge 0040F7A5h
  loc_0040F790: mov ecx, var_94
  loc_0040F796: push 00000074h
  loc_0040F798: push 00403C28h
  loc_0040F79D: push ecx
  loc_0040F79E: push eax
  loc_0040F79F: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F7A5: lea edx, var_20
  loc_0040F7A8: lea eax, var_1C
  loc_0040F7AB: push edx
  loc_0040F7AC: push eax
  loc_0040F7AD: push 00000002h
  loc_0040F7AF: call [0040102Ch] ; __vbaFreeObjList
  loc_0040F7B5: mov ecx, [esi]
  loc_0040F7B7: add esp, 0000000Ch
  loc_0040F7BA: push esi
  loc_0040F7BB: call [ecx+00000314h]
  loc_0040F7C1: lea edx, var_1C
  loc_0040F7C4: push eax
  loc_0040F7C5: push edx
  loc_0040F7C6: call edi
  loc_0040F7C8: mov ecx, [eax]
  loc_0040F7CA: lea edx, var_20
  loc_0040F7CD: push edx
  loc_0040F7CE: mov dx, [esi+00000034h]
  loc_0040F7D2: push edx
  loc_0040F7D3: push eax
  loc_0040F7D4: mov var_8C, eax
  loc_0040F7DA: call [ecx+00000040h]
  loc_0040F7DD: test eax, eax
  loc_0040F7DF: fnclex
  loc_0040F7E1: jge 0040F7F8h
  loc_0040F7E3: mov ecx, var_8C
  loc_0040F7E9: push 00000040h
  loc_0040F7EB: push 00403BE8h
  loc_0040F7F0: push ecx
  loc_0040F7F1: push eax
  loc_0040F7F2: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F7F8: mov eax, X
  loc_0040F7FB: mov ecx, var_20
  loc_0040F7FE: push ecx
  loc_0040F7FF: mov var_94, ecx
  loc_0040F805: fld real4 ptr [eax]
  loc_0040F807: fsub st0, real4 ptr [00401350h]
  loc_0040F80D: mov edx, [ecx]
  loc_0040F80F: fnstsw ax
  loc_0040F811: test al, 0Dh
  loc_0040F813: jnz 0040FCD3h
  loc_0040F819: fstp real4 ptr [esp]
  loc_0040F81C: push ecx
  loc_0040F81D: call [edx+0000006Ch]
  loc_0040F820: test eax, eax
  loc_0040F822: fnclex
  loc_0040F824: jge 0040F83Bh
  loc_0040F826: mov ecx, var_94
  loc_0040F82C: push 0000006Ch
  loc_0040F82E: push 00403C28h
  loc_0040F833: push ecx
  loc_0040F834: push eax
  loc_0040F835: call [00401040h] ; __vbaHresultCheckObj
  loc_0040F83B: lea edx, var_20
  loc_0040F83E: lea eax, var_1C
  loc_0040F841: push edx
  loc_0040F842: push eax
  loc_0040F843: push 00000002h
  loc_0040F845: call [0040102Ch] ; __vbaFreeObjList
  loc_0040F84B: mov ecx, 00000003h
  loc_0040F850: xor eax, eax
  loc_0040F852: push ecx
  loc_0040F853: mov edx, esp
  loc_0040F855: push 0000000Bh
  loc_0040F857: push esi
  loc_0040F858: mov [edx], ecx
  loc_0040F85A: mov ecx, var_5C
  loc_0040F85D: mov [edx+00000004h], ecx
  loc_0040F860: mov ecx, [esi]
  loc_0040F862: mov [edx+00000008h], eax
  loc_0040F865: mov eax, var_54
  loc_0040F868: mov [edx+0000000Ch], eax
  loc_0040F86B: call [ecx+00000410h]
  loc_0040F871: lea edx, var_1C
  loc_0040F874: push eax
  loc_0040F875: push edx
  loc_0040F876: call edi
  loc_0040F878: push eax
  loc_0040F879: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F87F: lea ecx, var_1C
  loc_0040F882: call ebx
  loc_0040F884: sub esp, 00000010h
  loc_0040F887: mov ecx, 00000003h
  loc_0040F88C: mov edx, esp
  loc_0040F88E: mov eax, 00000001h
  loc_0040F893: push 0000000Bh
  loc_0040F895: push esi
  loc_0040F896: mov [edx], ecx
  loc_0040F898: mov ecx, var_5C
  loc_0040F89B: mov [edx+00000004h], ecx
  loc_0040F89E: mov ecx, [esi]
  loc_0040F8A0: mov [edx+00000008h], eax
  loc_0040F8A3: mov eax, var_54
  loc_0040F8A6: mov [edx+0000000Ch], eax
  loc_0040F8A9: call [ecx+00000410h]
  loc_0040F8AF: lea edx, var_1C
  loc_0040F8B2: push eax
  loc_0040F8B3: push edx
  loc_0040F8B4: call edi
  loc_0040F8B6: push eax
  loc_0040F8B7: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F8BD: lea ecx, var_1C
  loc_0040F8C0: call ebx
  loc_0040F8C2: mov eax, Y
  loc_0040F8C5: lea ecx, var_30
  loc_0040F8C8: push 00000002h
  loc_0040F8CA: lea edx, var_40
  loc_0040F8CD: fld real4 ptr [eax]
  loc_0040F8CF: fsub st0, real4 ptr [esi+00000064h]
  loc_0040F8D2: push ecx
  loc_0040F8D3: push edx
  loc_0040F8D4: mov var_30, 00000004h
  loc_0040F8DB: cmp [00415000h], 00000000h
  loc_0040F8E2: jnz 0040F8E9h
  loc_0040F8E4: fdiv st0, real4 ptr [esi+00000060h]
  loc_0040F8E7: jmp 0040F8F1h
  loc_0040F8E9: push [esi+00000060h]
  loc_0040F8EC: call 00401558h ; _adj_fdiv_m32
  loc_0040F8F1: fstp real4 ptr var_28
  loc_0040F8F4: fnstsw ax
  loc_0040F8F6: test al, 0Dh
  loc_0040F8F8: jnz 0040FCD3h
  loc_0040F8FE: call [004010E8h] ; rtcRound
  loc_0040F904: lea eax, var_40
  loc_0040F907: push eax
  loc_0040F908: call [00401034h] ; __vbaStrErrVarCopy
  loc_0040F90E: sub esp, 00000010h
  loc_0040F911: mov ecx, 00000008h
  loc_0040F916: mov edx, esp
  loc_0040F918: mov var_50, ecx
  loc_0040F91B: mov var_48, eax
  loc_0040F91E: push 00000000h
  loc_0040F920: mov [edx], ecx
  loc_0040F922: mov ecx, var_4C
  loc_0040F925: mov [edx+00000004h], ecx
  loc_0040F928: mov [edx+00000008h], eax
  loc_0040F92B: mov eax, var_44
  loc_0040F92E: mov [edx+0000000Ch], eax
  loc_0040F931: mov ecx, [esi]
  loc_0040F933: push esi
  loc_0040F934: call [ecx+00000410h]
  loc_0040F93A: lea edx, var_1C
  loc_0040F93D: push eax
  loc_0040F93E: push edx
  loc_0040F93F: call edi
  loc_0040F941: push eax
  loc_0040F942: call [0040116Ch] ; __vbaLateIdSt
  loc_0040F948: lea ecx, var_1C
  loc_0040F94B: call ebx
  loc_0040F94D: lea eax, var_50
  loc_0040F950: lea ecx, var_40
  loc_0040F953: push eax
  loc_0040F954: lea edx, var_40
  loc_0040F957: push ecx
  loc_0040F958: lea eax, var_30
  loc_0040F95B: push edx
  loc_0040F95C: push eax
  loc_0040F95D: push 00000004h
  loc_0040F95F: call [00401024h] ; __vbaFreeVarList
  loc_0040F965: mov eax, [esi+0000003Ch]
  loc_0040F968: add esp, 00000014h
  loc_0040F96B: test eax, eax
  loc_0040F96D: jz 0040F9CFh
  loc_0040F96F: cmp [eax], 0002h
  loc_0040F973: jnz 0040F9CFh
  loc_0040F975: movsx edx, [esi+00000034h]
  loc_0040F979: mov ecx, [eax+0000001Ch]
  loc_0040F97C: sub edx, ecx
  loc_0040F97E: mov ecx, [eax+00000018h]
  loc_0040F981: cmp edx, ecx
  loc_0040F983: mov var_90, edx
  loc_0040F989: jb 0040F997h
  loc_0040F98B: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040F991: mov edx, var_90
  loc_0040F997: mov eax, [esi+0000003Ch]
  loc_0040F99A: mov ecx, 00000001h
  loc_0040F99F: sub ecx, [eax+00000014h]
  loc_0040F9A2: cmp ecx, [eax+00000010h]
  loc_0040F9A5: mov var_8C, ecx
  loc_0040F9AB: jb 0040F9BFh
  loc_0040F9AD: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040F9B3: mov ecx, var_8C
  loc_0040F9B9: mov edx, var_90
  loc_0040F9BF: mov eax, [esi+0000003Ch]
  loc_0040F9C2: mov eax, [eax+00000018h]
  loc_0040F9C5: imul eax, ecx
  loc_0040F9C8: add eax, edx
  loc_0040F9CA: shl eax, 02h
  loc_0040F9CD: jmp 0040F9D5h
  loc_0040F9CF: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040F9D5: mov ecx, [esi]
  loc_0040F9D7: push 00000000h
  loc_0040F9D9: push 00000000h
  loc_0040F9DB: push esi
  loc_0040F9DC: mov var_C8, eax
  loc_0040F9E2: call [ecx+00000410h]
  loc_0040F9E8: lea edx, var_1C
  loc_0040F9EB: push eax
  loc_0040F9EC: push edx
  loc_0040F9ED: call edi
  loc_0040F9EF: push eax
  loc_0040F9F0: lea eax, var_30
  loc_0040F9F3: push eax
  loc_0040F9F4: call [004010B4h] ; __vbaLateIdCallLd
  loc_0040F9FA: add esp, 00000010h
  loc_0040F9FD: push eax
  loc_0040F9FE: call [00401018h] ; __vbaStrVarMove
  loc_0040FA04: mov edx, eax
  loc_0040FA06: lea ecx, var_18
  loc_0040FA09: call [00401164h] ; __vbaStrMove
  loc_0040FA0F: push eax
  loc_0040FA10: call [004010A0h] ; __vbaR4Str
  loc_0040FA16: mov ecx, [esi+0000003Ch]
  loc_0040FA19: mov eax, var_C8
  loc_0040FA1F: mov edx, [ecx+0000000Ch]
  loc_0040FA22: lea ecx, var_18
  loc_0040FA25: fstp real4 ptr [edx+eax]
  loc_0040FA28: call [00401184h] ; __vbaFreeStr
  loc_0040FA2E: lea ecx, var_1C
  loc_0040FA31: call ebx
  loc_0040FA33: lea ecx, var_30
  loc_0040FA36: call [00401014h] ; __vbaFreeVar
  loc_0040FA3C: sub esp, 00000010h
  loc_0040FA3F: mov ecx, 00000003h
  loc_0040FA44: mov edx, esp
  loc_0040FA46: mov eax, 00000002h
  loc_0040FA4B: push 0000000Bh
  loc_0040FA4D: push esi
  loc_0040FA4E: mov [edx], ecx
  loc_0040FA50: mov ecx, var_5C
  loc_0040FA53: mov [edx+00000004h], ecx
  loc_0040FA56: mov ecx, [esi]
  loc_0040FA58: mov [edx+00000008h], eax
  loc_0040FA5B: mov eax, var_54
  loc_0040FA5E: mov [edx+0000000Ch], eax
  loc_0040FA61: call [ecx+00000410h]
  loc_0040FA67: lea edx, var_1C
  loc_0040FA6A: push eax
  loc_0040FA6B: push edx
  loc_0040FA6C: call edi
  loc_0040FA6E: push eax
  loc_0040FA6F: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FA75: lea ecx, var_1C
  loc_0040FA78: call ebx
  loc_0040FA7A: mov eax, X
  loc_0040FA7D: lea ecx, var_30
  loc_0040FA80: push 00000002h
  loc_0040FA82: lea edx, var_40
  loc_0040FA85: fld real4 ptr [eax]
  loc_0040FA87: fsub st0, real4 ptr [00401350h]
  loc_0040FA8D: push ecx
  loc_0040FA8E: push edx
  loc_0040FA8F: mov var_30, 00000004h
  loc_0040FA96: fsub st0, real4 ptr [esi+0000005Ch]
  loc_0040FA99: cmp [00415000h], 00000000h
  loc_0040FAA0: jnz 0040FAA7h
  loc_0040FAA2: fdiv st0, real4 ptr [esi+00000058h]
  loc_0040FAA5: jmp 0040FAAFh
  loc_0040FAA7: push [esi+00000058h]
  loc_0040FAAA: call 00401558h ; _adj_fdiv_m32
  loc_0040FAAF: fstp real4 ptr var_28
  loc_0040FAB2: fnstsw ax
  loc_0040FAB4: test al, 0Dh
  loc_0040FAB6: jnz 0040FCD3h
  loc_0040FABC: call [004010E8h] ; rtcRound
  loc_0040FAC2: lea eax, var_40
  loc_0040FAC5: push eax
  loc_0040FAC6: call [00401034h] ; __vbaStrErrVarCopy
  loc_0040FACC: sub esp, 00000010h
  loc_0040FACF: mov ecx, 00000008h
  loc_0040FAD4: mov edx, esp
  loc_0040FAD6: mov var_48, eax
  loc_0040FAD9: mov var_50, ecx
  loc_0040FADC: mov [edx], ecx
  loc_0040FADE: mov ecx, var_4C
  loc_0040FAE1: push 00000000h
  loc_0040FAE3: mov [edx+00000004h], ecx
  loc_0040FAE6: mov ecx, [esi]
  loc_0040FAE8: push esi
  loc_0040FAE9: mov [edx+00000008h], eax
  loc_0040FAEC: mov eax, var_44
  loc_0040FAEF: mov [edx+0000000Ch], eax
  loc_0040FAF2: call [ecx+00000410h]
  loc_0040FAF8: lea edx, var_1C
  loc_0040FAFB: push eax
  loc_0040FAFC: push edx
  loc_0040FAFD: call edi
  loc_0040FAFF: push eax
  loc_0040FB00: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FB06: lea ecx, var_1C
  loc_0040FB09: call ebx
  loc_0040FB0B: lea eax, var_50
  loc_0040FB0E: lea ecx, var_40
  loc_0040FB11: push eax
  loc_0040FB12: lea edx, var_40
  loc_0040FB15: push ecx
  loc_0040FB16: lea eax, var_30
  loc_0040FB19: push edx
  loc_0040FB1A: push eax
  loc_0040FB1B: push 00000004h
  loc_0040FB1D: call [00401024h] ; __vbaFreeVarList
  loc_0040FB23: mov eax, [esi+0000003Ch]
  loc_0040FB26: add esp, 00000014h
  loc_0040FB29: test eax, eax
  loc_0040FB2B: jz 0040FB8Dh
  loc_0040FB2D: cmp [eax], 0002h
  loc_0040FB31: jnz 0040FB8Dh
  loc_0040FB33: movsx edx, [esi+00000034h]
  loc_0040FB37: mov ecx, [eax+0000001Ch]
  loc_0040FB3A: sub edx, ecx
  loc_0040FB3C: mov ecx, [eax+00000018h]
  loc_0040FB3F: cmp edx, ecx
  loc_0040FB41: mov var_90, edx
  loc_0040FB47: jb 0040FB55h
  loc_0040FB49: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040FB4F: mov edx, var_90
  loc_0040FB55: mov eax, [esi+0000003Ch]
  loc_0040FB58: mov ecx, 00000002h
  loc_0040FB5D: sub ecx, [eax+00000014h]
  loc_0040FB60: cmp ecx, [eax+00000010h]
  loc_0040FB63: mov var_8C, ecx
  loc_0040FB69: jb 0040FB7Dh
  loc_0040FB6B: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040FB71: mov ecx, var_8C
  loc_0040FB77: mov edx, var_90
  loc_0040FB7D: mov eax, [esi+0000003Ch]
  loc_0040FB80: mov eax, [eax+00000018h]
  loc_0040FB83: imul eax, ecx
  loc_0040FB86: add eax, edx
  loc_0040FB88: shl eax, 02h
  loc_0040FB8B: jmp 0040FB93h
  loc_0040FB8D: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040FB93: mov ecx, [esi]
  loc_0040FB95: push 00000000h
  loc_0040FB97: push 00000000h
  loc_0040FB99: push esi
  loc_0040FB9A: mov var_CC, eax
  loc_0040FBA0: call [ecx+00000410h]
  loc_0040FBA6: lea edx, var_1C
  loc_0040FBA9: push eax
  loc_0040FBAA: push edx
  loc_0040FBAB: call edi
  loc_0040FBAD: push eax
  loc_0040FBAE: lea eax, var_30
  loc_0040FBB1: push eax
  loc_0040FBB2: call [004010B4h] ; __vbaLateIdCallLd
  loc_0040FBB8: add esp, 00000010h
  loc_0040FBBB: push eax
  loc_0040FBBC: call [00401018h] ; __vbaStrVarMove
  loc_0040FBC2: mov edx, eax
  loc_0040FBC4: lea ecx, var_18
  loc_0040FBC7: call [00401164h] ; __vbaStrMove
  loc_0040FBCD: push eax
  loc_0040FBCE: call [004010A0h] ; __vbaR4Str
  loc_0040FBD4: mov ecx, [esi+0000003Ch]
  loc_0040FBD7: mov eax, var_CC
  loc_0040FBDD: lea edi, [esi+0000003Ch]
  loc_0040FBE0: mov edx, [ecx+0000000Ch]
  loc_0040FBE3: lea ecx, var_18
  loc_0040FBE6: fstp real4 ptr [edx+eax]
  loc_0040FBE9: call [00401184h] ; __vbaFreeStr
  loc_0040FBEF: lea ecx, var_1C
  loc_0040FBF2: call ebx
  loc_0040FBF4: lea ecx, var_30
  loc_0040FBF7: call [00401014h] ; __vbaFreeVar
  loc_0040FBFD: mov ecx, [00415028h]
  loc_0040FC03: mov edx, [0041502Ch]
  loc_0040FC09: push ecx
  loc_0040FC0A: push edx
  loc_0040FC0B: call [004010A0h] ; __vbaR4Str
  loc_0040FC11: push ecx
  loc_0040FC12: fstp real4 ptr [esp]
  loc_0040FC15: push 00000001h
  loc_0040FC17: push 00000002h
  loc_0040FC19: push edi
  loc_0040FC1A: call 00409C80h
  loc_0040FC1F: mov eax, [esi]
  loc_0040FC21: push esi
  loc_0040FC22: call [eax+00000780h]
  loc_0040FC28: mov ecx, [esi]
  loc_0040FC2A: push esi
  loc_0040FC2B: call [ecx+0000078Ch]
  loc_0040FC31: lea edx, [esi+00000080h]
  loc_0040FC37: lea eax, [esi+00000088h]
  loc_0040FC3D: push edx
  loc_0040FC3E: lea ecx, [esi+00000084h]
  loc_0040FC44: push eax
  loc_0040FC45: lea edx, [esi+0000007Ch]
  loc_0040FC48: push ecx
  loc_0040FC49: lea eax, [esi+00000078h]
  loc_0040FC4C: push edx
  loc_0040FC4D: lea ecx, [esi+00000074h]
  loc_0040FC50: push eax
  loc_0040FC51: push ecx
  loc_0040FC52: call 00409DD0h
  loc_0040FC57: mov edx, [esi]
  loc_0040FC59: push esi
  loc_0040FC5A: call [edx+0000071Ch]
  loc_0040FC60: mov eax, [esi]
  loc_0040FC62: push esi
  loc_0040FC63: call [eax+00000718h]
  loc_0040FC69: xor eax, eax
  loc_0040FC6B: mov var_4, eax
  loc_0040FC6E: fwait
  loc_0040FC6F: push 0040FCB4h
  loc_0040FC74: jmp 0040FCA7h
  loc_0040FC76: lea ecx, var_18
  loc_0040FC79: call [00401184h] ; __vbaFreeStr
  loc_0040FC7F: lea ecx, var_20
  loc_0040FC82: lea edx, var_1C
  loc_0040FC85: push ecx
  loc_0040FC86: push edx
  loc_0040FC87: push 00000002h
  loc_0040FC89: call [0040102Ch] ; __vbaFreeObjList
  loc_0040FC8F: lea eax, var_50
  loc_0040FC92: lea ecx, var_40
  loc_0040FC95: push eax
  loc_0040FC96: lea edx, var_30
  loc_0040FC99: push ecx
  loc_0040FC9A: push edx
  loc_0040FC9B: push 00000003h
  loc_0040FC9D: call [00401024h] ; __vbaFreeVarList
  loc_0040FCA3: add esp, 0000001Ch
  loc_0040FCA6: ret
  loc_0040FCA7: lea ecx, var_A4
  loc_0040FCAD: call [00401180h] ; __vbaFreeObj
  loc_0040FCB3: ret
  loc_0040FCB4: mov eax, Me
  loc_0040FCB7: push eax
  loc_0040FCB8: mov ecx, [eax]
  loc_0040FCBA: call [ecx+00000008h]
  loc_0040FCBD: mov eax, var_4
  loc_0040FCC0: mov ecx, var_14
  loc_0040FCC3: pop edi
  loc_0040FCC4: pop esi
  loc_0040FCC5: mov fs:[00000000h], ecx
  loc_0040FCCC: pop ebx
  loc_0040FCCD: mov esp, ebp
  loc_0040FCCF: pop ebp
  loc_0040FCD0: retn 0014h
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '40FCE0
  loc_0040FCE0: push ebp
  loc_0040FCE1: mov ebp, esp
  loc_0040FCE3: sub esp, 0000000Ch
  loc_0040FCE6: push 00401546h ; __vbaExceptHandler
  loc_0040FCEB: mov eax, fs:[00000000h]
  loc_0040FCF1: push eax
  loc_0040FCF2: mov fs:[00000000h], esp
  loc_0040FCF9: sub esp, 0000002Ch
  loc_0040FCFC: push ebx
  loc_0040FCFD: push esi
  loc_0040FCFE: push edi
  loc_0040FCFF: mov var_C, esp
  loc_0040FD02: mov var_8, 00401368h
  loc_0040FD09: mov esi, Me
  loc_0040FD0C: mov eax, esi
  loc_0040FD0E: and eax, 00000001h
  loc_0040FD11: mov var_4, eax
  loc_0040FD14: and esi, FFFFFFFEh
  loc_0040FD17: push esi
  loc_0040FD18: mov Me, esi
  loc_0040FD1B: mov ecx, [esi]
  loc_0040FD1D: call [ecx+00000004h]
  loc_0040FD20: mov edx, Button
  loc_0040FD23: xor eax, eax
  loc_0040FD25: mov var_18, eax
  loc_0040FD28: mov var_28, eax
  loc_0040FD2B: cmp [edx], 0001h
  loc_0040FD2F: jnz 0040FEB1h
  loc_0040FD35: mov edi, var_24
  loc_0040FD38: sub esp, 00000010h
  loc_0040FD3B: mov edx, esp
  loc_0040FD3D: mov ecx, 00000003h
  loc_0040FD42: mov ebx, var_1C
  loc_0040FD45: mov [esi+00000036h], ax
  loc_0040FD49: mov [edx], ecx
  loc_0040FD4B: push 0000000Bh
  loc_0040FD4D: push esi
  loc_0040FD4E: mov [edx+00000004h], edi
  loc_0040FD51: mov [edx+00000008h], eax
  loc_0040FD54: mov eax, [esi]
  loc_0040FD56: mov [edx+0000000Ch], ebx
  loc_0040FD59: call [eax+00000410h]
  loc_0040FD5F: lea ecx, var_18
  loc_0040FD62: push eax
  loc_0040FD63: push ecx
  loc_0040FD64: call [0040105Ch] ; __vbaObjSet
  loc_0040FD6A: push eax
  loc_0040FD6B: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FD71: lea ecx, var_18
  loc_0040FD74: call [00401180h] ; __vbaFreeObj
  loc_0040FD7A: movsx eax, [esi+00000034h]
  loc_0040FD7E: sub esp, 00000010h
  loc_0040FD81: mov ecx, 00000003h
  loc_0040FD86: mov edx, esp
  loc_0040FD88: push 0000000Ah
  loc_0040FD8A: push esi
  loc_0040FD8B: mov [edx], ecx
  loc_0040FD8D: mov [edx+00000004h], edi
  loc_0040FD90: mov [edx+00000008h], eax
  loc_0040FD93: mov eax, [esi]
  loc_0040FD95: mov [edx+0000000Ch], ebx
  loc_0040FD98: call [eax+00000410h]
  loc_0040FD9E: lea ecx, var_18
  loc_0040FDA1: push eax
  loc_0040FDA2: push ecx
  loc_0040FDA3: call [0040105Ch] ; __vbaObjSet
  loc_0040FDA9: push eax
  loc_0040FDAA: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FDB0: lea ecx, var_18
  loc_0040FDB3: call [00401180h] ; __vbaFreeObj
  loc_0040FDB9: sub esp, 00000010h
  loc_0040FDBC: mov ecx, 00000003h
  loc_0040FDC1: mov edx, esp
  loc_0040FDC3: xor eax, eax
  loc_0040FDC5: push 00000026h
  loc_0040FDC7: push esi
  loc_0040FDC8: mov [edx], ecx
  loc_0040FDCA: mov [edx+00000004h], edi
  loc_0040FDCD: mov [edx+00000008h], eax
  loc_0040FDD0: mov eax, [esi]
  loc_0040FDD2: mov [edx+0000000Ch], ebx
  loc_0040FDD5: call [eax+00000410h]
  loc_0040FDDB: lea ecx, var_18
  loc_0040FDDE: push eax
  loc_0040FDDF: push ecx
  loc_0040FDE0: call [0040105Ch] ; __vbaObjSet
  loc_0040FDE6: push eax
  loc_0040FDE7: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FDED: lea ecx, var_18
  loc_0040FDF0: call [00401180h] ; __vbaFreeObj
  loc_0040FDF6: sub esp, 00000010h
  loc_0040FDF9: mov ecx, 00000003h
  loc_0040FDFE: mov edx, esp
  loc_0040FE00: xor eax, eax
  loc_0040FE02: push 0000000Bh
  loc_0040FE04: push esi
  loc_0040FE05: mov [edx], ecx
  loc_0040FE07: mov [edx+00000004h], edi
  loc_0040FE0A: mov [edx+00000008h], eax
  loc_0040FE0D: mov eax, [esi]
  loc_0040FE0F: mov [edx+0000000Ch], ebx
  loc_0040FE12: call [eax+0000040Ch]
  loc_0040FE18: lea ecx, var_18
  loc_0040FE1B: push eax
  loc_0040FE1C: push ecx
  loc_0040FE1D: call [0040105Ch] ; __vbaObjSet
  loc_0040FE23: push eax
  loc_0040FE24: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FE2A: lea ecx, var_18
  loc_0040FE2D: call [00401180h] ; __vbaFreeObj
  loc_0040FE33: movsx eax, [esi+00000034h]
  loc_0040FE37: sub esp, 00000010h
  loc_0040FE3A: mov ecx, 00000003h
  loc_0040FE3F: mov edx, esp
  loc_0040FE41: push 0000000Ah
  loc_0040FE43: push esi
  loc_0040FE44: mov [edx], ecx
  loc_0040FE46: mov [edx+00000004h], edi
  loc_0040FE49: mov [edx+00000008h], eax
  loc_0040FE4C: mov eax, [esi]
  loc_0040FE4E: mov [edx+0000000Ch], ebx
  loc_0040FE51: call [eax+0000040Ch]
  loc_0040FE57: lea ecx, var_18
  loc_0040FE5A: push eax
  loc_0040FE5B: push ecx
  loc_0040FE5C: call [0040105Ch] ; __vbaObjSet
  loc_0040FE62: push eax
  loc_0040FE63: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FE69: lea ecx, var_18
  loc_0040FE6C: call [00401180h] ; __vbaFreeObj
  loc_0040FE72: sub esp, 00000010h
  loc_0040FE75: mov ecx, 00000003h
  loc_0040FE7A: mov edx, esp
  loc_0040FE7C: xor eax, eax
  loc_0040FE7E: push 00000026h
  loc_0040FE80: push esi
  loc_0040FE81: mov [edx], ecx
  loc_0040FE83: mov [edx+00000004h], edi
  loc_0040FE86: mov [edx+00000008h], eax
  loc_0040FE89: mov eax, [esi]
  loc_0040FE8B: mov [edx+0000000Ch], ebx
  loc_0040FE8E: call [eax+0000040Ch]
  loc_0040FE94: lea ecx, var_18
  loc_0040FE97: push eax
  loc_0040FE98: push ecx
  loc_0040FE99: call [0040105Ch] ; __vbaObjSet
  loc_0040FE9F: push eax
  loc_0040FEA0: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FEA6: lea ecx, var_18
  loc_0040FEA9: call [00401180h] ; __vbaFreeObj
  loc_0040FEAF: xor eax, eax
  loc_0040FEB1: mov var_4, eax
  loc_0040FEB4: push 0040FEC6h
  loc_0040FEB9: jmp 0040FEC5h
  loc_0040FEBB: lea ecx, var_18
  loc_0040FEBE: call [00401180h] ; __vbaFreeObj
  loc_0040FEC4: ret
  loc_0040FEC5: ret
  loc_0040FEC6: mov eax, Me
  loc_0040FEC9: push eax
  loc_0040FECA: mov edx, [eax]
  loc_0040FECC: call [edx+00000008h]
  loc_0040FECF: mov eax, var_4
  loc_0040FED2: mov ecx, var_14
  loc_0040FED5: pop edi
  loc_0040FED6: pop esi
  loc_0040FED7: mov fs:[00000000h], ecx
  loc_0040FEDE: pop ebx
  loc_0040FEDF: mov esp, ebp
  loc_0040FEE1: pop ebp
  loc_0040FEE2: retn 0014h
End Sub

Private Sub mnuShowXg_Click() '40EA50
  loc_0040EA50: push ebp
  loc_0040EA51: mov ebp, esp
  loc_0040EA53: sub esp, 0000000Ch
  loc_0040EA56: push 00401546h ; __vbaExceptHandler
  loc_0040EA5B: mov eax, fs:[00000000h]
  loc_0040EA61: push eax
  loc_0040EA62: mov fs:[00000000h], esp
  loc_0040EA69: sub esp, 00000024h
  loc_0040EA6C: push ebx
  loc_0040EA6D: push esi
  loc_0040EA6E: push edi
  loc_0040EA6F: mov var_C, esp
  loc_0040EA72: mov var_8, 00401320h
  loc_0040EA79: mov esi, Me
  loc_0040EA7C: mov eax, esi
  loc_0040EA7E: and eax, 00000001h
  loc_0040EA81: mov var_4, eax
  loc_0040EA84: and esi, FFFFFFFEh
  loc_0040EA87: push esi
  loc_0040EA88: mov Me, esi
  loc_0040EA8B: mov ecx, [esi]
  loc_0040EA8D: call [ecx+00000004h]
  loc_0040EA90: mov edx, [esi]
  loc_0040EA92: xor eax, eax
  loc_0040EA94: push esi
  loc_0040EA95: mov var_18, eax
  loc_0040EA98: mov var_1C, eax
  loc_0040EA9B: mov var_20, eax
  loc_0040EA9E: call [edx+000003DCh]
  loc_0040EAA4: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_0040EAAA: push eax
  loc_0040EAAB: lea eax, var_1C
  loc_0040EAAE: push eax
  loc_0040EAAF: call ebx
  loc_0040EAB1: mov ecx, [esi]
  loc_0040EAB3: push esi
  loc_0040EAB4: mov var_2C, eax
  loc_0040EAB7: call [ecx+000003DCh]
  loc_0040EABD: lea edx, var_18
  loc_0040EAC0: push eax
  loc_0040EAC1: push edx
  loc_0040EAC2: call ebx
  loc_0040EAC4: mov edi, eax
  loc_0040EAC6: lea ecx, var_20
  loc_0040EAC9: push ecx
  loc_0040EACA: push edi
  loc_0040EACB: mov eax, [edi]
  loc_0040EACD: call [eax+00000068h]
  loc_0040EAD0: test eax, eax
  loc_0040EAD2: fnclex
  loc_0040EAD4: jge 0040EAE5h
  loc_0040EAD6: push 00000068h
  loc_0040EAD8: push 00403C38h
  loc_0040EADD: push edi
  loc_0040EADE: push eax
  loc_0040EADF: call [00401040h] ; __vbaHresultCheckObj
  loc_0040EAE5: mov eax, var_20
  loc_0040EAE8: mov edi, var_2C
  loc_0040EAEB: not eax
  loc_0040EAED: mov edx, [edi]
  loc_0040EAEF: push eax
  loc_0040EAF0: push edi
  loc_0040EAF1: call [edx+0000006Ch]
  loc_0040EAF4: test eax, eax
  loc_0040EAF6: fnclex
  loc_0040EAF8: jge 0040EB09h
  loc_0040EAFA: push 0000006Ch
  loc_0040EAFC: push 00403C38h
  loc_0040EB01: push edi
  loc_0040EB02: push eax
  loc_0040EB03: call [00401040h] ; __vbaHresultCheckObj
  loc_0040EB09: lea ecx, var_1C
  loc_0040EB0C: lea edx, var_18
  loc_0040EB0F: push ecx
  loc_0040EB10: push edx
  loc_0040EB11: push 00000002h
  loc_0040EB13: call [0040102Ch] ; __vbaFreeObjList
  loc_0040EB19: mov eax, [esi]
  loc_0040EB1B: add esp, 0000000Ch
  loc_0040EB1E: push esi
  loc_0040EB1F: call [eax+000003DCh]
  loc_0040EB25: lea ecx, var_18
  loc_0040EB28: push eax
  loc_0040EB29: push ecx
  loc_0040EB2A: call ebx
  loc_0040EB2C: mov edi, eax
  loc_0040EB2E: lea eax, var_20
  loc_0040EB31: push eax
  loc_0040EB32: push edi
  loc_0040EB33: mov edx, [edi]
  loc_0040EB35: call [edx+00000068h]
  loc_0040EB38: test eax, eax
  loc_0040EB3A: fnclex
  loc_0040EB3C: jge 0040EB4Dh
  loc_0040EB3E: push 00000068h
  loc_0040EB40: push 00403C38h
  loc_0040EB45: push edi
  loc_0040EB46: push eax
  loc_0040EB47: call [00401040h] ; __vbaHresultCheckObj
  loc_0040EB4D: mov cx, var_20
  loc_0040EB51: mov [esi+0000006Eh], cx
  loc_0040EB55: lea ecx, var_18
  loc_0040EB58: call [00401180h] ; __vbaFreeObj
  loc_0040EB5E: mov edx, [esi]
  loc_0040EB60: push esi
  loc_0040EB61: call [edx+0000030Ch]
  loc_0040EB67: push eax
  loc_0040EB68: lea eax, var_18
  loc_0040EB6B: push eax
  loc_0040EB6C: call ebx
  loc_0040EB6E: mov dx, [esi+0000006Eh]
  loc_0040EB72: mov edi, eax
  loc_0040EB74: push edx
  loc_0040EB75: push edi
  loc_0040EB76: mov ecx, [edi]
  loc_0040EB78: call [ecx+0000008Ch]
  loc_0040EB7E: test eax, eax
  loc_0040EB80: fnclex
  loc_0040EB82: jge 0040EB96h
  loc_0040EB84: push 0000008Ch
  loc_0040EB89: push 00403C28h
  loc_0040EB8E: push edi
  loc_0040EB8F: push eax
  loc_0040EB90: call [00401040h] ; __vbaHresultCheckObj
  loc_0040EB96: lea ecx, var_18
  loc_0040EB99: call [00401180h] ; __vbaFreeObj
  loc_0040EB9F: cmp [esi+0000006Eh], FFFFFFh
  loc_0040EBA4: jnz 0040EBAFh
  loc_0040EBA6: mov eax, [esi]
  loc_0040EBA8: push esi
  loc_0040EBA9: call [eax+00000790h]
  loc_0040EBAF: mov var_4, 00000000h
  loc_0040EBB6: push 0040EBD2h
  loc_0040EBBB: jmp 0040EBD1h
  loc_0040EBBD: lea ecx, var_1C
  loc_0040EBC0: lea edx, var_18
  loc_0040EBC3: push ecx
  loc_0040EBC4: push edx
  loc_0040EBC5: push 00000002h
  loc_0040EBC7: call [0040102Ch] ; __vbaFreeObjList
  loc_0040EBCD: add esp, 0000000Ch
  loc_0040EBD0: ret
  loc_0040EBD1: ret
  loc_0040EBD2: mov eax, Me
  loc_0040EBD5: push eax
  loc_0040EBD6: mov ecx, [eax]
  loc_0040EBD8: call [ecx+00000008h]
  loc_0040EBDB: mov eax, var_4
  loc_0040EBDE: mov ecx, var_14
  loc_0040EBE1: pop edi
  loc_0040EBE2: pop esi
  loc_0040EBE3: mov fs:[00000000h], ecx
  loc_0040EBEA: pop ebx
  loc_0040EBEB: mov esp, ebp
  loc_0040EBED: pop ebp
  loc_0040EBEE: retn 0004h
End Sub

Private Sub mnuShowEst_Click() '40E3B0
  loc_0040E3B0: push ebp
  loc_0040E3B1: mov ebp, esp
  loc_0040E3B3: sub esp, 0000000Ch
  loc_0040E3B6: push 00401546h ; __vbaExceptHandler
  loc_0040E3BB: mov eax, fs:[00000000h]
  loc_0040E3C1: push eax
  loc_0040E3C2: mov fs:[00000000h], esp
  loc_0040E3C9: sub esp, 00000024h
  loc_0040E3CC: push ebx
  loc_0040E3CD: push esi
  loc_0040E3CE: push edi
  loc_0040E3CF: mov var_C, esp
  loc_0040E3D2: mov var_8, 004012E0h
  loc_0040E3D9: mov esi, Me
  loc_0040E3DC: mov eax, esi
  loc_0040E3DE: and eax, 00000001h
  loc_0040E3E1: mov var_4, eax
  loc_0040E3E4: and esi, FFFFFFFEh
  loc_0040E3E7: push esi
  loc_0040E3E8: mov Me, esi
  loc_0040E3EB: mov ecx, [esi]
  loc_0040E3ED: call [ecx+00000004h]
  loc_0040E3F0: mov edx, [esi]
  loc_0040E3F2: xor eax, eax
  loc_0040E3F4: push esi
  loc_0040E3F5: mov var_18, eax
  loc_0040E3F8: mov var_1C, eax
  loc_0040E3FB: mov var_20, eax
  loc_0040E3FE: call [edx+000003E0h]
  loc_0040E404: mov edi, [0040105Ch] ; __vbaObjSet
  loc_0040E40A: push eax
  loc_0040E40B: lea eax, var_1C
  loc_0040E40E: push eax
  loc_0040E40F: call edi
  loc_0040E411: mov ecx, [esi]
  loc_0040E413: push esi
  loc_0040E414: mov ebx, eax
  loc_0040E416: call [ecx+000003E0h]
  loc_0040E41C: lea edx, var_18
  loc_0040E41F: push eax
  loc_0040E420: push edx
  loc_0040E421: call edi
  loc_0040E423: mov edi, eax
  loc_0040E425: lea ecx, var_20
  loc_0040E428: push ecx
  loc_0040E429: push edi
  loc_0040E42A: mov eax, [edi]
  loc_0040E42C: call [eax+00000068h]
  loc_0040E42F: test eax, eax
  loc_0040E431: fnclex
  loc_0040E433: jge 0040E444h
  loc_0040E435: push 00000068h
  loc_0040E437: push 00403C38h
  loc_0040E43C: push edi
  loc_0040E43D: push eax
  loc_0040E43E: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E444: mov eax, var_20
  loc_0040E447: mov edx, [ebx]
  loc_0040E449: not eax
  loc_0040E44B: push eax
  loc_0040E44C: push ebx
  loc_0040E44D: call [edx+0000006Ch]
  loc_0040E450: test eax, eax
  loc_0040E452: fnclex
  loc_0040E454: jge 0040E465h
  loc_0040E456: push 0000006Ch
  loc_0040E458: push 00403C38h
  loc_0040E45D: push ebx
  loc_0040E45E: push eax
  loc_0040E45F: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E465: lea ecx, var_1C
  loc_0040E468: lea edx, var_18
  loc_0040E46B: push ecx
  loc_0040E46C: push edx
  loc_0040E46D: push 00000002h
  loc_0040E46F: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E475: mov eax, [esi]
  loc_0040E477: add esp, 0000000Ch
  loc_0040E47A: push esi
  loc_0040E47B: call [eax+000003E0h]
  loc_0040E481: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_0040E487: lea ecx, var_18
  loc_0040E48A: push eax
  loc_0040E48B: push ecx
  loc_0040E48C: call ebx
  loc_0040E48E: mov edi, eax
  loc_0040E490: lea eax, var_20
  loc_0040E493: push eax
  loc_0040E494: push edi
  loc_0040E495: mov edx, [edi]
  loc_0040E497: call [edx+00000068h]
  loc_0040E49A: test eax, eax
  loc_0040E49C: fnclex
  loc_0040E49E: jge 0040E4AFh
  loc_0040E4A0: push 00000068h
  loc_0040E4A2: push 00403C38h
  loc_0040E4A7: push edi
  loc_0040E4A8: push eax
  loc_0040E4A9: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E4AF: mov cx, var_20
  loc_0040E4B3: mov [esi+00000072h], cx
  loc_0040E4B7: lea ecx, var_18
  loc_0040E4BA: call [00401180h] ; __vbaFreeObj
  loc_0040E4C0: cmp [esi+00000072h], FFFFFFh
  loc_0040E4C5: jnz 0040E4D0h
  loc_0040E4C7: mov edx, [esi]
  loc_0040E4C9: push esi
  loc_0040E4CA: call [edx+00000794h]
  loc_0040E4D0: mov eax, [esi]
  loc_0040E4D2: push esi
  loc_0040E4D3: call [eax+00000308h]
  loc_0040E4D9: lea ecx, var_18
  loc_0040E4DC: push eax
  loc_0040E4DD: push ecx
  loc_0040E4DE: call ebx
  loc_0040E4E0: mov edi, eax
  loc_0040E4E2: mov ax, [esi+00000072h]
  loc_0040E4E6: push eax
  loc_0040E4E7: push edi
  loc_0040E4E8: mov edx, [edi]
  loc_0040E4EA: call [edx+00000084h]
  loc_0040E4F0: test eax, eax
  loc_0040E4F2: fnclex
  loc_0040E4F4: jge 0040E508h
  loc_0040E4F6: push 00000084h
  loc_0040E4FB: push 00403C48h
  loc_0040E500: push edi
  loc_0040E501: push eax
  loc_0040E502: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E508: lea ecx, var_18
  loc_0040E50B: call [00401180h] ; __vbaFreeObj
  loc_0040E511: mov var_4, 00000000h
  loc_0040E518: push 0040E534h
  loc_0040E51D: jmp 0040E533h
  loc_0040E51F: lea ecx, var_1C
  loc_0040E522: lea edx, var_18
  loc_0040E525: push ecx
  loc_0040E526: push edx
  loc_0040E527: push 00000002h
  loc_0040E529: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E52F: add esp, 0000000Ch
  loc_0040E532: ret
  loc_0040E533: ret
  loc_0040E534: mov eax, Me
  loc_0040E537: push eax
  loc_0040E538: mov ecx, [eax]
  loc_0040E53A: call [ecx+00000008h]
  loc_0040E53D: mov eax, var_4
  loc_0040E540: mov ecx, var_14
  loc_0040E543: pop edi
  loc_0040E544: pop esi
  loc_0040E545: mov fs:[00000000h], ecx
  loc_0040E54C: pop ebx
  loc_0040E54D: mov esp, ebp
  loc_0040E54F: pop ebp
  loc_0040E550: retn 0004h
End Sub

Private Sub Form_Load() '40B780
  loc_0040B780: push ebp
  loc_0040B781: mov ebp, esp
  loc_0040B783: sub esp, 0000000Ch
  loc_0040B786: push 00401546h ; __vbaExceptHandler
  loc_0040B78B: mov eax, fs:[00000000h]
  loc_0040B791: push eax
  loc_0040B792: mov fs:[00000000h], esp
  loc_0040B799: sub esp, 00000050h
  loc_0040B79C: push ebx
  loc_0040B79D: push esi
  loc_0040B79E: push edi
  loc_0040B79F: mov var_C, esp
  loc_0040B7A2: mov var_8, 004011E0h
  loc_0040B7A9: mov esi, Me
  loc_0040B7AC: mov eax, esi
  loc_0040B7AE: and eax, 00000001h
  loc_0040B7B1: mov var_4, eax
  loc_0040B7B4: and esi, FFFFFFFEh
  loc_0040B7B7: push esi
  loc_0040B7B8: mov Me, esi
  loc_0040B7BB: mov ecx, [esi]
  loc_0040B7BD: call [ecx+00000004h]
  loc_0040B7C0: mov edx, [esi]
  loc_0040B7C2: xor edi, edi
  loc_0040B7C4: push esi
  loc_0040B7C5: mov var_18, edi
  loc_0040B7C8: mov var_1C, edi
  loc_0040B7CB: mov var_20, edi
  loc_0040B7CE: mov var_30, edi
  loc_0040B7D1: mov var_40, edi
  loc_0040B7D4: call [edx+00000700h]
  loc_0040B7DA: mov eax, [esi]
  loc_0040B7DC: push esi
  loc_0040B7DD: call [eax+000006F8h]
  loc_0040B7E3: cmp eax, edi
  loc_0040B7E5: jge 0040B7F9h
  loc_0040B7E7: push 000006F8h
  loc_0040B7EC: push 004030CCh
  loc_0040B7F1: push esi
  loc_0040B7F2: push eax
  loc_0040B7F3: call [00401040h] ; __vbaHresultCheckObj
  loc_0040B7F9: mov ecx, [esi]
  loc_0040B7FB: push esi
  loc_0040B7FC: call [ecx+00000730h]
  loc_0040B802: mov edx, [esi]
  loc_0040B804: push esi
  loc_0040B805: call [edx+0000073Ch]
  loc_0040B80B: mov eax, [esi]
  loc_0040B80D: push esi
  loc_0040B80E: call [eax+00000770h]
  loc_0040B814: fld real4 ptr [esi+00000044h]
  loc_0040B817: fsub st0, real4 ptr [esi+00000040h]
  loc_0040B81A: lea ecx, var_30
  loc_0040B81D: push 00000001h
  loc_0040B81F: lea edx, var_40
  loc_0040B822: push ecx
  loc_0040B823: cmp [00415000h], 00000000h
  loc_0040B82A: jnz 0040B834h
  loc_0040B82C: fdiv st0, real4 ptr [004011D8h]
  loc_0040B832: jmp 0040B83Fh
  loc_0040B834: push [004011D8h]
  loc_0040B83A: call 00401558h ; _adj_fdiv_m32
  loc_0040B83F: push edx
  loc_0040B840: mov var_30, 00000004h
  loc_0040B847: fstp real4 ptr var_28
  loc_0040B84A: fnstsw ax
  loc_0040B84C: test al, 0Dh
  loc_0040B84E: jnz 0040BA1Bh
  loc_0040B854: call [004010E8h] ; rtcRound
  loc_0040B85A: lea eax, var_40
  loc_0040B85D: push eax
  loc_0040B85E: call [00401018h] ; __vbaStrVarMove
  loc_0040B864: mov edi, [00401164h] ; __vbaStrMove
  loc_0040B86A: mov edx, eax
  loc_0040B86C: mov ecx, 0041502Ch
  loc_0040B871: call edi
  loc_0040B873: lea ecx, var_40
  loc_0040B876: lea edx, var_30
  loc_0040B879: push ecx
  loc_0040B87A: push edx
  loc_0040B87B: push 00000002h
  loc_0040B87D: call [00401024h] ; __vbaFreeVarList
  loc_0040B883: mov eax, [esi]
  loc_0040B885: add esp, 0000000Ch
  loc_0040B888: push esi
  loc_0040B889: call [eax+0000031Ch]
  loc_0040B88F: lea ecx, var_20
  loc_0040B892: push eax
  loc_0040B893: push ecx
  loc_0040B894: call [0040105Ch] ; __vbaObjSet
  loc_0040B89A: mov ebx, eax
  loc_0040B89C: mov eax, [0041502Ch]
  loc_0040B8A1: push eax
  loc_0040B8A2: push ebx
  loc_0040B8A3: mov edx, [ebx]
  loc_0040B8A5: call [edx+000000A4h]
  loc_0040B8AB: test eax, eax
  loc_0040B8AD: fnclex
  loc_0040B8AF: jge 0040B8C3h
  loc_0040B8B1: push 000000A4h
  loc_0040B8B6: push 00403848h
  loc_0040B8BB: push ebx
  loc_0040B8BC: push eax
  loc_0040B8BD: call [00401040h] ; __vbaHresultCheckObj
  loc_0040B8C3: lea ecx, var_20
  loc_0040B8C6: call [00401180h] ; __vbaFreeObj
  loc_0040B8CC: mov ecx, [esi]
  loc_0040B8CE: push esi
  loc_0040B8CF: call [ecx+00000398h]
  loc_0040B8D5: lea edx, var_20
  loc_0040B8D8: push eax
  loc_0040B8D9: push edx
  loc_0040B8DA: call [0040105Ch] ; __vbaObjSet
  loc_0040B8E0: mov ebx, [eax]
  loc_0040B8E2: push 000000DFh
  loc_0040B8E7: mov var_54, eax
  loc_0040B8EA: call [00401108h] ; rtcBstrFromAnsi
  loc_0040B8F0: mov edx, eax
  loc_0040B8F2: lea ecx, var_18
  loc_0040B8F5: call edi
  loc_0040B8F7: push eax
  loc_0040B8F8: push 004035A0h
  loc_0040B8FD: call [0040103Ch] ; __vbaStrCat
  loc_0040B903: mov edx, eax
  loc_0040B905: lea ecx, var_1C
  loc_0040B908: call edi
  loc_0040B90A: mov var_64, ebx
  loc_0040B90D: mov ebx, var_54
  loc_0040B910: push eax
  loc_0040B911: mov eax, var_64
  loc_0040B914: push ebx
  loc_0040B915: call [eax+00000054h]
  loc_0040B918: test eax, eax
  loc_0040B91A: fnclex
  loc_0040B91C: jge 0040B92Dh
  loc_0040B91E: push 00000054h
  loc_0040B920: push 00403858h
  loc_0040B925: push ebx
  loc_0040B926: push eax
  loc_0040B927: call [00401040h] ; __vbaHresultCheckObj
  loc_0040B92D: lea ecx, var_1C
  loc_0040B930: lea edx, var_18
  loc_0040B933: push ecx
  loc_0040B934: push edx
  loc_0040B935: push 00000002h
  loc_0040B937: call [00401130h] ; __vbaFreeStrList
  loc_0040B93D: add esp, 0000000Ch
  loc_0040B940: lea ecx, var_20
  loc_0040B943: call [00401180h] ; __vbaFreeObj
  loc_0040B949: mov eax, [esi]
  loc_0040B94B: push esi
  loc_0040B94C: call [eax+00000394h]
  loc_0040B952: lea ecx, var_20
  loc_0040B955: push eax
  loc_0040B956: push ecx
  loc_0040B957: call [0040105Ch] ; __vbaObjSet
  loc_0040B95D: mov esi, eax
  loc_0040B95F: push 000000DFh
  loc_0040B964: mov ebx, [esi]
  loc_0040B966: call [00401108h] ; rtcBstrFromAnsi
  loc_0040B96C: mov edx, eax
  loc_0040B96E: lea ecx, var_18
  loc_0040B971: call edi
  loc_0040B973: push eax
  loc_0040B974: push 0040386Ch
  loc_0040B979: call [0040103Ch] ; __vbaStrCat
  loc_0040B97F: mov edx, eax
  loc_0040B981: lea ecx, var_1C
  loc_0040B984: call edi
  loc_0040B986: push eax
  loc_0040B987: push esi
  loc_0040B988: call [ebx+00000054h]
  loc_0040B98B: test eax, eax
  loc_0040B98D: fnclex
  loc_0040B98F: jge 0040B9A0h
  loc_0040B991: push 00000054h
  loc_0040B993: push 00403858h
  loc_0040B998: push esi
  loc_0040B999: push eax
  loc_0040B99A: call [00401040h] ; __vbaHresultCheckObj
  loc_0040B9A0: lea edx, var_1C
  loc_0040B9A3: lea eax, var_18
  loc_0040B9A6: push edx
  loc_0040B9A7: push eax
  loc_0040B9A8: push 00000002h
  loc_0040B9AA: call [00401130h] ; __vbaFreeStrList
  loc_0040B9B0: add esp, 0000000Ch
  loc_0040B9B3: lea ecx, var_20
  loc_0040B9B6: call [00401180h] ; __vbaFreeObj
  loc_0040B9BC: mov var_4, 00000000h
  loc_0040B9C3: fwait
  loc_0040B9C4: push 0040B9FCh
  loc_0040B9C9: jmp 0040B9FBh
  loc_0040B9CB: lea ecx, var_1C
  loc_0040B9CE: lea edx, var_18
  loc_0040B9D1: push ecx
  loc_0040B9D2: push edx
  loc_0040B9D3: push 00000002h
  loc_0040B9D5: call [00401130h] ; __vbaFreeStrList
  loc_0040B9DB: add esp, 0000000Ch
  loc_0040B9DE: lea ecx, var_20
  loc_0040B9E1: call [00401180h] ; __vbaFreeObj
  loc_0040B9E7: lea eax, var_40
  loc_0040B9EA: lea ecx, var_30
  loc_0040B9ED: push eax
  loc_0040B9EE: push ecx
  loc_0040B9EF: push 00000002h
  loc_0040B9F1: call [00401024h] ; __vbaFreeVarList
  loc_0040B9F7: add esp, 0000000Ch
  loc_0040B9FA: ret
  loc_0040B9FB: ret
  loc_0040B9FC: mov eax, Me
  loc_0040B9FF: push eax
  loc_0040BA00: mov edx, [eax]
  loc_0040BA02: call [edx+00000008h]
  loc_0040BA05: mov eax, var_4
  loc_0040BA08: mov ecx, var_14
  loc_0040BA0B: pop edi
  loc_0040BA0C: pop esi
  loc_0040BA0D: mov fs:[00000000h], ecx
  loc_0040BA14: pop ebx
  loc_0040BA15: mov esp, ebp
  loc_0040BA17: pop ebp
  loc_0040BA18: retn 0004h
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '412210
  loc_00412210: push ebp
  loc_00412211: mov ebp, esp
  loc_00412213: sub esp, 0000000Ch
  loc_00412216: push 00401546h ; __vbaExceptHandler
  loc_0041221B: mov eax, fs:[00000000h]
  loc_00412221: push eax
  loc_00412222: mov fs:[00000000h], esp
  loc_00412229: sub esp, 0000009Ch
  loc_0041222F: push ebx
  loc_00412230: push esi
  loc_00412231: push edi
  loc_00412232: mov var_C, esp
  loc_00412235: mov var_8, 00401468h
  loc_0041223C: mov edi, Me
  loc_0041223F: mov eax, edi
  loc_00412241: and eax, 00000001h
  loc_00412244: mov var_4, eax
  loc_00412247: and edi, FFFFFFFEh
  loc_0041224A: push edi
  loc_0041224B: mov Me, edi
  loc_0041224E: mov ecx, [edi]
  loc_00412250: call [ecx+00000004h]
  loc_00412253: mov ebx, [00401150h] ; __vbaVarDup
  loc_00412259: mov ecx, 80020004h
  loc_0041225E: xor esi, esi
  loc_00412260: mov var_54, ecx
  loc_00412263: mov eax, 0000000Ah
  loc_00412268: mov var_44, ecx
  loc_0041226B: mov var_4C, esi
  loc_0041226E: mov var_5C, esi
  loc_00412271: mov var_7C, esi
  loc_00412274: lea edx, var_7C
  loc_00412277: lea ecx, var_3C
  loc_0041227A: mov var_1C, esi
  loc_0041227D: mov var_2C, esi
  loc_00412280: mov var_3C, esi
  loc_00412283: mov var_6C, esi
  loc_00412286: mov var_5C, eax
  loc_00412289: mov var_4C, eax
  loc_0041228C: mov var_74, 00403DB0h ; "Exit Check"
  loc_00412293: mov var_7C, 00000008h
  loc_0041229A: call ebx
  loc_0041229C: lea edx, var_6C
  loc_0041229F: lea ecx, var_2C
  loc_004122A2: mov var_64, 00403D6Ch ; "Are you sure you want to exit?"
  loc_004122A9: mov var_6C, 00000008h
  loc_004122B0: call ebx
  loc_004122B2: lea edx, var_5C
  loc_004122B5: lea eax, var_4C
  loc_004122B8: push edx
  loc_004122B9: lea ecx, var_3C
  loc_004122BC: push eax
  loc_004122BD: push ecx
  loc_004122BE: lea edx, var_2C
  loc_004122C1: push 00000004h
  loc_004122C3: push edx
  loc_004122C4: call [00401060h] ; rtcMsgBox
  loc_004122CA: mov ecx, eax
  loc_004122CC: call [004010A4h] ; __vbaI2I4
  loc_004122D2: mov ebx, eax
  loc_004122D4: lea eax, var_5C
  loc_004122D7: lea ecx, var_4C
  loc_004122DA: push eax
  loc_004122DB: lea edx, var_3C
  loc_004122DE: push ecx
  loc_004122DF: lea eax, var_2C
  loc_004122E2: push edx
  loc_004122E3: push eax
  loc_004122E4: push 00000004h
  loc_004122E6: call [00401024h] ; __vbaFreeVarList
  loc_004122EC: add esp, 00000014h
  loc_004122EF: cmp bx, 0007h
  loc_004122F3: jnz 004122FFh
  loc_004122F5: mov ecx, Cancel
  loc_004122F8: mov [ecx], FFFFFFh
  loc_004122FD: jmp 0041235Fh
  loc_004122FF: cmp [0041567Ch], esi
  loc_00412305: jnz 00412317h
  loc_00412307: push 0041567Ch
  loc_0041230C: push 00403C18h
  loc_00412311: call [0040111Ch] ; __vbaNew2
  loc_00412317: mov ebx, [0041567Ch]
  loc_0041231D: lea eax, var_1C
  loc_00412320: push edi
  loc_00412321: push eax
  loc_00412322: mov edx, [ebx]
  loc_00412324: mov var_B0, edx
  loc_0041232A: call [0040106Ch] ; __vbaObjSetAddref
  loc_00412330: mov ecx, var_B0
  loc_00412336: push eax
  loc_00412337: push ebx
  loc_00412338: call [ecx+00000010h]
  loc_0041233B: cmp eax, esi
  loc_0041233D: fnclex
  loc_0041233F: jge 00412350h
  loc_00412341: push 00000010h
  loc_00412343: push 00403C08h
  loc_00412348: push ebx
  loc_00412349: push eax
  loc_0041234A: call [00401040h] ; __vbaHresultCheckObj
  loc_00412350: lea ecx, var_1C
  loc_00412353: call [00401180h] ; __vbaFreeObj
  loc_00412359: call [00401020h] ; __vbaEnd
  loc_0041235F: mov var_4, esi
  loc_00412362: push 0041238Fh
  loc_00412367: jmp 0041238Eh
  loc_00412369: lea ecx, var_1C
  loc_0041236C: call [00401180h] ; __vbaFreeObj
  loc_00412372: lea edx, var_5C
  loc_00412375: lea eax, var_4C
  loc_00412378: push edx
  loc_00412379: lea ecx, var_3C
  loc_0041237C: push eax
  loc_0041237D: lea edx, var_2C
  loc_00412380: push ecx
  loc_00412381: push edx
  loc_00412382: push 00000004h
  loc_00412384: call [00401024h] ; __vbaFreeVarList
  loc_0041238A: add esp, 00000014h
  loc_0041238D: ret
  loc_0041238E: ret
  loc_0041238F: mov eax, Me
  loc_00412392: push eax
  loc_00412393: mov ecx, [eax]
  loc_00412395: call [ecx+00000008h]
  loc_00412398: mov eax, var_4
  loc_0041239B: mov ecx, var_14
  loc_0041239E: pop edi
  loc_0041239F: pop esi
  loc_004123A0: mov fs:[00000000h], ecx
  loc_004123A7: pop ebx
  loc_004123A8: mov esp, ebp
  loc_004123AA: pop ebp
  loc_004123AB: retn 000Ch
End Sub

Private Sub Form_Activate() '40C9A0
  loc_0040C9A0: push ebp
  loc_0040C9A1: mov ebp, esp
  loc_0040C9A3: sub esp, 0000000Ch
  loc_0040C9A6: push 00401546h ; __vbaExceptHandler
  loc_0040C9AB: mov eax, fs:[00000000h]
  loc_0040C9B1: push eax
  loc_0040C9B2: mov fs:[00000000h], esp
  loc_0040C9B9: sub esp, 00000008h
  loc_0040C9BC: push ebx
  loc_0040C9BD: push esi
  loc_0040C9BE: push edi
  loc_0040C9BF: mov var_C, esp
  loc_0040C9C2: mov var_8, 00401238h
  loc_0040C9C9: mov esi, Me
  loc_0040C9CC: mov eax, esi
  loc_0040C9CE: and eax, 00000001h
  loc_0040C9D1: mov var_4, eax
  loc_0040C9D4: and esi, FFFFFFFEh
  loc_0040C9D7: push esi
  loc_0040C9D8: mov Me, esi
  loc_0040C9DB: mov ecx, [esi]
  loc_0040C9DD: call [ecx+00000004h]
  loc_0040C9E0: mov edx, [00415028h]
  loc_0040C9E6: mov eax, [0041502Ch]
  loc_0040C9EB: push edx
  loc_0040C9EC: push eax
  loc_0040C9ED: call [004010A0h] ; __vbaR4Str
  loc_0040C9F3: push ecx
  loc_0040C9F4: lea ecx, [esi+0000003Ch]
  loc_0040C9F7: fstp real4 ptr [esp]
  loc_0040C9FA: push 00000001h
  loc_0040C9FC: push 00000002h
  loc_0040C9FE: push ecx
  loc_0040C9FF: call 00409C80h
  loc_0040CA04: mov edx, [esi]
  loc_0040CA06: push esi
  loc_0040CA07: call [edx+00000780h]
  loc_0040CA0D: mov eax, [esi]
  loc_0040CA0F: push esi
  loc_0040CA10: call [eax+0000078Ch]
  loc_0040CA16: lea ecx, [esi+00000080h]
  loc_0040CA1C: lea edx, [esi+00000088h]
  loc_0040CA22: push ecx
  loc_0040CA23: lea eax, [esi+00000084h]
  loc_0040CA29: push edx
  loc_0040CA2A: lea ecx, [esi+0000007Ch]
  loc_0040CA2D: push eax
  loc_0040CA2E: lea edx, [esi+00000078h]
  loc_0040CA31: push ecx
  loc_0040CA32: lea eax, [esi+00000074h]
  loc_0040CA35: push edx
  loc_0040CA36: push eax
  loc_0040CA37: call 00409DD0h
  loc_0040CA3C: mov ecx, [esi]
  loc_0040CA3E: push esi
  loc_0040CA3F: call [ecx+0000071Ch]
  loc_0040CA45: mov edx, [esi]
  loc_0040CA47: push esi
  loc_0040CA48: call [edx+00000718h]
  loc_0040CA4E: mov var_4, 00000000h
  loc_0040CA55: mov eax, Me
  loc_0040CA58: push eax
  loc_0040CA59: mov ecx, [eax]
  loc_0040CA5B: call [ecx+00000008h]
  loc_0040CA5E: mov eax, var_4
  loc_0040CA61: mov ecx, var_14
  loc_0040CA64: pop edi
  loc_0040CA65: pop esi
  loc_0040CA66: mov fs:[00000000h], ecx
  loc_0040CA6D: pop ebx
  loc_0040CA6E: mov esp, ebp
  loc_0040CA70: pop ebp
  loc_0040CA71: retn 0004h
End Sub

Private Sub mnuShowLSline_Click() '40E560
  loc_0040E560: push ebp
  loc_0040E561: mov ebp, esp
  loc_0040E563: sub esp, 0000000Ch
  loc_0040E566: push 00401546h ; __vbaExceptHandler
  loc_0040E56B: mov eax, fs:[00000000h]
  loc_0040E571: push eax
  loc_0040E572: mov fs:[00000000h], esp
  loc_0040E579: sub esp, 00000024h
  loc_0040E57C: push ebx
  loc_0040E57D: push esi
  loc_0040E57E: push edi
  loc_0040E57F: mov var_C, esp
  loc_0040E582: mov var_8, 004012F0h
  loc_0040E589: mov esi, Me
  loc_0040E58C: mov eax, esi
  loc_0040E58E: and eax, 00000001h
  loc_0040E591: mov var_4, eax
  loc_0040E594: and esi, FFFFFFFEh
  loc_0040E597: push esi
  loc_0040E598: mov Me, esi
  loc_0040E59B: mov ecx, [esi]
  loc_0040E59D: call [ecx+00000004h]
  loc_0040E5A0: mov edx, [esi]
  loc_0040E5A2: xor eax, eax
  loc_0040E5A4: push esi
  loc_0040E5A5: mov var_18, eax
  loc_0040E5A8: mov var_1C, eax
  loc_0040E5AB: mov var_20, eax
  loc_0040E5AE: call [edx+000003D8h]
  loc_0040E5B4: mov edi, [0040105Ch] ; __vbaObjSet
  loc_0040E5BA: push eax
  loc_0040E5BB: lea eax, var_1C
  loc_0040E5BE: push eax
  loc_0040E5BF: call edi
  loc_0040E5C1: mov ecx, [esi]
  loc_0040E5C3: push esi
  loc_0040E5C4: mov ebx, eax
  loc_0040E5C6: call [ecx+000003D8h]
  loc_0040E5CC: lea edx, var_18
  loc_0040E5CF: push eax
  loc_0040E5D0: push edx
  loc_0040E5D1: call edi
  loc_0040E5D3: mov edi, eax
  loc_0040E5D5: lea ecx, var_20
  loc_0040E5D8: push ecx
  loc_0040E5D9: push edi
  loc_0040E5DA: mov eax, [edi]
  loc_0040E5DC: call [eax+00000068h]
  loc_0040E5DF: test eax, eax
  loc_0040E5E1: fnclex
  loc_0040E5E3: jge 0040E5F4h
  loc_0040E5E5: push 00000068h
  loc_0040E5E7: push 00403C38h
  loc_0040E5EC: push edi
  loc_0040E5ED: push eax
  loc_0040E5EE: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E5F4: mov eax, var_20
  loc_0040E5F7: mov edx, [ebx]
  loc_0040E5F9: not eax
  loc_0040E5FB: push eax
  loc_0040E5FC: push ebx
  loc_0040E5FD: call [edx+0000006Ch]
  loc_0040E600: test eax, eax
  loc_0040E602: fnclex
  loc_0040E604: jge 0040E615h
  loc_0040E606: push 0000006Ch
  loc_0040E608: push 00403C38h
  loc_0040E60D: push ebx
  loc_0040E60E: push eax
  loc_0040E60F: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E615: lea ecx, var_1C
  loc_0040E618: lea edx, var_18
  loc_0040E61B: push ecx
  loc_0040E61C: push edx
  loc_0040E61D: push 00000002h
  loc_0040E61F: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E625: mov eax, [esi]
  loc_0040E627: add esp, 0000000Ch
  loc_0040E62A: push esi
  loc_0040E62B: call [eax+000003D8h]
  loc_0040E631: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_0040E637: lea ecx, var_18
  loc_0040E63A: push eax
  loc_0040E63B: push ecx
  loc_0040E63C: call ebx
  loc_0040E63E: mov edi, eax
  loc_0040E640: lea eax, var_20
  loc_0040E643: push eax
  loc_0040E644: push edi
  loc_0040E645: mov edx, [edi]
  loc_0040E647: call [edx+00000068h]
  loc_0040E64A: test eax, eax
  loc_0040E64C: fnclex
  loc_0040E64E: jge 0040E65Fh
  loc_0040E650: push 00000068h
  loc_0040E652: push 00403C38h
  loc_0040E657: push edi
  loc_0040E658: push eax
  loc_0040E659: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E65F: mov cx, var_20
  loc_0040E663: mov [esi+0000006Ch], cx
  loc_0040E667: lea ecx, var_18
  loc_0040E66A: call [00401180h] ; __vbaFreeObj
  loc_0040E670: cmp [esi+0000006Ch], FFFFFFh
  loc_0040E675: jnz 0040E680h
  loc_0040E677: mov edx, [esi]
  loc_0040E679: push esi
  loc_0040E67A: call [edx+00000774h]
  loc_0040E680: mov eax, [esi]
  loc_0040E682: push esi
  loc_0040E683: call [eax+00000310h]
  loc_0040E689: lea ecx, var_18
  loc_0040E68C: push eax
  loc_0040E68D: push ecx
  loc_0040E68E: call ebx
  loc_0040E690: mov edi, eax
  loc_0040E692: mov ax, [esi+0000006Ch]
  loc_0040E696: push eax
  loc_0040E697: push edi
  loc_0040E698: mov edx, [edi]
  loc_0040E69A: call [edx+00000084h]
  loc_0040E6A0: test eax, eax
  loc_0040E6A2: fnclex
  loc_0040E6A4: jge 0040E6B8h
  loc_0040E6A6: push 00000084h
  loc_0040E6AB: push 00403C48h
  loc_0040E6B0: push edi
  loc_0040E6B1: push eax
  loc_0040E6B2: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E6B8: lea ecx, var_18
  loc_0040E6BB: call [00401180h] ; __vbaFreeObj
  loc_0040E6C1: mov var_4, 00000000h
  loc_0040E6C8: push 0040E6E4h
  loc_0040E6CD: jmp 0040E6E3h
  loc_0040E6CF: lea ecx, var_1C
  loc_0040E6D2: lea edx, var_18
  loc_0040E6D5: push ecx
  loc_0040E6D6: push edx
  loc_0040E6D7: push 00000002h
  loc_0040E6D9: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E6DF: add esp, 0000000Ch
  loc_0040E6E2: ret
  loc_0040E6E3: ret
  loc_0040E6E4: mov eax, Me
  loc_0040E6E7: push eax
  loc_0040E6E8: mov ecx, [eax]
  loc_0040E6EA: call [ecx+00000008h]
  loc_0040E6ED: mov eax, var_4
  loc_0040E6F0: mov ecx, var_14
  loc_0040E6F3: pop edi
  loc_0040E6F4: pop esi
  loc_0040E6F5: mov fs:[00000000h], ecx
  loc_0040E6FC: pop ebx
  loc_0040E6FD: mov esp, ebp
  loc_0040E6FF: pop ebp
  loc_0040E700: retn 0004h
End Sub

Private Sub mnuChangeTTable_Click() '40C8B0
  loc_0040C8B0: push ebp
  loc_0040C8B1: mov ebp, esp
  loc_0040C8B3: sub esp, 0000000Ch
  loc_0040C8B6: push 00401546h ; __vbaExceptHandler
  loc_0040C8BB: mov eax, fs:[00000000h]
  loc_0040C8C1: push eax
  loc_0040C8C2: mov fs:[00000000h], esp
  loc_0040C8C9: sub esp, 00000030h
  loc_0040C8CC: push ebx
  loc_0040C8CD: push esi
  loc_0040C8CE: push edi
  loc_0040C8CF: mov var_C, esp
  loc_0040C8D2: mov var_8, 00401230h
  loc_0040C8D9: mov eax, Me
  loc_0040C8DC: mov ecx, eax
  loc_0040C8DE: and ecx, 00000001h
  loc_0040C8E1: mov var_4, ecx
  loc_0040C8E4: and al, FEh
  loc_0040C8E6: push eax
  loc_0040C8E7: mov Me, eax
  loc_0040C8EA: mov edx, [eax]
  loc_0040C8EC: call [edx+00000004h]
  loc_0040C8EF: mov eax, [0041520Ch]
  loc_0040C8F4: test eax, eax
  loc_0040C8F6: jnz 0040C908h
  loc_0040C8F8: push 0041520Ch
  loc_0040C8FD: push 00402644h
  loc_0040C902: call [0040111Ch] ; __vbaNew2
  loc_0040C908: sub esp, 00000010h
  loc_0040C90B: mov ecx, 0000000Ah
  loc_0040C910: mov ebx, esp
  loc_0040C912: mov var_24, ecx
  loc_0040C915: mov eax, 80020004h
  loc_0040C91A: sub esp, 00000010h
  loc_0040C91D: mov [ebx], ecx
  loc_0040C91F: mov ecx, var_30
  loc_0040C922: mov edx, eax
  loc_0040C924: mov esi, [0041520Ch]
  loc_0040C92A: mov [ebx+00000004h], ecx
  loc_0040C92D: mov ecx, esp
  loc_0040C92F: mov edi, [esi]
  loc_0040C931: push esi
  loc_0040C932: mov [ebx+00000008h], eax
  loc_0040C935: mov eax, var_28
  loc_0040C938: mov [ebx+0000000Ch], eax
  loc_0040C93B: mov eax, var_24
  loc_0040C93E: mov [ecx], eax
  loc_0040C940: mov eax, var_20
  loc_0040C943: mov [ecx+00000004h], eax
  loc_0040C946: mov [ecx+00000008h], edx
  loc_0040C949: mov edx, var_18
  loc_0040C94C: mov [ecx+0000000Ch], edx
  loc_0040C94F: call [edi+000002B0h]
  loc_0040C955: test eax, eax
  loc_0040C957: fnclex
  loc_0040C959: jge 0040C96Dh
  loc_0040C95B: push 000002B0h
  loc_0040C960: push 00403A74h
  loc_0040C965: push esi
  loc_0040C966: push eax
  loc_0040C967: call [00401040h] ; __vbaHresultCheckObj
  loc_0040C96D: mov var_4, 00000000h
  loc_0040C974: mov eax, Me
  loc_0040C977: push eax
  loc_0040C978: mov ecx, [eax]
  loc_0040C97A: call [ecx+00000008h]
  loc_0040C97D: mov eax, var_4
  loc_0040C980: mov ecx, var_14
  loc_0040C983: pop edi
  loc_0040C984: pop esi
  loc_0040C985: mov fs:[00000000h], ecx
  loc_0040C98C: pop ebx
  loc_0040C98D: mov esp, ebp
  loc_0040C98F: pop ebp
  loc_0040C990: retn 0004h
End Sub

Private Sub mnuAbout_Click() '40C6B0
  loc_0040C6B0: push ebp
  loc_0040C6B1: mov ebp, esp
  loc_0040C6B3: sub esp, 0000000Ch
  loc_0040C6B6: push 00401546h ; __vbaExceptHandler
  loc_0040C6BB: mov eax, fs:[00000000h]
  loc_0040C6C1: push eax
  loc_0040C6C2: mov fs:[00000000h], esp
  loc_0040C6C9: sub esp, 00000088h
  loc_0040C6CF: push ebx
  loc_0040C6D0: push esi
  loc_0040C6D1: push edi
  loc_0040C6D2: mov var_C, esp
  loc_0040C6D5: mov var_8, 00401218h
  loc_0040C6DC: mov eax, Me
  loc_0040C6DF: mov ecx, eax
  loc_0040C6E1: and ecx, 00000001h
  loc_0040C6E4: mov var_4, ecx
  loc_0040C6E7: and al, FEh
  loc_0040C6E9: push eax
  loc_0040C6EA: mov Me, eax
  loc_0040C6ED: mov edx, [eax]
  loc_0040C6EF: call [edx+00000004h]
  loc_0040C6F2: mov ecx, 80020004h
  loc_0040C6F7: xor esi, esi
  loc_0040C6F9: mov var_4C, ecx
  loc_0040C6FC: mov eax, 0000000Ah
  loc_0040C701: mov var_3C, ecx
  loc_0040C704: mov var_2C, ecx
  loc_0040C707: mov var_34, esi
  loc_0040C70A: mov var_44, esi
  loc_0040C70D: mov var_54, esi
  loc_0040C710: mov var_64, esi
  loc_0040C713: lea edx, var_64
  loc_0040C716: lea ecx, var_24
  loc_0040C719: mov var_24, esi
  loc_0040C71C: mov var_54, eax
  loc_0040C71F: mov var_44, eax
  loc_0040C722: mov var_34, eax
  loc_0040C725: mov var_5C, 00403914h ; "This program was written by Dr. Mark Eakin, Associate Professor, UTA, 817-272-3529"
  loc_0040C72C: mov var_64, 00000008h
  loc_0040C733: call [00401150h] ; __vbaVarDup
  loc_0040C739: lea eax, var_54
  loc_0040C73C: lea ecx, var_44
  loc_0040C73F: push eax
  loc_0040C740: lea edx, var_34
  loc_0040C743: push ecx
  loc_0040C744: push edx
  loc_0040C745: lea eax, var_24
  loc_0040C748: push esi
  loc_0040C749: push eax
  loc_0040C74A: call [00401060h] ; rtcMsgBox
  loc_0040C750: lea ecx, var_54
  loc_0040C753: lea edx, var_44
  loc_0040C756: push ecx
  loc_0040C757: lea eax, var_34
  loc_0040C75A: push edx
  loc_0040C75B: lea ecx, var_24
  loc_0040C75E: push eax
  loc_0040C75F: push ecx
  loc_0040C760: push 00000004h
  loc_0040C762: call [00401024h] ; __vbaFreeVarList
  loc_0040C768: add esp, 00000014h
  loc_0040C76B: mov var_4, esi
  loc_0040C76E: push 0040C792h
  loc_0040C773: jmp 0040C791h
  loc_0040C775: lea edx, var_54
  loc_0040C778: lea eax, var_44
  loc_0040C77B: push edx
  loc_0040C77C: lea ecx, var_34
  loc_0040C77F: push eax
  loc_0040C780: lea edx, var_24
  loc_0040C783: push ecx
  loc_0040C784: push edx
  loc_0040C785: push 00000004h
  loc_0040C787: call [00401024h] ; __vbaFreeVarList
  loc_0040C78D: add esp, 00000014h
  loc_0040C790: ret
  loc_0040C791: ret
  loc_0040C792: mov eax, Me
  loc_0040C795: push eax
  loc_0040C796: mov ecx, [eax]
  loc_0040C798: call [ecx+00000008h]
  loc_0040C79B: mov eax, var_4
  loc_0040C79E: mov ecx, var_14
  loc_0040C7A1: pop edi
  loc_0040C7A2: pop esi
  loc_0040C7A3: mov fs:[00000000h], ecx
  loc_0040C7AA: pop ebx
  loc_0040C7AB: mov esp, ebp
  loc_0040C7AD: pop ebp
  loc_0040C7AE: retn 0004h
End Sub

Private Sub mnuFileSave_Click() '40FEF0
  loc_0040FEF0: push ebp
  loc_0040FEF1: mov ebp, esp
  loc_0040FEF3: sub esp, 00000014h
  loc_0040FEF6: push 00401546h ; __vbaExceptHandler
  loc_0040FEFB: mov eax, fs:[00000000h]
  loc_0040FF01: push eax
  loc_0040FF02: mov fs:[00000000h], esp
  loc_0040FF09: sub esp, 00000064h
  loc_0040FF0C: push ebx
  loc_0040FF0D: push esi
  loc_0040FF0E: push edi
  loc_0040FF0F: mov var_14, esp
  loc_0040FF12: mov var_10, 00401378h
  loc_0040FF19: mov esi, Me
  loc_0040FF1C: mov eax, esi
  loc_0040FF1E: and eax, 00000001h
  loc_0040FF21: mov var_C, eax
  loc_0040FF24: and esi, FFFFFFFEh
  loc_0040FF27: mov Me, esi
  loc_0040FF2A: xor edi, edi
  loc_0040FF2C: mov var_8, edi
  loc_0040FF2F: mov ecx, [esi]
  loc_0040FF31: push esi
  loc_0040FF32: call [ecx+00000004h]
  loc_0040FF35: mov var_20, edi
  loc_0040FF38: mov var_28, edi
  loc_0040FF3B: mov var_38, edi
  loc_0040FF3E: mov var_48, edi
  loc_0040FF41: push 00000001h
  loc_0040FF43: call [00401064h] ; __vbaOnError
  loc_0040FF49: mov eax, 00403CC8h ; "Text files (*.txt)|*.txt"
  loc_0040FF4E: mov ecx, 00000008h
  loc_0040FF53: sub esp, 00000010h
  loc_0040FF56: mov edx, esp
  loc_0040FF58: mov [edx], ecx
  loc_0040FF5A: mov ecx, var_44
  loc_0040FF5D: mov [edx+00000004h], ecx
  loc_0040FF60: mov [edx+00000008h], eax
  loc_0040FF63: mov eax, var_3C
  loc_0040FF66: mov [edx+0000000Ch], eax
  loc_0040FF69: push 00000003h
  loc_0040FF6B: mov ecx, [esi]
  loc_0040FF6D: push esi
  loc_0040FF6E: call [ecx+00000404h]
  loc_0040FF74: push eax
  loc_0040FF75: lea edx, var_28
  loc_0040FF78: push edx
  loc_0040FF79: mov edi, [0040105Ch] ; __vbaObjSet
  loc_0040FF7F: call edi
  loc_0040FF81: push eax
  loc_0040FF82: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FF88: lea ecx, var_28
  loc_0040FF8B: mov ebx, [00401180h] ; __vbaFreeObj
  loc_0040FF91: call ebx
  loc_0040FF93: mov eax, 00403D00h ; "Save the data file"
  loc_0040FF98: mov ecx, 00000008h
  loc_0040FF9D: sub esp, 00000010h
  loc_0040FFA0: mov edx, esp
  loc_0040FFA2: mov [edx], ecx
  loc_0040FFA4: mov ecx, var_44
  loc_0040FFA7: mov [edx+00000004h], ecx
  loc_0040FFAA: mov [edx+00000008h], eax
  loc_0040FFAD: mov eax, var_3C
  loc_0040FFB0: mov [edx+0000000Ch], eax
  loc_0040FFB3: push 00000002h
  loc_0040FFB5: mov ecx, [esi]
  loc_0040FFB7: push esi
  loc_0040FFB8: call [ecx+00000404h]
  loc_0040FFBE: push eax
  loc_0040FFBF: lea edx, var_28
  loc_0040FFC2: push edx
  loc_0040FFC3: call edi
  loc_0040FFC5: push eax
  loc_0040FFC6: call [0040116Ch] ; __vbaLateIdSt
  loc_0040FFCC: lea ecx, var_28
  loc_0040FFCF: call ebx
  loc_0040FFD1: push 00000000h
  loc_0040FFD3: push 0000001Fh
  loc_0040FFD5: mov eax, [esi]
  loc_0040FFD7: push esi
  loc_0040FFD8: call [eax+00000404h]
  loc_0040FFDE: push eax
  loc_0040FFDF: lea ecx, var_28
  loc_0040FFE2: push ecx
  loc_0040FFE3: call edi
  loc_0040FFE5: push eax
  loc_0040FFE6: call [0040101Ch] ; __vbaLateIdCall
  loc_0040FFEC: add esp, 0000000Ch
  loc_0040FFEF: lea ecx, var_28
  loc_0040FFF2: call ebx
  loc_0040FFF4: push 00000000h
  loc_0040FFF6: push 00000001h
  loc_0040FFF8: mov edx, [esi]
  loc_0040FFFA: push esi
  loc_0040FFFB: call [edx+00000404h]
  loc_00410001: push eax
  loc_00410002: lea eax, var_28
  loc_00410005: push eax
  loc_00410006: call edi
  loc_00410008: push eax
  loc_00410009: lea ecx, var_38
  loc_0041000C: push ecx
  loc_0041000D: call [004010B4h] ; __vbaLateIdCallLd
  loc_00410013: add esp, 00000010h
  loc_00410016: push eax
  loc_00410017: call [00401018h] ; __vbaStrVarMove
  loc_0041001D: mov edx, eax
  loc_0041001F: lea ecx, var_20
  loc_00410022: call [00401164h] ; __vbaStrMove
  loc_00410028: lea ecx, var_28
  loc_0041002B: call ebx
  loc_0041002D: lea ecx, var_38
  loc_00410030: call [00401014h] ; __vbaFreeVar
  loc_00410036: mov edx, var_20
  loc_00410039: push edx
  loc_0041003A: push 00000001h
  loc_0041003C: push FFFFFFFFh
  loc_0041003E: push 00000002h
  loc_00410040: call [00401114h] ; __vbaFileOpen
  loc_00410046: mov ax, [esi+00000038h]
  loc_0041004A: mov var_70, ax
  loc_0041004E: mov ebx, 00000001h
  loc_00410053: mov var_24, ebx
  loc_00410056: cmp bx, var_70
  loc_0041005A: jg 00410136h
  loc_00410060: mov eax, [esi+0000003Ch]
  loc_00410063: test eax, eax
  loc_00410065: jz 004100ABh
  loc_00410067: cmp [eax], 0002h
  loc_0041006B: jnz 004100ABh
  loc_0041006D: movsx ecx, bx
  loc_00410070: sub ecx, [eax+0000001Ch]
  loc_00410073: mov var_68, ecx
  loc_00410076: cmp ecx, [eax+00000018h]
  loc_00410079: jb 00410081h
  loc_0041007B: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410081: mov eax, [esi+0000003Ch]
  loc_00410084: mov ebx, 00000002h
  loc_00410089: sub ebx, [eax+00000014h]
  loc_0041008C: cmp ebx, [eax+00000010h]
  loc_0041008F: jb 00410097h
  loc_00410091: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410097: mov ecx, [esi+0000003Ch]
  loc_0041009A: mov edi, [ecx+00000018h]
  loc_0041009D: imul edi, ebx
  loc_004100A0: add edi, var_68
  loc_004100A3: shl edi, 02h
  loc_004100A6: mov ebx, var_24
  loc_004100A9: jmp 004100B3h
  loc_004100AB: call [00401094h] ; __vbaGenerateBoundsError
  loc_004100B1: mov edi, eax
  loc_004100B3: mov eax, [esi+0000003Ch]
  loc_004100B6: test eax, eax
  loc_004100B8: jz 004100FEh
  loc_004100BA: cmp [eax], 0002h
  loc_004100BE: jnz 004100FEh
  loc_004100C0: movsx ecx, bx
  loc_004100C3: sub ecx, [eax+0000001Ch]
  loc_004100C6: mov var_60, ecx
  loc_004100C9: cmp ecx, [eax+00000018h]
  loc_004100CC: jb 004100D4h
  loc_004100CE: call [00401094h] ; __vbaGenerateBoundsError
  loc_004100D4: mov eax, [esi+0000003Ch]
  loc_004100D7: mov ebx, 00000001h
  loc_004100DC: sub ebx, [eax+00000014h]
  loc_004100DF: cmp ebx, [eax+00000010h]
  loc_004100E2: jb 004100EAh
  loc_004100E4: call [00401094h] ; __vbaGenerateBoundsError
  loc_004100EA: mov edx, [esi+0000003Ch]
  loc_004100ED: mov eax, [edx+00000018h]
  loc_004100F0: imul eax, ebx
  loc_004100F3: add eax, var_60
  loc_004100F6: shl eax, 02h
  loc_004100F9: mov ebx, var_24
  loc_004100FC: jmp 00410104h
  loc_004100FE: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410104: mov ecx, [esi+0000003Ch]
  loc_00410107: mov ecx, [ecx+0000000Ch]
  loc_0041010A: mov edx, [ecx+edi]
  loc_0041010D: push edx
  loc_0041010E: mov eax, [ecx+eax]
  loc_00410111: push eax
  loc_00410112: push 00000001h
  loc_00410114: push 00403D2Ch
  loc_00410119: call [004010D8h] ; __vbaPrintFile
  loc_0041011F: add esp, 00000010h
  loc_00410122: mov eax, 00000001h
  loc_00410127: add ax, bx
  loc_0041012A: jo 00410188h
  loc_0041012C: mov var_24, eax
  loc_0041012F: mov ebx, eax
  loc_00410131: jmp 00410056h
  loc_00410136: push 00000001h
  loc_00410138: call [0040108Ch] ; __vbaFileClose
  loc_0041013E: call [00401054h] ; __vbaExitProc
  loc_00410144: fwait
  loc_00410145: push 00410169h
  loc_0041014A: jmp 0041015Fh
  loc_0041014C: lea ecx, var_28
  loc_0041014F: call [00401180h] ; __vbaFreeObj
  loc_00410155: lea ecx, var_38
  loc_00410158: call [00401014h] ; __vbaFreeVar
  loc_0041015E: ret
  loc_0041015F: lea ecx, var_20
  loc_00410162: call [00401184h] ; __vbaFreeStr
  loc_00410168: ret
  loc_00410169: mov eax, Me
  loc_0041016C: mov ecx, [eax]
  loc_0041016E: push eax
  loc_0041016F: call [ecx+00000008h]
  loc_00410172: mov eax, var_C
  loc_00410175: mov ecx, var_1C
  loc_00410178: mov fs:[00000000h], ecx
  loc_0041017F: pop edi
  loc_00410180: pop esi
  loc_00410181: pop ebx
  loc_00410182: mov esp, ebp
  loc_00410184: pop ebp
  loc_00410185: retn 0004h
End Sub

Private Sub mnuViewHelp_Click() '40EC00
  loc_0040EC00: push ebp
  loc_0040EC01: mov ebp, esp
  loc_0040EC03: sub esp, 0000000Ch
  loc_0040EC06: push 00401546h ; __vbaExceptHandler
  loc_0040EC0B: mov eax, fs:[00000000h]
  loc_0040EC11: push eax
  loc_0040EC12: mov fs:[00000000h], esp
  loc_0040EC19: sub esp, 00000024h
  loc_0040EC1C: push ebx
  loc_0040EC1D: push esi
  loc_0040EC1E: push edi
  loc_0040EC1F: mov var_C, esp
  loc_0040EC22: mov var_8, 00401330h
  loc_0040EC29: mov esi, Me
  loc_0040EC2C: mov eax, esi
  loc_0040EC2E: and eax, 00000001h
  loc_0040EC31: mov var_4, eax
  loc_0040EC34: and esi, FFFFFFFEh
  loc_0040EC37: push esi
  loc_0040EC38: mov Me, esi
  loc_0040EC3B: mov ecx, [esi]
  loc_0040EC3D: call [ecx+00000004h]
  loc_0040EC40: mov edx, [esi]
  loc_0040EC42: xor eax, eax
  loc_0040EC44: push esi
  loc_0040EC45: mov var_18, eax
  loc_0040EC48: mov var_1C, eax
  loc_0040EC4B: mov var_20, eax
  loc_0040EC4E: call [edx+000003B0h]
  loc_0040EC54: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_0040EC5A: push eax
  loc_0040EC5B: lea eax, var_1C
  loc_0040EC5E: push eax
  loc_0040EC5F: call ebx
  loc_0040EC61: mov ecx, [esi]
  loc_0040EC63: push esi
  loc_0040EC64: mov var_2C, eax
  loc_0040EC67: call [ecx+000003B0h]
  loc_0040EC6D: lea edx, var_18
  loc_0040EC70: push eax
  loc_0040EC71: push edx
  loc_0040EC72: call ebx
  loc_0040EC74: mov edi, eax
  loc_0040EC76: lea ecx, var_20
  loc_0040EC79: push ecx
  loc_0040EC7A: push edi
  loc_0040EC7B: mov eax, [edi]
  loc_0040EC7D: call [eax+00000098h]
  loc_0040EC83: test eax, eax
  loc_0040EC85: fnclex
  loc_0040EC87: jge 0040EC9Bh
  loc_0040EC89: push 00000098h
  loc_0040EC8E: push 00403858h
  loc_0040EC93: push edi
  loc_0040EC94: push eax
  loc_0040EC95: call [00401040h] ; __vbaHresultCheckObj
  loc_0040EC9B: mov eax, var_20
  loc_0040EC9E: mov edi, var_2C
  loc_0040ECA1: not eax
  loc_0040ECA3: mov edx, [edi]
  loc_0040ECA5: push eax
  loc_0040ECA6: push edi
  loc_0040ECA7: call [edx+0000009Ch]
  loc_0040ECAD: test eax, eax
  loc_0040ECAF: fnclex
  loc_0040ECB1: jge 0040ECC5h
  loc_0040ECB3: push 0000009Ch
  loc_0040ECB8: push 00403858h
  loc_0040ECBD: push edi
  loc_0040ECBE: push eax
  loc_0040ECBF: call [00401040h] ; __vbaHresultCheckObj
  loc_0040ECC5: lea ecx, var_1C
  loc_0040ECC8: lea edx, var_18
  loc_0040ECCB: push ecx
  loc_0040ECCC: push edx
  loc_0040ECCD: push 00000002h
  loc_0040ECCF: call [0040102Ch] ; __vbaFreeObjList
  loc_0040ECD5: mov eax, [esi]
  loc_0040ECD7: add esp, 0000000Ch
  loc_0040ECDA: push esi
  loc_0040ECDB: call [eax+000003B0h]
  loc_0040ECE1: lea ecx, var_18
  loc_0040ECE4: push eax
  loc_0040ECE5: push ecx
  loc_0040ECE6: call ebx
  loc_0040ECE8: mov edi, eax
  loc_0040ECEA: lea eax, var_20
  loc_0040ECED: push eax
  loc_0040ECEE: push edi
  loc_0040ECEF: mov edx, [edi]
  loc_0040ECF1: call [edx+00000098h]
  loc_0040ECF7: test eax, eax
  loc_0040ECF9: fnclex
  loc_0040ECFB: jge 0040ED0Fh
  loc_0040ECFD: push 00000098h
  loc_0040ED02: push 00403858h
  loc_0040ED07: push edi
  loc_0040ED08: push eax
  loc_0040ED09: call [00401040h] ; __vbaHresultCheckObj
  loc_0040ED0F: xor ecx, ecx
  loc_0040ED11: cmp var_20, FFFFFFh
  loc_0040ED16: setz cl
  loc_0040ED19: neg ecx
  loc_0040ED1B: mov edi, ecx
  loc_0040ED1D: lea ecx, var_18
  loc_0040ED20: call [00401180h] ; __vbaFreeObj
  loc_0040ED26: mov edx, [esi]
  loc_0040ED28: push esi
  loc_0040ED29: test di, di
  loc_0040ED2C: jz 0040ED50h
  loc_0040ED2E: call [edx+000003FCh]
  loc_0040ED34: push eax
  loc_0040ED35: lea eax, var_18
  loc_0040ED38: push eax
  loc_0040ED39: call ebx
  loc_0040ED3B: mov esi, eax
  loc_0040ED3D: push 00403C78h ; "&Instructions Off"
  loc_0040ED42: push esi
  loc_0040ED43: mov ecx, [esi]
  loc_0040ED45: call [ecx+00000064h]
  loc_0040ED48: test eax, eax
  loc_0040ED4A: fnclex
  loc_0040ED4C: jge 0040ED7Fh
  loc_0040ED4E: jmp 0040ED70h
  loc_0040ED50: call [edx+000003FCh]
  loc_0040ED56: push eax
  loc_0040ED57: lea eax, var_18
  loc_0040ED5A: push eax
  loc_0040ED5B: call ebx
  loc_0040ED5D: mov esi, eax
  loc_0040ED5F: push 00403CA0h ; "&Instructions On"
  loc_0040ED64: push esi
  loc_0040ED65: mov ecx, [esi]
  loc_0040ED67: call [ecx+00000064h]
  loc_0040ED6A: test eax, eax
  loc_0040ED6C: fnclex
  loc_0040ED6E: jge 0040ED7Fh
  loc_0040ED70: push 00000064h
  loc_0040ED72: push 00403C38h
  loc_0040ED77: push esi
  loc_0040ED78: push eax
  loc_0040ED79: call [00401040h] ; __vbaHresultCheckObj
  loc_0040ED7F: lea ecx, var_18
  loc_0040ED82: call [00401180h] ; __vbaFreeObj
  loc_0040ED88: mov var_4, 00000000h
  loc_0040ED8F: push 0040EDABh
  loc_0040ED94: jmp 0040EDAAh
  loc_0040ED96: lea edx, var_1C
  loc_0040ED99: lea eax, var_18
  loc_0040ED9C: push edx
  loc_0040ED9D: push eax
  loc_0040ED9E: push 00000002h
  loc_0040EDA0: call [0040102Ch] ; __vbaFreeObjList
  loc_0040EDA6: add esp, 0000000Ch
  loc_0040EDA9: ret
  loc_0040EDAA: ret
  loc_0040EDAB: mov eax, Me
  loc_0040EDAE: push eax
  loc_0040EDAF: mov ecx, [eax]
  loc_0040EDB1: call [ecx+00000008h]
  loc_0040EDB4: mov eax, var_4
  loc_0040EDB7: mov ecx, var_14
  loc_0040EDBA: pop edi
  loc_0040EDBB: pop esi
  loc_0040EDBC: mov fs:[00000000h], ecx
  loc_0040EDC3: pop ebx
  loc_0040EDC4: mov esp, ebp
  loc_0040EDC6: pop ebp
  loc_0040EDC7: retn 0004h
End Sub

Private Sub mnuExit_Click() '412140
  loc_00412140: push ebp
  loc_00412141: mov ebp, esp
  loc_00412143: sub esp, 0000000Ch
  loc_00412146: push 00401546h ; __vbaExceptHandler
  loc_0041214B: mov eax, fs:[00000000h]
  loc_00412151: push eax
  loc_00412152: mov fs:[00000000h], esp
  loc_00412159: sub esp, 00000018h
  loc_0041215C: push ebx
  loc_0041215D: push esi
  loc_0041215E: push edi
  loc_0041215F: mov var_C, esp
  loc_00412162: mov var_8, 00401458h
  loc_00412169: mov edi, Me
  loc_0041216C: mov eax, edi
  loc_0041216E: and eax, 00000001h
  loc_00412171: mov var_4, eax
  loc_00412174: and edi, FFFFFFFEh
  loc_00412177: push edi
  loc_00412178: mov Me, edi
  loc_0041217B: mov ecx, [edi]
  loc_0041217D: call [ecx+00000004h]
  loc_00412180: mov eax, [0041567Ch]
  loc_00412185: xor ebx, ebx
  loc_00412187: cmp eax, ebx
  loc_00412189: mov var_18, ebx
  loc_0041218C: jnz 0041219Eh
  loc_0041218E: push 0041567Ch
  loc_00412193: push 00403C18h
  loc_00412198: call [0040111Ch] ; __vbaNew2
  loc_0041219E: mov esi, [0041567Ch]
  loc_004121A4: lea eax, var_18
  loc_004121A7: push edi
  loc_004121A8: push eax
  loc_004121A9: mov edx, [esi]
  loc_004121AB: mov var_2C, edx
  loc_004121AE: call [0040106Ch] ; __vbaObjSetAddref
  loc_004121B4: mov ecx, var_2C
  loc_004121B7: push eax
  loc_004121B8: push esi
  loc_004121B9: call [ecx+00000010h]
  loc_004121BC: cmp eax, ebx
  loc_004121BE: fnclex
  loc_004121C0: jge 004121D1h
  loc_004121C2: push 00000010h
  loc_004121C4: push 00403C08h
  loc_004121C9: push esi
  loc_004121CA: push eax
  loc_004121CB: call [00401040h] ; __vbaHresultCheckObj
  loc_004121D1: lea ecx, var_18
  loc_004121D4: call [00401180h] ; __vbaFreeObj
  loc_004121DA: mov var_4, ebx
  loc_004121DD: push 004121EFh
  loc_004121E2: jmp 004121EEh
  loc_004121E4: lea ecx, var_18
  loc_004121E7: call [00401180h] ; __vbaFreeObj
  loc_004121ED: ret
  loc_004121EE: ret
  loc_004121EF: mov eax, Me
  loc_004121F2: push eax
  loc_004121F3: mov edx, [eax]
  loc_004121F5: call [edx+00000008h]
  loc_004121F8: mov eax, var_4
  loc_004121FB: mov ecx, var_14
  loc_004121FE: pop edi
  loc_004121FF: pop esi
  loc_00412200: mov fs:[00000000h], ecx
  loc_00412207: pop ebx
  loc_00412208: mov esp, ebp
  loc_0041220A: pop ebp
  loc_0041220B: retn 0004h
End Sub

Private Sub mnuChangeHypValue_Click() '40C7C0
  loc_0040C7C0: push ebp
  loc_0040C7C1: mov ebp, esp
  loc_0040C7C3: sub esp, 0000000Ch
  loc_0040C7C6: push 00401546h ; __vbaExceptHandler
  loc_0040C7CB: mov eax, fs:[00000000h]
  loc_0040C7D1: push eax
  loc_0040C7D2: mov fs:[00000000h], esp
  loc_0040C7D9: sub esp, 00000030h
  loc_0040C7DC: push ebx
  loc_0040C7DD: push esi
  loc_0040C7DE: push edi
  loc_0040C7DF: mov var_C, esp
  loc_0040C7E2: mov var_8, 00401228h
  loc_0040C7E9: mov eax, Me
  loc_0040C7EC: mov ecx, eax
  loc_0040C7EE: and ecx, 00000001h
  loc_0040C7F1: mov var_4, ecx
  loc_0040C7F4: and al, FEh
  loc_0040C7F6: push eax
  loc_0040C7F7: mov Me, eax
  loc_0040C7FA: mov edx, [eax]
  loc_0040C7FC: call [edx+00000004h]
  loc_0040C7FF: mov eax, [004151F8h]
  loc_0040C804: test eax, eax
  loc_0040C806: jnz 0040C818h
  loc_0040C808: push 004151F8h
  loc_0040C80D: push 00401C68h
  loc_0040C812: call [0040111Ch] ; __vbaNew2
  loc_0040C818: sub esp, 00000010h
  loc_0040C81B: mov ecx, 0000000Ah
  loc_0040C820: mov ebx, esp
  loc_0040C822: mov var_24, ecx
  loc_0040C825: mov eax, 80020004h
  loc_0040C82A: sub esp, 00000010h
  loc_0040C82D: mov [ebx], ecx
  loc_0040C82F: mov ecx, var_30
  loc_0040C832: mov edx, eax
  loc_0040C834: mov esi, [004151F8h]
  loc_0040C83A: mov [ebx+00000004h], ecx
  loc_0040C83D: mov ecx, esp
  loc_0040C83F: mov edi, [esi]
  loc_0040C841: push esi
  loc_0040C842: mov [ebx+00000008h], eax
  loc_0040C845: mov eax, var_28
  loc_0040C848: mov [ebx+0000000Ch], eax
  loc_0040C84B: mov eax, var_24
  loc_0040C84E: mov [ecx], eax
  loc_0040C850: mov eax, var_20
  loc_0040C853: mov [ecx+00000004h], eax
  loc_0040C856: mov [ecx+00000008h], edx
  loc_0040C859: mov edx, var_18
  loc_0040C85C: mov [ecx+0000000Ch], edx
  loc_0040C85F: call [edi+000002B0h]
  loc_0040C865: test eax, eax
  loc_0040C867: fnclex
  loc_0040C869: jge 0040C87Dh
  loc_0040C86B: push 000002B0h
  loc_0040C870: push 004039BCh
  loc_0040C875: push esi
  loc_0040C876: push eax
  loc_0040C877: call [00401040h] ; __vbaHresultCheckObj
  loc_0040C87D: mov var_4, 00000000h
  loc_0040C884: mov eax, Me
  loc_0040C887: push eax
  loc_0040C888: mov ecx, [eax]
  loc_0040C88A: call [ecx+00000008h]
  loc_0040C88D: mov eax, var_4
  loc_0040C890: mov ecx, var_14
  loc_0040C893: pop edi
  loc_0040C894: pop esi
  loc_0040C895: mov fs:[00000000h], ecx
  loc_0040C89C: pop ebx
  loc_0040C89D: mov esp, ebp
  loc_0040C89F: pop ebp
  loc_0040C8A0: retn 0004h
End Sub

Private Sub mnuShowPred_Click() '40E710
  loc_0040E710: push ebp
  loc_0040E711: mov ebp, esp
  loc_0040E713: sub esp, 0000000Ch
  loc_0040E716: push 00401546h ; __vbaExceptHandler
  loc_0040E71B: mov eax, fs:[00000000h]
  loc_0040E721: push eax
  loc_0040E722: mov fs:[00000000h], esp
  loc_0040E729: sub esp, 00000024h
  loc_0040E72C: push ebx
  loc_0040E72D: push esi
  loc_0040E72E: push edi
  loc_0040E72F: mov var_C, esp
  loc_0040E732: mov var_8, 00401300h
  loc_0040E739: mov esi, Me
  loc_0040E73C: mov eax, esi
  loc_0040E73E: and eax, 00000001h
  loc_0040E741: mov var_4, eax
  loc_0040E744: and esi, FFFFFFFEh
  loc_0040E747: push esi
  loc_0040E748: mov Me, esi
  loc_0040E74B: mov ecx, [esi]
  loc_0040E74D: call [ecx+00000004h]
  loc_0040E750: mov edx, [esi]
  loc_0040E752: xor eax, eax
  loc_0040E754: push esi
  loc_0040E755: mov var_18, eax
  loc_0040E758: mov var_1C, eax
  loc_0040E75B: mov var_20, eax
  loc_0040E75E: call [edx+000003E4h]
  loc_0040E764: mov edi, [0040105Ch] ; __vbaObjSet
  loc_0040E76A: push eax
  loc_0040E76B: lea eax, var_1C
  loc_0040E76E: push eax
  loc_0040E76F: call edi
  loc_0040E771: mov ecx, [esi]
  loc_0040E773: push esi
  loc_0040E774: mov ebx, eax
  loc_0040E776: call [ecx+000003E4h]
  loc_0040E77C: lea edx, var_18
  loc_0040E77F: push eax
  loc_0040E780: push edx
  loc_0040E781: call edi
  loc_0040E783: mov edi, eax
  loc_0040E785: lea ecx, var_20
  loc_0040E788: push ecx
  loc_0040E789: push edi
  loc_0040E78A: mov eax, [edi]
  loc_0040E78C: call [eax+00000068h]
  loc_0040E78F: test eax, eax
  loc_0040E791: fnclex
  loc_0040E793: jge 0040E7A4h
  loc_0040E795: push 00000068h
  loc_0040E797: push 00403C38h
  loc_0040E79C: push edi
  loc_0040E79D: push eax
  loc_0040E79E: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E7A4: mov eax, var_20
  loc_0040E7A7: mov edx, [ebx]
  loc_0040E7A9: not eax
  loc_0040E7AB: push eax
  loc_0040E7AC: push ebx
  loc_0040E7AD: call [edx+0000006Ch]
  loc_0040E7B0: test eax, eax
  loc_0040E7B2: fnclex
  loc_0040E7B4: jge 0040E7C5h
  loc_0040E7B6: push 0000006Ch
  loc_0040E7B8: push 00403C38h
  loc_0040E7BD: push ebx
  loc_0040E7BE: push eax
  loc_0040E7BF: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E7C5: lea ecx, var_1C
  loc_0040E7C8: lea edx, var_18
  loc_0040E7CB: push ecx
  loc_0040E7CC: push edx
  loc_0040E7CD: push 00000002h
  loc_0040E7CF: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E7D5: mov eax, [esi]
  loc_0040E7D7: add esp, 0000000Ch
  loc_0040E7DA: push esi
  loc_0040E7DB: call [eax+000003E4h]
  loc_0040E7E1: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_0040E7E7: lea ecx, var_18
  loc_0040E7EA: push eax
  loc_0040E7EB: push ecx
  loc_0040E7EC: call ebx
  loc_0040E7EE: mov edi, eax
  loc_0040E7F0: lea eax, var_20
  loc_0040E7F3: push eax
  loc_0040E7F4: push edi
  loc_0040E7F5: mov edx, [edi]
  loc_0040E7F7: call [edx+00000068h]
  loc_0040E7FA: test eax, eax
  loc_0040E7FC: fnclex
  loc_0040E7FE: jge 0040E80Fh
  loc_0040E800: push 00000068h
  loc_0040E802: push 00403C38h
  loc_0040E807: push edi
  loc_0040E808: push eax
  loc_0040E809: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E80F: mov cx, var_20
  loc_0040E813: mov [esi+00000070h], cx
  loc_0040E817: lea ecx, var_18
  loc_0040E81A: call [00401180h] ; __vbaFreeObj
  loc_0040E820: cmp [esi+00000070h], FFFFFFh
  loc_0040E825: jnz 0040E830h
  loc_0040E827: mov edx, [esi]
  loc_0040E829: push esi
  loc_0040E82A: call [edx+00000798h]
  loc_0040E830: mov eax, [esi]
  loc_0040E832: push esi
  loc_0040E833: call [eax+00000304h]
  loc_0040E839: lea ecx, var_18
  loc_0040E83C: push eax
  loc_0040E83D: push ecx
  loc_0040E83E: call ebx
  loc_0040E840: mov edi, eax
  loc_0040E842: mov ax, [esi+00000070h]
  loc_0040E846: push eax
  loc_0040E847: push edi
  loc_0040E848: mov edx, [edi]
  loc_0040E84A: call [edx+00000084h]
  loc_0040E850: test eax, eax
  loc_0040E852: fnclex
  loc_0040E854: jge 0040E868h
  loc_0040E856: push 00000084h
  loc_0040E85B: push 00403C48h
  loc_0040E860: push edi
  loc_0040E861: push eax
  loc_0040E862: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E868: lea ecx, var_18
  loc_0040E86B: call [00401180h] ; __vbaFreeObj
  loc_0040E871: mov var_4, 00000000h
  loc_0040E878: push 0040E894h
  loc_0040E87D: jmp 0040E893h
  loc_0040E87F: lea ecx, var_1C
  loc_0040E882: lea edx, var_18
  loc_0040E885: push ecx
  loc_0040E886: push edx
  loc_0040E887: push 00000002h
  loc_0040E889: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E88F: add esp, 0000000Ch
  loc_0040E892: ret
  loc_0040E893: ret
  loc_0040E894: mov eax, Me
  loc_0040E897: push eax
  loc_0040E898: mov ecx, [eax]
  loc_0040E89A: call [ecx+00000008h]
  loc_0040E89D: mov eax, var_4
  loc_0040E8A0: mov ecx, var_14
  loc_0040E8A3: pop edi
  loc_0040E8A4: pop esi
  loc_0040E8A5: mov fs:[00000000h], ecx
  loc_0040E8AC: pop ebx
  loc_0040E8AD: mov esp, ebp
  loc_0040E8AF: pop ebp
  loc_0040E8B0: retn 0004h
End Sub

Private Sub mnuShowSlider_Click() '40E8C0
  loc_0040E8C0: push ebp
  loc_0040E8C1: mov ebp, esp
  loc_0040E8C3: sub esp, 0000000Ch
  loc_0040E8C6: push 00401546h ; __vbaExceptHandler
  loc_0040E8CB: mov eax, fs:[00000000h]
  loc_0040E8D1: push eax
  loc_0040E8D2: mov fs:[00000000h], esp
  loc_0040E8D9: sub esp, 00000024h
  loc_0040E8DC: push ebx
  loc_0040E8DD: push esi
  loc_0040E8DE: push edi
  loc_0040E8DF: mov var_C, esp
  loc_0040E8E2: mov var_8, 00401310h
  loc_0040E8E9: mov esi, Me
  loc_0040E8EC: mov eax, esi
  loc_0040E8EE: and eax, 00000001h
  loc_0040E8F1: mov var_4, eax
  loc_0040E8F4: and esi, FFFFFFFEh
  loc_0040E8F7: push esi
  loc_0040E8F8: mov Me, esi
  loc_0040E8FB: mov ecx, [esi]
  loc_0040E8FD: call [ecx+00000004h]
  loc_0040E900: mov edx, [esi]
  loc_0040E902: xor eax, eax
  loc_0040E904: push esi
  loc_0040E905: mov var_18, eax
  loc_0040E908: mov var_1C, eax
  loc_0040E90B: mov var_20, eax
  loc_0040E90E: call [edx+000003E8h]
  loc_0040E914: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_0040E91A: push eax
  loc_0040E91B: lea eax, var_1C
  loc_0040E91E: push eax
  loc_0040E91F: call ebx
  loc_0040E921: mov ecx, [esi]
  loc_0040E923: push esi
  loc_0040E924: mov var_2C, eax
  loc_0040E927: call [ecx+000003E8h]
  loc_0040E92D: lea edx, var_18
  loc_0040E930: push eax
  loc_0040E931: push edx
  loc_0040E932: call ebx
  loc_0040E934: mov edi, eax
  loc_0040E936: lea ecx, var_20
  loc_0040E939: push ecx
  loc_0040E93A: push edi
  loc_0040E93B: mov eax, [edi]
  loc_0040E93D: call [eax+00000068h]
  loc_0040E940: test eax, eax
  loc_0040E942: fnclex
  loc_0040E944: jge 0040E955h
  loc_0040E946: push 00000068h
  loc_0040E948: push 00403C38h
  loc_0040E94D: push edi
  loc_0040E94E: push eax
  loc_0040E94F: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E955: mov eax, var_20
  loc_0040E958: mov edi, var_2C
  loc_0040E95B: not eax
  loc_0040E95D: mov edx, [edi]
  loc_0040E95F: push eax
  loc_0040E960: push edi
  loc_0040E961: call [edx+0000006Ch]
  loc_0040E964: test eax, eax
  loc_0040E966: fnclex
  loc_0040E968: jge 0040E979h
  loc_0040E96A: push 0000006Ch
  loc_0040E96C: push 00403C38h
  loc_0040E971: push edi
  loc_0040E972: push eax
  loc_0040E973: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E979: lea ecx, var_1C
  loc_0040E97C: lea edx, var_18
  loc_0040E97F: push ecx
  loc_0040E980: push edx
  loc_0040E981: push 00000002h
  loc_0040E983: call [0040102Ch] ; __vbaFreeObjList
  loc_0040E989: mov eax, [esi]
  loc_0040E98B: add esp, 0000000Ch
  loc_0040E98E: push esi
  loc_0040E98F: call [eax+000002FCh]
  loc_0040E995: lea ecx, var_1C
  loc_0040E998: push eax
  loc_0040E999: push ecx
  loc_0040E99A: call ebx
  loc_0040E99C: mov edx, [esi]
  loc_0040E99E: push esi
  loc_0040E99F: mov edi, eax
  loc_0040E9A1: call [edx+000003E8h]
  loc_0040E9A7: push eax
  loc_0040E9A8: lea eax, var_18
  loc_0040E9AB: push eax
  loc_0040E9AC: call ebx
  loc_0040E9AE: mov esi, eax
  loc_0040E9B0: lea edx, var_20
  loc_0040E9B3: push edx
  loc_0040E9B4: push esi
  loc_0040E9B5: mov ecx, [esi]
  loc_0040E9B7: call [ecx+00000068h]
  loc_0040E9BA: test eax, eax
  loc_0040E9BC: fnclex
  loc_0040E9BE: jge 0040E9CFh
  loc_0040E9C0: push 00000068h
  loc_0040E9C2: push 00403C38h
  loc_0040E9C7: push esi
  loc_0040E9C8: push eax
  loc_0040E9C9: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E9CF: mov ecx, var_20
  loc_0040E9D2: mov eax, [edi]
  loc_0040E9D4: push ecx
  loc_0040E9D5: push edi
  loc_0040E9D6: call [eax+00000084h]
  loc_0040E9DC: test eax, eax
  loc_0040E9DE: fnclex
  loc_0040E9E0: jge 0040E9F4h
  loc_0040E9E2: push 00000084h
  loc_0040E9E7: push 00403814h
  loc_0040E9EC: push edi
  loc_0040E9ED: push eax
  loc_0040E9EE: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E9F4: lea edx, var_1C
  loc_0040E9F7: lea eax, var_18
  loc_0040E9FA: push edx
  loc_0040E9FB: push eax
  loc_0040E9FC: push 00000002h
  loc_0040E9FE: call [0040102Ch] ; __vbaFreeObjList
  loc_0040EA04: add esp, 0000000Ch
  loc_0040EA07: mov var_4, 00000000h
  loc_0040EA0E: push 0040EA2Ah
  loc_0040EA13: jmp 0040EA29h
  loc_0040EA15: lea ecx, var_1C
  loc_0040EA18: lea edx, var_18
  loc_0040EA1B: push ecx
  loc_0040EA1C: push edx
  loc_0040EA1D: push 00000002h
  loc_0040EA1F: call [0040102Ch] ; __vbaFreeObjList
  loc_0040EA25: add esp, 0000000Ch
  loc_0040EA28: ret
  loc_0040EA29: ret
  loc_0040EA2A: mov eax, Me
  loc_0040EA2D: push eax
  loc_0040EA2E: mov ecx, [eax]
  loc_0040EA30: call [ecx+00000008h]
  loc_0040EA33: mov eax, var_4
  loc_0040EA36: mov ecx, var_14
  loc_0040EA39: pop edi
  loc_0040EA3A: pop esi
  loc_0040EA3B: mov fs:[00000000h], ecx
  loc_0040EA42: pop ebx
  loc_0040EA43: mov esp, ebp
  loc_0040EA45: pop ebp
  loc_0040EA46: retn 0004h
End Sub

Private Sub txtXg_Change() '411840
  loc_00411840: push ebp
  loc_00411841: mov ebp, esp
  loc_00411843: sub esp, 0000000Ch
  loc_00411846: push 00401546h ; __vbaExceptHandler
  loc_0041184B: mov eax, fs:[00000000h]
  loc_00411851: push eax
  loc_00411852: mov fs:[00000000h], esp
  loc_00411859: sub esp, 0000001Ch
  loc_0041185C: push ebx
  loc_0041185D: push esi
  loc_0041185E: push edi
  loc_0041185F: mov var_C, esp
  loc_00411862: mov var_8, 00401418h
  loc_00411869: mov esi, Me
  loc_0041186C: mov eax, esi
  loc_0041186E: and eax, 00000001h
  loc_00411871: mov var_4, eax
  loc_00411874: and esi, FFFFFFFEh
  loc_00411877: push esi
  loc_00411878: mov Me, esi
  loc_0041187B: mov ecx, [esi]
  loc_0041187D: call [ecx+00000004h]
  loc_00411880: mov edx, [esi]
  loc_00411882: xor eax, eax
  loc_00411884: push esi
  loc_00411885: mov var_18, eax
  loc_00411888: mov var_1C, eax
  loc_0041188B: call [edx+0000031Ch]
  loc_00411891: mov ebx, [0040105Ch] ; __vbaObjSet
  loc_00411897: push eax
  loc_00411898: lea eax, var_1C
  loc_0041189B: push eax
  loc_0041189C: call ebx
  loc_0041189E: mov edi, eax
  loc_004118A0: lea edx, var_18
  loc_004118A3: push edx
  loc_004118A4: push edi
  loc_004118A5: mov ecx, [edi]
  loc_004118A7: call [ecx+000000A0h]
  loc_004118AD: test eax, eax
  loc_004118AF: fnclex
  loc_004118B1: jge 004118C5h
  loc_004118B3: push 000000A0h
  loc_004118B8: push 00403848h
  loc_004118BD: push edi
  loc_004118BE: push eax
  loc_004118BF: call [00401040h] ; __vbaHresultCheckObj
  loc_004118C5: mov eax, var_18
  loc_004118C8: push eax
  loc_004118C9: push 00403D64h
  loc_004118CE: call [00401098h] ; __vbaStrCmp
  loc_004118D4: mov edi, eax
  loc_004118D6: lea ecx, var_18
  loc_004118D9: neg edi
  loc_004118DB: sbb edi, edi
  loc_004118DD: inc edi
  loc_004118DE: neg edi
  loc_004118E0: call [00401184h] ; __vbaFreeStr
  loc_004118E6: lea ecx, var_1C
  loc_004118E9: call [00401180h] ; __vbaFreeObj
  loc_004118EF: test di, di
  loc_004118F2: jnz 00411998h
  loc_004118F8: mov ecx, [esi]
  loc_004118FA: push esi
  loc_004118FB: call [ecx+0000031Ch]
  loc_00411901: lea edx, var_1C
  loc_00411904: push eax
  loc_00411905: push edx
  loc_00411906: call ebx
  loc_00411908: mov edi, eax
  loc_0041190A: lea ecx, var_18
  loc_0041190D: push ecx
  loc_0041190E: push edi
  loc_0041190F: mov eax, [edi]
  loc_00411911: call [eax+000000A0h]
  loc_00411917: test eax, eax
  loc_00411919: fnclex
  loc_0041191B: jge 0041192Fh
  loc_0041191D: push 000000A0h
  loc_00411922: push 00403848h
  loc_00411927: push edi
  loc_00411928: push eax
  loc_00411929: call [00401040h] ; __vbaHresultCheckObj
  loc_0041192F: mov edx, var_18
  loc_00411932: push edx
  loc_00411933: call [004010A0h] ; __vbaR4Str
  loc_00411939: push ecx
  loc_0041193A: fstp real4 ptr [esp]
  loc_0041193D: call [004010A8h] ; __vbaStrR4
  loc_00411943: mov edx, eax
  loc_00411945: mov ecx, 0041502Ch
  loc_0041194A: call [00401164h] ; __vbaStrMove
  loc_00411950: lea ecx, var_18
  loc_00411953: call [00401184h] ; __vbaFreeStr
  loc_00411959: lea ecx, var_1C
  loc_0041195C: call [00401180h] ; __vbaFreeObj
  loc_00411962: mov eax, [00415028h]
  loc_00411967: mov ecx, [0041502Ch]
  loc_0041196D: push eax
  loc_0041196E: push ecx
  loc_0041196F: call [004010A0h] ; __vbaR4Str
  loc_00411975: push ecx
  loc_00411976: lea edx, [esi+0000003Ch]
  loc_00411979: fstp real4 ptr [esp]
  loc_0041197C: push 00000001h
  loc_0041197E: push 00000002h
  loc_00411980: push edx
  loc_00411981: call 00409C80h
  loc_00411986: mov eax, [esi]
  loc_00411988: push esi
  loc_00411989: call [eax+00000780h]
  loc_0041198F: mov ecx, [esi]
  loc_00411991: push esi
  loc_00411992: call [ecx+0000078Ch]
  loc_00411998: mov var_4, 00000000h
  loc_0041199F: fwait
  loc_004119A0: push 004119BBh
  loc_004119A5: jmp 004119BAh
  loc_004119A7: lea ecx, var_18
  loc_004119AA: call [00401184h] ; __vbaFreeStr
  loc_004119B0: lea ecx, var_1C
  loc_004119B3: call [00401180h] ; __vbaFreeObj
  loc_004119B9: ret
  loc_004119BA: ret
  loc_004119BB: mov eax, Me
  loc_004119BE: push eax
  loc_004119BF: mov edx, [eax]
  loc_004119C1: call [edx+00000008h]
  loc_004119C4: mov eax, var_4
  loc_004119C7: mov ecx, var_14
  loc_004119CA: pop edi
  loc_004119CB: pop esi
  loc_004119CC: mov fs:[00000000h], ecx
  loc_004119D3: pop ebx
  loc_004119D4: mov esp, ebp
  loc_004119D6: pop ebp
  loc_004119D7: retn 0004h
End Sub

Public Sub RescaleData() '410190
  loc_00410190: push ebp
  loc_00410191: mov ebp, esp
  loc_00410193: sub esp, 0000000Ch
  loc_00410196: push 00401546h ; __vbaExceptHandler
  loc_0041019B: mov eax, fs:[00000000h]
  loc_004101A1: push eax
  loc_004101A2: mov fs:[00000000h], esp
  loc_004101A9: sub esp, 00000064h
  loc_004101AC: push ebx
  loc_004101AD: push esi
  loc_004101AE: push edi
  loc_004101AF: mov var_C, esp
  loc_004101B2: mov var_8, 004013A0h
  loc_004101B9: xor ebx, ebx
  loc_004101BB: mov var_4, ebx
  loc_004101BE: mov esi, Me
  loc_004101C1: push esi
  loc_004101C2: mov eax, [esi]
  loc_004101C4: call [eax+00000004h]
  loc_004101C7: movsx ecx, [esi+00000038h]
  loc_004101CB: mov edi, [004010B8h] ; __vbaRedim
  loc_004101D1: push ebx
  loc_004101D2: push ecx
  loc_004101D3: push 00000001h
  loc_004101D5: lea edx, [esi+00000050h]
  loc_004101D8: push 00000004h
  loc_004101DA: push edx
  loc_004101DB: push 00000004h
  loc_004101DD: push 00000080h
  loc_004101E2: mov var_2C, ebx
  loc_004101E5: mov var_3C, ebx
  loc_004101E8: call edi
  loc_004101EA: movsx eax, [esi+00000038h]
  loc_004101EE: push ebx
  loc_004101EF: push eax
  loc_004101F0: push 00000001h
  loc_004101F2: lea ecx, [esi+00000054h]
  loc_004101F5: push 00000004h
  loc_004101F7: push ecx
  loc_004101F8: push 00000004h
  loc_004101FA: push 00000080h
  loc_004101FF: call edi
  loc_00410201: mov eax, [esi+0000003Ch]
  loc_00410204: add esp, 00000038h
  loc_00410207: cmp eax, ebx
  loc_00410209: jz 00410252h
  loc_0041020B: cmp [eax], 0002h
  loc_0041020F: jnz 00410252h
  loc_00410211: mov edx, [eax+0000001Ch]
  loc_00410214: mov ecx, [eax+00000018h]
  loc_00410217: mov edi, 00000001h
  loc_0041021C: sub edi, edx
  loc_0041021E: cmp edi, ecx
  loc_00410220: jb 00410228h
  loc_00410222: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410228: mov eax, [esi+0000003Ch]
  loc_0041022B: mov ebx, 00000002h
  loc_00410230: mov edx, [eax+00000014h]
  loc_00410233: mov ecx, [eax+00000010h]
  loc_00410236: sub ebx, edx
  loc_00410238: cmp ebx, ecx
  loc_0041023A: jb 00410242h
  loc_0041023C: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410242: mov edx, [esi+0000003Ch]
  loc_00410245: mov eax, [edx+00000018h]
  loc_00410248: imul eax, ebx
  loc_0041024B: add eax, edi
  loc_0041024D: shl eax, 02h
  loc_00410250: jmp 00410258h
  loc_00410252: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410258: mov ecx, [esi+0000003Ch]
  loc_0041025B: test ecx, ecx
  loc_0041025D: mov edx, [ecx+0000000Ch]
  loc_00410260: mov eax, [edx+eax]
  loc_00410263: mov [esi+00000044h], eax
  loc_00410266: jz 004102B5h
  loc_00410268: cmp [ecx], 0002h
  loc_0041026C: jnz 004102B5h
  loc_0041026E: mov edx, [ecx+0000001Ch]
  loc_00410271: mov eax, [ecx+00000018h]
  loc_00410274: mov edi, 00000001h
  loc_00410279: sub edi, edx
  loc_0041027B: cmp edi, eax
  loc_0041027D: jb 00410285h
  loc_0041027F: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410285: mov eax, [esi+0000003Ch]
  loc_00410288: mov ebx, 00000001h
  loc_0041028D: mov edx, [eax+00000014h]
  loc_00410290: mov ecx, [eax+00000010h]
  loc_00410293: sub ebx, edx
  loc_00410295: cmp ebx, ecx
  loc_00410297: jb 0041029Fh
  loc_00410299: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041029F: mov ecx, [esi+0000003Ch]
  loc_004102A2: mov eax, [ecx+00000018h]
  loc_004102A5: imul eax, ebx
  loc_004102A8: add eax, edi
  loc_004102AA: mov edi, [00401094h] ; __vbaGenerateBoundsError
  loc_004102B0: shl eax, 02h
  loc_004102B3: jmp 004102BDh
  loc_004102B5: mov edi, [00401094h] ; __vbaGenerateBoundsError
  loc_004102BB: call edi
  loc_004102BD: mov ecx, [esi+0000003Ch]
  loc_004102C0: mov edx, [ecx+0000000Ch]
  loc_004102C3: mov eax, [edx+eax]
  loc_004102C6: mov edx, [esi+00000044h]
  loc_004102C9: mov [esi+00000048h], eax
  loc_004102CC: mov [esi+00000040h], edx
  loc_004102CF: mov dx, [esi+00000038h]
  loc_004102D3: mov [esi+0000004Ch], eax
  loc_004102D6: mov eax, 00000002h
  loc_004102DB: mov var_50, dx
  loc_004102DF: mov var_18, eax
  loc_004102E2: cmp ax, var_50
  loc_004102E6: jg 004105FCh
  loc_004102EC: test ecx, ecx
  loc_004102EE: jz 00410331h
  loc_004102F0: cmp [ecx], 0002h
  loc_004102F4: jnz 00410331h
  loc_004102F6: mov edx, [ecx+0000001Ch]
  loc_004102F9: movsx ebx, ax
  loc_004102FC: mov eax, [ecx+00000018h]
  loc_004102FF: sub ebx, edx
  loc_00410301: cmp ebx, eax
  loc_00410303: jb 00410307h
  loc_00410305: call edi
  loc_00410307: mov eax, [esi+0000003Ch]
  loc_0041030A: mov edi, 00000002h
  loc_0041030F: mov edx, [eax+00000014h]
  loc_00410312: mov ecx, [eax+00000010h]
  loc_00410315: sub edi, edx
  loc_00410317: cmp edi, ecx
  loc_00410319: jb 00410321h
  loc_0041031B: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410321: mov eax, [esi+0000003Ch]
  loc_00410324: mov eax, [eax+00000018h]
  loc_00410327: imul eax, edi
  loc_0041032A: add eax, ebx
  loc_0041032C: shl eax, 02h
  loc_0041032F: jmp 00410333h
  loc_00410331: call edi
  loc_00410333: mov ecx, [esi+0000003Ch]
  loc_00410336: fld real4 ptr [esi+00000044h]
  loc_00410339: mov edx, [ecx+0000000Ch]
  loc_0041033C: fcomp real4 ptr [edx+eax]
  loc_0041033F: fnstsw ax
  loc_00410341: test ah, 01h
  loc_00410344: jz 004103A2h
  loc_00410346: test ecx, ecx
  loc_00410348: jz 00410390h
  loc_0041034A: cmp [ecx], 0002h
  loc_0041034E: jnz 00410390h
  loc_00410350: movsx ebx, var_18
  loc_00410354: mov edx, [ecx+0000001Ch]
  loc_00410357: mov eax, [ecx+00000018h]
  loc_0041035A: sub ebx, edx
  loc_0041035C: cmp ebx, eax
  loc_0041035E: jb 00410366h
  loc_00410360: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410366: mov eax, [esi+0000003Ch]
  loc_00410369: mov edi, 00000002h
  loc_0041036E: mov edx, [eax+00000014h]
  loc_00410371: mov ecx, [eax+00000010h]
  loc_00410374: sub edi, edx
  loc_00410376: cmp edi, ecx
  loc_00410378: jb 00410380h
  loc_0041037A: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410380: mov eax, [esi+0000003Ch]
  loc_00410383: mov eax, [eax+00000018h]
  loc_00410386: imul eax, edi
  loc_00410389: add eax, ebx
  loc_0041038B: shl eax, 02h
  loc_0041038E: jmp 00410396h
  loc_00410390: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410396: mov ecx, [esi+0000003Ch]
  loc_00410399: mov edx, [ecx+0000000Ch]
  loc_0041039C: mov eax, [edx+eax]
  loc_0041039F: mov [esi+00000044h], eax
  loc_004103A2: test ecx, ecx
  loc_004103A4: jz 004103ECh
  loc_004103A6: cmp [ecx], 0002h
  loc_004103AA: jnz 004103ECh
  loc_004103AC: movsx ebx, var_18
  loc_004103B0: mov edx, [ecx+0000001Ch]
  loc_004103B3: mov eax, [ecx+00000018h]
  loc_004103B6: sub ebx, edx
  loc_004103B8: cmp ebx, eax
  loc_004103BA: jb 004103C2h
  loc_004103BC: call [00401094h] ; __vbaGenerateBoundsError
  loc_004103C2: mov eax, [esi+0000003Ch]
  loc_004103C5: mov edi, 00000002h
  loc_004103CA: mov edx, [eax+00000014h]
  loc_004103CD: mov ecx, [eax+00000010h]
  loc_004103D0: sub edi, edx
  loc_004103D2: cmp edi, ecx
  loc_004103D4: jb 004103DCh
  loc_004103D6: call [00401094h] ; __vbaGenerateBoundsError
  loc_004103DC: mov ecx, [esi+0000003Ch]
  loc_004103DF: mov eax, [ecx+00000018h]
  loc_004103E2: imul eax, edi
  loc_004103E5: add eax, ebx
  loc_004103E7: shl eax, 02h
  loc_004103EA: jmp 004103F2h
  loc_004103EC: call [00401094h] ; __vbaGenerateBoundsError
  loc_004103F2: mov ecx, [esi+0000003Ch]
  loc_004103F5: fld real4 ptr [esi+00000040h]
  loc_004103F8: mov edx, [ecx+0000000Ch]
  loc_004103FB: fcomp real4 ptr [edx+eax]
  loc_004103FE: fnstsw ax
  loc_00410400: test ah, 41h
  loc_00410403: jnz 00410461h
  loc_00410405: test ecx, ecx
  loc_00410407: jz 0041044Fh
  loc_00410409: cmp [ecx], 0002h
  loc_0041040D: jnz 0041044Fh
  loc_0041040F: movsx ebx, var_18
  loc_00410413: mov edx, [ecx+0000001Ch]
  loc_00410416: mov eax, [ecx+00000018h]
  loc_00410419: sub ebx, edx
  loc_0041041B: cmp ebx, eax
  loc_0041041D: jb 00410425h
  loc_0041041F: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410425: mov eax, [esi+0000003Ch]
  loc_00410428: mov edi, 00000002h
  loc_0041042D: mov edx, [eax+00000014h]
  loc_00410430: mov ecx, [eax+00000010h]
  loc_00410433: sub edi, edx
  loc_00410435: cmp edi, ecx
  loc_00410437: jb 0041043Fh
  loc_00410439: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041043F: mov eax, [esi+0000003Ch]
  loc_00410442: mov eax, [eax+00000018h]
  loc_00410445: imul eax, edi
  loc_00410448: add eax, ebx
  loc_0041044A: shl eax, 02h
  loc_0041044D: jmp 00410455h
  loc_0041044F: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410455: mov ecx, [esi+0000003Ch]
  loc_00410458: mov edx, [ecx+0000000Ch]
  loc_0041045B: mov eax, [edx+eax]
  loc_0041045E: mov [esi+00000040h], eax
  loc_00410461: test ecx, ecx
  loc_00410463: jz 004104ABh
  loc_00410465: cmp [ecx], 0002h
  loc_00410469: jnz 004104ABh
  loc_0041046B: movsx ebx, var_18
  loc_0041046F: mov edx, [ecx+0000001Ch]
  loc_00410472: mov eax, [ecx+00000018h]
  loc_00410475: sub ebx, edx
  loc_00410477: cmp ebx, eax
  loc_00410479: jb 00410481h
  loc_0041047B: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410481: mov eax, [esi+0000003Ch]
  loc_00410484: mov edi, 00000001h
  loc_00410489: mov edx, [eax+00000014h]
  loc_0041048C: mov ecx, [eax+00000010h]
  loc_0041048F: sub edi, edx
  loc_00410491: cmp edi, ecx
  loc_00410493: jb 0041049Bh
  loc_00410495: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041049B: mov ecx, [esi+0000003Ch]
  loc_0041049E: mov eax, [ecx+00000018h]
  loc_004104A1: imul eax, edi
  loc_004104A4: add eax, ebx
  loc_004104A6: shl eax, 02h
  loc_004104A9: jmp 004104B1h
  loc_004104AB: call [00401094h] ; __vbaGenerateBoundsError
  loc_004104B1: mov ecx, [esi+0000003Ch]
  loc_004104B4: fld real4 ptr [esi+00000048h]
  loc_004104B7: mov edx, [ecx+0000000Ch]
  loc_004104BA: fcomp real4 ptr [edx+eax]
  loc_004104BD: fnstsw ax
  loc_004104BF: test ah, 01h
  loc_004104C2: jz 00410520h
  loc_004104C4: test ecx, ecx
  loc_004104C6: jz 0041050Eh
  loc_004104C8: cmp [ecx], 0002h
  loc_004104CC: jnz 0041050Eh
  loc_004104CE: movsx ebx, var_18
  loc_004104D2: mov edx, [ecx+0000001Ch]
  loc_004104D5: mov eax, [ecx+00000018h]
  loc_004104D8: sub ebx, edx
  loc_004104DA: cmp ebx, eax
  loc_004104DC: jb 004104E4h
  loc_004104DE: call [00401094h] ; __vbaGenerateBoundsError
  loc_004104E4: mov eax, [esi+0000003Ch]
  loc_004104E7: mov edi, 00000001h
  loc_004104EC: mov edx, [eax+00000014h]
  loc_004104EF: mov ecx, [eax+00000010h]
  loc_004104F2: sub edi, edx
  loc_004104F4: cmp edi, ecx
  loc_004104F6: jb 004104FEh
  loc_004104F8: call [00401094h] ; __vbaGenerateBoundsError
  loc_004104FE: mov eax, [esi+0000003Ch]
  loc_00410501: mov eax, [eax+00000018h]
  loc_00410504: imul eax, edi
  loc_00410507: add eax, ebx
  loc_00410509: shl eax, 02h
  loc_0041050C: jmp 00410514h
  loc_0041050E: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410514: mov ecx, [esi+0000003Ch]
  loc_00410517: mov edx, [ecx+0000000Ch]
  loc_0041051A: mov eax, [edx+eax]
  loc_0041051D: mov [esi+00000048h], eax
  loc_00410520: test ecx, ecx
  loc_00410522: jz 0041056Ah
  loc_00410524: cmp [ecx], 0002h
  loc_00410528: jnz 0041056Ah
  loc_0041052A: movsx ebx, var_18
  loc_0041052E: mov edx, [ecx+0000001Ch]
  loc_00410531: mov eax, [ecx+00000018h]
  loc_00410534: sub ebx, edx
  loc_00410536: cmp ebx, eax
  loc_00410538: jb 00410540h
  loc_0041053A: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410540: mov eax, [esi+0000003Ch]
  loc_00410543: mov edi, 00000001h
  loc_00410548: mov edx, [eax+00000014h]
  loc_0041054B: mov ecx, [eax+00000010h]
  loc_0041054E: sub edi, edx
  loc_00410550: cmp edi, ecx
  loc_00410552: jb 0041055Ah
  loc_00410554: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041055A: mov ecx, [esi+0000003Ch]
  loc_0041055D: mov eax, [ecx+00000018h]
  loc_00410560: imul eax, edi
  loc_00410563: add eax, ebx
  loc_00410565: shl eax, 02h
  loc_00410568: jmp 00410570h
  loc_0041056A: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410570: mov ecx, [esi+0000003Ch]
  loc_00410573: fld real4 ptr [esi+0000004Ch]
  loc_00410576: mov edx, [ecx+0000000Ch]
  loc_00410579: fcomp real4 ptr [edx+eax]
  loc_0041057C: fnstsw ax
  loc_0041057E: test ah, 41h
  loc_00410581: jnz 004105DFh
  loc_00410583: test ecx, ecx
  loc_00410585: jz 004105CDh
  loc_00410587: cmp [ecx], 0002h
  loc_0041058B: jnz 004105CDh
  loc_0041058D: movsx ebx, var_18
  loc_00410591: mov edx, [ecx+0000001Ch]
  loc_00410594: mov eax, [ecx+00000018h]
  loc_00410597: sub ebx, edx
  loc_00410599: cmp ebx, eax
  loc_0041059B: jb 004105A3h
  loc_0041059D: call [00401094h] ; __vbaGenerateBoundsError
  loc_004105A3: mov eax, [esi+0000003Ch]
  loc_004105A6: mov edi, 00000001h
  loc_004105AB: mov edx, [eax+00000014h]
  loc_004105AE: mov ecx, [eax+00000010h]
  loc_004105B1: sub edi, edx
  loc_004105B3: cmp edi, ecx
  loc_004105B5: jb 004105BDh
  loc_004105B7: call [00401094h] ; __vbaGenerateBoundsError
  loc_004105BD: mov eax, [esi+0000003Ch]
  loc_004105C0: mov eax, [eax+00000018h]
  loc_004105C3: imul eax, edi
  loc_004105C6: add eax, ebx
  loc_004105C8: shl eax, 02h
  loc_004105CB: jmp 004105D3h
  loc_004105CD: call [00401094h] ; __vbaGenerateBoundsError
  loc_004105D3: mov ecx, [esi+0000003Ch]
  loc_004105D6: mov edx, [ecx+0000000Ch]
  loc_004105D9: mov eax, [edx+eax]
  loc_004105DC: mov [esi+0000004Ch], eax
  loc_004105DF: mov edi, [00401094h] ; __vbaGenerateBoundsError
  loc_004105E5: mov eax, 00000001h
  loc_004105EA: add ax, var_18
  loc_004105EE: jo 004108E4h
  loc_004105F4: mov var_18, eax
  loc_004105F7: jmp 004102E2h
  loc_004105FC: fld real4 ptr [esi+00000044h]
  loc_004105FF: fsub st0, real4 ptr [esi+00000040h]
  loc_00410602: mov cx, [esi+00000068h]
  loc_00410606: mov ebx, 00000001h
  loc_0041060B: mov var_18, ebx
  loc_0041060E: fstp real4 ptr var_24
  loc_00410611: fnstsw ax
  loc_00410613: test al, 0Dh
  loc_00410615: jnz 004108DFh
  loc_0041061B: fld real4 ptr [esi+00000048h]
  loc_0041061E: fsub st0, real4 ptr [esi+0000004Ch]
  loc_00410621: fstp real4 ptr var_28
  loc_00410624: fnstsw ax
  loc_00410626: test al, 0Dh
  loc_00410628: jnz 004108DFh
  loc_0041062E: sub cx, 00C8h
  loc_00410633: jo 004108E4h
  loc_00410639: movsx edx, cx
  loc_0041063C: mov var_64, edx
  loc_0041063F: mov cx, 00C8h
  loc_00410643: fild real4 ptr var_64
  loc_00410646: fstp real4 ptr var_68
  loc_00410649: fld real4 ptr var_68
  loc_0041064C: cmp [00415000h], 00000000h
  loc_00410653: jnz 0041065Ah
  loc_00410655: fdiv st0, real4 ptr var_24
  loc_00410658: jmp 00410662h
  loc_0041065A: push var_24
  loc_0041065D: call 00401558h ; _adj_fdiv_m32
  loc_00410662: fnstsw ax
  loc_00410664: test al, 0Dh
  loc_00410666: jnz 004108DFh
  loc_0041066C: fstp real4 ptr var_6C
  loc_0041066F: fld real4 ptr var_6C
  loc_00410672: mov eax, var_6C
  loc_00410675: fchs
  loc_00410677: fmul st0, real4 ptr [esi+00000040h]
  loc_0041067A: mov [esi+00000058h], eax
  loc_0041067D: fstp real4 ptr [esi+0000005Ch]
  loc_00410680: fnstsw ax
  loc_00410682: test al, 0Dh
  loc_00410684: jnz 004108DFh
  loc_0041068A: sub cx, [esi+0000006Ah]
  loc_0041068E: jo 004108E4h
  loc_00410694: movsx edx, cx
  loc_00410697: mov var_70, edx
  loc_0041069A: mov cx, [esi+00000038h]
  loc_0041069E: fild real4 ptr var_70
  loc_004106A1: mov var_58, cx
  loc_004106A5: fstp real4 ptr var_74
  loc_004106A8: fld real4 ptr var_74
  loc_004106AB: cmp [00415000h], 00000000h
  loc_004106B2: jnz 004106B9h
  loc_004106B4: fdiv st0, real4 ptr var_28
  loc_004106B7: jmp 004106C1h
  loc_004106B9: push var_28
  loc_004106BC: call 00401558h ; _adj_fdiv_m32
  loc_004106C1: fnstsw ax
  loc_004106C3: test al, 0Dh
  loc_004106C5: jnz 004108DFh
  loc_004106CB: fstp real4 ptr var_78
  loc_004106CE: fld real4 ptr var_78
  loc_004106D1: fmul st0, real4 ptr [esi+00000048h]
  loc_004106D4: mov eax, var_78
  loc_004106D7: mov [esi+00000060h], eax
  loc_004106DA: fsubr st0, real4 ptr [00401350h]
  loc_004106E0: fstp real4 ptr [esi+00000064h]
  loc_004106E3: fnstsw ax
  loc_004106E5: test al, 0Dh
  loc_004106E7: jnz 004108DFh
  loc_004106ED: cmp bx, var_58
  loc_004106F1: jg 004108A3h
  loc_004106F7: mov edx, [esi+0000003Ch]
  loc_004106FA: lea eax, var_2C
  loc_004106FD: push edx
  loc_004106FE: push eax
  loc_004106FF: call [0040114Ch] ; __vbaAryLock
  loc_00410705: mov ecx, var_2C
  loc_00410708: test ecx, ecx
  loc_0041070A: jz 0041074Dh
  loc_0041070C: cmp [ecx], 0002h
  loc_00410710: jnz 0041074Dh
  loc_00410712: mov edx, [ecx+0000001Ch]
  loc_00410715: mov eax, [ecx+00000018h]
  loc_00410718: movsx ebx, bx
  loc_0041071B: sub ebx, edx
  loc_0041071D: cmp ebx, eax
  loc_0041071F: jb 00410726h
  loc_00410721: call edi
  loc_00410723: mov ecx, var_2C
  loc_00410726: mov edx, [ecx+00000014h]
  loc_00410729: mov eax, [ecx+00000010h]
  loc_0041072C: mov edi, 00000002h
  loc_00410731: sub edi, edx
  loc_00410733: cmp edi, eax
  loc_00410735: jb 00410740h
  loc_00410737: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041073D: mov ecx, var_2C
  loc_00410740: mov eax, [ecx+00000018h]
  loc_00410743: imul eax, edi
  loc_00410746: add eax, ebx
  loc_00410748: shl eax, 02h
  loc_0041074B: jmp 00410752h
  loc_0041074D: call edi
  loc_0041074F: mov ecx, var_2C
  loc_00410752: mov ecx, [ecx+0000000Ch]
  loc_00410755: mov edx, [esi]
  loc_00410757: lea edi, var_3C
  loc_0041075A: add ecx, eax
  loc_0041075C: push edi
  loc_0041075D: push ecx
  loc_0041075E: push esi
  loc_0041075F: call [edx+00000778h]
  loc_00410765: lea edx, var_2C
  loc_00410768: push edx
  loc_00410769: call [00401174h] ; __vbaAryUnlock
  loc_0041076F: mov eax, [esi+00000050h]
  loc_00410772: test eax, eax
  loc_00410774: jz 00410797h
  loc_00410776: cmp [eax], 0001h
  loc_0041077A: jnz 00410797h
  loc_0041077C: movsx edi, var_18
  loc_00410780: mov edx, [eax+00000014h]
  loc_00410783: mov ecx, [eax+00000010h]
  loc_00410786: sub edi, edx
  loc_00410788: cmp edi, ecx
  loc_0041078A: jb 00410792h
  loc_0041078C: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410792: shl edi, 02h
  loc_00410795: jmp 0041079Fh
  loc_00410797: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041079D: mov edi, eax
  loc_0041079F: lea eax, var_3C
  loc_004107A2: push eax
  loc_004107A3: call [004010B0h] ; __vbaR4Var
  loc_004107A9: mov ecx, [esi+00000050h]
  loc_004107AC: mov edx, [ecx+0000000Ch]
  loc_004107AF: lea ecx, var_3C
  loc_004107B2: fstp real4 ptr [edx+edi]
  loc_004107B5: call [00401014h] ; __vbaFreeVar
  loc_004107BB: mov eax, [esi+0000003Ch]
  loc_004107BE: lea ecx, var_2C
  loc_004107C1: push eax
  loc_004107C2: push ecx
  loc_004107C3: call [0040114Ch] ; __vbaAryLock
  loc_004107C9: mov ecx, var_2C
  loc_004107CC: test ecx, ecx
  loc_004107CE: jz 00410816h
  loc_004107D0: cmp [ecx], 0002h
  loc_004107D4: jnz 00410816h
  loc_004107D6: movsx ebx, var_18
  loc_004107DA: mov edx, [ecx+0000001Ch]
  loc_004107DD: mov eax, [ecx+00000018h]
  loc_004107E0: sub ebx, edx
  loc_004107E2: cmp ebx, eax
  loc_004107E4: jb 004107EFh
  loc_004107E6: call [00401094h] ; __vbaGenerateBoundsError
  loc_004107EC: mov ecx, var_2C
  loc_004107EF: mov edx, [ecx+00000014h]
  loc_004107F2: mov eax, [ecx+00000010h]
  loc_004107F5: mov edi, 00000001h
  loc_004107FA: sub edi, edx
  loc_004107FC: cmp edi, eax
  loc_004107FE: jb 00410809h
  loc_00410800: call [00401094h] ; __vbaGenerateBoundsError
  loc_00410806: mov ecx, var_2C
  loc_00410809: mov eax, [ecx+00000018h]
  loc_0041080C: imul eax, edi
  loc_0041080F: add eax, ebx
  loc_00410811: shl eax, 02h
  loc_00410814: jmp 0041081Fh
  loc_00410816: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041081C: mov ecx, var_2C
  loc_0041081F: mov ecx, [ecx+0000000Ch]
  loc_00410822: mov edx, [esi]
  loc_00410824: lea edi, var_3C
  loc_00410827: add ecx, eax
  loc_00410829: push edi
  loc_0041082A: push ecx
  loc_0041082B: push esi
  loc_0041082C: call [edx+0000077Ch]
  loc_00410832: lea edx, var_2C
  loc_00410835: push edx
  loc_00410836: call [00401174h] ; __vbaAryUnlock
  loc_0041083C: mov eax, [esi+00000054h]
  loc_0041083F: test eax, eax
  loc_00410841: jz 00410864h
  loc_00410843: cmp [eax], 0001h
  loc_00410847: jnz 00410864h
  loc_00410849: movsx edi, var_18
  loc_0041084D: mov edx, [eax+00000014h]
  loc_00410850: mov ecx, [eax+00000010h]
  loc_00410853: sub edi, edx
  loc_00410855: cmp edi, ecx
  loc_00410857: jb 0041085Fh
  loc_00410859: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041085F: shl edi, 02h
  loc_00410862: jmp 0041086Ch
  loc_00410864: call [00401094h] ; __vbaGenerateBoundsError
  loc_0041086A: mov edi, eax
  loc_0041086C: lea eax, var_3C
  loc_0041086F: push eax
  loc_00410870: call [004010B0h] ; __vbaR4Var
  loc_00410876: mov ecx, [esi+00000054h]
  loc_00410879: mov edx, [ecx+0000000Ch]
  loc_0041087C: lea ecx, var_3C
  loc_0041087F: fstp real4 ptr [edx+edi]
  loc_00410882: call [00401014h] ; __vbaFreeVar
  loc_00410888: mov edi, [00401094h] ; __vbaGenerateBoundsError
  loc_0041088E: mov eax, 00000001h
  loc_00410893: add ax, var_18
  loc_00410897: jo 004108E4h
  loc_00410899: mov var_18, eax
  loc_0041089C: mov ebx, eax
  loc_0041089E: jmp 004106EDh
  loc_004108A3: fwait
  loc_004108A4: push 004108C0h
  loc_004108A9: jmp 004108BFh
  loc_004108AB: lea eax, var_2C
  loc_004108AE: push eax
  loc_004108AF: call [00401174h] ; __vbaAryUnlock
  loc_004108B5: lea ecx, var_3C
  loc_004108B8: call [00401014h] ; __vbaFreeVar
  loc_004108BE: ret
  loc_004108BF: ret
  loc_004108C0: mov eax, Me
  loc_004108C3: push eax
  loc_004108C4: mov ecx, [eax]
  loc_004108C6: call [ecx+00000008h]
  loc_004108C9: mov eax, var_4
  loc_004108CC: mov ecx, var_14
  loc_004108CF: pop edi
  loc_004108D0: pop esi
  loc_004108D1: mov fs:[00000000h], ecx
  loc_004108D8: pop ebx
  loc_004108D9: mov esp, ebp
  loc_004108DB: pop ebp
  loc_004108DC: retn 0004h
End Sub

Private Sub Proc_3_23_40BA20
  loc_0040BA20: push ebp
  loc_0040BA21: mov ebp, esp
  loc_0040BA23: sub esp, 00000008h
  loc_0040BA26: push 00401546h ; __vbaExceptHandler
  loc_0040BA2B: mov eax, fs:[00000000h]
  loc_0040BA31: push eax
  loc_0040BA32: mov fs:[00000000h], esp
  loc_0040BA39: sub esp, 000000B0h
  loc_0040BA3F: push ebx
  loc_0040BA40: push esi
  loc_0040BA41: push edi
  loc_0040BA42: mov var_8, esp
  loc_0040BA45: mov var_4, 004011F8h
  loc_0040BA4C: mov esi, Me
  loc_0040BA4F: xor eax, eax
  loc_0040BA51: mov var_18, eax
  loc_0040BA54: mov var_20, eax
  loc_0040BA57: mov var_30, eax
  loc_0040BA5A: mov var_40, eax
  loc_0040BA5D: mov var_50, eax
  loc_0040BA60: mov var_A0, eax
  loc_0040BA66: mov var_A4, eax
  loc_0040BA6C: mov eax, [esi]
  loc_0040BA6E: push esi
  loc_0040BA6F: mov [esi+00000038h], 000Ah
  loc_0040BA75: mov [esi+00000068h], 0FC3h
  loc_0040BA7B: mov [esi+0000006Ah], 0C8Ah
  loc_0040BA81: call [eax+0000040Ch]
  loc_0040BA87: lea ecx, var_A0
  loc_0040BA8D: push eax
  loc_0040BA8E: push ecx
  loc_0040BA8F: call [0040105Ch] ; __vbaObjSet
  loc_0040BA95: mov dx, [esi+00000038h]
  loc_0040BA99: mov esi, var_5C
  loc_0040BA9C: add dx, 0001h
  loc_0040BAA0: mov ecx, 00000003h
  loc_0040BAA5: jo 0040C479h
  loc_0040BAAB: sub esp, 00000010h
  loc_0040BAAE: mov edi, var_54
  loc_0040BAB1: movsx eax, dx
  loc_0040BAB4: mov edx, esp
  loc_0040BAB6: mov ebx, [0040116Ch] ; __vbaLateIdSt
  loc_0040BABC: push 00000004h
  loc_0040BABE: mov [edx], ecx
  loc_0040BAC0: mov [edx+00000004h], esi
  loc_0040BAC3: mov [edx+00000008h], eax
  loc_0040BAC6: mov eax, var_A0
  loc_0040BACC: push eax
  loc_0040BACD: mov [edx+0000000Ch], edi
  loc_0040BAD0: call ebx
  loc_0040BAD2: sub esp, 00000010h
  loc_0040BAD5: mov ecx, 00000003h
  loc_0040BADA: mov edx, esp
  loc_0040BADC: xor eax, eax
  loc_0040BADE: push 0000000Ah
  loc_0040BAE0: mov [edx], ecx
  loc_0040BAE2: mov [edx+00000004h], esi
  loc_0040BAE5: mov [edx+00000008h], eax
  loc_0040BAE8: mov eax, var_A0
  loc_0040BAEE: push eax
  loc_0040BAEF: mov [edx+0000000Ch], edi
  loc_0040BAF2: call ebx
  loc_0040BAF4: sub esp, 00000010h
  loc_0040BAF7: mov ecx, 00000003h
  loc_0040BAFC: mov edx, esp
  loc_0040BAFE: xor eax, eax
  loc_0040BB00: push 0000000Bh
  loc_0040BB02: mov [edx], ecx
  loc_0040BB04: mov [edx+00000004h], esi
  loc_0040BB07: mov [edx+00000008h], eax
  loc_0040BB0A: mov eax, var_A0
  loc_0040BB10: push eax
  loc_0040BB11: mov [edx+0000000Ch], edi
  loc_0040BB14: call ebx
  loc_0040BB16: sub esp, 00000010h
  loc_0040BB19: mov ecx, 00000003h
  loc_0040BB1E: mov edx, esp
  loc_0040BB20: xor eax, eax
  loc_0040BB22: sub esp, 00000010h
  loc_0040BB25: mov var_80, ecx
  loc_0040BB28: mov [edx], ecx
  loc_0040BB2A: mov ecx, esp
  loc_0040BB2C: mov [edx+00000004h], esi
  loc_0040BB2F: mov [edx+00000008h], eax
  loc_0040BB32: mov eax, var_7C
  loc_0040BB35: mov [edx+0000000Ch], edi
  loc_0040BB38: mov edx, var_80
  loc_0040BB3B: mov [ecx], edx
  loc_0040BB3D: mov [ecx+00000004h], eax
  loc_0040BB40: mov edx, var_74
  loc_0040BB43: mov eax, 000003E8h
  loc_0040BB48: mov [ecx+00000008h], eax
  loc_0040BB4B: mov eax, var_A0
  loc_0040BB51: push 00000001h
  loc_0040BB53: push 00000039h
  loc_0040BB55: push eax
  loc_0040BB56: mov [ecx+0000000Ch], edx
  loc_0040BB59: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BB5F: add esp, 0000001Ch
  loc_0040BB62: mov ecx, 00000008h
  loc_0040BB67: mov edx, esp
  loc_0040BB69: mov eax, 00403874h ; "Obs"
  loc_0040BB6E: push 00000000h
  loc_0040BB70: mov [edx], ecx
  loc_0040BB72: mov [edx+00000004h], esi
  loc_0040BB75: mov [edx+00000008h], eax
  loc_0040BB78: mov eax, var_A0
  loc_0040BB7E: push eax
  loc_0040BB7F: mov [edx+0000000Ch], edi
  loc_0040BB82: call ebx
  loc_0040BB84: sub esp, 00000010h
  loc_0040BB87: mov ecx, 00000003h
  loc_0040BB8C: mov edx, esp
  loc_0040BB8E: mov eax, 00000001h
  loc_0040BB93: push 0000000Bh
  loc_0040BB95: mov [edx], ecx
  loc_0040BB97: mov [edx+00000004h], esi
  loc_0040BB9A: mov [edx+00000008h], eax
  loc_0040BB9D: mov eax, var_A0
  loc_0040BBA3: push eax
  loc_0040BBA4: mov [edx+0000000Ch], edi
  loc_0040BBA7: call ebx
  loc_0040BBA9: sub esp, 00000010h
  loc_0040BBAC: mov ecx, 00000008h
  loc_0040BBB1: mov edx, esp
  loc_0040BBB3: mov eax, 00403880h ; " Yhat"
  loc_0040BBB8: push 00000000h
  loc_0040BBBA: mov [edx], ecx
  loc_0040BBBC: mov [edx+00000004h], esi
  loc_0040BBBF: mov [edx+00000008h], eax
  loc_0040BBC2: mov eax, var_A0
  loc_0040BBC8: push eax
  loc_0040BBC9: mov [edx+0000000Ch], edi
  loc_0040BBCC: call ebx
  loc_0040BBCE: sub esp, 00000010h
  loc_0040BBD1: mov ecx, 00000003h
  loc_0040BBD6: mov edx, esp
  loc_0040BBD8: mov eax, 00000001h
  loc_0040BBDD: sub esp, 00000010h
  loc_0040BBE0: mov var_80, ecx
  loc_0040BBE3: mov [edx], ecx
  loc_0040BBE5: mov ecx, esp
  loc_0040BBE7: push 00000001h
  loc_0040BBE9: push 00000039h
  loc_0040BBEB: mov [edx+00000004h], esi
  loc_0040BBEE: mov [edx+00000008h], eax
  loc_0040BBF1: mov eax, var_7C
  loc_0040BBF4: mov [edx+0000000Ch], edi
  loc_0040BBF7: mov edx, var_80
  loc_0040BBFA: mov [ecx], edx
  loc_0040BBFC: mov edx, var_74
  loc_0040BBFF: mov [ecx+00000004h], eax
  loc_0040BC02: mov eax, 00000708h
  loc_0040BC07: mov [ecx+00000008h], eax
  loc_0040BC0A: mov eax, var_A0
  loc_0040BC10: push eax
  loc_0040BC11: mov [ecx+0000000Ch], edx
  loc_0040BC14: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BC1A: add esp, 0000001Ch
  loc_0040BC1D: mov ecx, 00000003h
  loc_0040BC22: mov edx, esp
  loc_0040BC24: mov eax, 00000002h
  loc_0040BC29: push 0000000Bh
  loc_0040BC2B: mov [edx], ecx
  loc_0040BC2D: mov [edx+00000004h], esi
  loc_0040BC30: mov [edx+00000008h], eax
  loc_0040BC33: mov eax, var_A0
  loc_0040BC39: push eax
  loc_0040BC3A: mov [edx+0000000Ch], edi
  loc_0040BC3D: call ebx
  loc_0040BC3F: sub esp, 00000010h
  loc_0040BC42: mov ecx, 00000008h
  loc_0040BC47: mov edx, esp
  loc_0040BC49: mov eax, 00403890h ; " Resid "
  loc_0040BC4E: push 00000000h
  loc_0040BC50: mov [edx], ecx
  loc_0040BC52: mov [edx+00000004h], esi
  loc_0040BC55: mov [edx+00000008h], eax
  loc_0040BC58: mov eax, var_A0
  loc_0040BC5E: push eax
  loc_0040BC5F: mov [edx+0000000Ch], edi
  loc_0040BC62: call ebx
  loc_0040BC64: sub esp, 00000010h
  loc_0040BC67: mov ecx, 00000003h
  loc_0040BC6C: mov edx, esp
  loc_0040BC6E: mov eax, 00000002h
  loc_0040BC73: sub esp, 00000010h
  loc_0040BC76: mov var_80, ecx
  loc_0040BC79: mov [edx], ecx
  loc_0040BC7B: mov ecx, esp
  loc_0040BC7D: push 00000001h
  loc_0040BC7F: push 00000039h
  loc_0040BC81: mov [edx+00000004h], esi
  loc_0040BC84: mov [edx+00000008h], eax
  loc_0040BC87: mov eax, var_7C
  loc_0040BC8A: mov [edx+0000000Ch], edi
  loc_0040BC8D: mov edx, var_80
  loc_0040BC90: mov [ecx], edx
  loc_0040BC92: mov edx, var_74
  loc_0040BC95: mov [ecx+00000004h], eax
  loc_0040BC98: mov eax, 00000708h
  loc_0040BC9D: mov [ecx+00000008h], eax
  loc_0040BCA0: mov eax, var_A0
  loc_0040BCA6: push eax
  loc_0040BCA7: mov [ecx+0000000Ch], edx
  loc_0040BCAA: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BCB0: add esp, 0000001Ch
  loc_0040BCB3: mov eax, 00000003h
  loc_0040BCB8: mov edx, esp
  loc_0040BCBA: mov ecx, eax
  loc_0040BCBC: push 0000000Bh
  loc_0040BCBE: mov [edx], ecx
  loc_0040BCC0: mov [edx+00000004h], esi
  loc_0040BCC3: mov [edx+00000008h], eax
  loc_0040BCC6: mov eax, var_A0
  loc_0040BCCC: push eax
  loc_0040BCCD: mov [edx+0000000Ch], edi
  loc_0040BCD0: call ebx
  loc_0040BCD2: sub esp, 00000010h
  loc_0040BCD5: mov ecx, 00000003h
  loc_0040BCDA: mov edx, esp
  loc_0040BCDC: mov eax, ecx
  loc_0040BCDE: sub esp, 00000010h
  loc_0040BCE1: mov var_80, ecx
  loc_0040BCE4: mov [edx], ecx
  loc_0040BCE6: mov ecx, esp
  loc_0040BCE8: push 00000001h
  loc_0040BCEA: push 00000039h
  loc_0040BCEC: mov [edx+00000004h], esi
  loc_0040BCEF: mov [edx+00000008h], eax
  loc_0040BCF2: mov [edx+0000000Ch], edi
  loc_0040BCF5: mov edx, eax
  loc_0040BCF7: mov eax, var_7C
  loc_0040BCFA: mov [ecx], edx
  loc_0040BCFC: mov edx, var_74
  loc_0040BCFF: mov [ecx+00000004h], eax
  loc_0040BD02: mov eax, 000003E8h
  loc_0040BD07: mov [ecx+00000008h], eax
  loc_0040BD0A: mov eax, var_A0
  loc_0040BD10: push eax
  loc_0040BD11: mov [ecx+0000000Ch], edx
  loc_0040BD14: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BD1A: add esp, 0000001Ch
  loc_0040BD1D: mov ecx, 00000008h
  loc_0040BD22: mov edx, esp
  loc_0040BD24: mov eax, 004038A4h ; "  Hii  "
  loc_0040BD29: push 00000000h
  loc_0040BD2B: mov [edx], ecx
  loc_0040BD2D: mov [edx+00000004h], esi
  loc_0040BD30: mov [edx+00000008h], eax
  loc_0040BD33: mov [edx+0000000Ch], edi
  loc_0040BD36: mov eax, var_A0
  loc_0040BD3C: push eax
  loc_0040BD3D: call ebx
  loc_0040BD3F: sub esp, 00000010h
  loc_0040BD42: mov ecx, 00000003h
  loc_0040BD47: mov edx, esp
  loc_0040BD49: mov eax, 00000004h
  loc_0040BD4E: push 0000000Bh
  loc_0040BD50: mov [edx], ecx
  loc_0040BD52: mov [edx+00000004h], esi
  loc_0040BD55: mov [edx+00000008h], eax
  loc_0040BD58: mov eax, var_A0
  loc_0040BD5E: push eax
  loc_0040BD5F: mov [edx+0000000Ch], edi
  loc_0040BD62: call ebx
  loc_0040BD64: sub esp, 00000010h
  loc_0040BD67: mov ecx, 00000008h
  loc_0040BD6C: mov edx, esp
  loc_0040BD6E: mov eax, 004038B8h ; "Rstudent"
  loc_0040BD73: push 00000000h
  loc_0040BD75: mov [edx], ecx
  loc_0040BD77: mov [edx+00000004h], esi
  loc_0040BD7A: mov [edx+00000008h], eax
  loc_0040BD7D: mov eax, var_A0
  loc_0040BD83: push eax
  loc_0040BD84: mov [edx+0000000Ch], edi
  loc_0040BD87: call ebx
  loc_0040BD89: sub esp, 00000010h
  loc_0040BD8C: mov ecx, 00000003h
  loc_0040BD91: mov edx, esp
  loc_0040BD93: mov eax, 00000004h
  loc_0040BD98: sub esp, 00000010h
  loc_0040BD9B: mov var_80, ecx
  loc_0040BD9E: mov [edx], ecx
  loc_0040BDA0: mov ecx, esp
  loc_0040BDA2: push 00000001h
  loc_0040BDA4: push 00000039h
  loc_0040BDA6: mov [edx+00000004h], esi
  loc_0040BDA9: mov [edx+00000008h], eax
  loc_0040BDAC: mov eax, var_7C
  loc_0040BDAF: mov [edx+0000000Ch], edi
  loc_0040BDB2: mov edx, var_80
  loc_0040BDB5: mov [ecx], edx
  loc_0040BDB7: mov edx, var_74
  loc_0040BDBA: mov [ecx+00000004h], eax
  loc_0040BDBD: mov eax, 00000708h
  loc_0040BDC2: mov [ecx+00000008h], eax
  loc_0040BDC5: mov eax, var_A0
  loc_0040BDCB: push eax
  loc_0040BDCC: mov [ecx+0000000Ch], edx
  loc_0040BDCF: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BDD5: add esp, 0000001Ch
  loc_0040BDD8: mov ecx, 00000003h
  loc_0040BDDD: mov edx, esp
  loc_0040BDDF: mov eax, 00000005h
  loc_0040BDE4: push 0000000Bh
  loc_0040BDE6: mov [edx], ecx
  loc_0040BDE8: mov [edx+00000004h], esi
  loc_0040BDEB: mov [edx+00000008h], eax
  loc_0040BDEE: mov eax, var_A0
  loc_0040BDF4: push eax
  loc_0040BDF5: mov [edx+0000000Ch], edi
  loc_0040BDF8: call ebx
  loc_0040BDFA: sub esp, 00000010h
  loc_0040BDFD: mov ecx, 00000008h
  loc_0040BE02: mov edx, esp
  loc_0040BE04: mov eax, 004038D0h ; "Cooks D"
  loc_0040BE09: push 00000000h
  loc_0040BE0B: mov [edx], ecx
  loc_0040BE0D: mov [edx+00000004h], esi
  loc_0040BE10: mov [edx+00000008h], eax
  loc_0040BE13: mov eax, var_A0
  loc_0040BE19: push eax
  loc_0040BE1A: mov [edx+0000000Ch], edi
  loc_0040BE1D: call ebx
  loc_0040BE1F: sub esp, 00000010h
  loc_0040BE22: mov ecx, 00000003h
  loc_0040BE27: mov edx, esp
  loc_0040BE29: mov eax, 00000005h
  loc_0040BE2E: mov var_80, ecx
  loc_0040BE31: mov [edx], ecx
  loc_0040BE33: mov [edx+00000004h], esi
  loc_0040BE36: sub esp, 00000010h
  loc_0040BE39: mov ecx, esp
  loc_0040BE3B: mov [edx+00000008h], eax
  loc_0040BE3E: mov eax, var_7C
  loc_0040BE41: push 00000001h
  loc_0040BE43: push 00000039h
  loc_0040BE45: mov [edx+0000000Ch], edi
  loc_0040BE48: mov edx, var_80
  loc_0040BE4B: mov [ecx], edx
  loc_0040BE4D: mov edx, var_74
  loc_0040BE50: mov [ecx+00000004h], eax
  loc_0040BE53: mov eax, 00000708h
  loc_0040BE58: mov [ecx+00000008h], eax
  loc_0040BE5B: mov eax, var_A0
  loc_0040BE61: push eax
  loc_0040BE62: mov [ecx+0000000Ch], edx
  loc_0040BE65: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BE6B: add esp, 0000001Ch
  loc_0040BE6E: mov ecx, 00000003h
  loc_0040BE73: mov edx, esp
  loc_0040BE75: mov eax, 00000006h
  loc_0040BE7A: push 0000000Bh
  loc_0040BE7C: mov [edx], ecx
  loc_0040BE7E: mov [edx+00000004h], esi
  loc_0040BE81: mov [edx+00000008h], eax
  loc_0040BE84: mov eax, var_A0
  loc_0040BE8A: push eax
  loc_0040BE8B: mov [edx+0000000Ch], edi
  loc_0040BE8E: call ebx
  loc_0040BE90: sub esp, 00000010h
  loc_0040BE93: mov ecx, 00000008h
  loc_0040BE98: mov edx, esp
  loc_0040BE9A: mov eax, 004038E4h ; "Dffits"
  loc_0040BE9F: push 00000000h
  loc_0040BEA1: mov [edx], ecx
  loc_0040BEA3: mov [edx+00000004h], esi
  loc_0040BEA6: mov [edx+00000008h], eax
  loc_0040BEA9: mov eax, var_A0
  loc_0040BEAF: push eax
  loc_0040BEB0: mov [edx+0000000Ch], edi
  loc_0040BEB3: call ebx
  loc_0040BEB5: sub esp, 00000010h
  loc_0040BEB8: mov ecx, 00000003h
  loc_0040BEBD: mov edx, esp
  loc_0040BEBF: mov eax, 00000006h
  loc_0040BEC4: sub esp, 00000010h
  loc_0040BEC7: mov var_80, ecx
  loc_0040BECA: mov [edx], ecx
  loc_0040BECC: mov ecx, esp
  loc_0040BECE: push 00000001h
  loc_0040BED0: push 00000039h
  loc_0040BED2: mov [edx+00000004h], esi
  loc_0040BED5: mov [edx+00000008h], eax
  loc_0040BED8: mov eax, var_7C
  loc_0040BEDB: mov [edx+0000000Ch], edi
  loc_0040BEDE: mov edx, var_80
  loc_0040BEE1: mov [ecx], edx
  loc_0040BEE3: mov edx, var_74
  loc_0040BEE6: mov [ecx+00000004h], eax
  loc_0040BEE9: mov eax, 00000708h
  loc_0040BEEE: mov [ecx+00000008h], eax
  loc_0040BEF1: mov eax, var_A0
  loc_0040BEF7: push eax
  loc_0040BEF8: mov [ecx+0000000Ch], edx
  loc_0040BEFB: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BF01: add esp, 0000002Ch
  loc_0040BF04: lea ecx, var_A0
  loc_0040BF0A: push 00000000h
  loc_0040BF0C: push ecx
  loc_0040BF0D: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040BF13: mov eax, Me
  loc_0040BF16: push eax
  loc_0040BF17: mov edx, [eax]
  loc_0040BF19: call [edx+00000410h]
  loc_0040BF1F: push eax
  loc_0040BF20: lea eax, var_A4
  loc_0040BF26: push eax
  loc_0040BF27: call [0040105Ch] ; __vbaObjSet
  loc_0040BF2D: mov ecx, Me
  loc_0040BF30: mov dx, [ecx+00000038h]
  loc_0040BF34: add dx, 0001h
  loc_0040BF38: mov ecx, 00000003h
  loc_0040BF3D: jo 0040C479h
  loc_0040BF43: sub esp, 00000010h
  loc_0040BF46: movsx eax, dx
  loc_0040BF49: mov edx, esp
  loc_0040BF4B: push 00000004h
  loc_0040BF4D: mov [edx], ecx
  loc_0040BF4F: mov [edx+00000004h], esi
  loc_0040BF52: mov [edx+00000008h], eax
  loc_0040BF55: mov eax, var_A4
  loc_0040BF5B: push eax
  loc_0040BF5C: mov [edx+0000000Ch], edi
  loc_0040BF5F: call ebx
  loc_0040BF61: sub esp, 00000010h
  loc_0040BF64: mov ecx, 00000003h
  loc_0040BF69: mov edx, esp
  loc_0040BF6B: xor eax, eax
  loc_0040BF6D: push 0000000Ah
  loc_0040BF6F: mov [edx], ecx
  loc_0040BF71: mov [edx+00000004h], esi
  loc_0040BF74: mov [edx+00000008h], eax
  loc_0040BF77: mov eax, var_A4
  loc_0040BF7D: push eax
  loc_0040BF7E: mov [edx+0000000Ch], edi
  loc_0040BF81: call ebx
  loc_0040BF83: sub esp, 00000010h
  loc_0040BF86: mov ecx, 00000003h
  loc_0040BF8B: mov edx, esp
  loc_0040BF8D: xor eax, eax
  loc_0040BF8F: push 0000000Bh
  loc_0040BF91: mov [edx], ecx
  loc_0040BF93: mov [edx+00000004h], esi
  loc_0040BF96: mov [edx+00000008h], eax
  loc_0040BF99: mov eax, var_A4
  loc_0040BF9F: push eax
  loc_0040BFA0: mov [edx+0000000Ch], edi
  loc_0040BFA3: call ebx
  loc_0040BFA5: sub esp, 00000010h
  loc_0040BFA8: mov ecx, 00000003h
  loc_0040BFAD: mov edx, esp
  loc_0040BFAF: xor eax, eax
  loc_0040BFB1: sub esp, 00000010h
  loc_0040BFB4: mov var_80, ecx
  loc_0040BFB7: mov [edx], ecx
  loc_0040BFB9: mov ecx, esp
  loc_0040BFBB: push 00000001h
  loc_0040BFBD: push 00000039h
  loc_0040BFBF: mov [edx+00000004h], esi
  loc_0040BFC2: mov [edx+00000008h], eax
  loc_0040BFC5: mov eax, var_7C
  loc_0040BFC8: mov [edx+0000000Ch], edi
  loc_0040BFCB: mov edx, var_80
  loc_0040BFCE: mov [ecx], edx
  loc_0040BFD0: mov edx, var_74
  loc_0040BFD3: mov [ecx+00000004h], eax
  loc_0040BFD6: mov eax, 000003E8h
  loc_0040BFDB: mov [ecx+00000008h], eax
  loc_0040BFDE: mov eax, var_A4
  loc_0040BFE4: push eax
  loc_0040BFE5: mov [ecx+0000000Ch], edx
  loc_0040BFE8: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040BFEE: add esp, 0000001Ch
  loc_0040BFF1: mov ecx, 00000008h
  loc_0040BFF6: mov edx, esp
  loc_0040BFF8: mov eax, 00403874h ; "Obs"
  loc_0040BFFD: push 00000000h
  loc_0040BFFF: mov [edx], ecx
  loc_0040C001: mov [edx+00000004h], esi
  loc_0040C004: mov [edx+00000008h], eax
  loc_0040C007: mov eax, var_A4
  loc_0040C00D: push eax
  loc_0040C00E: mov [edx+0000000Ch], edi
  loc_0040C011: call ebx
  loc_0040C013: sub esp, 00000010h
  loc_0040C016: mov ecx, 00000003h
  loc_0040C01B: mov edx, esp
  loc_0040C01D: mov eax, 00000001h
  loc_0040C022: mov [edx], ecx
  loc_0040C024: mov [edx+00000004h], esi
  loc_0040C027: mov [edx+00000008h], eax
  loc_0040C02A: mov eax, var_A4
  loc_0040C030: push 0000000Bh
  loc_0040C032: push eax
  loc_0040C033: mov [edx+0000000Ch], edi
  loc_0040C036: call ebx
  loc_0040C038: sub esp, 00000010h
  loc_0040C03B: mov ecx, 00000008h
  loc_0040C040: mov edx, esp
  loc_0040C042: mov eax, 004038F8h ; "  Y  "
  loc_0040C047: push 00000000h
  loc_0040C049: mov [edx], ecx
  loc_0040C04B: mov [edx+00000004h], esi
  loc_0040C04E: mov [edx+00000008h], eax
  loc_0040C051: mov eax, var_A4
  loc_0040C057: push eax
  loc_0040C058: mov [edx+0000000Ch], edi
  loc_0040C05B: call ebx
  loc_0040C05D: sub esp, 00000010h
  loc_0040C060: mov ecx, 00000003h
  loc_0040C065: mov edx, esp
  loc_0040C067: mov eax, 00000001h
  loc_0040C06C: sub esp, 00000010h
  loc_0040C06F: mov var_80, ecx
  loc_0040C072: mov [edx], ecx
  loc_0040C074: mov ecx, esp
  loc_0040C076: push 00000001h
  loc_0040C078: push 00000039h
  loc_0040C07A: mov [edx+00000004h], esi
  loc_0040C07D: mov [edx+00000008h], eax
  loc_0040C080: mov eax, var_7C
  loc_0040C083: mov [edx+0000000Ch], edi
  loc_0040C086: mov edx, var_80
  loc_0040C089: mov [ecx], edx
  loc_0040C08B: mov edx, var_74
  loc_0040C08E: mov [ecx+00000004h], eax
  loc_0040C091: mov eax, 00000708h
  loc_0040C096: mov [ecx+00000008h], eax
  loc_0040C099: mov eax, var_A4
  loc_0040C09F: push eax
  loc_0040C0A0: mov [ecx+0000000Ch], edx
  loc_0040C0A3: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040C0A9: add esp, 0000001Ch
  loc_0040C0AC: mov ecx, 00000003h
  loc_0040C0B1: mov edx, esp
  loc_0040C0B3: mov eax, 00000002h
  loc_0040C0B8: push 0000000Bh
  loc_0040C0BA: mov [edx], ecx
  loc_0040C0BC: mov [edx+00000004h], esi
  loc_0040C0BF: mov [edx+00000008h], eax
  loc_0040C0C2: mov eax, var_A4
  loc_0040C0C8: push eax
  loc_0040C0C9: mov [edx+0000000Ch], edi
  loc_0040C0CC: call ebx
  loc_0040C0CE: sub esp, 00000010h
  loc_0040C0D1: mov ecx, 00000008h
  loc_0040C0D6: mov edx, esp
  loc_0040C0D8: mov eax, 00403808h ; "  X  "
  loc_0040C0DD: push 00000000h
  loc_0040C0DF: mov [edx], ecx
  loc_0040C0E1: mov [edx+00000004h], esi
  loc_0040C0E4: mov [edx+00000008h], eax
  loc_0040C0E7: mov eax, var_A4
  loc_0040C0ED: push eax
  loc_0040C0EE: mov [edx+0000000Ch], edi
  loc_0040C0F1: call ebx
  loc_0040C0F3: sub esp, 00000010h
  loc_0040C0F6: mov ecx, 00000003h
  loc_0040C0FB: mov ebx, esp
  loc_0040C0FD: mov eax, 00000002h
  loc_0040C102: sub esp, 00000010h
  loc_0040C105: mov edx, 00000708h
  loc_0040C10A: mov [ebx], ecx
  loc_0040C10C: mov ecx, esp
  loc_0040C10E: mov [ebx+00000004h], esi
  loc_0040C111: mov [ebx+00000008h], eax
  loc_0040C114: mov eax, 00000003h
  loc_0040C119: mov [ebx+0000000Ch], edi
  loc_0040C11C: mov [ecx], eax
  loc_0040C11E: mov eax, var_7C
  loc_0040C121: mov [ecx+00000004h], eax
  loc_0040C124: mov [ecx+00000008h], edx
  loc_0040C127: mov eax, var_A4
  loc_0040C12D: mov edx, var_74
  loc_0040C130: push 00000001h
  loc_0040C132: push 00000039h
  loc_0040C134: push eax
  loc_0040C135: mov [ecx+0000000Ch], edx
  loc_0040C138: call [004010C8h] ; __vbaLateIdCallSt
  loc_0040C13E: mov esi, Me
  loc_0040C141: push 00000000h
  loc_0040C143: mov edi, [004010B8h] ; __vbaRedim
  loc_0040C149: movsx ecx, [esi+00000038h]
  loc_0040C14D: push ecx
  loc_0040C14E: push 00000000h
  loc_0040C150: push 00000002h
  loc_0040C152: lea ebx, [esi+0000003Ch]
  loc_0040C155: push 00000002h
  loc_0040C157: push 00000004h
  loc_0040C159: push ebx
  loc_0040C15A: push 00000004h
  loc_0040C15C: push 00000080h
  loc_0040C161: call edi
  loc_0040C163: movsx edx, [esi+00000038h]
  loc_0040C167: add esp, 00000050h
  loc_0040C16A: lea eax, var_20
  loc_0040C16D: push 00000000h
  loc_0040C16F: push edx
  loc_0040C170: push 00000001h
  loc_0040C172: push 0000000Ch
  loc_0040C174: push eax
  loc_0040C175: push 00000010h
  loc_0040C177: push 00000880h
  loc_0040C17C: call edi
  loc_0040C17E: movsx ecx, [esi+00000038h]
  loc_0040C182: push 00000000h
  loc_0040C184: push ecx
  loc_0040C185: push 00000001h
  loc_0040C187: lea edx, var_18
  loc_0040C18A: push 0000000Ch
  loc_0040C18C: push edx
  loc_0040C18D: push 00000010h
  loc_0040C18F: push 00000880h
  loc_0040C194: call edi
  loc_0040C196: movsx eax, [esi+00000038h]
  loc_0040C19A: push 00000000h
  loc_0040C19C: push eax
  loc_0040C19D: push 00000001h
  loc_0040C19F: lea ecx, [esi+0000007Ch]
  loc_0040C1A2: push 00000004h
  loc_0040C1A4: push ecx
  loc_0040C1A5: push 00000004h
  loc_0040C1A7: push 00000080h
  loc_0040C1AC: call edi
  loc_0040C1AE: movsx edx, [esi+00000038h]
  loc_0040C1B2: add esp, 00000054h
  loc_0040C1B5: lea eax, [esi+00000084h]
  loc_0040C1BB: push 00000000h
  loc_0040C1BD: push edx
  loc_0040C1BE: push 00000001h
  loc_0040C1C0: push 00000004h
  loc_0040C1C2: push eax
  loc_0040C1C3: push 00000004h
  loc_0040C1C5: push 00000080h
  loc_0040C1CA: call edi
  loc_0040C1CC: movsx ecx, [esi+00000038h]
  loc_0040C1D0: push 00000000h
  loc_0040C1D2: push ecx
  loc_0040C1D3: push 00000001h
  loc_0040C1D5: lea edx, [esi+00000080h]
  loc_0040C1DB: push 00000004h
  loc_0040C1DD: push edx
  loc_0040C1DE: push 00000004h
  loc_0040C1E0: push 00000080h
  loc_0040C1E5: call edi
  loc_0040C1E7: movsx eax, [esi+00000038h]
  loc_0040C1EB: push 00000000h
  loc_0040C1ED: push eax
  loc_0040C1EE: push 00000001h
  loc_0040C1F0: lea ecx, [esi+00000088h]
  loc_0040C1F6: push 00000004h
  loc_0040C1F8: push ecx
  loc_0040C1F9: push 00000004h
  loc_0040C1FB: push 00000080h
  loc_0040C200: call edi
  loc_0040C202: mov dx, [esi+00000038h]
  loc_0040C206: add esp, 00000054h
  loc_0040C209: mov var_AC, dx
  loc_0040C210: mov edi, 00000001h
  loc_0040C215: cmp di, var_AC
  loc_0040C21C: jg 0040C3F7h
  loc_0040C222: mov eax, [ebx]
  loc_0040C224: test eax, eax
  loc_0040C226: jz 0040C285h
  loc_0040C228: cmp [eax], 0002h
  loc_0040C22C: jnz 0040C285h
  loc_0040C22E: mov ecx, [eax+0000001Ch]
  loc_0040C231: movsx edx, di
  loc_0040C234: sub edx, ecx
  loc_0040C236: mov ecx, [eax+00000018h]
  loc_0040C239: cmp edx, ecx
  loc_0040C23B: mov var_9C, edx
  loc_0040C241: jb 0040C24Fh
  loc_0040C243: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040C249: mov edx, var_9C
  loc_0040C24F: mov eax, [ebx]
  loc_0040C251: mov ecx, 00000002h
  loc_0040C256: sub ecx, [eax+00000014h]
  loc_0040C259: cmp ecx, [eax+00000010h]
  loc_0040C25C: mov var_98, ecx
  loc_0040C262: jb 0040C276h
  loc_0040C264: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040C26A: mov edx, var_9C
  loc_0040C270: mov ecx, var_98
  loc_0040C276: mov eax, [ebx]
  loc_0040C278: mov eax, [eax+00000018h]
  loc_0040C27B: imul eax, ecx
  loc_0040C27E: add eax, edx
  loc_0040C280: shl eax, 02h
  loc_0040C283: jmp 0040C28Bh
  loc_0040C285: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040C28B: mov edx, [ebx]
  loc_0040C28D: movsx ecx, di
  loc_0040C290: mov var_B4, ecx
  loc_0040C296: mov ecx, [edx+0000000Ch]
  loc_0040C299: fild real4 ptr var_B4
  loc_0040C29F: lea edx, var_30
  loc_0040C2A2: push edx
  loc_0040C2A3: fstp real4 ptr [ecx+eax]
  loc_0040C2A6: mov var_28, 80020004h
  loc_0040C2AD: mov var_30, 0000000Ah
  loc_0040C2B4: call [00401058h] ; rtcRandomize
  loc_0040C2BA: lea ecx, var_30
  loc_0040C2BD: call [00401014h] ; __vbaFreeVar
  loc_0040C2C3: lea eax, var_30
  loc_0040C2C6: mov var_28, 00001BD3h
  loc_0040C2CD: push eax
  loc_0040C2CE: mov var_30, 00000002h
  loc_0040C2D5: call [00401050h] ; rtcRandomNext
  loc_0040C2DB: mov cx, di
  loc_0040C2DE: push 00000002h
  loc_0040C2E0: imul cx, cx, 0004h
  loc_0040C2E4: fstp real4 ptr var_94
  loc_0040C2EA: jo 0040C479h
  loc_0040C2F0: sub cx, 0002h
  loc_0040C2F4: mov var_40, 00000004h
  loc_0040C2FB: jo 0040C479h
  loc_0040C301: movsx edx, cx
  loc_0040C304: mov var_B8, edx
  loc_0040C30A: lea ecx, var_50
  loc_0040C30D: fild real4 ptr var_B8
  loc_0040C313: fstp real4 ptr var_BC
  loc_0040C319: fld real4 ptr var_94
  loc_0040C31F: fmul st0, real4 ptr [004011F0h]
  loc_0040C325: fadd st0, real4 ptr var_BC
  loc_0040C32B: fstp real4 ptr var_38
  loc_0040C32E: fnstsw ax
  loc_0040C330: test al, 0Dh
  loc_0040C332: jnz 0040C474h
  loc_0040C338: lea eax, var_40
  loc_0040C33B: push eax
  loc_0040C33C: push ecx
  loc_0040C33D: call [004010E8h] ; rtcRound
  loc_0040C343: mov eax, [ebx]
  loc_0040C345: test eax, eax
  loc_0040C347: jz 0040C3A7h
  loc_0040C349: cmp [eax], 0002h
  loc_0040C34D: jnz 0040C3A7h
  loc_0040C34F: mov ecx, var_B4
  loc_0040C355: mov edx, [eax+0000001Ch]
  loc_0040C358: sub ecx, edx
  loc_0040C35A: mov edx, [eax+00000018h]
  loc_0040C35D: cmp ecx, edx
  loc_0040C35F: mov var_9C, ecx
  loc_0040C365: jb 0040C36Dh
  loc_0040C367: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040C36D: mov eax, [ebx]
  loc_0040C36F: mov ecx, 00000001h
  loc_0040C374: mov edx, [eax+00000014h]
  loc_0040C377: sub ecx, edx
  loc_0040C379: mov edx, [eax+00000010h]
  loc_0040C37C: cmp ecx, edx
  loc_0040C37E: mov var_98, ecx
  loc_0040C384: jb 0040C392h
  loc_0040C386: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040C38C: mov ecx, var_98
  loc_0040C392: mov edx, [ebx]
  loc_0040C394: mov eax, [edx+00000018h]
  loc_0040C397: mov edx, var_9C
  loc_0040C39D: imul eax, ecx
  loc_0040C3A0: add eax, edx
  loc_0040C3A2: shl eax, 02h
  loc_0040C3A5: jmp 0040C3ADh
  loc_0040C3A7: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040C3AD: mov var_C0, eax
  loc_0040C3B3: lea eax, var_50
  loc_0040C3B6: push eax
  loc_0040C3B7: call [004010B0h] ; __vbaR4Var
  loc_0040C3BD: mov ecx, [ebx]
  loc_0040C3BF: mov eax, var_C0
  loc_0040C3C5: mov edx, [ecx+0000000Ch]
  loc_0040C3C8: lea ecx, var_50
  loc_0040C3CB: push ecx
  loc_0040C3CC: fstp real4 ptr [edx+eax]
  loc_0040C3CF: lea edx, var_40
  loc_0040C3D2: lea eax, var_30
  loc_0040C3D5: push edx
  loc_0040C3D6: push eax
  loc_0040C3D7: push 00000003h
  loc_0040C3D9: call [00401024h] ; __vbaFreeVarList
  loc_0040C3DF: mov eax, 00000001h
  loc_0040C3E4: add esp, 00000010h
  loc_0040C3E7: add ax, di
  loc_0040C3EA: jo 0040C479h
  loc_0040C3F0: mov edi, eax
  loc_0040C3F2: jmp 0040C215h
  loc_0040C3F7: lea ecx, var_A4
  loc_0040C3FD: push 00000000h
  loc_0040C3FF: push ecx
  loc_0040C400: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040C406: mov edx, [esi]
  loc_0040C408: push esi
  loc_0040C409: call [edx+00000728h]
  loc_0040C40F: fwait
  loc_0040C410: push 0040C45Fh
  loc_0040C415: jmp 0040C42Fh
  loc_0040C417: lea eax, var_50
  loc_0040C41A: lea ecx, var_40
  loc_0040C41D: push eax
  loc_0040C41E: lea edx, var_30
  loc_0040C421: push ecx
  loc_0040C422: push edx
  loc_0040C423: push 00000003h
  loc_0040C425: call [00401024h] ; __vbaFreeVarList
  loc_0040C42B: add esp, 00000010h
  loc_0040C42E: ret
  loc_0040C42F: lea eax, var_A4
  loc_0040C435: lea ecx, var_A0
  loc_0040C43B: push eax
  loc_0040C43C: push ecx
  loc_0040C43D: push 00000002h
  loc_0040C43F: call [0040102Ch] ; __vbaFreeObjList
  loc_0040C445: mov esi, [00401048h] ; __vbaAryDestruct
  loc_0040C44B: add esp, 0000000Ch
  loc_0040C44E: lea edx, var_18
  loc_0040C451: push edx
  loc_0040C452: push 00000000h
  loc_0040C454: call __vbaAryDestruct
  loc_0040C456: lea eax, var_20
  loc_0040C459: push eax
  loc_0040C45A: push 00000000h
  loc_0040C45C: call __vbaAryDestruct
  loc_0040C45E: ret
  loc_0040C45F: mov ecx, var_10
  loc_0040C462: pop edi
  loc_0040C463: pop esi
  loc_0040C464: xor eax, eax
  loc_0040C466: mov fs:[00000000h], ecx
  loc_0040C46D: pop ebx
  loc_0040C46E: mov esp, ebp
  loc_0040C470: pop ebp
  loc_0040C471: retn 0004h
End Sub

Private Sub Proc_3_24_40CA80
  loc_0040CA80: push ebp
  loc_0040CA81: mov ebp, esp
  loc_0040CA83: sub esp, 00000008h
  loc_0040CA86: push 00401546h ; __vbaExceptHandler
  loc_0040CA8B: mov eax, fs:[00000000h]
  loc_0040CA91: push eax
  loc_0040CA92: mov fs:[00000000h], esp
  loc_0040CA99: sub esp, 0000003Ch
  loc_0040CA9C: push ebx
  loc_0040CA9D: push esi
  loc_0040CA9E: push edi
  loc_0040CA9F: mov var_8, esp
  loc_0040CAA2: mov var_4, 00401240h
  loc_0040CAA9: mov ebx, Me
  loc_0040CAAC: xor eax, eax
  loc_0040CAAE: push eax
  loc_0040CAAF: lea edx, var_18
  loc_0040CAB2: movsx ecx, [ebx+00000038h]
  loc_0040CAB6: push ecx
  loc_0040CAB7: push eax
  loc_0040CAB8: push 00000002h
  loc_0040CABA: push 00000002h
  loc_0040CABC: push 00000005h
  loc_0040CABE: push edx
  loc_0040CABF: push 00000008h
  loc_0040CAC1: push 00000080h
  loc_0040CAC6: mov var_18, eax
  loc_0040CAC9: mov var_1C, eax
  loc_0040CACC: mov var_20, eax
  loc_0040CACF: mov var_24, eax
  loc_0040CAD2: mov var_28, eax
  loc_0040CAD5: call [004010B8h] ; __vbaRedim
  loc_0040CADB: mov ax, [ebx+00000038h]
  loc_0040CADF: mov edi, 00000001h
  loc_0040CAE4: add esp, 00000024h
  loc_0040CAE7: mov var_3C, ax
  loc_0040CAEB: mov var_14, edi
  loc_0040CAEE: cmp di, var_3C
  loc_0040CAF2: jg 0040CC68h
  loc_0040CAF8: mov eax, [ebx+00000074h]
  loc_0040CAFB: test eax, eax
  loc_0040CAFD: jz 0040CB26h
  loc_0040CAFF: cmp [eax], 0001h
  loc_0040CB03: jnz 0040CB26h
  loc_0040CB05: mov edx, [eax+00000014h]
  loc_0040CB08: mov ecx, [eax+00000010h]
  loc_0040CB0B: movsx esi, di
  loc_0040CB0E: sub esi, edx
  loc_0040CB10: cmp esi, ecx
  loc_0040CB12: jb 0040CB1Ah
  loc_0040CB14: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CB1A: lea ecx, [esi*4]
  loc_0040CB21: mov var_48, ecx
  loc_0040CB24: jmp 0040CB2Fh
  loc_0040CB26: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CB2C: mov var_48, eax
  loc_0040CB2F: mov eax, var_18
  loc_0040CB32: test eax, eax
  loc_0040CB34: jz 0040CB81h
  loc_0040CB36: cmp [eax], 0002h
  loc_0040CB3A: jnz 0040CB81h
  loc_0040CB3C: mov edx, [eax+0000001Ch]
  loc_0040CB3F: mov ecx, [eax+00000018h]
  loc_0040CB42: movsx ebx, di
  loc_0040CB45: sub ebx, edx
  loc_0040CB47: cmp ebx, ecx
  loc_0040CB49: jb 0040CB54h
  loc_0040CB4B: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CB51: mov eax, var_18
  loc_0040CB54: mov edx, [eax+00000014h]
  loc_0040CB57: mov ecx, [eax+00000010h]
  loc_0040CB5A: mov edi, 00000002h
  loc_0040CB5F: sub edi, edx
  loc_0040CB61: cmp edi, ecx
  loc_0040CB63: jb 0040CB6Eh
  loc_0040CB65: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CB6B: mov eax, var_18
  loc_0040CB6E: mov esi, [eax+00000018h]
  loc_0040CB71: imul esi, edi
  loc_0040CB74: mov edi, var_14
  loc_0040CB77: add esi, ebx
  loc_0040CB79: mov ebx, Me
  loc_0040CB7C: shl esi, 03h
  loc_0040CB7F: jmp 0040CB89h
  loc_0040CB81: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CB87: mov esi, eax
  loc_0040CB89: mov edx, [ebx+00000074h]
  loc_0040CB8C: mov ecx, var_48
  loc_0040CB8F: mov eax, [edx+0000000Ch]
  loc_0040CB92: fld real4 ptr [eax+ecx]
  loc_0040CB95: call [00401010h] ; __vbaFpCDblR4
  loc_0040CB9B: mov edx, var_18
  loc_0040CB9E: mov eax, [edx+0000000Ch]
  loc_0040CBA1: fstp real8 ptr [eax+esi]
  loc_0040CBA4: mov eax, [ebx+00000078h]
  loc_0040CBA7: test eax, eax
  loc_0040CBA9: jz 0040CBD2h
  loc_0040CBAB: cmp [eax], 0001h
  loc_0040CBAF: jnz 0040CBD2h
  loc_0040CBB1: mov edx, [eax+00000014h]
  loc_0040CBB4: mov ecx, [eax+00000010h]
  loc_0040CBB7: movsx esi, di
  loc_0040CBBA: sub esi, edx
  loc_0040CBBC: cmp esi, ecx
  loc_0040CBBE: jb 0040CBC6h
  loc_0040CBC0: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CBC6: lea ecx, [esi*4]
  loc_0040CBCD: mov var_4C, ecx
  loc_0040CBD0: jmp 0040CBDBh
  loc_0040CBD2: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CBD8: mov var_4C, eax
  loc_0040CBDB: mov eax, var_18
  loc_0040CBDE: test eax, eax
  loc_0040CBE0: jz 0040CC2Dh
  loc_0040CBE2: cmp [eax], 0002h
  loc_0040CBE6: jnz 0040CC2Dh
  loc_0040CBE8: mov edx, [eax+0000001Ch]
  loc_0040CBEB: mov ecx, [eax+00000018h]
  loc_0040CBEE: movsx ebx, di
  loc_0040CBF1: sub ebx, edx
  loc_0040CBF3: cmp ebx, ecx
  loc_0040CBF5: jb 0040CC00h
  loc_0040CBF7: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CBFD: mov eax, var_18
  loc_0040CC00: mov edx, [eax+00000014h]
  loc_0040CC03: mov ecx, [eax+00000010h]
  loc_0040CC06: mov edi, 00000001h
  loc_0040CC0B: sub edi, edx
  loc_0040CC0D: cmp edi, ecx
  loc_0040CC0F: jb 0040CC1Ah
  loc_0040CC11: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CC17: mov eax, var_18
  loc_0040CC1A: mov esi, [eax+00000018h]
  loc_0040CC1D: imul esi, edi
  loc_0040CC20: mov edi, var_14
  loc_0040CC23: add esi, ebx
  loc_0040CC25: mov ebx, Me
  loc_0040CC28: shl esi, 03h
  loc_0040CC2B: jmp 0040CC35h
  loc_0040CC2D: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CC33: mov esi, eax
  loc_0040CC35: mov edx, [ebx+00000078h]
  loc_0040CC38: mov ecx, var_4C
  loc_0040CC3B: mov eax, [edx+0000000Ch]
  loc_0040CC3E: fld real4 ptr [eax+ecx]
  loc_0040CC41: call [00401010h] ; __vbaFpCDblR4
  loc_0040CC47: mov edx, var_18
  loc_0040CC4A: mov eax, [edx+0000000Ch]
  loc_0040CC4D: fstp real8 ptr [eax+esi]
  loc_0040CC50: mov eax, 00000001h
  loc_0040CC55: add ax, di
  loc_0040CC58: jo 0040CD03h
  loc_0040CC5E: mov var_14, eax
  loc_0040CC61: mov edi, eax
  loc_0040CC63: jmp 0040CAEEh
  loc_0040CC68: mov ecx, [ebx]
  loc_0040CC6A: push ebx
  loc_0040CC6B: call [ecx+00000318h]
  loc_0040CC71: mov esi, [0040105Ch] ; __vbaObjSet
  loc_0040CC77: lea edx, var_20
  loc_0040CC7A: push eax
  loc_0040CC7B: push edx
  loc_0040CC7C: call __vbaObjSet
  loc_0040CC7E: mov eax, var_20
  loc_0040CC81: mov var_28, 00000002h
  loc_0040CC88: push eax
  loc_0040CC89: lea eax, var_1C
  loc_0040CC8C: push eax
  loc_0040CC8D: mov var_24, 00000001h
  loc_0040CC94: mov var_20, 00000000h
  loc_0040CC9B: call __vbaObjSet
  loc_0040CC9D: lea ecx, var_28
  loc_0040CCA0: lea edx, var_24
  loc_0040CCA3: push ecx
  loc_0040CCA4: lea eax, var_18
  loc_0040CCA7: push edx
  loc_0040CCA8: lea ecx, var_1C
  loc_0040CCAB: push eax
  loc_0040CCAC: push ecx
  loc_0040CCAD: call 00409030h
  loc_0040CCB2: lea edx, var_20
  loc_0040CCB5: lea eax, var_1C
  loc_0040CCB8: push edx
  loc_0040CCB9: push eax
  loc_0040CCBA: push 00000002h
  loc_0040CCBC: call [0040102Ch] ; __vbaFreeObjList
  loc_0040CCC2: add esp, 0000000Ch
  loc_0040CCC5: fwait
  loc_0040CCC6: push 0040CCEEh
  loc_0040CCCB: jmp 0040CCE1h
  loc_0040CCCD: lea ecx, var_20
  loc_0040CCD0: lea edx, var_1C
  loc_0040CCD3: push ecx
  loc_0040CCD4: push edx
  loc_0040CCD5: push 00000002h
  loc_0040CCD7: call [0040102Ch] ; __vbaFreeObjList
  loc_0040CCDD: add esp, 0000000Ch
  loc_0040CCE0: ret
  loc_0040CCE1: lea eax, var_18
  loc_0040CCE4: push eax
  loc_0040CCE5: push 00000000h
  loc_0040CCE7: call [00401048h] ; __vbaAryDestruct
  loc_0040CCED: ret
  loc_0040CCEE: mov ecx, var_10
  loc_0040CCF1: pop edi
  loc_0040CCF2: pop esi
  loc_0040CCF3: xor eax, eax
  loc_0040CCF5: mov fs:[00000000h], ecx
  loc_0040CCFC: pop ebx
  loc_0040CCFD: mov esp, ebp
  loc_0040CCFF: pop ebp
  loc_0040CD00: retn 0004h
End Sub

Private Sub Proc_3_25_40CD10
  loc_0040CD10: push ebp
  loc_0040CD11: mov ebp, esp
  loc_0040CD13: sub esp, 00000008h
  loc_0040CD16: push 00401546h ; __vbaExceptHandler
  loc_0040CD1B: mov eax, fs:[00000000h]
  loc_0040CD21: push eax
  loc_0040CD22: mov fs:[00000000h], esp
  loc_0040CD29: sub esp, 00000060h
  loc_0040CD2C: push ebx
  loc_0040CD2D: push esi
  loc_0040CD2E: push edi
  loc_0040CD2F: mov var_8, esp
  loc_0040CD32: mov var_4, 00401250h
  loc_0040CD39: mov edi, Me
  loc_0040CD3C: mov esi, [0040116Ch] ; __vbaLateIdSt
  loc_0040CD42: mov ebx, [0040114Ch] ; __vbaAryLock
  loc_0040CD48: xor eax, eax
  loc_0040CD4A: mov var_18, eax
  loc_0040CD4D: mov var_28, eax
  loc_0040CD50: mov var_38, eax
  loc_0040CD53: mov var_48, eax
  loc_0040CD56: mov var_60, eax
  loc_0040CD59: mov ax, [edi+00000038h]
  loc_0040CD5D: mov var_68, ax
  loc_0040CD61: mov var_14, 00000001h
  loc_0040CD68: mov cx, var_68
  loc_0040CD6C: cmp var_14, cx
  loc_0040CD70: jg 0040D36Dh
  loc_0040CD76: mov edx, [edi]
  loc_0040CD78: push edi
  loc_0040CD79: call [edx+0000040Ch]
  loc_0040CD7F: push eax
  loc_0040CD80: lea eax, var_60
  loc_0040CD83: push eax
  loc_0040CD84: call [0040105Ch] ; __vbaObjSet
  loc_0040CD8A: movsx eax, var_14
  loc_0040CD8E: sub esp, 00000010h
  loc_0040CD91: mov ecx, 00000003h
  loc_0040CD96: mov edx, esp
  loc_0040CD98: mov var_48, ecx
  loc_0040CD9B: mov var_70, eax
  loc_0040CD9E: mov var_40, eax
  loc_0040CDA1: mov [edx], ecx
  loc_0040CDA3: mov ecx, var_44
  loc_0040CDA6: push 0000000Ah
  loc_0040CDA8: mov [edx+00000004h], ecx
  loc_0040CDAB: mov ecx, var_60
  loc_0040CDAE: push ecx
  loc_0040CDAF: mov [edx+00000008h], eax
  loc_0040CDB2: mov eax, var_3C
  loc_0040CDB5: mov [edx+0000000Ch], eax
  loc_0040CDB8: call __vbaLateIdSt
  loc_0040CDBA: sub esp, 00000010h
  loc_0040CDBD: mov ecx, 00000003h
  loc_0040CDC2: mov edx, esp
  loc_0040CDC4: mov var_48, ecx
  loc_0040CDC7: xor eax, eax
  loc_0040CDC9: push 0000000Bh
  loc_0040CDCB: mov [edx], ecx
  loc_0040CDCD: mov ecx, var_44
  loc_0040CDD0: mov var_40, eax
  loc_0040CDD3: mov [edx+00000004h], ecx
  loc_0040CDD6: mov ecx, var_60
  loc_0040CDD9: push ecx
  loc_0040CDDA: mov [edx+00000008h], eax
  loc_0040CDDD: mov eax, var_3C
  loc_0040CDE0: mov [edx+0000000Ch], eax
  loc_0040CDE3: call __vbaLateIdSt
  loc_0040CDE5: mov edx, var_14
  loc_0040CDE8: push edx
  loc_0040CDE9: call [00401000h] ; __vbaStrI2
  loc_0040CDEF: sub esp, 00000010h
  loc_0040CDF2: mov ecx, 00000008h
  loc_0040CDF7: mov edx, esp
  loc_0040CDF9: mov var_28, ecx
  loc_0040CDFC: mov var_20, eax
  loc_0040CDFF: push 00000000h
  loc_0040CE01: mov [edx], ecx
  loc_0040CE03: mov ecx, var_24
  loc_0040CE06: mov [edx+00000004h], ecx
  loc_0040CE09: mov ecx, var_60
  loc_0040CE0C: push ecx
  loc_0040CE0D: mov [edx+00000008h], eax
  loc_0040CE10: mov eax, var_1C
  loc_0040CE13: mov [edx+0000000Ch], eax
  loc_0040CE16: call __vbaLateIdSt
  loc_0040CE18: lea ecx, var_28
  loc_0040CE1B: call [00401014h] ; __vbaFreeVar
  loc_0040CE21: sub esp, 00000010h
  loc_0040CE24: mov ecx, 00000003h
  loc_0040CE29: mov edx, esp
  loc_0040CE2B: mov var_48, ecx
  loc_0040CE2E: mov eax, 00000001h
  loc_0040CE33: push 0000000Bh
  loc_0040CE35: mov [edx], ecx
  loc_0040CE37: mov ecx, var_44
  loc_0040CE3A: mov var_40, eax
  loc_0040CE3D: mov [edx+00000004h], ecx
  loc_0040CE40: mov ecx, var_60
  loc_0040CE43: push ecx
  loc_0040CE44: mov [edx+00000008h], eax
  loc_0040CE47: mov eax, var_3C
  loc_0040CE4A: mov [edx+0000000Ch], eax
  loc_0040CE4D: call __vbaLateIdSt
  loc_0040CE4F: mov edx, [edi+00000074h]
  loc_0040CE52: lea eax, var_18
  loc_0040CE55: push edx
  loc_0040CE56: push eax
  loc_0040CE57: call ebx
  loc_0040CE59: mov ecx, var_18
  loc_0040CE5C: test ecx, ecx
  loc_0040CE5E: jz 0040CE89h
  loc_0040CE60: cmp [ecx], 0001h
  loc_0040CE64: jnz 0040CE89h
  loc_0040CE66: mov eax, var_70
  loc_0040CE69: mov edx, [ecx+00000014h]
  loc_0040CE6C: sub eax, edx
  loc_0040CE6E: mov edx, [ecx+00000010h]
  loc_0040CE71: cmp eax, edx
  loc_0040CE73: mov var_5C, eax
  loc_0040CE76: jb 0040CE84h
  loc_0040CE78: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CE7E: mov ecx, var_18
  loc_0040CE81: mov eax, var_5C
  loc_0040CE84: shl eax, 02h
  loc_0040CE87: jmp 0040CE92h
  loc_0040CE89: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CE8F: mov ecx, var_18
  loc_0040CE92: mov ecx, [ecx+0000000Ch]
  loc_0040CE95: lea edx, var_48
  loc_0040CE98: add ecx, eax
  loc_0040CE9A: push 00000002h
  loc_0040CE9C: lea eax, var_28
  loc_0040CE9F: push edx
  loc_0040CEA0: push eax
  loc_0040CEA1: mov var_40, ecx
  loc_0040CEA4: mov var_48, 00004004h
  loc_0040CEAB: call [004010E8h] ; rtcRound
  loc_0040CEB1: lea ecx, var_18
  loc_0040CEB4: push ecx
  loc_0040CEB5: call [00401174h] ; __vbaAryUnlock
  loc_0040CEBB: lea edx, var_28
  loc_0040CEBE: push edx
  loc_0040CEBF: call [00401018h] ; __vbaStrVarMove
  loc_0040CEC5: sub esp, 00000010h
  loc_0040CEC8: mov ecx, 00000008h
  loc_0040CECD: mov edx, esp
  loc_0040CECF: mov var_38, ecx
  loc_0040CED2: mov var_30, eax
  loc_0040CED5: push 00000000h
  loc_0040CED7: mov [edx], ecx
  loc_0040CED9: mov ecx, var_34
  loc_0040CEDC: mov [edx+00000004h], ecx
  loc_0040CEDF: mov ecx, var_60
  loc_0040CEE2: push ecx
  loc_0040CEE3: mov [edx+00000008h], eax
  loc_0040CEE6: mov eax, var_2C
  loc_0040CEE9: mov [edx+0000000Ch], eax
  loc_0040CEEC: call __vbaLateIdSt
  loc_0040CEEE: lea edx, var_38
  loc_0040CEF1: lea eax, var_28
  loc_0040CEF4: push edx
  loc_0040CEF5: push eax
  loc_0040CEF6: push 00000002h
  loc_0040CEF8: call [00401024h] ; __vbaFreeVarList
  loc_0040CEFE: mov ecx, 00000003h
  loc_0040CF03: mov eax, 00000002h
  loc_0040CF08: push ecx
  loc_0040CF09: mov var_48, ecx
  loc_0040CF0C: mov edx, esp
  loc_0040CF0E: mov var_40, eax
  loc_0040CF11: push 0000000Bh
  loc_0040CF13: mov [edx], ecx
  loc_0040CF15: mov ecx, var_44
  loc_0040CF18: mov [edx+00000004h], ecx
  loc_0040CF1B: mov ecx, var_60
  loc_0040CF1E: push ecx
  loc_0040CF1F: mov [edx+00000008h], eax
  loc_0040CF22: mov eax, var_3C
  loc_0040CF25: mov [edx+0000000Ch], eax
  loc_0040CF28: call __vbaLateIdSt
  loc_0040CF2A: mov edx, [edi+00000078h]
  loc_0040CF2D: lea eax, var_18
  loc_0040CF30: push edx
  loc_0040CF31: push eax
  loc_0040CF32: call ebx
  loc_0040CF34: mov ecx, var_18
  loc_0040CF37: test ecx, ecx
  loc_0040CF39: jz 0040CF64h
  loc_0040CF3B: cmp [ecx], 0001h
  loc_0040CF3F: jnz 0040CF64h
  loc_0040CF41: mov eax, var_70
  loc_0040CF44: mov edx, [ecx+00000014h]
  loc_0040CF47: sub eax, edx
  loc_0040CF49: mov edx, [ecx+00000010h]
  loc_0040CF4C: cmp eax, edx
  loc_0040CF4E: mov var_5C, eax
  loc_0040CF51: jb 0040CF5Fh
  loc_0040CF53: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CF59: mov ecx, var_18
  loc_0040CF5C: mov eax, var_5C
  loc_0040CF5F: shl eax, 02h
  loc_0040CF62: jmp 0040CF6Dh
  loc_0040CF64: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040CF6A: mov ecx, var_18
  loc_0040CF6D: mov ecx, [ecx+0000000Ch]
  loc_0040CF70: lea edx, var_48
  loc_0040CF73: add ecx, eax
  loc_0040CF75: push 00000002h
  loc_0040CF77: lea eax, var_28
  loc_0040CF7A: push edx
  loc_0040CF7B: push eax
  loc_0040CF7C: mov var_40, ecx
  loc_0040CF7F: mov var_48, 00004004h
  loc_0040CF86: call [004010E8h] ; rtcRound
  loc_0040CF8C: lea ecx, var_18
  loc_0040CF8F: push ecx
  loc_0040CF90: call [00401174h] ; __vbaAryUnlock
  loc_0040CF96: lea edx, var_28
  loc_0040CF99: push edx
  loc_0040CF9A: call [00401018h] ; __vbaStrVarMove
  loc_0040CFA0: sub esp, 00000010h
  loc_0040CFA3: mov ecx, 00000008h
  loc_0040CFA8: mov edx, esp
  loc_0040CFAA: mov var_38, ecx
  loc_0040CFAD: mov var_30, eax
  loc_0040CFB0: push 00000000h
  loc_0040CFB2: mov [edx], ecx
  loc_0040CFB4: mov ecx, var_34
  loc_0040CFB7: mov [edx+00000004h], ecx
  loc_0040CFBA: mov ecx, var_60
  loc_0040CFBD: push ecx
  loc_0040CFBE: mov [edx+00000008h], eax
  loc_0040CFC1: mov eax, var_2C
  loc_0040CFC4: mov [edx+0000000Ch], eax
  loc_0040CFC7: call __vbaLateIdSt
  loc_0040CFC9: lea edx, var_38
  loc_0040CFCC: lea eax, var_28
  loc_0040CFCF: push edx
  loc_0040CFD0: push eax
  loc_0040CFD1: push 00000002h
  loc_0040CFD3: call [00401024h] ; __vbaFreeVarList
  loc_0040CFD9: mov eax, 00000003h
  loc_0040CFDE: mov ecx, eax
  loc_0040CFE0: mov var_40, eax
  loc_0040CFE3: push ecx
  loc_0040CFE4: mov var_48, ecx
  loc_0040CFE7: mov edx, esp
  loc_0040CFE9: push 0000000Bh
  loc_0040CFEB: mov [edx], ecx
  loc_0040CFED: mov ecx, var_44
  loc_0040CFF0: mov [edx+00000004h], ecx
  loc_0040CFF3: mov ecx, var_60
  loc_0040CFF6: push ecx
  loc_0040CFF7: mov [edx+00000008h], eax
  loc_0040CFFA: mov eax, var_3C
  loc_0040CFFD: mov [edx+0000000Ch], eax
  loc_0040D000: call __vbaLateIdSt
  loc_0040D002: mov edx, [edi+0000007Ch]
  loc_0040D005: lea eax, var_18
  loc_0040D008: push edx
  loc_0040D009: push eax
  loc_0040D00A: call ebx
  loc_0040D00C: mov ecx, var_18
  loc_0040D00F: test ecx, ecx
  loc_0040D011: jz 0040D03Ch
  loc_0040D013: cmp [ecx], 0001h
  loc_0040D017: jnz 0040D03Ch
  loc_0040D019: mov eax, var_70
  loc_0040D01C: mov edx, [ecx+00000014h]
  loc_0040D01F: sub eax, edx
  loc_0040D021: mov edx, [ecx+00000010h]
  loc_0040D024: cmp eax, edx
  loc_0040D026: mov var_5C, eax
  loc_0040D029: jb 0040D037h
  loc_0040D02B: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D031: mov ecx, var_18
  loc_0040D034: mov eax, var_5C
  loc_0040D037: shl eax, 02h
  loc_0040D03A: jmp 0040D045h
  loc_0040D03C: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D042: mov ecx, var_18
  loc_0040D045: mov ecx, [ecx+0000000Ch]
  loc_0040D048: lea edx, var_48
  loc_0040D04B: add ecx, eax
  loc_0040D04D: push 00000002h
  loc_0040D04F: lea eax, var_28
  loc_0040D052: push edx
  loc_0040D053: push eax
  loc_0040D054: mov var_40, ecx
  loc_0040D057: mov var_48, 00004004h
  loc_0040D05E: call [004010E8h] ; rtcRound
  loc_0040D064: lea ecx, var_18
  loc_0040D067: push ecx
  loc_0040D068: call [00401174h] ; __vbaAryUnlock
  loc_0040D06E: lea edx, var_28
  loc_0040D071: push edx
  loc_0040D072: call [00401018h] ; __vbaStrVarMove
  loc_0040D078: sub esp, 00000010h
  loc_0040D07B: mov ecx, 00000008h
  loc_0040D080: mov edx, esp
  loc_0040D082: mov var_38, ecx
  loc_0040D085: mov var_30, eax
  loc_0040D088: push 00000000h
  loc_0040D08A: mov [edx], ecx
  loc_0040D08C: mov ecx, var_34
  loc_0040D08F: mov [edx+00000004h], ecx
  loc_0040D092: mov ecx, var_60
  loc_0040D095: push ecx
  loc_0040D096: mov [edx+00000008h], eax
  loc_0040D099: mov eax, var_2C
  loc_0040D09C: mov [edx+0000000Ch], eax
  loc_0040D09F: call __vbaLateIdSt
  loc_0040D0A1: lea edx, var_38
  loc_0040D0A4: lea eax, var_28
  loc_0040D0A7: push edx
  loc_0040D0A8: push eax
  loc_0040D0A9: push 00000002h
  loc_0040D0AB: call [00401024h] ; __vbaFreeVarList
  loc_0040D0B1: mov ecx, 00000003h
  loc_0040D0B6: mov eax, 00000004h
  loc_0040D0BB: push ecx
  loc_0040D0BC: mov var_48, ecx
  loc_0040D0BF: mov edx, esp
  loc_0040D0C1: mov var_40, eax
  loc_0040D0C4: push 0000000Bh
  loc_0040D0C6: mov [edx], ecx
  loc_0040D0C8: mov ecx, var_44
  loc_0040D0CB: mov [edx+00000004h], ecx
  loc_0040D0CE: mov ecx, var_60
  loc_0040D0D1: push ecx
  loc_0040D0D2: mov [edx+00000008h], eax
  loc_0040D0D5: mov eax, var_3C
  loc_0040D0D8: mov [edx+0000000Ch], eax
  loc_0040D0DB: call __vbaLateIdSt
  loc_0040D0DD: mov edx, [edi+00000084h]
  loc_0040D0E3: lea eax, var_18
  loc_0040D0E6: push edx
  loc_0040D0E7: push eax
  loc_0040D0E8: call ebx
  loc_0040D0EA: mov ecx, var_18
  loc_0040D0ED: test ecx, ecx
  loc_0040D0EF: jz 0040D11Ah
  loc_0040D0F1: cmp [ecx], 0001h
  loc_0040D0F5: jnz 0040D11Ah
  loc_0040D0F7: mov eax, var_70
  loc_0040D0FA: mov edx, [ecx+00000014h]
  loc_0040D0FD: sub eax, edx
  loc_0040D0FF: mov edx, [ecx+00000010h]
  loc_0040D102: cmp eax, edx
  loc_0040D104: mov var_5C, eax
  loc_0040D107: jb 0040D115h
  loc_0040D109: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D10F: mov ecx, var_18
  loc_0040D112: mov eax, var_5C
  loc_0040D115: shl eax, 02h
  loc_0040D118: jmp 0040D123h
  loc_0040D11A: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D120: mov ecx, var_18
  loc_0040D123: mov ecx, [ecx+0000000Ch]
  loc_0040D126: lea edx, var_48
  loc_0040D129: add ecx, eax
  loc_0040D12B: push 00000002h
  loc_0040D12D: lea eax, var_28
  loc_0040D130: push edx
  loc_0040D131: push eax
  loc_0040D132: mov var_40, ecx
  loc_0040D135: mov var_48, 00004004h
  loc_0040D13C: call [004010E8h] ; rtcRound
  loc_0040D142: lea ecx, var_18
  loc_0040D145: push ecx
  loc_0040D146: call [00401174h] ; __vbaAryUnlock
  loc_0040D14C: lea edx, var_28
  loc_0040D14F: push edx
  loc_0040D150: call [00401018h] ; __vbaStrVarMove
  loc_0040D156: sub esp, 00000010h
  loc_0040D159: mov ecx, 00000008h
  loc_0040D15E: mov edx, esp
  loc_0040D160: mov var_38, ecx
  loc_0040D163: mov var_30, eax
  loc_0040D166: push 00000000h
  loc_0040D168: mov [edx], ecx
  loc_0040D16A: mov ecx, var_34
  loc_0040D16D: mov [edx+00000004h], ecx
  loc_0040D170: mov ecx, var_60
  loc_0040D173: push ecx
  loc_0040D174: mov [edx+00000008h], eax
  loc_0040D177: mov eax, var_2C
  loc_0040D17A: mov [edx+0000000Ch], eax
  loc_0040D17D: call __vbaLateIdSt
  loc_0040D17F: lea edx, var_38
  loc_0040D182: lea eax, var_28
  loc_0040D185: push edx
  loc_0040D186: push eax
  loc_0040D187: push 00000002h
  loc_0040D189: call [00401024h] ; __vbaFreeVarList
  loc_0040D18F: mov ecx, 00000003h
  loc_0040D194: mov eax, 00000005h
  loc_0040D199: push ecx
  loc_0040D19A: mov var_48, ecx
  loc_0040D19D: mov edx, esp
  loc_0040D19F: mov var_40, eax
  loc_0040D1A2: push 0000000Bh
  loc_0040D1A4: mov [edx], ecx
  loc_0040D1A6: mov ecx, var_44
  loc_0040D1A9: mov [edx+00000004h], ecx
  loc_0040D1AC: mov ecx, var_60
  loc_0040D1AF: push ecx
  loc_0040D1B0: mov [edx+00000008h], eax
  loc_0040D1B3: mov eax, var_3C
  loc_0040D1B6: mov [edx+0000000Ch], eax
  loc_0040D1B9: call __vbaLateIdSt
  loc_0040D1BB: mov edx, [edi+00000080h]
  loc_0040D1C1: lea eax, var_18
  loc_0040D1C4: push edx
  loc_0040D1C5: push eax
  loc_0040D1C6: call ebx
  loc_0040D1C8: mov ecx, var_18
  loc_0040D1CB: test ecx, ecx
  loc_0040D1CD: jz 0040D1F8h
  loc_0040D1CF: cmp [ecx], 0001h
  loc_0040D1D3: jnz 0040D1F8h
  loc_0040D1D5: mov eax, var_70
  loc_0040D1D8: mov edx, [ecx+00000014h]
  loc_0040D1DB: sub eax, edx
  loc_0040D1DD: mov edx, [ecx+00000010h]
  loc_0040D1E0: cmp eax, edx
  loc_0040D1E2: mov var_5C, eax
  loc_0040D1E5: jb 0040D1F3h
  loc_0040D1E7: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D1ED: mov ecx, var_18
  loc_0040D1F0: mov eax, var_5C
  loc_0040D1F3: shl eax, 02h
  loc_0040D1F6: jmp 0040D201h
  loc_0040D1F8: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D1FE: mov ecx, var_18
  loc_0040D201: mov ecx, [ecx+0000000Ch]
  loc_0040D204: lea edx, var_48
  loc_0040D207: add ecx, eax
  loc_0040D209: push 00000003h
  loc_0040D20B: lea eax, var_28
  loc_0040D20E: push edx
  loc_0040D20F: push eax
  loc_0040D210: mov var_40, ecx
  loc_0040D213: mov var_48, 00004004h
  loc_0040D21A: call [004010E8h] ; rtcRound
  loc_0040D220: lea ecx, var_18
  loc_0040D223: push ecx
  loc_0040D224: call [00401174h] ; __vbaAryUnlock
  loc_0040D22A: lea edx, var_28
  loc_0040D22D: push edx
  loc_0040D22E: call [00401018h] ; __vbaStrVarMove
  loc_0040D234: sub esp, 00000010h
  loc_0040D237: mov ecx, 00000008h
  loc_0040D23C: mov edx, esp
  loc_0040D23E: mov var_38, ecx
  loc_0040D241: mov var_30, eax
  loc_0040D244: push 00000000h
  loc_0040D246: mov [edx], ecx
  loc_0040D248: mov ecx, var_34
  loc_0040D24B: mov [edx+00000004h], ecx
  loc_0040D24E: mov ecx, var_60
  loc_0040D251: push ecx
  loc_0040D252: mov [edx+00000008h], eax
  loc_0040D255: mov eax, var_2C
  loc_0040D258: mov [edx+0000000Ch], eax
  loc_0040D25B: call __vbaLateIdSt
  loc_0040D25D: lea edx, var_38
  loc_0040D260: lea eax, var_28
  loc_0040D263: push edx
  loc_0040D264: push eax
  loc_0040D265: push 00000002h
  loc_0040D267: call [00401024h] ; __vbaFreeVarList
  loc_0040D26D: mov ecx, 00000003h
  loc_0040D272: mov eax, 00000006h
  loc_0040D277: push ecx
  loc_0040D278: mov var_48, ecx
  loc_0040D27B: mov edx, esp
  loc_0040D27D: mov var_40, eax
  loc_0040D280: push 0000000Bh
  loc_0040D282: mov [edx], ecx
  loc_0040D284: mov ecx, var_44
  loc_0040D287: mov [edx+00000004h], ecx
  loc_0040D28A: mov ecx, var_60
  loc_0040D28D: push ecx
  loc_0040D28E: mov [edx+00000008h], eax
  loc_0040D291: mov eax, var_3C
  loc_0040D294: mov [edx+0000000Ch], eax
  loc_0040D297: call __vbaLateIdSt
  loc_0040D299: mov edx, [edi+00000088h]
  loc_0040D29F: lea eax, var_18
  loc_0040D2A2: push edx
  loc_0040D2A3: push eax
  loc_0040D2A4: call ebx
  loc_0040D2A6: mov ecx, var_18
  loc_0040D2A9: test ecx, ecx
  loc_0040D2AB: jz 0040D2D6h
  loc_0040D2AD: cmp [ecx], 0001h
  loc_0040D2B1: jnz 0040D2D6h
  loc_0040D2B3: mov eax, var_70
  loc_0040D2B6: mov edx, [ecx+00000014h]
  loc_0040D2B9: sub eax, edx
  loc_0040D2BB: mov edx, [ecx+00000010h]
  loc_0040D2BE: cmp eax, edx
  loc_0040D2C0: mov var_5C, eax
  loc_0040D2C3: jb 0040D2D1h
  loc_0040D2C5: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D2CB: mov ecx, var_18
  loc_0040D2CE: mov eax, var_5C
  loc_0040D2D1: shl eax, 02h
  loc_0040D2D4: jmp 0040D2DFh
  loc_0040D2D6: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D2DC: mov ecx, var_18
  loc_0040D2DF: mov ecx, [ecx+0000000Ch]
  loc_0040D2E2: lea edx, var_48
  loc_0040D2E5: add ecx, eax
  loc_0040D2E7: push 00000002h
  loc_0040D2E9: lea eax, var_28
  loc_0040D2EC: push edx
  loc_0040D2ED: push eax
  loc_0040D2EE: mov var_40, ecx
  loc_0040D2F1: mov var_48, 00004004h
  loc_0040D2F8: call [004010E8h] ; rtcRound
  loc_0040D2FE: lea ecx, var_18
  loc_0040D301: push ecx
  loc_0040D302: call [00401174h] ; __vbaAryUnlock
  loc_0040D308: lea edx, var_28
  loc_0040D30B: push edx
  loc_0040D30C: call [00401018h] ; __vbaStrVarMove
  loc_0040D312: sub esp, 00000010h
  loc_0040D315: mov ecx, 00000008h
  loc_0040D31A: mov edx, esp
  loc_0040D31C: mov var_38, ecx
  loc_0040D31F: mov var_30, eax
  loc_0040D322: push 00000000h
  loc_0040D324: mov [edx], ecx
  loc_0040D326: mov ecx, var_34
  loc_0040D329: mov [edx+00000004h], ecx
  loc_0040D32C: mov ecx, var_60
  loc_0040D32F: push ecx
  loc_0040D330: mov [edx+00000008h], eax
  loc_0040D333: mov eax, var_2C
  loc_0040D336: mov [edx+0000000Ch], eax
  loc_0040D339: call __vbaLateIdSt
  loc_0040D33B: lea edx, var_38
  loc_0040D33E: lea eax, var_28
  loc_0040D341: push edx
  loc_0040D342: push eax
  loc_0040D343: push 00000002h
  loc_0040D345: call [00401024h] ; __vbaFreeVarList
  loc_0040D34B: add esp, 0000000Ch
  loc_0040D34E: lea ecx, var_60
  loc_0040D351: push 00000000h
  loc_0040D353: push ecx
  loc_0040D354: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040D35A: mov eax, 00000001h
  loc_0040D35F: add ax, var_14
  loc_0040D363: jo 0040D3B1h
  loc_0040D365: mov var_14, eax
  loc_0040D368: jmp 0040CD68h
  loc_0040D36D: push 0040D39Ch
  loc_0040D372: jmp 0040D392h
  loc_0040D374: lea edx, var_18
  loc_0040D377: push edx
  loc_0040D378: call [00401174h] ; __vbaAryUnlock
  loc_0040D37E: lea eax, var_38
  loc_0040D381: lea ecx, var_28
  loc_0040D384: push eax
  loc_0040D385: push ecx
  loc_0040D386: push 00000002h
  loc_0040D388: call [00401024h] ; __vbaFreeVarList
  loc_0040D38E: add esp, 0000000Ch
  loc_0040D391: ret
  loc_0040D392: lea ecx, var_60
  loc_0040D395: call [00401180h] ; __vbaFreeObj
  loc_0040D39B: ret
  loc_0040D39C: mov ecx, var_10
  loc_0040D39F: pop edi
  loc_0040D3A0: pop esi
  loc_0040D3A1: xor eax, eax
  loc_0040D3A3: mov fs:[00000000h], ecx
  loc_0040D3AA: pop ebx
  loc_0040D3AB: mov esp, ebp
  loc_0040D3AD: pop ebp
  loc_0040D3AE: retn 0004h
End Sub

Private Sub Proc_3_26_40D630(arg_C) '40D630
  loc_0040D630: push ebp
  loc_0040D631: mov ebp, esp
  loc_0040D633: sub esp, 00000008h
  loc_0040D636: push 00401546h ; __vbaExceptHandler
  loc_0040D63B: mov eax, fs:[00000000h]
  loc_0040D641: push eax
  loc_0040D642: mov fs:[00000000h], esp
  loc_0040D649: sub esp, 00000060h
  loc_0040D64C: push ebx
  loc_0040D64D: push esi
  loc_0040D64E: push edi
  loc_0040D64F: mov var_8, esp
  loc_0040D652: mov var_4, 00401270h
  loc_0040D659: mov esi, Me
  loc_0040D65C: xor edi, edi
  loc_0040D65E: push esi
  loc_0040D65F: mov var_1C, edi
  loc_0040D662: mov eax, [esi]
  loc_0040D664: mov var_20, edi
  loc_0040D667: mov var_28, edi
  loc_0040D66A: mov var_38, edi
  loc_0040D66D: mov var_6C, edi
  loc_0040D670: call [eax+00000404h]
  loc_0040D676: lea ecx, var_6C
  loc_0040D679: push eax
  loc_0040D67A: push ecx
  loc_0040D67B: call [0040105Ch] ; __vbaObjSet
  loc_0040D681: mov ebx, var_44
  loc_0040D684: sub esp, 00000010h
  loc_0040D687: mov edx, esp
  loc_0040D689: mov ecx, 00000008h
  loc_0040D68E: mov eax, 00403B90h ; "Text files |*.txt"
  loc_0040D693: push 00000003h
  loc_0040D695: mov [edx], ecx
  loc_0040D697: mov ecx, var_6C
  loc_0040D69A: push ecx
  loc_0040D69B: mov [edx+00000004h], ebx
  loc_0040D69E: mov [edx+00000008h], eax
  loc_0040D6A1: mov eax, var_3C
  loc_0040D6A4: mov [edx+0000000Ch], eax
  loc_0040D6A7: call [0040116Ch] ; __vbaLateIdSt
  loc_0040D6AD: sub esp, 00000010h
  loc_0040D6B0: mov ecx, 00000008h
  loc_0040D6B5: mov edx, esp
  loc_0040D6B7: mov eax, 00403BB8h ; "Open a text file"
  loc_0040D6BC: push 00000002h
  loc_0040D6BE: mov [edx], ecx
  loc_0040D6C0: mov ecx, var_6C
  loc_0040D6C3: push ecx
  loc_0040D6C4: mov [edx+00000004h], ebx
  loc_0040D6C7: mov [edx+00000008h], eax
  loc_0040D6CA: mov eax, var_3C
  loc_0040D6CD: mov [edx+0000000Ch], eax
  loc_0040D6D0: call [0040116Ch] ; __vbaLateIdSt
  loc_0040D6D6: mov edx, var_6C
  loc_0040D6D9: push edi
  loc_0040D6DA: push 0000001Eh
  loc_0040D6DC: push edx
  loc_0040D6DD: call [0040101Ch] ; __vbaLateIdCall
  loc_0040D6E3: mov eax, var_6C
  loc_0040D6E6: mov ebx, [004010B4h] ; __vbaLateIdCallLd
  loc_0040D6EC: push edi
  loc_0040D6ED: push 00000001h
  loc_0040D6EF: lea ecx, var_38
  loc_0040D6F2: push eax
  loc_0040D6F3: push ecx
  loc_0040D6F4: call ebx
  loc_0040D6F6: add esp, 0000001Ch
  loc_0040D6F9: push eax
  loc_0040D6FA: call [00401018h] ; __vbaStrVarMove
  loc_0040D700: mov edx, eax
  loc_0040D702: lea ecx, var_28
  loc_0040D705: call [00401164h] ; __vbaStrMove
  loc_0040D70B: push eax
  loc_0040D70C: push 00000001h
  loc_0040D70E: push FFFFFFFFh
  loc_0040D710: push 00000001h
  loc_0040D712: call [00401114h] ; __vbaFileOpen
  loc_0040D718: lea ecx, var_28
  loc_0040D71B: call [00401184h] ; __vbaFreeStr
  loc_0040D721: lea ecx, var_38
  loc_0040D724: call [00401014h] ; __vbaFreeVar
  loc_0040D72A: push 00000001h
  loc_0040D72C: call [00401120h] ; rtcEndOfFile
  loc_0040D732: test ax, ax
  loc_0040D735: jnz 0040D75Bh
  loc_0040D737: lea edx, var_20
  loc_0040D73A: lea eax, var_1C
  loc_0040D73D: push edx
  loc_0040D73E: push eax
  loc_0040D73F: add di, 0001h
  loc_0040D743: push 00000001h
  loc_0040D745: push 00403BE0h
  loc_0040D74A: jo 0040D909h
  loc_0040D750: call [004010D4h] ; __vbaInputFile
  loc_0040D756: add esp, 00000010h
  loc_0040D759: jmp 0040D72Ah
  loc_0040D75B: push 00000001h
  loc_0040D75D: call [0040108Ch] ; __vbaFileClose
  loc_0040D763: movsx ecx, di
  loc_0040D766: push 00000000h
  loc_0040D768: push ecx
  loc_0040D769: push 00000000h
  loc_0040D76B: push 00000002h
  loc_0040D76D: add esi, 0000003Ch
  loc_0040D770: push 00000002h
  loc_0040D772: push 00000004h
  loc_0040D774: push esi
  loc_0040D775: push 00000004h
  loc_0040D777: push 00000080h
  loc_0040D77C: call [004010B8h] ; __vbaRedim
  loc_0040D782: mov edx, var_6C
  loc_0040D785: xor edi, edi
  loc_0040D787: push edi
  loc_0040D788: push 00000001h
  loc_0040D78A: lea eax, var_38
  loc_0040D78D: push edx
  loc_0040D78E: push eax
  loc_0040D78F: call ebx
  loc_0040D791: add esp, 00000034h
  loc_0040D794: push eax
  loc_0040D795: call [00401018h] ; __vbaStrVarMove
  loc_0040D79B: mov edx, eax
  loc_0040D79D: lea ecx, var_28
  loc_0040D7A0: call [00401164h] ; __vbaStrMove
  loc_0040D7A6: push eax
  loc_0040D7A7: push 00000001h
  loc_0040D7A9: push FFFFFFFFh
  loc_0040D7AB: push 00000001h
  loc_0040D7AD: call [00401114h] ; __vbaFileOpen
  loc_0040D7B3: lea ecx, var_28
  loc_0040D7B6: call [00401184h] ; __vbaFreeStr
  loc_0040D7BC: lea ecx, var_38
  loc_0040D7BF: call [00401014h] ; __vbaFreeVar
  loc_0040D7C5: push 00000001h
  loc_0040D7C7: call [00401120h] ; rtcEndOfFile
  loc_0040D7CD: test ax, ax
  loc_0040D7D0: jnz 0040D8B6h
  loc_0040D7D6: mov eax, [esi]
  loc_0040D7D8: add di, 0001h
  loc_0040D7DC: jo 0040D909h
  loc_0040D7E2: mov ebx, edi
  loc_0040D7E4: test eax, eax
  loc_0040D7E6: mov var_18, ebx
  loc_0040D7E9: jz 0040D837h
  loc_0040D7EB: cmp [eax], 0002h
  loc_0040D7EF: jnz 0040D837h
  loc_0040D7F1: mov edi, [eax+0000001Ch]
  loc_0040D7F4: mov edx, [eax+00000018h]
  loc_0040D7F7: movsx ecx, bx
  loc_0040D7FA: sub ecx, edi
  loc_0040D7FC: cmp ecx, edx
  loc_0040D7FE: mov var_68, ecx
  loc_0040D801: jb 0040D809h
  loc_0040D803: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D809: mov eax, [esi]
  loc_0040D80B: mov ebx, 00000002h
  loc_0040D810: mov edx, [eax+00000014h]
  loc_0040D813: mov ecx, [eax+00000010h]
  loc_0040D816: sub ebx, edx
  loc_0040D818: cmp ebx, ecx
  loc_0040D81A: jb 0040D822h
  loc_0040D81C: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D822: mov ecx, [esi]
  loc_0040D824: mov edi, [ecx+00000018h]
  loc_0040D827: mov ecx, var_68
  loc_0040D82A: imul edi, ebx
  loc_0040D82D: mov ebx, var_18
  loc_0040D830: add edi, ecx
  loc_0040D832: shl edi, 02h
  loc_0040D835: jmp 0040D83Fh
  loc_0040D837: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D83D: mov edi, eax
  loc_0040D83F: mov eax, [esi]
  loc_0040D841: test eax, eax
  loc_0040D843: jz 0040D88Ch
  loc_0040D845: cmp [eax], 0002h
  loc_0040D849: jnz 0040D88Ch
  loc_0040D84B: mov edx, [eax+00000018h]
  loc_0040D84E: movsx ecx, bx
  loc_0040D851: sub ecx, [eax+0000001Ch]
  loc_0040D854: cmp ecx, edx
  loc_0040D856: mov var_60, ecx
  loc_0040D859: jb 0040D861h
  loc_0040D85B: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D861: mov eax, [esi]
  loc_0040D863: mov ebx, 00000001h
  loc_0040D868: mov edx, [eax+00000014h]
  loc_0040D86B: mov ecx, [eax+00000010h]
  loc_0040D86E: sub ebx, edx
  loc_0040D870: cmp ebx, ecx
  loc_0040D872: jb 0040D87Ah
  loc_0040D874: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D87A: mov edx, [esi]
  loc_0040D87C: mov eax, [edx+00000018h]
  loc_0040D87F: mov edx, var_60
  loc_0040D882: imul eax, ebx
  loc_0040D885: add eax, edx
  loc_0040D887: shl eax, 02h
  loc_0040D88A: jmp 0040D892h
  loc_0040D88C: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040D892: mov ecx, [esi]
  loc_0040D894: mov ecx, [ecx+0000000Ch]
  loc_0040D897: lea edx, [ecx+edi]
  loc_0040D89A: add ecx, eax
  loc_0040D89C: push edx
  loc_0040D89D: push ecx
  loc_0040D89E: push 00000001h
  loc_0040D8A0: push 00403BE0h
  loc_0040D8A5: call [004010D4h] ; __vbaInputFile
  loc_0040D8AB: mov edi, var_18
  loc_0040D8AE: add esp, 00000010h
  loc_0040D8B1: jmp 0040D7C5h
  loc_0040D8B6: push 00000001h
  loc_0040D8B8: call [0040108Ch] ; __vbaFileClose
  loc_0040D8BE: lea eax, var_6C
  loc_0040D8C1: push 00000000h
  loc_0040D8C3: push eax
  loc_0040D8C4: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040D8CA: mov ecx, arg_C
  loc_0040D8CD: push 0040D8F4h
  loc_0040D8D2: mov [ecx], di
  loc_0040D8D5: jmp 0040D8EAh
  loc_0040D8D7: lea ecx, var_28
  loc_0040D8DA: call [00401184h] ; __vbaFreeStr
  loc_0040D8E0: lea ecx, var_38
  loc_0040D8E3: call [00401014h] ; __vbaFreeVar
  loc_0040D8E9: ret
  loc_0040D8EA: lea ecx, var_6C
  loc_0040D8ED: call [00401180h] ; __vbaFreeObj
  loc_0040D8F3: ret
  loc_0040D8F4: mov ecx, var_10
  loc_0040D8F7: pop edi
  loc_0040D8F8: pop esi
  loc_0040D8F9: xor eax, eax
  loc_0040D8FB: mov fs:[00000000h], ecx
  loc_0040D902: pop ebx
  loc_0040D903: mov esp, ebp
  loc_0040D905: pop ebp
  loc_0040D906: retn 0008h
End Sub

Private Function Proc_3_27_40D910
  loc_0040D910: mov eax, [00415028h]
  loc_0040D915: mov ecx, [0041502Ch]
  loc_0040D91B: push esi
  loc_0040D91C: push eax
  loc_0040D91D: push ecx
  loc_0040D91E: call [004010A0h] ; __vbaR4Str
  loc_0040D924: push ecx
  loc_0040D925: mov esi, Me
  loc_0040D929: fstp real4 ptr [esp]
  loc_0040D92C: push 00000001h
  loc_0040D92E: lea edx, [esi+0000003Ch]
  loc_0040D931: push 00000002h
  loc_0040D933: push edx
  loc_0040D934: call 00409C80h
  loc_0040D939: mov eax, [esi]
  loc_0040D93B: push esi
  loc_0040D93C: call [eax+00000780h]
  loc_0040D942: lea ecx, [esi+00000080h]
  loc_0040D948: lea edx, [esi+00000088h]
  loc_0040D94E: push ecx
  loc_0040D94F: lea eax, [esi+00000084h]
  loc_0040D955: push edx
  loc_0040D956: lea ecx, [esi+0000007Ch]
  loc_0040D959: push eax
  loc_0040D95A: lea edx, [esi+00000078h]
  loc_0040D95D: push ecx
  loc_0040D95E: lea eax, [esi+00000074h]
  loc_0040D961: push edx
  loc_0040D962: push eax
  loc_0040D963: call 00409DD0h
  loc_0040D968: mov ecx, [esi]
  loc_0040D96A: push esi
  loc_0040D96B: call [ecx+0000071Ch]
  loc_0040D971: mov edx, [esi]
  loc_0040D973: push esi
  loc_0040D974: call [edx+00000718h]
  loc_0040D97A: xor eax, eax
  loc_0040D97C: pop esi
  loc_0040D97D: retn 0004h
End Sub

Private Sub Proc_3_28_40D980
  loc_0040D980: push ebp
  loc_0040D981: mov ebp, esp
  loc_0040D983: sub esp, 00000008h
  loc_0040D986: push 00401546h ; __vbaExceptHandler
  loc_0040D98B: mov eax, fs:[00000000h]
  loc_0040D991: push eax
  loc_0040D992: mov fs:[00000000h], esp
  loc_0040D999: sub esp, 00000030h
  loc_0040D99C: push ebx
  loc_0040D99D: push esi
  loc_0040D99E: push edi
  loc_0040D99F: mov var_8, esp
  loc_0040D9A2: mov var_4, 00401280h
  loc_0040D9A9: mov esi, Me
  loc_0040D9AC: xor eax, eax
  loc_0040D9AE: mov var_18, eax
  loc_0040D9B1: mov var_1C, eax
  loc_0040D9B4: mov var_20, eax
  loc_0040D9B7: mov ax, [esi+00000038h]
  loc_0040D9BB: mov var_38, ax
  loc_0040D9BF: mov ebx, 00000001h
  loc_0040D9C4: cmp bx, var_38
  loc_0040D9C8: jg 0040DA7Bh
  loc_0040D9CE: mov eax, [0041567Ch]
  loc_0040D9D3: test eax, eax
  loc_0040D9D5: jnz 0040D9E7h
  loc_0040D9D7: push 0041567Ch
  loc_0040D9DC: push 00403C18h
  loc_0040D9E1: call [0040111Ch] ; __vbaNew2
  loc_0040D9E7: mov ecx, [esi]
  loc_0040D9E9: mov edi, [0041567Ch]
  loc_0040D9EF: push esi
  loc_0040D9F0: call [ecx+00000314h]
  loc_0040D9F6: lea edx, var_18
  loc_0040D9F9: push eax
  loc_0040D9FA: push edx
  loc_0040D9FB: call [0040105Ch] ; __vbaObjSet
  loc_0040DA01: mov esi, eax
  loc_0040DA03: lea ecx, var_1C
  loc_0040DA06: push ecx
  loc_0040DA07: push ebx
  loc_0040DA08: mov eax, [esi]
  loc_0040DA0A: push esi
  loc_0040DA0B: call [eax+00000040h]
  loc_0040DA0E: test eax, eax
  loc_0040DA10: fnclex
  loc_0040DA12: jge 0040DA23h
  loc_0040DA14: push 00000040h
  loc_0040DA16: push 00403BE8h
  loc_0040DA1B: push esi
  loc_0040DA1C: push eax
  loc_0040DA1D: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DA23: mov eax, var_1C
  loc_0040DA26: lea edx, var_20
  loc_0040DA29: push eax
  loc_0040DA2A: mov var_1C, 00000000h
  loc_0040DA31: mov esi, [edi]
  loc_0040DA33: push edx
  loc_0040DA34: call [0040105Ch] ; __vbaObjSet
  loc_0040DA3A: push eax
  loc_0040DA3B: push edi
  loc_0040DA3C: call [esi+00000010h]
  loc_0040DA3F: test eax, eax
  loc_0040DA41: fnclex
  loc_0040DA43: jge 0040DA54h
  loc_0040DA45: push 00000010h
  loc_0040DA47: push 00403C08h
  loc_0040DA4C: push edi
  loc_0040DA4D: push eax
  loc_0040DA4E: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DA54: lea eax, var_20
  loc_0040DA57: lea ecx, var_18
  loc_0040DA5A: push eax
  loc_0040DA5B: push ecx
  loc_0040DA5C: push 00000002h
  loc_0040DA5E: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DA64: mov esi, Me
  loc_0040DA67: mov eax, 00000001h
  loc_0040DA6C: add esp, 0000000Ch
  loc_0040DA6F: add ax, bx
  loc_0040DA72: jo 0040DAB0h
  loc_0040DA74: mov ebx, eax
  loc_0040DA76: jmp 0040D9C4h
  loc_0040DA7B: push 0040DA9Bh
  loc_0040DA80: jmp 0040DA9Ah
  loc_0040DA82: lea edx, var_20
  loc_0040DA85: lea eax, var_1C
  loc_0040DA88: push edx
  loc_0040DA89: lea ecx, var_18
  loc_0040DA8C: push eax
  loc_0040DA8D: push ecx
  loc_0040DA8E: push 00000003h
  loc_0040DA90: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DA96: add esp, 00000010h
  loc_0040DA99: ret
  loc_0040DA9A: ret
  loc_0040DA9B: mov ecx, var_10
  loc_0040DA9E: pop edi
  loc_0040DA9F: pop esi
  loc_0040DAA0: xor eax, eax
  loc_0040DAA2: mov fs:[00000000h], ecx
  loc_0040DAA9: pop ebx
  loc_0040DAAA: mov esp, ebp
  loc_0040DAAC: pop ebp
  loc_0040DAAD: retn 0004h
End Sub

Private Sub Proc_3_29_40DAC0
  loc_0040DAC0: push ebp
  loc_0040DAC1: mov ebp, esp
  loc_0040DAC3: sub esp, 00000008h
  loc_0040DAC6: push 00401546h ; __vbaExceptHandler
  loc_0040DACB: mov eax, fs:[00000000h]
  loc_0040DAD1: push eax
  loc_0040DAD2: mov fs:[00000000h], esp
  loc_0040DAD9: sub esp, 00000074h
  loc_0040DADC: push ebx
  loc_0040DADD: push esi
  loc_0040DADE: push edi
  loc_0040DADF: mov var_8, esp
  loc_0040DAE2: mov var_4, 00401298h
  loc_0040DAE9: mov ebx, Me
  loc_0040DAEC: xor esi, esi
  loc_0040DAEE: push ebx
  loc_0040DAEF: mov var_1C, esi
  loc_0040DAF2: mov eax, [ebx]
  loc_0040DAF4: mov var_20, esi
  loc_0040DAF7: mov var_24, esi
  loc_0040DAFA: mov var_34, esi
  loc_0040DAFD: mov var_44, esi
  loc_0040DB00: mov var_6C, esi
  loc_0040DB03: call [eax+00000410h]
  loc_0040DB09: lea ecx, var_6C
  loc_0040DB0C: push eax
  loc_0040DB0D: push ecx
  loc_0040DB0E: call [0040105Ch] ; __vbaObjSet
  loc_0040DB14: mov dx, [ebx+00000038h]
  loc_0040DB18: mov var_18, 00000001h
  loc_0040DB1F: mov var_74, dx
  loc_0040DB23: mov ax, var_74
  loc_0040DB27: cmp var_18, ax
  loc_0040DB2B: jg 0040DEE7h
  loc_0040DB31: cmp [0041567Ch], esi
  loc_0040DB37: jnz 0040DB49h
  loc_0040DB39: push 0041567Ch
  loc_0040DB3E: push 00403C18h
  loc_0040DB43: call [0040111Ch] ; __vbaNew2
  loc_0040DB49: mov ecx, [ebx]
  loc_0040DB4B: mov edi, [0041567Ch]
  loc_0040DB51: push ebx
  loc_0040DB52: call [ecx+00000314h]
  loc_0040DB58: lea edx, var_1C
  loc_0040DB5B: push eax
  loc_0040DB5C: push edx
  loc_0040DB5D: call [0040105Ch] ; __vbaObjSet
  loc_0040DB63: mov edx, var_18
  loc_0040DB66: mov esi, eax
  loc_0040DB68: lea ecx, var_20
  loc_0040DB6B: mov eax, [esi]
  loc_0040DB6D: push ecx
  loc_0040DB6E: push edx
  loc_0040DB6F: push esi
  loc_0040DB70: call [eax+00000040h]
  loc_0040DB73: test eax, eax
  loc_0040DB75: fnclex
  loc_0040DB77: jge 0040DB88h
  loc_0040DB79: push 00000040h
  loc_0040DB7B: push 00403BE8h
  loc_0040DB80: push esi
  loc_0040DB81: push eax
  loc_0040DB82: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DB88: mov eax, var_20
  loc_0040DB8B: mov var_20, 00000000h
  loc_0040DB92: mov esi, [edi]
  loc_0040DB94: push eax
  loc_0040DB95: lea eax, var_24
  loc_0040DB98: push eax
  loc_0040DB99: call [0040105Ch] ; __vbaObjSet
  loc_0040DB9F: push eax
  loc_0040DBA0: push edi
  loc_0040DBA1: call [esi+0000000Ch]
  loc_0040DBA4: test eax, eax
  loc_0040DBA6: fnclex
  loc_0040DBA8: jge 0040DBB9h
  loc_0040DBAA: push 0000000Ch
  loc_0040DBAC: push 00403C08h
  loc_0040DBB1: push edi
  loc_0040DBB2: push eax
  loc_0040DBB3: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DBB9: lea ecx, var_24
  loc_0040DBBC: lea edx, var_1C
  loc_0040DBBF: push ecx
  loc_0040DBC0: push edx
  loc_0040DBC1: push 00000002h
  loc_0040DBC3: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DBC9: mov eax, [ebx]
  loc_0040DBCB: add esp, 0000000Ch
  loc_0040DBCE: push ebx
  loc_0040DBCF: call [eax+00000314h]
  loc_0040DBD5: lea ecx, var_1C
  loc_0040DBD8: push eax
  loc_0040DBD9: push ecx
  loc_0040DBDA: call [0040105Ch] ; __vbaObjSet
  loc_0040DBE0: mov edi, var_18
  loc_0040DBE3: mov esi, eax
  loc_0040DBE5: lea eax, var_20
  loc_0040DBE8: mov edx, [esi]
  loc_0040DBEA: push eax
  loc_0040DBEB: push edi
  loc_0040DBEC: push esi
  loc_0040DBED: call [edx+00000040h]
  loc_0040DBF0: test eax, eax
  loc_0040DBF2: fnclex
  loc_0040DBF4: jge 0040DC05h
  loc_0040DBF6: push 00000040h
  loc_0040DBF8: push 00403BE8h
  loc_0040DBFD: push esi
  loc_0040DBFE: push eax
  loc_0040DBFF: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DC05: mov eax, var_20
  loc_0040DC08: push FFFFFFFFh
  loc_0040DC0A: push eax
  loc_0040DC0B: mov esi, eax
  loc_0040DC0D: mov ecx, [eax]
  loc_0040DC0F: call [ecx+0000008Ch]
  loc_0040DC15: test eax, eax
  loc_0040DC17: fnclex
  loc_0040DC19: jge 0040DC2Dh
  loc_0040DC1B: push 0000008Ch
  loc_0040DC20: push 00403C28h
  loc_0040DC25: push esi
  loc_0040DC26: push eax
  loc_0040DC27: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DC2D: lea edx, var_20
  loc_0040DC30: lea eax, var_1C
  loc_0040DC33: push edx
  loc_0040DC34: push eax
  loc_0040DC35: push 00000002h
  loc_0040DC37: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DC3D: mov ecx, [ebx]
  loc_0040DC3F: add esp, 0000000Ch
  loc_0040DC42: push ebx
  loc_0040DC43: call [ecx+00000314h]
  loc_0040DC49: lea edx, var_1C
  loc_0040DC4C: push eax
  loc_0040DC4D: push edx
  loc_0040DC4E: call [0040105Ch] ; __vbaObjSet
  loc_0040DC54: mov esi, eax
  loc_0040DC56: lea ecx, var_20
  loc_0040DC59: push ecx
  loc_0040DC5A: push edi
  loc_0040DC5B: mov eax, [esi]
  loc_0040DC5D: push esi
  loc_0040DC5E: call [eax+00000040h]
  loc_0040DC61: test eax, eax
  loc_0040DC63: fnclex
  loc_0040DC65: jge 0040DC76h
  loc_0040DC67: push 00000040h
  loc_0040DC69: push 00403BE8h
  loc_0040DC6E: push esi
  loc_0040DC6F: push eax
  loc_0040DC70: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DC76: mov eax, [ebx+00000054h]
  loc_0040DC79: mov esi, var_20
  loc_0040DC7C: test eax, eax
  loc_0040DC7E: jz 0040DCAAh
  loc_0040DC80: cmp [eax], 0001h
  loc_0040DC84: jnz 0040DCAAh
  loc_0040DC86: mov edx, [eax+00000014h]
  loc_0040DC89: movsx ecx, di
  loc_0040DC8C: sub ecx, edx
  loc_0040DC8E: mov edx, [eax+00000010h]
  loc_0040DC91: cmp ecx, edx
  loc_0040DC93: mov var_58, ecx
  loc_0040DC96: jb 0040DCA1h
  loc_0040DC98: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DC9E: mov ecx, var_58
  loc_0040DCA1: lea eax, [ecx*4]
  loc_0040DCA8: jmp 0040DCB0h
  loc_0040DCAA: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DCB0: mov ecx, [ebx+00000054h]
  loc_0040DCB3: mov edx, [esi]
  loc_0040DCB5: mov ecx, [ecx+0000000Ch]
  loc_0040DCB8: push ecx
  loc_0040DCB9: fld real4 ptr [ecx+eax]
  loc_0040DCBC: fsub st0, real4 ptr [00401290h]
  loc_0040DCC2: fnstsw ax
  loc_0040DCC4: test al, 0Dh
  loc_0040DCC6: jnz 0040DF3Ah
  loc_0040DCCC: fstp real4 ptr [esp]
  loc_0040DCCF: push esi
  loc_0040DCD0: call [edx+00000074h]
  loc_0040DCD3: test eax, eax
  loc_0040DCD5: fnclex
  loc_0040DCD7: jge 0040DCE8h
  loc_0040DCD9: push 00000074h
  loc_0040DCDB: push 00403C28h
  loc_0040DCE0: push esi
  loc_0040DCE1: push eax
  loc_0040DCE2: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DCE8: lea edx, var_20
  loc_0040DCEB: lea eax, var_1C
  loc_0040DCEE: push edx
  loc_0040DCEF: push eax
  loc_0040DCF0: push 00000002h
  loc_0040DCF2: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DCF8: mov ecx, [ebx]
  loc_0040DCFA: add esp, 0000000Ch
  loc_0040DCFD: push ebx
  loc_0040DCFE: call [ecx+00000314h]
  loc_0040DD04: lea edx, var_1C
  loc_0040DD07: push eax
  loc_0040DD08: push edx
  loc_0040DD09: call [0040105Ch] ; __vbaObjSet
  loc_0040DD0F: mov esi, eax
  loc_0040DD11: lea ecx, var_20
  loc_0040DD14: push ecx
  loc_0040DD15: push edi
  loc_0040DD16: mov eax, [esi]
  loc_0040DD18: push esi
  loc_0040DD19: call [eax+00000040h]
  loc_0040DD1C: test eax, eax
  loc_0040DD1E: fnclex
  loc_0040DD20: jge 0040DD31h
  loc_0040DD22: push 00000040h
  loc_0040DD24: push 00403BE8h
  loc_0040DD29: push esi
  loc_0040DD2A: push eax
  loc_0040DD2B: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DD31: mov eax, [ebx+00000050h]
  loc_0040DD34: mov esi, var_20
  loc_0040DD37: test eax, eax
  loc_0040DD39: jz 0040DD62h
  loc_0040DD3B: cmp [eax], 0001h
  loc_0040DD3F: jnz 0040DD62h
  loc_0040DD41: mov edx, [eax+00000014h]
  loc_0040DD44: mov ecx, [eax+00000010h]
  loc_0040DD47: movsx edi, di
  loc_0040DD4A: sub edi, edx
  loc_0040DD4C: cmp edi, ecx
  loc_0040DD4E: jb 0040DD56h
  loc_0040DD50: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DD56: lea eax, [edi*4]
  loc_0040DD5D: mov edi, var_18
  loc_0040DD60: jmp 0040DD68h
  loc_0040DD62: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DD68: mov ecx, [ebx+00000050h]
  loc_0040DD6B: mov edx, [esi]
  loc_0040DD6D: mov ecx, [ecx+0000000Ch]
  loc_0040DD70: push ecx
  loc_0040DD71: fld real4 ptr [ecx+eax]
  loc_0040DD74: fsub st0, real4 ptr [00401290h]
  loc_0040DD7A: fnstsw ax
  loc_0040DD7C: test al, 0Dh
  loc_0040DD7E: jnz 0040DF3Ah
  loc_0040DD84: fstp real4 ptr [esp]
  loc_0040DD87: push esi
  loc_0040DD88: call [edx+0000006Ch]
  loc_0040DD8B: test eax, eax
  loc_0040DD8D: fnclex
  loc_0040DD8F: jge 0040DDA0h
  loc_0040DD91: push 0000006Ch
  loc_0040DD93: push 00403C28h
  loc_0040DD98: push esi
  loc_0040DD99: push eax
  loc_0040DD9A: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DDA0: lea edx, var_20
  loc_0040DDA3: lea eax, var_1C
  loc_0040DDA6: push edx
  loc_0040DDA7: push eax
  loc_0040DDA8: push 00000002h
  loc_0040DDAA: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DDB0: add esp, 0000000Ch
  loc_0040DDB3: xor eax, eax
  loc_0040DDB5: mov var_14, eax
  loc_0040DDB8: mov ecx, 00000002h
  loc_0040DDBD: cmp ax, cx
  loc_0040DDC0: Unknown_33E84589()
  loc_0040DDC6: mov edx, var_40
  loc_0040DDC9: sub esp, 00000010h
  loc_0040DDCC: movsx esi, ax
  loc_0040DDCF: mov ecx, esp
  loc_0040DDD1: mov eax, 00000003h
  loc_0040DDD6: push 0000000Bh
  loc_0040DDD8: mov [ecx], eax
  loc_0040DDDA: mov eax, var_38
  loc_0040DDDD: mov [ecx+00000004h], edx
  loc_0040DDE0: mov [ecx+00000008h], esi
  loc_0040DDE3: mov [ecx+0000000Ch], eax
  loc_0040DDE6: mov ecx, var_6C
  loc_0040DDE9: push ecx
  loc_0040DDEA: call [0040116Ch] ; __vbaLateIdSt
  loc_0040DDF0: sub esp, 00000010h
  loc_0040DDF3: mov eax, 00000003h
  loc_0040DDF8: mov edx, esp
  loc_0040DDFA: mov ecx, var_38
  loc_0040DDFD: movsx edi, di
  loc_0040DE00: mov [edx], eax
  loc_0040DE02: mov eax, var_40
  loc_0040DE05: push 0000000Ah
  loc_0040DE07: mov [edx+00000004h], eax
  loc_0040DE0A: mov [edx+00000008h], edi
  loc_0040DE0D: mov [edx+0000000Ch], ecx
  loc_0040DE10: mov edx, var_6C
  loc_0040DE13: push edx
  loc_0040DE14: call [0040116Ch] ; __vbaLateIdSt
  loc_0040DE1A: cmp var_14, 0000h
  loc_0040DE1F: jnz 0040DE2Dh
  loc_0040DE21: mov eax, var_18
  loc_0040DE24: push eax
  loc_0040DE25: call [00401000h] ; __vbaStrI2
  loc_0040DE2B: jmp 0040DE87h
  loc_0040DE2D: mov eax, [ebx+0000003Ch]
  loc_0040DE30: test eax, eax
  loc_0040DE32: jz 0040DE71h
  loc_0040DE34: cmp [eax], 0002h
  loc_0040DE38: jnz 0040DE71h
  loc_0040DE3A: mov edx, [eax+0000001Ch]
  loc_0040DE3D: mov ecx, [eax+00000018h]
  loc_0040DE40: sub edi, edx
  loc_0040DE42: cmp edi, ecx
  loc_0040DE44: jb 0040DE4Ch
  loc_0040DE46: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DE4C: mov eax, [ebx+0000003Ch]
  loc_0040DE4F: mov edx, [eax+00000014h]
  loc_0040DE52: mov ecx, [eax+00000010h]
  loc_0040DE55: sub esi, edx
  loc_0040DE57: cmp esi, ecx
  loc_0040DE59: jb 0040DE61h
  loc_0040DE5B: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DE61: mov edx, [ebx+0000003Ch]
  loc_0040DE64: mov eax, [edx+00000018h]
  loc_0040DE67: imul eax, esi
  loc_0040DE6A: add eax, edi
  loc_0040DE6C: shl eax, 02h
  loc_0040DE6F: jmp 0040DE77h
  loc_0040DE71: call [00401094h] ; __vbaGenerateBoundsError
  loc_0040DE77: mov ecx, [ebx+0000003Ch]
  loc_0040DE7A: mov edx, [ecx+0000000Ch]
  loc_0040DE7D: mov eax, [edx+eax]
  loc_0040DE80: push eax
  loc_0040DE81: call [004010A8h] ; __vbaStrR4
  loc_0040DE87: sub esp, 00000010h
  loc_0040DE8A: mov ecx, 00000008h
  loc_0040DE8F: mov edx, esp
  loc_0040DE91: mov var_34, ecx
  loc_0040DE94: mov var_2C, eax
  loc_0040DE97: push 00000000h
  loc_0040DE99: mov [edx], ecx
  loc_0040DE9B: mov ecx, var_30
  loc_0040DE9E: mov [edx+00000004h], ecx
  loc_0040DEA1: mov ecx, var_6C
  loc_0040DEA4: push ecx
  loc_0040DEA5: mov [edx+00000008h], eax
  loc_0040DEA8: mov eax, var_28
  loc_0040DEAB: mov [edx+0000000Ch], eax
  loc_0040DEAE: call [0040116Ch] ; __vbaLateIdSt
  loc_0040DEB4: lea ecx, var_34
  loc_0040DEB7: call [00401014h] ; __vbaFreeVar
  loc_0040DEBD: mov edi, var_18
  loc_0040DEC0: mov eax, 00000001h
  loc_0040DEC5: add ax, var_14
  loc_0040DEC9: jo 0040DF3Fh
  loc_0040DECB: mov var_14, eax
  loc_0040DECE: jmp 0040DDB8h
  loc_0040DED3: mov eax, 00000001h
  loc_0040DED8: add ax, di
  loc_0040DEDB: jo 0040DF3Fh
  loc_0040DEDD: mov var_18, eax
  loc_0040DEE0: xor esi, esi
  loc_0040DEE2: jmp 0040DB23h
  loc_0040DEE7: lea edx, var_6C
  loc_0040DEEA: push esi
  loc_0040DEEB: push edx
  loc_0040DEEC: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040DEF2: fwait
  loc_0040DEF3: push 0040DF25h
  loc_0040DEF8: jmp 0040DF1Bh
  loc_0040DEFA: lea eax, var_24
  loc_0040DEFD: lea ecx, var_20
  loc_0040DF00: push eax
  loc_0040DF01: lea edx, var_1C
  loc_0040DF04: push ecx
  loc_0040DF05: push edx
  loc_0040DF06: push 00000003h
  loc_0040DF08: call [0040102Ch] ; __vbaFreeObjList
  loc_0040DF0E: add esp, 00000010h
  loc_0040DF11: lea ecx, var_34
  loc_0040DF14: call [00401014h] ; __vbaFreeVar
  loc_0040DF1A: ret
  loc_0040DF1B: lea ecx, var_6C
  loc_0040DF1E: call [00401180h] ; __vbaFreeObj
  loc_0040DF24: ret
  loc_0040DF25: mov ecx, var_10
  loc_0040DF28: pop edi
  loc_0040DF29: pop esi
  loc_0040DF2A: xor eax, eax
  loc_0040DF2C: mov fs:[00000000h], ecx
  loc_0040DF33: pop ebx
  loc_0040DF34: mov esp, ebp
  loc_0040DF36: pop ebp
  loc_0040DF37: retn 0004h
End Sub

Private Function Proc_3_30_40DF50
  loc_0040DF50: mov eax, var_4
  loc_0040DF54: mov cx, [eax+00000038h]
  loc_0040DF58: mov eax, 00000001h
  loc_0040DF5D: cmp cx, ax
  loc_0040DF60: jl 0040DF6Dh
  loc_0040DF62: add ax, 0001h
  loc_0040DF66: jo 0040DF72h
  loc_0040DF68: cmp ax, cx
  loc_0040DF6B: jle 0040DF62h
  loc_0040DF6D: xor eax, eax
  loc_0040DF6F: retn 0004h
End Sub

Private Sub Proc_3_31_40DF80
  loc_0040DF80: push ebp
  loc_0040DF81: mov ebp, esp
  loc_0040DF83: sub esp, 00000008h
  loc_0040DF86: push 00401546h ; __vbaExceptHandler
  loc_0040DF8B: mov eax, fs:[00000000h]
  loc_0040DF91: push eax
  loc_0040DF92: mov fs:[00000000h], esp
  loc_0040DF99: sub esp, 00000010h
  loc_0040DF9C: push ebx
  loc_0040DF9D: push esi
  loc_0040DF9E: push edi
  loc_0040DF9F: mov var_8, esp
  loc_0040DFA2: mov var_4, 004012A8h
  loc_0040DFA9: mov eax, [0041567Ch]
  loc_0040DFAE: mov var_14, 00000000h
  loc_0040DFB5: test eax, eax
  loc_0040DFB7: jnz 0040DFC9h
  loc_0040DFB9: push 0041567Ch
  loc_0040DFBE: push 00403C18h
  loc_0040DFC3: call [0040111Ch] ; __vbaNew2
  loc_0040DFC9: mov eax, Me
  loc_0040DFCC: mov esi, [0041567Ch]
  loc_0040DFD2: lea ecx, var_14
  loc_0040DFD5: push eax
  loc_0040DFD6: mov edi, [esi]
  loc_0040DFD8: push ecx
  loc_0040DFD9: call [0040106Ch] ; __vbaObjSetAddref
  loc_0040DFDF: push eax
  loc_0040DFE0: push esi
  loc_0040DFE1: call [edi+00000010h]
  loc_0040DFE4: test eax, eax
  loc_0040DFE6: fnclex
  loc_0040DFE8: jge 0040DFF9h
  loc_0040DFEA: push 00000010h
  loc_0040DFEC: push 00403C08h
  loc_0040DFF1: push esi
  loc_0040DFF2: push eax
  loc_0040DFF3: call [00401040h] ; __vbaHresultCheckObj
  loc_0040DFF9: lea ecx, var_14
  loc_0040DFFC: call [00401180h] ; __vbaFreeObj
  loc_0040E002: call [00401020h] ; __vbaEnd
  loc_0040E008: push 0040E01Ah
  loc_0040E00D: jmp 0040E019h
  loc_0040E00F: lea ecx, var_14
  loc_0040E012: call [00401180h] ; __vbaFreeObj
  loc_0040E018: ret
  loc_0040E019: ret
  loc_0040E01A: mov ecx, var_10
  loc_0040E01D: pop edi
  loc_0040E01E: pop esi
  loc_0040E01F: xor eax, eax
  loc_0040E021: mov fs:[00000000h], ecx
  loc_0040E028: pop ebx
  loc_0040E029: mov esp, ebp
  loc_0040E02B: pop ebp
  loc_0040E02C: retn 0004h
End Sub

Private Sub Proc_3_32_40E030
  loc_0040E030: push ebp
  loc_0040E031: mov ebp, esp
  loc_0040E033: sub esp, 00000008h
  loc_0040E036: push 00401546h ; __vbaExceptHandler
  loc_0040E03B: mov eax, fs:[00000000h]
  loc_0040E041: push eax
  loc_0040E042: mov fs:[00000000h], esp
  loc_0040E049: sub esp, 00000034h
  loc_0040E04C: push ebx
  loc_0040E04D: push esi
  loc_0040E04E: push edi
  loc_0040E04F: mov var_8, esp
  loc_0040E052: mov var_4, 004012B8h
  loc_0040E059: mov esi, Me
  loc_0040E05C: xor eax, eax
  loc_0040E05E: mov var_14, eax
  loc_0040E061: mov var_18, eax
  loc_0040E064: mov var_28, eax
  loc_0040E067: mov var_38, eax
  loc_0040E06A: mov eax, [esi]
  loc_0040E06C: push esi
  loc_0040E06D: call [eax+000003B4h]
  loc_0040E073: lea ecx, var_18
  loc_0040E076: push eax
  loc_0040E077: push ecx
  loc_0040E078: call [0040105Ch] ; __vbaObjSet
  loc_0040E07E: mov edi, eax
  loc_0040E080: lea eax, var_38
  loc_0040E083: push 00000003h
  loc_0040E085: lea ecx, var_28
  loc_0040E088: lea edx, [esi+00000044h]
  loc_0040E08B: push eax
  loc_0040E08C: push ecx
  loc_0040E08D: mov var_30, edx
  loc_0040E090: mov var_38, 00004004h
  loc_0040E097: call [004010E8h] ; rtcRound
  loc_0040E09D: mov ebx, [edi]
  loc_0040E09F: lea edx, var_28
  loc_0040E0A2: lea eax, var_14
  loc_0040E0A5: push edx
  loc_0040E0A6: push eax
  loc_0040E0A7: call [004010F8h] ; __vbaStrVarVal
  loc_0040E0AD: push eax
  loc_0040E0AE: push edi
  loc_0040E0AF: call [ebx+00000054h]
  loc_0040E0B2: test eax, eax
  loc_0040E0B4: fnclex
  loc_0040E0B6: jge 0040E0C7h
  loc_0040E0B8: push 00000054h
  loc_0040E0BA: push 00403858h
  loc_0040E0BF: push edi
  loc_0040E0C0: push eax
  loc_0040E0C1: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E0C7: lea ecx, var_14
  loc_0040E0CA: call [00401184h] ; __vbaFreeStr
  loc_0040E0D0: lea ecx, var_18
  loc_0040E0D3: call [00401180h] ; __vbaFreeObj
  loc_0040E0D9: lea ecx, var_28
  loc_0040E0DC: call [00401014h] ; __vbaFreeVar
  loc_0040E0E2: mov ecx, [esi]
  loc_0040E0E4: push esi
  loc_0040E0E5: call [ecx+000003B8h]
  loc_0040E0EB: lea edx, var_18
  loc_0040E0EE: push eax
  loc_0040E0EF: push edx
  loc_0040E0F0: call [0040105Ch] ; __vbaObjSet
  loc_0040E0F6: lea ecx, var_38
  loc_0040E0F9: mov edi, eax
  loc_0040E0FB: push 00000003h
  loc_0040E0FD: lea edx, var_28
  loc_0040E100: lea eax, [esi+00000040h]
  loc_0040E103: push ecx
  loc_0040E104: push edx
  loc_0040E105: mov var_30, eax
  loc_0040E108: mov var_38, 00004004h
  loc_0040E10F: call [004010E8h] ; rtcRound
  loc_0040E115: mov ebx, [edi]
  loc_0040E117: lea eax, var_28
  loc_0040E11A: lea ecx, var_14
  loc_0040E11D: push eax
  loc_0040E11E: push ecx
  loc_0040E11F: call [004010F8h] ; __vbaStrVarVal
  loc_0040E125: push eax
  loc_0040E126: push edi
  loc_0040E127: call [ebx+00000054h]
  loc_0040E12A: test eax, eax
  loc_0040E12C: fnclex
  loc_0040E12E: jge 0040E13Fh
  loc_0040E130: push 00000054h
  loc_0040E132: push 00403858h
  loc_0040E137: push edi
  loc_0040E138: push eax
  loc_0040E139: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E13F: lea ecx, var_14
  loc_0040E142: call [00401184h] ; __vbaFreeStr
  loc_0040E148: lea ecx, var_18
  loc_0040E14B: call [00401180h] ; __vbaFreeObj
  loc_0040E151: lea ecx, var_28
  loc_0040E154: call [00401014h] ; __vbaFreeVar
  loc_0040E15A: mov edx, [esi]
  loc_0040E15C: push esi
  loc_0040E15D: call [edx+000003C0h]
  loc_0040E163: push eax
  loc_0040E164: lea eax, var_18
  loc_0040E167: push eax
  loc_0040E168: call [0040105Ch] ; __vbaObjSet
  loc_0040E16E: mov edi, eax
  loc_0040E170: lea edx, var_38
  loc_0040E173: push 00000003h
  loc_0040E175: lea eax, var_28
  loc_0040E178: lea ecx, [esi+00000048h]
  loc_0040E17B: push edx
  loc_0040E17C: push eax
  loc_0040E17D: mov var_30, ecx
  loc_0040E180: mov var_38, 00004004h
  loc_0040E187: call [004010E8h] ; rtcRound
  loc_0040E18D: mov ebx, [edi]
  loc_0040E18F: lea ecx, var_28
  loc_0040E192: lea edx, var_14
  loc_0040E195: push ecx
  loc_0040E196: push edx
  loc_0040E197: call [004010F8h] ; __vbaStrVarVal
  loc_0040E19D: push eax
  loc_0040E19E: push edi
  loc_0040E19F: call [ebx+00000054h]
  loc_0040E1A2: test eax, eax
  loc_0040E1A4: fnclex
  loc_0040E1A6: jge 0040E1B7h
  loc_0040E1A8: push 00000054h
  loc_0040E1AA: push 00403858h
  loc_0040E1AF: push edi
  loc_0040E1B0: push eax
  loc_0040E1B1: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E1B7: mov ebx, [00401184h] ; __vbaFreeStr
  loc_0040E1BD: lea ecx, var_14
  loc_0040E1C0: call ebx
  loc_0040E1C2: lea ecx, var_18
  loc_0040E1C5: call [00401180h] ; __vbaFreeObj
  loc_0040E1CB: lea ecx, var_28
  loc_0040E1CE: call [00401014h] ; __vbaFreeVar
  loc_0040E1D4: mov eax, [esi]
  loc_0040E1D6: push esi
  loc_0040E1D7: call [eax+000003BCh]
  loc_0040E1DD: lea ecx, var_18
  loc_0040E1E0: push eax
  loc_0040E1E1: push ecx
  loc_0040E1E2: call [0040105Ch] ; __vbaObjSet
  loc_0040E1E8: mov edi, eax
  loc_0040E1EA: lea edx, var_38
  loc_0040E1ED: push 00000003h
  loc_0040E1EF: lea eax, var_28
  loc_0040E1F2: add esi, 0000004Ch
  loc_0040E1F5: push edx
  loc_0040E1F6: push eax
  loc_0040E1F7: mov var_30, esi
  loc_0040E1FA: mov var_38, 00004004h
  loc_0040E201: call [004010E8h] ; rtcRound
  loc_0040E207: mov esi, [edi]
  loc_0040E209: lea ecx, var_28
  loc_0040E20C: lea edx, var_14
  loc_0040E20F: push ecx
  loc_0040E210: push edx
  loc_0040E211: call [004010F8h] ; __vbaStrVarVal
  loc_0040E217: push eax
  loc_0040E218: push edi
  loc_0040E219: call [esi+00000054h]
  loc_0040E21C: test eax, eax
  loc_0040E21E: fnclex
  loc_0040E220: jge 0040E231h
  loc_0040E222: push 00000054h
  loc_0040E224: push 00403858h
  loc_0040E229: push edi
  loc_0040E22A: push eax
  loc_0040E22B: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E231: lea ecx, var_14
  loc_0040E234: call ebx
  loc_0040E236: lea ecx, var_18
  loc_0040E239: call [00401180h] ; __vbaFreeObj
  loc_0040E23F: lea ecx, var_28
  loc_0040E242: call [00401014h] ; __vbaFreeVar
  loc_0040E248: push 0040E26Ch
  loc_0040E24D: jmp 0040E26Bh
  loc_0040E24F: lea ecx, var_14
  loc_0040E252: call [00401184h] ; __vbaFreeStr
  loc_0040E258: lea ecx, var_18
  loc_0040E25B: call [00401180h] ; __vbaFreeObj
  loc_0040E261: lea ecx, var_28
  loc_0040E264: call [00401014h] ; __vbaFreeVar
  loc_0040E26A: ret
  loc_0040E26B: ret
  loc_0040E26C: mov ecx, var_10
  loc_0040E26F: pop edi
  loc_0040E270: pop esi
  loc_0040E271: xor eax, eax
  loc_0040E273: mov fs:[00000000h], ecx
  loc_0040E27A: pop ebx
  loc_0040E27B: mov esp, ebp
  loc_0040E27D: pop ebp
  loc_0040E27E: retn 0004h
End Sub

Private Sub Proc_3_33_40E300
  loc_0040E300: push ebp
  loc_0040E301: mov ebp, esp
  loc_0040E303: sub esp, 00000008h
  loc_0040E306: push 00401546h ; __vbaExceptHandler
  loc_0040E30B: mov eax, fs:[00000000h]
  loc_0040E311: push eax
  loc_0040E312: mov fs:[00000000h], esp
  loc_0040E319: sub esp, 00000014h
  loc_0040E31C: push ebx
  loc_0040E31D: push esi
  loc_0040E31E: push edi
  loc_0040E31F: mov var_8, esp
  loc_0040E322: mov var_4, 004012D0h
  loc_0040E329: mov edi, Me
  loc_0040E32C: xor ebx, ebx
  loc_0040E32E: push edi
  loc_0040E32F: mov var_14, ebx
  loc_0040E332: mov eax, [edi]
  loc_0040E334: mov var_18, ebx
  loc_0040E337: call [eax+000003DCh]
  loc_0040E33D: lea ecx, var_14
  loc_0040E340: push eax
  loc_0040E341: push ecx
  loc_0040E342: call [0040105Ch] ; __vbaObjSet
  loc_0040E348: mov esi, eax
  loc_0040E34A: lea eax, var_18
  loc_0040E34D: push eax
  loc_0040E34E: push esi
  loc_0040E34F: mov edx, [esi]
  loc_0040E351: call [edx+00000068h]
  loc_0040E354: cmp eax, ebx
  loc_0040E356: fnclex
  loc_0040E358: jge 0040E369h
  loc_0040E35A: push 00000068h
  loc_0040E35C: push 00403C38h
  loc_0040E361: push esi
  loc_0040E362: push eax
  loc_0040E363: call [00401040h] ; __vbaHresultCheckObj
  loc_0040E369: mov cx, var_18
  loc_0040E36D: mov [edi+0000006Eh], cx
  loc_0040E371: lea ecx, var_14
  loc_0040E374: call [00401180h] ; __vbaFreeObj
  loc_0040E37A: push 0040E38Ch
  loc_0040E37F: jmp 0040E38Bh
  loc_0040E381: lea ecx, var_14
  loc_0040E384: call [00401180h] ; __vbaFreeObj
  loc_0040E38A: ret
  loc_0040E38B: ret
  loc_0040E38C: mov ecx, var_10
  loc_0040E38F: pop edi
  loc_0040E390: pop esi
  loc_0040E391: xor eax, eax
  loc_0040E393: mov fs:[00000000h], ecx
  loc_0040E39A: pop ebx
  loc_0040E39B: mov esp, ebp
  loc_0040E39D: pop ebp
  loc_0040E39E: retn 0004h
End Sub

Private Sub Proc_3_34_4108F0
  loc_004108F0: push ebp
  loc_004108F1: mov ebp, esp
  loc_004108F3: sub esp, 00000008h
  loc_004108F6: push 00401546h ; __vbaExceptHandler
  loc_004108FB: mov eax, fs:[00000000h]
  loc_00410901: push eax
  loc_00410902: mov fs:[00000000h], esp
  loc_00410909: sub esp, 0000001Ch
  loc_0041090C: push ebx
  loc_0041090D: push esi
  loc_0041090E: push edi
  loc_0041090F: mov var_8, esp
  loc_00410912: mov var_4, 004013B0h
  loc_00410919: mov esi, Me
  loc_0041091C: xor eax, eax
  loc_0041091E: mov var_20, eax
  loc_00410921: mov var_28, eax
  loc_00410924: mov eax, [esi]
  loc_00410926: push esi
  loc_00410927: call [eax+000002FCh]
  loc_0041092D: lea ecx, var_28
  loc_00410930: push eax
  loc_00410931: push ecx
  loc_00410932: call [0040105Ch] ; __vbaObjSet
  loc_00410938: mov edx, [esi]
  loc_0041093A: lea eax, var_20
  loc_0041093D: lea ecx, [esi+00000040h]
  loc_00410940: push eax
  loc_00410941: push ecx
  loc_00410942: push esi
  loc_00410943: call [edx+00000778h]
  loc_00410949: mov edx, var_28
  loc_0041094C: mov edi, [00401104h] ; __vbaI2Var
  loc_00410952: lea eax, var_20
  loc_00410955: mov ebx, [edx]
  loc_00410957: push eax
  loc_00410958: call edi
  loc_0041095A: mov ecx, var_28
  loc_0041095D: push eax
  loc_0041095E: push ecx
  loc_0041095F: call [ebx+0000009Ch]
  loc_00410965: test eax, eax
  loc_00410967: fnclex
  loc_00410969: jge 00410980h
  loc_0041096B: mov edx, var_28
  loc_0041096E: push 0000009Ch
  loc_00410973: push 00403814h
  loc_00410978: push edx
  loc_00410979: push eax
  loc_0041097A: call [00401040h] ; __vbaHresultCheckObj
  loc_00410980: mov ebx, [00401014h] ; __vbaFreeVar
  loc_00410986: lea ecx, var_20
  loc_00410989: call ebx
  loc_0041098B: mov eax, [esi]
  loc_0041098D: lea ecx, var_20
  loc_00410990: lea edx, [esi+00000044h]
  loc_00410993: push ecx
  loc_00410994: push edx
  loc_00410995: push esi
  loc_00410996: call [eax+00000778h]
  loc_0041099C: mov eax, var_28
  loc_0041099F: lea ecx, var_20
  loc_004109A2: push ecx
  loc_004109A3: mov esi, [eax]
  loc_004109A5: call edi
  loc_004109A7: mov edx, var_28
  loc_004109AA: push eax
  loc_004109AB: push edx
  loc_004109AC: call [esi+000000A4h]
  loc_004109B2: test eax, eax
  loc_004109B4: fnclex
  loc_004109B6: jge 004109CDh
  loc_004109B8: mov ecx, var_28
  loc_004109BB: push 000000A4h
  loc_004109C0: push 00403814h
  loc_004109C5: push ecx
  loc_004109C6: push eax
  loc_004109C7: call [00401040h] ; __vbaHresultCheckObj
  loc_004109CD: lea ecx, var_20
  loc_004109D0: call ebx
  loc_004109D2: lea edx, var_28
  loc_004109D5: push 00000000h
  loc_004109D7: push edx
  loc_004109D8: call [0040106Ch] ; __vbaObjSetAddref
  loc_004109DE: push 004109F9h
  loc_004109E3: jmp 004109EFh
  loc_004109E5: lea ecx, var_20
  loc_004109E8: call [00401014h] ; __vbaFreeVar
  loc_004109EE: ret
  loc_004109EF: lea ecx, var_28
  loc_004109F2: call [00401180h] ; __vbaFreeObj
  loc_004109F8: ret
  loc_004109F9: mov ecx, var_10
  loc_004109FC: pop edi
  loc_004109FD: pop esi
  loc_004109FE: xor eax, eax
  loc_00410A00: mov fs:[00000000h], ecx
  loc_00410A07: pop ebx
  loc_00410A08: mov esp, ebp
  loc_00410A0A: pop ebp
  loc_00410A0B: retn 0004h
End Sub

Private Sub Proc_3_35_410A10
  loc_00410A10: push ebp
  loc_00410A11: mov ebp, esp
  loc_00410A13: sub esp, 00000008h
  loc_00410A16: push 00401546h ; __vbaExceptHandler
  loc_00410A1B: mov eax, fs:[00000000h]
  loc_00410A21: push eax
  loc_00410A22: mov fs:[00000000h], esp
  loc_00410A29: sub esp, 00000048h
  loc_00410A2C: push ebx
  loc_00410A2D: push esi
  loc_00410A2E: push edi
  loc_00410A2F: mov var_8, esp
  loc_00410A32: mov var_4, 004013C0h
  loc_00410A39: mov esi, Me
  loc_00410A3C: xor edi, edi
  loc_00410A3E: push esi
  loc_00410A3F: mov var_20, edi
  loc_00410A42: mov eax, [esi]
  loc_00410A44: mov var_30, edi
  loc_00410A47: mov var_40, edi
  loc_00410A4A: mov var_44, edi
  loc_00410A4D: mov var_54, edi
  loc_00410A50: call [eax+00000310h]
  loc_00410A56: lea ecx, var_54
  loc_00410A59: push eax
  loc_00410A5A: push ecx
  loc_00410A5B: call [0040105Ch] ; __vbaObjSet
  loc_00410A61: mov eax, var_54
  loc_00410A64: push FFFFFFFFh
  loc_00410A66: push eax
  loc_00410A67: mov edx, [eax]
  loc_00410A69: call [edx+00000084h]
  loc_00410A6F: cmp eax, edi
  loc_00410A71: fnclex
  loc_00410A73: jge 00410A8Ah
  loc_00410A75: mov ecx, var_54
  loc_00410A78: push 00000084h
  loc_00410A7D: push 00403C48h
  loc_00410A82: push ecx
  loc_00410A83: push eax
  loc_00410A84: call [00401040h] ; __vbaHresultCheckObj
  loc_00410A8A: mov edx, [esi]
  loc_00410A8C: lea eax, var_20
  loc_00410A8F: lea edi, [esi+00000040h]
  loc_00410A92: push eax
  loc_00410A93: push edi
  loc_00410A94: push esi
  loc_00410A95: call [edx+00000778h]
  loc_00410A9B: mov ecx, var_54
  loc_00410A9E: lea edx, var_20
  loc_00410AA1: push edx
  loc_00410AA2: mov ebx, [ecx]
  loc_00410AA4: call [004010B0h] ; __vbaR4Var
  loc_00410AAA: mov eax, var_54
  loc_00410AAD: push ecx
  loc_00410AAE: fstp real4 ptr [esp]
  loc_00410AB1: push eax
  loc_00410AB2: call [ebx+00000064h]
  loc_00410AB5: test eax, eax
  loc_00410AB7: fnclex
  loc_00410AB9: jge 00410ACDh
  loc_00410ABB: mov ecx, var_54
  loc_00410ABE: push 00000064h
  loc_00410AC0: push 00403C48h
  loc_00410AC5: push ecx
  loc_00410AC6: push eax
  loc_00410AC7: call [00401040h] ; __vbaHresultCheckObj
  loc_00410ACD: lea ecx, var_20
  loc_00410AD0: call [00401014h] ; __vbaFreeVar
  loc_00410AD6: call 0040B600h
  loc_00410ADB: fstp real4 ptr var_48
  loc_00410ADE: call 0040B5E0h
  loc_00410AE3: fstp real4 ptr var_4C
  loc_00410AE6: fld real4 ptr var_4C
  loc_00410AE9: fmul st0, real4 ptr [edi]
  loc_00410AEB: mov edx, [esi]
  loc_00410AED: lea ecx, var_44
  loc_00410AF0: fadd st0, real4 ptr var_48
  loc_00410AF3: fstp real4 ptr var_44
  loc_00410AF6: fnstsw ax
  loc_00410AF8: test al, 0Dh
  loc_00410AFA: jnz 00410C99h
  loc_00410B00: lea eax, var_20
  loc_00410B03: push eax
  loc_00410B04: push ecx
  loc_00410B05: push esi
  loc_00410B06: call [edx+0000077Ch]
  loc_00410B0C: mov edx, var_54
  loc_00410B0F: mov var_38, 0000006Eh
  loc_00410B16: mov var_40, 00000002h
  loc_00410B1D: lea eax, var_20
  loc_00410B20: mov edi, [edx]
  loc_00410B22: lea ecx, var_40
  loc_00410B25: push eax
  loc_00410B26: lea edx, var_30
  loc_00410B29: push ecx
  loc_00410B2A: push edx
  loc_00410B2B: call [00401144h] ; __vbaVarAdd
  loc_00410B31: push eax
  loc_00410B32: call [004010B0h] ; __vbaR4Var
  loc_00410B38: mov eax, var_54
  loc_00410B3B: push ecx
  loc_00410B3C: fstp real4 ptr [esp]
  loc_00410B3F: push eax
  loc_00410B40: call [edi+0000006Ch]
  loc_00410B43: test eax, eax
  loc_00410B45: fnclex
  loc_00410B47: jge 00410B5Bh
  loc_00410B49: mov ecx, var_54
  loc_00410B4C: push 0000006Ch
  loc_00410B4E: push 00403C48h
  loc_00410B53: push ecx
  loc_00410B54: push eax
  loc_00410B55: call [00401040h] ; __vbaHresultCheckObj
  loc_00410B5B: lea edx, var_30
  loc_00410B5E: lea eax, var_20
  loc_00410B61: push edx
  loc_00410B62: push eax
  loc_00410B63: push 00000002h
  loc_00410B65: call [00401024h] ; __vbaFreeVarList
  loc_00410B6B: mov ecx, [esi]
  loc_00410B6D: add esp, 0000000Ch
  loc_00410B70: lea edx, var_20
  loc_00410B73: lea edi, [esi+00000044h]
  loc_00410B76: push edx
  loc_00410B77: push edi
  loc_00410B78: push esi
  loc_00410B79: call [ecx+00000778h]
  loc_00410B7F: mov eax, var_54
  loc_00410B82: lea ecx, var_20
  loc_00410B85: push ecx
  loc_00410B86: mov ebx, [eax]
  loc_00410B88: call [004010B0h] ; __vbaR4Var
  loc_00410B8E: mov edx, var_54
  loc_00410B91: push ecx
  loc_00410B92: fstp real4 ptr [esp]
  loc_00410B95: push edx
  loc_00410B96: call [ebx+00000074h]
  loc_00410B99: test eax, eax
  loc_00410B9B: fnclex
  loc_00410B9D: jge 00410BB1h
  loc_00410B9F: mov ecx, var_54
  loc_00410BA2: push 00000074h
  loc_00410BA4: push 00403C48h
  loc_00410BA9: push ecx
  loc_00410BAA: push eax
  loc_00410BAB: call [00401040h] ; __vbaHresultCheckObj
  loc_00410BB1: lea ecx, var_20
  loc_00410BB4: call [00401014h] ; __vbaFreeVar
  loc_00410BBA: call 0040B600h
  loc_00410BBF: fstp real4 ptr var_48
  loc_00410BC2: call 0040B5E0h
  loc_00410BC7: fstp real4 ptr var_4C
  loc_00410BCA: fld real4 ptr var_4C
  loc_00410BCD: fmul st0, real4 ptr [edi]
  loc_00410BCF: mov edx, [esi]
  loc_00410BD1: lea ecx, var_44
  loc_00410BD4: fadd st0, real4 ptr var_48
  loc_00410BD7: fstp real4 ptr var_44
  loc_00410BDA: fnstsw ax
  loc_00410BDC: test al, 0Dh
  loc_00410BDE: jnz 00410C99h
  loc_00410BE4: lea eax, var_20
  loc_00410BE7: push eax
  loc_00410BE8: push ecx
  loc_00410BE9: push esi
  loc_00410BEA: call [edx+0000077Ch]
  loc_00410BF0: mov edx, var_54
  loc_00410BF3: mov var_38, 0000006Eh
  loc_00410BFA: mov var_40, 00000002h
  loc_00410C01: lea eax, var_20
  loc_00410C04: mov esi, [edx]
  loc_00410C06: lea ecx, var_40
  loc_00410C09: push eax
  loc_00410C0A: lea edx, var_30
  loc_00410C0D: push ecx
  loc_00410C0E: push edx
  loc_00410C0F: call [00401144h] ; __vbaVarAdd
  loc_00410C15: push eax
  loc_00410C16: call [004010B0h] ; __vbaR4Var
  loc_00410C1C: mov eax, var_54
  loc_00410C1F: push ecx
  loc_00410C20: fstp real4 ptr [esp]
  loc_00410C23: push eax
  loc_00410C24: call [esi+0000007Ch]
  loc_00410C27: test eax, eax
  loc_00410C29: fnclex
  loc_00410C2B: jge 00410C3Fh
  loc_00410C2D: mov ecx, var_54
  loc_00410C30: push 0000007Ch
  loc_00410C32: push 00403C48h
  loc_00410C37: push ecx
  loc_00410C38: push eax
  loc_00410C39: call [00401040h] ; __vbaHresultCheckObj
  loc_00410C3F: lea edx, var_30
  loc_00410C42: lea eax, var_20
  loc_00410C45: push edx
  loc_00410C46: push eax
  loc_00410C47: push 00000002h
  loc_00410C49: call [00401024h] ; __vbaFreeVarList
  loc_00410C4F: add esp, 0000000Ch
  loc_00410C52: lea ecx, var_54
  loc_00410C55: push 00000000h
  loc_00410C57: push ecx
  loc_00410C58: call [0040106Ch] ; __vbaObjSetAddref
  loc_00410C5E: fwait
  loc_00410C5F: push 00410C84h
  loc_00410C64: jmp 00410C7Ah
  loc_00410C66: lea edx, var_30
  loc_00410C69: lea eax, var_20
  loc_00410C6C: push edx
  loc_00410C6D: push eax
  loc_00410C6E: push 00000002h
  loc_00410C70: call [00401024h] ; __vbaFreeVarList
  loc_00410C76: add esp, 0000000Ch
  loc_00410C79: ret
  loc_00410C7A: lea ecx, var_54
  loc_00410C7D: call [00401180h] ; __vbaFreeObj
  loc_00410C83: ret
  loc_00410C84: mov ecx, var_10
  loc_00410C87: pop edi
  loc_00410C88: pop esi
  loc_00410C89: xor eax, eax
  loc_00410C8B: mov fs:[00000000h], ecx
  loc_00410C92: pop ebx
  loc_00410C93: mov esp, ebp
  loc_00410C95: pop ebp
  loc_00410C96: retn 0004h
End Sub

Private Sub Proc_3_36_410CA0(arg_C, arg_10) '410CA0
  loc_00410CA0: push ebp
  loc_00410CA1: mov ebp, esp
  loc_00410CA3: sub esp, 0000000Ch
  loc_00410CA6: push 00401546h ; __vbaExceptHandler
  loc_00410CAB: mov eax, fs:[00000000h]
  loc_00410CB1: push eax
  loc_00410CB2: mov fs:[00000000h], esp
  loc_00410CB9: sub esp, 00000038h
  loc_00410CBC: push ebx
  loc_00410CBD: push esi
  loc_00410CBE: push edi
  loc_00410CBF: mov var_C, esp
  loc_00410CC2: mov var_8, 004013D8h
  loc_00410CC9: mov ecx, arg_10
  loc_00410CCC: mov edx, arg_C
  loc_00410CCF: xor eax, eax
  loc_00410CD1: mov esi, [0040100Ch] ; __vbaVarMove
  loc_00410CD7: mov var_24, eax
  loc_00410CDA: mov var_34, eax
  loc_00410CDD: mov [ecx], eax
  loc_00410CDF: mov eax, Me
  loc_00410CE2: lea ecx, var_24
  loc_00410CE5: mov var_44, 00000004h
  loc_00410CEC: fld real4 ptr [eax+00000058h]
  loc_00410CEF: fmul st0, real4 ptr [edx]
  loc_00410CF1: lea edx, var_44
  loc_00410CF4: fadd st0, real4 ptr [eax+0000005Ch]
  loc_00410CF7: fadd st0, real4 ptr [004013D0h]
  loc_00410CFD: fstp real4 ptr var_3C
  loc_00410D00: fnstsw ax
  loc_00410D02: test al, 0Dh
  loc_00410D04: jnz 00410D71h
  loc_00410D06: call __vbaVarMove
  loc_00410D08: lea eax, var_24
  loc_00410D0B: push 00000002h
  loc_00410D0D: lea ecx, var_34
  loc_00410D10: push eax
  loc_00410D11: push ecx
  loc_00410D12: call [004010E8h] ; rtcRound
  loc_00410D18: lea edx, var_34
  loc_00410D1B: lea ecx, var_24
  loc_00410D1E: call __vbaVarMove
  loc_00410D20: fwait
  loc_00410D21: push 00410D42h
  loc_00410D26: jmp 00410D41h
  loc_00410D28: test var_4, 04h
  loc_00410D2C: jz 00410D37h
  loc_00410D2E: lea ecx, var_24
  loc_00410D31: call [00401014h] ; __vbaFreeVar
  loc_00410D37: lea ecx, var_34
  loc_00410D3A: call [00401014h] ; __vbaFreeVar
  loc_00410D40: ret
  loc_00410D41: ret
  loc_00410D42: mov edx, arg_10
  loc_00410D45: mov eax, var_24
  loc_00410D48: mov ecx, var_20
  loc_00410D4B: pop edi
  loc_00410D4C: mov [edx], eax
  loc_00410D4E: mov eax, var_1C
  loc_00410D51: pop esi
  loc_00410D52: pop ebx
  loc_00410D53: mov [edx+00000004h], ecx
  loc_00410D56: mov ecx, var_18
  loc_00410D59: mov [edx+00000008h], eax
  loc_00410D5C: xor eax, eax
  loc_00410D5E: mov [edx+0000000Ch], ecx
  loc_00410D61: mov ecx, var_14
  loc_00410D64: mov fs:[00000000h], ecx
  loc_00410D6B: mov esp, ebp
  loc_00410D6D: pop ebp
  loc_00410D6E: retn 000Ch
End Sub

Private Sub Proc_3_37_410D80(arg_C, arg_10) '410D80
  loc_00410D80: push ebp
  loc_00410D81: mov ebp, esp
  loc_00410D83: sub esp, 0000000Ch
  loc_00410D86: push 00401546h ; __vbaExceptHandler
  loc_00410D8B: mov eax, fs:[00000000h]
  loc_00410D91: push eax
  loc_00410D92: mov fs:[00000000h], esp
  loc_00410D99: sub esp, 00000038h
  loc_00410D9C: push ebx
  loc_00410D9D: push esi
  loc_00410D9E: push edi
  loc_00410D9F: mov var_C, esp
  loc_00410DA2: mov var_8, 004013E8h
  loc_00410DA9: mov ecx, arg_10
  loc_00410DAC: mov edx, arg_C
  loc_00410DAF: xor eax, eax
  loc_00410DB1: mov esi, [0040100Ch] ; __vbaVarMove
  loc_00410DB7: mov var_24, eax
  loc_00410DBA: mov var_34, eax
  loc_00410DBD: mov [ecx], eax
  loc_00410DBF: mov eax, Me
  loc_00410DC2: lea ecx, var_24
  loc_00410DC5: mov var_44, 00000004h
  loc_00410DCC: fld real4 ptr [eax+00000060h]
  loc_00410DCF: fmul st0, real4 ptr [edx]
  loc_00410DD1: lea edx, var_44
  loc_00410DD4: fadd st0, real4 ptr [eax+00000064h]
  loc_00410DD7: fsub st0, real4 ptr [004013D0h]
  loc_00410DDD: fstp real4 ptr var_3C
  loc_00410DE0: fnstsw ax
  loc_00410DE2: test al, 0Dh
  loc_00410DE4: jnz 00410E51h
  loc_00410DE6: call __vbaVarMove
  loc_00410DE8: lea eax, var_24
  loc_00410DEB: push 00000002h
  loc_00410DED: lea ecx, var_34
  loc_00410DF0: push eax
  loc_00410DF1: push ecx
  loc_00410DF2: call [004010E8h] ; rtcRound
  loc_00410DF8: lea edx, var_34
  loc_00410DFB: lea ecx, var_24
  loc_00410DFE: call __vbaVarMove
  loc_00410E00: fwait
  loc_00410E01: push 00410E22h
  loc_00410E06: jmp 00410E21h
  loc_00410E08: test var_4, 04h
  loc_00410E0C: jz 00410E17h
  loc_00410E0E: lea ecx, var_24
  loc_00410E11: call [00401014h] ; __vbaFreeVar
  loc_00410E17: lea ecx, var_34
  loc_00410E1A: call [00401014h] ; __vbaFreeVar
  loc_00410E20: ret
  loc_00410E21: ret
  loc_00410E22: mov edx, arg_10
  loc_00410E25: mov eax, var_24
  loc_00410E28: mov ecx, var_20
  loc_00410E2B: pop edi
  loc_00410E2C: mov [edx], eax
  loc_00410E2E: mov eax, var_1C
  loc_00410E31: pop esi
  loc_00410E32: pop ebx
  loc_00410E33: mov [edx+00000004h], ecx
  loc_00410E36: mov ecx, var_18
  loc_00410E39: mov [edx+00000008h], eax
  loc_00410E3C: xor eax, eax
  loc_00410E3E: mov [edx+0000000Ch], ecx
  loc_00410E41: mov ecx, var_14
  loc_00410E44: mov fs:[00000000h], ecx
  loc_00410E4B: mov esp, ebp
  loc_00410E4D: pop ebp
  loc_00410E4E: retn 000Ch
End Sub

Private Sub Proc_3_38_410E60
  loc_00410E60: push ebp
  loc_00410E61: mov ebp, esp
  loc_00410E63: sub esp, 00000008h
  loc_00410E66: push 00401546h ; __vbaExceptHandler
  loc_00410E6B: mov eax, fs:[00000000h]
  loc_00410E71: push eax
  loc_00410E72: mov fs:[00000000h], esp
  loc_00410E79: sub esp, 00000098h
  loc_00410E7F: push ebx
  loc_00410E80: push esi
  loc_00410E81: push edi
  loc_00410E82: mov var_8, esp
  loc_00410E85: mov var_4, 004013F8h
  loc_00410E8C: xor eax, eax
  loc_00410E8E: mov var_14, eax
  loc_00410E91: mov var_18, eax
  loc_00410E94: mov var_1C, eax
  loc_00410E97: mov var_20, eax
  loc_00410E9A: mov var_24, eax
  loc_00410E9D: mov var_34, eax
  loc_00410EA0: mov var_44, eax
  loc_00410EA3: mov var_54, eax
  loc_00410EA6: mov var_64, eax
  loc_00410EA9: call 0040B5E0h
  loc_00410EAE: mov esi, Me
  loc_00410EB1: fstp real4 ptr var_68
  loc_00410EB4: mov eax, [esi]
  loc_00410EB6: push esi
  loc_00410EB7: call [eax+0000038Ch]
  loc_00410EBD: mov edi, [0040105Ch] ; __vbaObjSet
  loc_00410EC3: lea ecx, var_24
  loc_00410EC6: push eax
  loc_00410EC7: push ecx
  loc_00410EC8: call edi
  loc_00410ECA: mov edx, var_68
  loc_00410ECD: mov ebx, eax
  loc_00410ECF: lea eax, var_34
  loc_00410ED2: push 00000002h
  loc_00410ED4: lea ecx, var_44
  loc_00410ED7: push eax
  loc_00410ED8: push ecx
  loc_00410ED9: mov var_6C, ebx
  loc_00410EDC: mov var_2C, edx
  loc_00410EDF: mov var_34, 00000004h
  loc_00410EE6: call [004010E8h] ; rtcRound
  loc_00410EEC: mov ebx, [ebx]
  loc_00410EEE: lea edx, var_44
  loc_00410EF1: lea eax, var_14
  loc_00410EF4: push edx
  loc_00410EF5: push eax
  loc_00410EF6: call [004010F8h] ; __vbaStrVarVal
  loc_00410EFC: mov ecx, ebx
  loc_00410EFE: mov ebx, var_6C
  loc_00410F01: push eax
  loc_00410F02: push ebx
  loc_00410F03: call [ecx+00000054h]
  loc_00410F06: test eax, eax
  loc_00410F08: fnclex
  loc_00410F0A: jge 00410F1Bh
  loc_00410F0C: push 00000054h
  loc_00410F0E: push 00403858h
  loc_00410F13: push ebx
  loc_00410F14: push eax
  loc_00410F15: call [00401040h] ; __vbaHresultCheckObj
  loc_00410F1B: lea ecx, var_14
  loc_00410F1E: call [00401184h] ; __vbaFreeStr
  loc_00410F24: lea ecx, var_24
  loc_00410F27: call [00401180h] ; __vbaFreeObj
  loc_00410F2D: lea edx, var_44
  loc_00410F30: lea eax, var_34
  loc_00410F33: push edx
  loc_00410F34: push eax
  loc_00410F35: push 00000002h
  loc_00410F37: call [00401024h] ; __vbaFreeVarList
  loc_00410F3D: add esp, 0000000Ch
  loc_00410F40: call 0040B600h
  loc_00410F45: mov ecx, [esi]
  loc_00410F47: push esi
  loc_00410F48: fstp real4 ptr var_68
  loc_00410F4B: call [ecx+00000390h]
  loc_00410F51: lea edx, var_24
  loc_00410F54: push eax
  loc_00410F55: push edx
  loc_00410F56: call edi
  loc_00410F58: lea ecx, var_34
  loc_00410F5B: mov ebx, eax
  loc_00410F5D: mov eax, var_68
  loc_00410F60: push 00000002h
  loc_00410F62: lea edx, var_44
  loc_00410F65: push ecx
  loc_00410F66: push edx
  loc_00410F67: mov var_6C, ebx
  loc_00410F6A: mov var_2C, eax
  loc_00410F6D: mov var_34, 00000004h
  loc_00410F74: call [004010E8h] ; rtcRound
  loc_00410F7A: mov ebx, [ebx]
  loc_00410F7C: lea eax, var_44
  loc_00410F7F: lea ecx, var_14
  loc_00410F82: push eax
  loc_00410F83: push ecx
  loc_00410F84: call [004010F8h] ; __vbaStrVarVal
  loc_00410F8A: mov edx, ebx
  loc_00410F8C: mov ebx, var_6C
  loc_00410F8F: push eax
  loc_00410F90: push ebx
  loc_00410F91: call [edx+00000054h]
  loc_00410F94: test eax, eax
  loc_00410F96: fnclex
  loc_00410F98: jge 00410FA9h
  loc_00410F9A: push 00000054h
  loc_00410F9C: push 00403858h
  loc_00410FA1: push ebx
  loc_00410FA2: push eax
  loc_00410FA3: call [00401040h] ; __vbaHresultCheckObj
  loc_00410FA9: lea ecx, var_14
  loc_00410FAC: call [00401184h] ; __vbaFreeStr
  loc_00410FB2: lea ecx, var_24
  loc_00410FB5: call [00401180h] ; __vbaFreeObj
  loc_00410FBB: lea eax, var_44
  loc_00410FBE: lea ecx, var_34
  loc_00410FC1: push eax
  loc_00410FC2: push ecx
  loc_00410FC3: push 00000002h
  loc_00410FC5: call [00401024h] ; __vbaFreeVarList
  loc_00410FCB: add esp, 0000000Ch
  loc_00410FCE: call 0040B620h
  loc_00410FD3: mov edx, [esi]
  loc_00410FD5: push esi
  loc_00410FD6: fstp real4 ptr var_68
  loc_00410FD9: call [edx+00000388h]
  loc_00410FDF: push eax
  loc_00410FE0: lea eax, var_24
  loc_00410FE3: push eax
  loc_00410FE4: call edi
  loc_00410FE6: mov ecx, var_68
  loc_00410FE9: mov ebx, eax
  loc_00410FEB: lea edx, var_34
  loc_00410FEE: push 00000002h
  loc_00410FF0: lea eax, var_44
  loc_00410FF3: push edx
  loc_00410FF4: push eax
  loc_00410FF5: mov var_6C, ebx
  loc_00410FF8: mov var_2C, ecx
  loc_00410FFB: mov var_34, 00000004h
  loc_00411002: call [004010E8h] ; rtcRound
  loc_00411008: mov ebx, [ebx]
  loc_0041100A: lea ecx, var_44
  loc_0041100D: lea edx, var_14
  loc_00411010: push ecx
  loc_00411011: push edx
  loc_00411012: call [004010F8h] ; __vbaStrVarVal
  loc_00411018: mov var_80, ebx
  loc_0041101B: mov ebx, var_6C
  loc_0041101E: push eax
  loc_0041101F: mov eax, var_80
  loc_00411022: push ebx
  loc_00411023: call [eax+00000054h]
  loc_00411026: test eax, eax
  loc_00411028: fnclex
  loc_0041102A: jge 0041103Bh
  loc_0041102C: push 00000054h
  loc_0041102E: push 00403858h
  loc_00411033: push ebx
  loc_00411034: push eax
  loc_00411035: call [00401040h] ; __vbaHresultCheckObj
  loc_0041103B: lea ecx, var_14
  loc_0041103E: call [00401184h] ; __vbaFreeStr
  loc_00411044: lea ecx, var_24
  loc_00411047: call [00401180h] ; __vbaFreeObj
  loc_0041104D: lea ecx, var_44
  loc_00411050: lea edx, var_34
  loc_00411053: push ecx
  loc_00411054: push edx
  loc_00411055: push 00000002h
  loc_00411057: call [00401024h] ; __vbaFreeVarList
  loc_0041105D: add esp, 0000000Ch
  loc_00411060: call 0040B640h
  loc_00411065: mov eax, [esi]
  loc_00411067: push esi
  loc_00411068: fstp real4 ptr var_68
  loc_0041106B: call [eax+00000380h]
  loc_00411071: lea ecx, var_24
  loc_00411074: push eax
  loc_00411075: push ecx
  loc_00411076: call edi
  loc_00411078: mov edx, var_68
  loc_0041107B: mov ebx, eax
  loc_0041107D: lea eax, var_34
  loc_00411080: push 00000002h
  loc_00411082: lea ecx, var_44
  loc_00411085: push eax
  loc_00411086: push ecx
  loc_00411087: mov var_6C, ebx
  loc_0041108A: mov var_2C, edx
  loc_0041108D: mov var_34, 00000004h
  loc_00411094: call [004010E8h] ; rtcRound
  loc_0041109A: mov ebx, [ebx]
  loc_0041109C: lea edx, var_44
  loc_0041109F: lea eax, var_14
  loc_004110A2: push edx
  loc_004110A3: push eax
  loc_004110A4: call [004010F8h] ; __vbaStrVarVal
  loc_004110AA: mov ecx, ebx
  loc_004110AC: mov ebx, var_6C
  loc_004110AF: push eax
  loc_004110B0: push ebx
  loc_004110B1: call [ecx+00000054h]
  loc_004110B4: test eax, eax
  loc_004110B6: fnclex
  loc_004110B8: jge 004110C9h
  loc_004110BA: push 00000054h
  loc_004110BC: push 00403858h
  loc_004110C1: push ebx
  loc_004110C2: push eax
  loc_004110C3: call [00401040h] ; __vbaHresultCheckObj
  loc_004110C9: lea ecx, var_14
  loc_004110CC: call [00401184h] ; __vbaFreeStr
  loc_004110D2: lea ecx, var_24
  loc_004110D5: call [00401180h] ; __vbaFreeObj
  loc_004110DB: lea edx, var_44
  loc_004110DE: lea eax, var_34
  loc_004110E1: push edx
  loc_004110E2: push eax
  loc_004110E3: push 00000002h
  loc_004110E5: call [00401024h] ; __vbaFreeVarList
  loc_004110EB: add esp, 0000000Ch
  loc_004110EE: call 0040B760h
  loc_004110F3: mov ecx, [esi]
  loc_004110F5: push esi
  loc_004110F6: fstp real4 ptr var_68
  loc_004110F9: call [ecx+0000036Ch]
  loc_004110FF: lea edx, var_24
  loc_00411102: push eax
  loc_00411103: push edx
  loc_00411104: call edi
  loc_00411106: lea ecx, var_34
  loc_00411109: mov ebx, eax
  loc_0041110B: mov eax, var_68
  loc_0041110E: push 00000002h
  loc_00411110: lea edx, var_44
  loc_00411113: push ecx
  loc_00411114: push edx
  loc_00411115: mov var_6C, ebx
  loc_00411118: mov var_2C, eax
  loc_0041111B: mov var_34, 00000004h
  loc_00411122: call [004010E8h] ; rtcRound
  loc_00411128: mov ebx, [ebx]
  loc_0041112A: lea eax, var_44
  loc_0041112D: lea ecx, var_14
  loc_00411130: push eax
  loc_00411131: push ecx
  loc_00411132: call [004010F8h] ; __vbaStrVarVal
  loc_00411138: mov edx, ebx
  loc_0041113A: mov ebx, var_6C
  loc_0041113D: push eax
  loc_0041113E: push ebx
  loc_0041113F: call [edx+00000054h]
  loc_00411142: test eax, eax
  loc_00411144: fnclex
  loc_00411146: jge 00411157h
  loc_00411148: push 00000054h
  loc_0041114A: push 00403858h
  loc_0041114F: push ebx
  loc_00411150: push eax
  loc_00411151: call [00401040h] ; __vbaHresultCheckObj
  loc_00411157: lea ecx, var_14
  loc_0041115A: call [00401184h] ; __vbaFreeStr
  loc_00411160: lea ecx, var_24
  loc_00411163: call [00401180h] ; __vbaFreeObj
  loc_00411169: lea eax, var_44
  loc_0041116C: lea ecx, var_34
  loc_0041116F: push eax
  loc_00411170: push ecx
  loc_00411171: push 00000002h
  loc_00411173: call [00401024h] ; __vbaFreeVarList
  loc_00411179: add esp, 0000000Ch
  loc_0041117C: call 0040B6A0h
  loc_00411181: mov edx, [esi]
  loc_00411183: push esi
  loc_00411184: fstp real4 ptr var_68
  loc_00411187: call [edx+0000034Ch]
  loc_0041118D: push eax
  loc_0041118E: lea eax, var_24
  loc_00411191: push eax
  loc_00411192: call edi
  loc_00411194: mov ecx, var_68
  loc_00411197: mov ebx, eax
  loc_00411199: lea edx, var_34
  loc_0041119C: push 00000002h
  loc_0041119E: lea eax, var_44
  loc_004111A1: push edx
  loc_004111A2: push eax
  loc_004111A3: mov var_6C, ebx
  loc_004111A6: mov var_2C, ecx
  loc_004111A9: mov var_34, 00000004h
  loc_004111B0: call [004010E8h] ; rtcRound
  loc_004111B6: mov ebx, [ebx]
  loc_004111B8: lea ecx, var_44
  loc_004111BB: lea edx, var_14
  loc_004111BE: push ecx
  loc_004111BF: push edx
  loc_004111C0: call [004010F8h] ; __vbaStrVarVal
  loc_004111C6: mov var_8C, ebx
  loc_004111CC: mov ebx, var_6C
  loc_004111CF: push eax
  loc_004111D0: mov eax, var_8C
  loc_004111D6: push ebx
  loc_004111D7: call [eax+00000054h]
  loc_004111DA: test eax, eax
  loc_004111DC: fnclex
  loc_004111DE: jge 004111EFh
  loc_004111E0: push 00000054h
  loc_004111E2: push 00403858h
  loc_004111E7: push ebx
  loc_004111E8: push eax
  loc_004111E9: call [00401040h] ; __vbaHresultCheckObj
  loc_004111EF: lea ecx, var_14
  loc_004111F2: call [00401184h] ; __vbaFreeStr
  loc_004111F8: lea ecx, var_24
  loc_004111FB: call [00401180h] ; __vbaFreeObj
  loc_00411201: lea ecx, var_44
  loc_00411204: lea edx, var_34
  loc_00411207: push ecx
  loc_00411208: push edx
  loc_00411209: push 00000002h
  loc_0041120B: call [00401024h] ; __vbaFreeVarList
  loc_00411211: add esp, 0000000Ch
  loc_00411214: call 0040B6C0h
  loc_00411219: mov eax, [esi]
  loc_0041121B: push esi
  loc_0041121C: fstp real4 ptr var_68
  loc_0041121F: call [eax+00000348h]
  loc_00411225: lea ecx, var_24
  loc_00411228: push eax
  loc_00411229: push ecx
  loc_0041122A: call edi
  loc_0041122C: mov edx, var_68
  loc_0041122F: mov ebx, eax
  loc_00411231: lea eax, var_34
  loc_00411234: push 00000002h
  loc_00411236: lea ecx, var_44
  loc_00411239: push eax
  loc_0041123A: push ecx
  loc_0041123B: mov var_6C, ebx
  loc_0041123E: mov var_2C, edx
  loc_00411241: mov var_34, 00000004h
  loc_00411248: call [004010E8h] ; rtcRound
  loc_0041124E: mov ebx, [ebx]
  loc_00411250: lea edx, var_44
  loc_00411253: lea eax, var_14
  loc_00411256: push edx
  loc_00411257: push eax
  loc_00411258: call [004010F8h] ; __vbaStrVarVal
  loc_0041125E: mov ecx, ebx
  loc_00411260: mov ebx, var_6C
  loc_00411263: push eax
  loc_00411264: push ebx
  loc_00411265: call [ecx+00000054h]
  loc_00411268: test eax, eax
  loc_0041126A: fnclex
  loc_0041126C: jge 0041127Dh
  loc_0041126E: push 00000054h
  loc_00411270: push 00403858h
  loc_00411275: push ebx
  loc_00411276: push eax
  loc_00411277: call [00401040h] ; __vbaHresultCheckObj
  loc_0041127D: lea ecx, var_14
  loc_00411280: call [00401184h] ; __vbaFreeStr
  loc_00411286: lea ecx, var_24
  loc_00411289: call [00401180h] ; __vbaFreeObj
  loc_0041128F: lea edx, var_44
  loc_00411292: lea eax, var_34
  loc_00411295: push edx
  loc_00411296: push eax
  loc_00411297: push 00000002h
  loc_00411299: call [00401024h] ; __vbaFreeVarList
  loc_0041129F: add esp, 0000000Ch
  loc_004112A2: call 0040B700h
  loc_004112A7: mov ecx, [esi]
  loc_004112A9: push esi
  loc_004112AA: fstp real4 ptr var_68
  loc_004112AD: call [ecx+00000344h]
  loc_004112B3: lea edx, var_24
  loc_004112B6: push eax
  loc_004112B7: push edx
  loc_004112B8: call edi
  loc_004112BA: lea ecx, var_34
  loc_004112BD: mov ebx, eax
  loc_004112BF: mov eax, var_68
  loc_004112C2: push 00000002h
  loc_004112C4: lea edx, var_44
  loc_004112C7: push ecx
  loc_004112C8: push edx
  loc_004112C9: mov var_6C, ebx
  loc_004112CC: mov var_2C, eax
  loc_004112CF: mov var_34, 00000004h
  loc_004112D6: call [004010E8h] ; rtcRound
  loc_004112DC: mov ebx, [ebx]
  loc_004112DE: lea eax, var_44
  loc_004112E1: lea ecx, var_14
  loc_004112E4: push eax
  loc_004112E5: push ecx
  loc_004112E6: call [004010F8h] ; __vbaStrVarVal
  loc_004112EC: mov edx, ebx
  loc_004112EE: mov ebx, var_6C
  loc_004112F1: push eax
  loc_004112F2: push ebx
  loc_004112F3: call [edx+00000054h]
  loc_004112F6: test eax, eax
  loc_004112F8: fnclex
  loc_004112FA: jge 0041130Bh
  loc_004112FC: push 00000054h
  loc_004112FE: push 00403858h
  loc_00411303: push ebx
  loc_00411304: push eax
  loc_00411305: call [00401040h] ; __vbaHresultCheckObj
  loc_0041130B: lea ecx, var_14
  loc_0041130E: call [00401184h] ; __vbaFreeStr
  loc_00411314: lea ecx, var_24
  loc_00411317: call [00401180h] ; __vbaFreeObj
  loc_0041131D: lea eax, var_44
  loc_00411320: lea ecx, var_34
  loc_00411323: push eax
  loc_00411324: push ecx
  loc_00411325: push 00000002h
  loc_00411327: call [00401024h] ; __vbaFreeVarList
  loc_0041132D: add esp, 0000000Ch
  loc_00411330: call 0040B720h
  loc_00411335: mov edx, [esi]
  loc_00411337: push esi
  loc_00411338: fstp real4 ptr var_68
  loc_0041133B: call [edx+00000340h]
  loc_00411341: push eax
  loc_00411342: lea eax, var_24
  loc_00411345: push eax
  loc_00411346: call edi
  loc_00411348: mov ecx, var_68
  loc_0041134B: mov ebx, eax
  loc_0041134D: lea edx, var_34
  loc_00411350: push 00000002h
  loc_00411352: lea eax, var_44
  loc_00411355: push edx
  loc_00411356: push eax
  loc_00411357: mov var_6C, ebx
  loc_0041135A: mov var_2C, ecx
  loc_0041135D: mov var_34, 00000004h
  loc_00411364: call [004010E8h] ; rtcRound
  loc_0041136A: mov ebx, [ebx]
  loc_0041136C: lea ecx, var_44
  loc_0041136F: lea edx, var_14
  loc_00411372: push ecx
  loc_00411373: push edx
  loc_00411374: call [004010F8h] ; __vbaStrVarVal
  loc_0041137A: mov var_98, ebx
  loc_00411380: mov ebx, var_6C
  loc_00411383: push eax
  loc_00411384: mov eax, var_98
  loc_0041138A: push ebx
  loc_0041138B: call [eax+00000054h]
  loc_0041138E: test eax, eax
  loc_00411390: fnclex
  loc_00411392: jge 004113A3h
  loc_00411394: push 00000054h
  loc_00411396: push 00403858h
  loc_0041139B: push ebx
  loc_0041139C: push eax
  loc_0041139D: call [00401040h] ; __vbaHresultCheckObj
  loc_004113A3: lea ecx, var_14
  loc_004113A6: call [00401184h] ; __vbaFreeStr
  loc_004113AC: lea ecx, var_24
  loc_004113AF: call [00401180h] ; __vbaFreeObj
  loc_004113B5: lea ecx, var_44
  loc_004113B8: lea edx, var_34
  loc_004113BB: push ecx
  loc_004113BC: push edx
  loc_004113BD: push 00000002h
  loc_004113BF: call [00401024h] ; __vbaFreeVarList
  loc_004113C5: add esp, 0000000Ch
  loc_004113C8: call 0040B660h
  loc_004113CD: mov eax, [esi]
  loc_004113CF: push esi
  loc_004113D0: fstp real4 ptr var_68
  loc_004113D3: call [eax+00000374h]
  loc_004113D9: lea ecx, var_24
  loc_004113DC: push eax
  loc_004113DD: push ecx
  loc_004113DE: call edi
  loc_004113E0: mov edx, var_68
  loc_004113E3: mov ebx, eax
  loc_004113E5: lea eax, var_34
  loc_004113E8: push 00000002h
  loc_004113EA: lea ecx, var_44
  loc_004113ED: push eax
  loc_004113EE: push ecx
  loc_004113EF: mov var_6C, ebx
  loc_004113F2: mov var_2C, edx
  loc_004113F5: mov var_34, 00000004h
  loc_004113FC: call [004010E8h] ; rtcRound
  loc_00411402: mov ebx, [ebx]
  loc_00411404: lea edx, var_44
  loc_00411407: lea eax, var_14
  loc_0041140A: push edx
  loc_0041140B: push eax
  loc_0041140C: call [004010F8h] ; __vbaStrVarVal
  loc_00411412: mov ecx, ebx
  loc_00411414: mov ebx, var_6C
  loc_00411417: push eax
  loc_00411418: push ebx
  loc_00411419: call [ecx+00000054h]
  loc_0041141C: test eax, eax
  loc_0041141E: fnclex
  loc_00411420: jge 00411431h
  loc_00411422: push 00000054h
  loc_00411424: push 00403858h
  loc_00411429: push ebx
  loc_0041142A: push eax
  loc_0041142B: call [00401040h] ; __vbaHresultCheckObj
  loc_00411431: lea ecx, var_14
  loc_00411434: call [00401184h] ; __vbaFreeStr
  loc_0041143A: lea ecx, var_24
  loc_0041143D: call [00401180h] ; __vbaFreeObj
  loc_00411443: lea edx, var_44
  loc_00411446: lea eax, var_34
  loc_00411449: push edx
  loc_0041144A: push eax
  loc_0041144B: push 00000002h
  loc_0041144D: call [00401024h] ; __vbaFreeVarList
  loc_00411453: add esp, 0000000Ch
  loc_00411456: call 0040B680h
  loc_0041145B: mov ecx, [esi]
  loc_0041145D: push esi
  loc_0041145E: fstp real4 ptr var_68
  loc_00411461: call [ecx+00000330h]
  loc_00411467: lea edx, var_24
  loc_0041146A: push eax
  loc_0041146B: push edx
  loc_0041146C: call edi
  loc_0041146E: lea ecx, var_34
  loc_00411471: mov ebx, eax
  loc_00411473: mov eax, var_68
  loc_00411476: push 00000002h
  loc_00411478: lea edx, var_44
  loc_0041147B: push ecx
  loc_0041147C: push edx
  loc_0041147D: mov var_6C, ebx
  loc_00411480: mov var_2C, eax
  loc_00411483: mov var_34, 00000004h
  loc_0041148A: call [004010E8h] ; rtcRound
  loc_00411490: mov ebx, [ebx]
  loc_00411492: lea eax, var_44
  loc_00411495: lea ecx, var_14
  loc_00411498: push eax
  loc_00411499: push ecx
  loc_0041149A: call [004010F8h] ; __vbaStrVarVal
  loc_004114A0: mov edx, ebx
  loc_004114A2: mov ebx, var_6C
  loc_004114A5: push eax
  loc_004114A6: push ebx
  loc_004114A7: call [edx+00000054h]
  loc_004114AA: test eax, eax
  loc_004114AC: fnclex
  loc_004114AE: jge 004114BFh
  loc_004114B0: push 00000054h
  loc_004114B2: push 00403858h
  loc_004114B7: push ebx
  loc_004114B8: push eax
  loc_004114B9: call [00401040h] ; __vbaHresultCheckObj
  loc_004114BF: lea ecx, var_14
  loc_004114C2: call [00401184h] ; __vbaFreeStr
  loc_004114C8: lea ecx, var_24
  loc_004114CB: call [00401180h] ; __vbaFreeObj
  loc_004114D1: lea eax, var_44
  loc_004114D4: lea ecx, var_34
  loc_004114D7: push eax
  loc_004114D8: push ecx
  loc_004114D9: push 00000002h
  loc_004114DB: call [00401024h] ; __vbaFreeVarList
  loc_004114E1: add esp, 0000000Ch
  loc_004114E4: call 0040B740h
  loc_004114E9: mov edx, [esi]
  loc_004114EB: push esi
  loc_004114EC: fstp real4 ptr var_68
  loc_004114EF: call [edx+00000328h]
  loc_004114F5: push eax
  loc_004114F6: lea eax, var_24
  loc_004114F9: push eax
  loc_004114FA: call edi
  loc_004114FC: mov ecx, var_68
  loc_004114FF: mov ebx, eax
  loc_00411501: lea edx, var_34
  loc_00411504: push 00000002h
  loc_00411506: lea eax, var_44
  loc_00411509: push edx
  loc_0041150A: push eax
  loc_0041150B: mov var_6C, ebx
  loc_0041150E: mov var_2C, ecx
  loc_00411511: mov var_34, 00000004h
  loc_00411518: call [004010E8h] ; rtcRound
  loc_0041151E: mov ebx, [ebx]
  loc_00411520: lea ecx, var_44
  loc_00411523: lea edx, var_14
  loc_00411526: push ecx
  loc_00411527: push edx
  loc_00411528: call [004010F8h] ; __vbaStrVarVal
  loc_0041152E: mov var_A4, ebx
  loc_00411534: mov ebx, var_6C
  loc_00411537: push eax
  loc_00411538: mov eax, var_A4
  loc_0041153E: push ebx
  loc_0041153F: call [eax+00000054h]
  loc_00411542: test eax, eax
  loc_00411544: fnclex
  loc_00411546: jge 00411557h
  loc_00411548: push 00000054h
  loc_0041154A: push 00403858h
  loc_0041154F: push ebx
  loc_00411550: push eax
  loc_00411551: call [00401040h] ; __vbaHresultCheckObj
  loc_00411557: lea ecx, var_14
  loc_0041155A: call [00401184h] ; __vbaFreeStr
  loc_00411560: lea ecx, var_24
  loc_00411563: call [00401180h] ; __vbaFreeObj
  loc_00411569: lea ecx, var_44
  loc_0041156C: lea edx, var_34
  loc_0041156F: push ecx
  loc_00411570: push edx
  loc_00411571: push 00000002h
  loc_00411573: call [00401024h] ; __vbaFreeVarList
  loc_00411579: add esp, 0000000Ch
  loc_0041157C: call 0040B6E0h
  loc_00411581: mov eax, [esi]
  loc_00411583: push esi
  loc_00411584: fstp real4 ptr var_68
  loc_00411587: call [eax+00000338h]
  loc_0041158D: lea ecx, var_24
  loc_00411590: push eax
  loc_00411591: push ecx
  loc_00411592: call edi
  loc_00411594: mov edx, var_68
  loc_00411597: mov ebx, eax
  loc_00411599: lea eax, var_34
  loc_0041159C: push 00000002h
  loc_0041159E: lea ecx, var_44
  loc_004115A1: push eax
  loc_004115A2: push ecx
  loc_004115A3: mov var_6C, ebx
  loc_004115A6: mov var_2C, edx
  loc_004115A9: mov var_34, 00000004h
  loc_004115B0: call [004010E8h] ; rtcRound
  loc_004115B6: mov ebx, [ebx]
  loc_004115B8: lea edx, var_44
  loc_004115BB: lea eax, var_14
  loc_004115BE: push edx
  loc_004115BF: push eax
  loc_004115C0: call [004010F8h] ; __vbaStrVarVal
  loc_004115C6: mov ecx, ebx
  loc_004115C8: mov ebx, var_6C
  loc_004115CB: push eax
  loc_004115CC: push ebx
  loc_004115CD: call [ecx+00000054h]
  loc_004115D0: test eax, eax
  loc_004115D2: fnclex
  loc_004115D4: jge 004115E5h
  loc_004115D6: push 00000054h
  loc_004115D8: push 00403858h
  loc_004115DD: push ebx
  loc_004115DE: push eax
  loc_004115DF: call [00401040h] ; __vbaHresultCheckObj
  loc_004115E5: lea ecx, var_14
  loc_004115E8: call [00401184h] ; __vbaFreeStr
  loc_004115EE: lea ecx, var_24
  loc_004115F1: call [00401180h] ; __vbaFreeObj
  loc_004115F7: lea edx, var_44
  loc_004115FA: lea eax, var_34
  loc_004115FD: push edx
  loc_004115FE: push eax
  loc_004115FF: push 00000002h
  loc_00411601: call [00401024h] ; __vbaFreeVarList
  loc_00411607: mov ecx, [esi]
  loc_00411609: add esp, 0000000Ch
  loc_0041160C: push esi
  loc_0041160D: call [ecx+000003A0h]
  loc_00411613: lea edx, var_24
  loc_00411616: push eax
  loc_00411617: push edx
  loc_00411618: call edi
  loc_0041161A: mov edi, [0040103Ch] ; __vbaStrCat
  loc_00411620: push 00403D38h ; "T-Test"
  loc_00411625: push 00403D4Ch ; vbCrLf
  loc_0041162A: mov esi, eax
  loc_0041162C: call edi
  loc_0041162E: mov ebx, [00401164h] ; __vbaStrMove
  loc_00411634: mov edx, eax
  loc_00411636: lea ecx, var_14
  loc_00411639: call ebx
  loc_0041163B: push eax
  loc_0041163C: push 000000DFh
  loc_00411641: call [00401108h] ; rtcBstrFromAnsi
  loc_00411647: mov edx, eax
  loc_00411649: lea ecx, var_18
  loc_0041164C: call ebx
  loc_0041164E: push eax
  loc_0041164F: call edi
  loc_00411651: mov edx, eax
  loc_00411653: lea ecx, var_1C
  loc_00411656: call ebx
  loc_00411658: push eax
  loc_00411659: push 00403D58h ; "1="
  loc_0041165E: call edi
  loc_00411660: mov var_3C, eax
  loc_00411663: lea eax, var_64
  loc_00411666: push 00000002h
  loc_00411668: lea ecx, var_34
  loc_0041166B: push eax
  loc_0041166C: push ecx
  loc_0041166D: mov var_44, 00000008h
  loc_00411674: mov var_5C, 0041505Ch
  loc_0041167B: mov var_64, 00004004h
  loc_00411682: call [004010E8h] ; rtcRound
  loc_00411688: mov edi, [esi]
  loc_0041168A: lea edx, var_44
  loc_0041168D: lea eax, var_34
  loc_00411690: push edx
  loc_00411691: lea ecx, var_54
  loc_00411694: push eax
  loc_00411695: push ecx
  loc_00411696: call [00401100h] ; __vbaVarCat
  loc_0041169C: lea edx, var_20
  loc_0041169F: push eax
  loc_004116A0: push edx
  loc_004116A1: call [004010F8h] ; __vbaStrVarVal
  loc_004116A7: push eax
  loc_004116A8: push esi
  loc_004116A9: call [edi+00000054h]
  loc_004116AC: test eax, eax
  loc_004116AE: fnclex
  loc_004116B0: jge 004116C1h
  loc_004116B2: push 00000054h
  loc_004116B4: push 00403858h
  loc_004116B9: push esi
  loc_004116BA: push eax
  loc_004116BB: call [00401040h] ; __vbaHresultCheckObj
  loc_004116C1: lea eax, var_20
  loc_004116C4: lea ecx, var_1C
  loc_004116C7: push eax
  loc_004116C8: lea edx, var_18
  loc_004116CB: push ecx
  loc_004116CC: lea eax, var_14
  loc_004116CF: push edx
  loc_004116D0: push eax
  loc_004116D1: push 00000004h
  loc_004116D3: call [00401130h] ; __vbaFreeStrList
  loc_004116D9: add esp, 00000014h
  loc_004116DC: lea ecx, var_24
  loc_004116DF: call [00401180h] ; __vbaFreeObj
  loc_004116E5: lea ecx, var_54
  loc_004116E8: lea edx, var_34
  loc_004116EB: push ecx
  loc_004116EC: lea eax, var_44
  loc_004116EF: push edx
  loc_004116F0: push eax
  loc_004116F1: push 00000003h
  loc_004116F3: call [00401024h] ; __vbaFreeVarList
  loc_004116F9: add esp, 00000010h
  loc_004116FC: fwait
  loc_004116FD: push 00411741h
  loc_00411702: jmp 00411740h
  loc_00411704: lea ecx, var_20
  loc_00411707: lea edx, var_1C
  loc_0041170A: push ecx
  loc_0041170B: lea eax, var_18
  loc_0041170E: push edx
  loc_0041170F: lea ecx, var_14
  loc_00411712: push eax
  loc_00411713: push ecx
  loc_00411714: push 00000004h
  loc_00411716: call [00401130h] ; __vbaFreeStrList
  loc_0041171C: add esp, 00000014h
  loc_0041171F: lea ecx, var_24
  loc_00411722: call [00401180h] ; __vbaFreeObj
  loc_00411728: lea edx, var_54
  loc_0041172B: lea eax, var_44
  loc_0041172E: push edx
  loc_0041172F: lea ecx, var_34
  loc_00411732: push eax
  loc_00411733: push ecx
  loc_00411734: push 00000003h
  loc_00411736: call [00401024h] ; __vbaFreeVarList
  loc_0041173C: add esp, 00000010h
  loc_0041173F: ret
  loc_00411740: ret
  loc_00411741: mov ecx, var_10
  loc_00411744: pop edi
  loc_00411745: pop esi
  loc_00411746: xor eax, eax
  loc_00411748: mov fs:[00000000h], ecx
  loc_0041174F: pop ebx
  loc_00411750: mov esp, ebp
  loc_00411752: pop ebp
  loc_00411753: retn 0004h
End Sub

Private Function Proc_3_39_4119E0
  loc_004119E0: push esi
  loc_004119E1: mov esi, [esp+00000008h]
  loc_004119E5: push edi
  loc_004119E6: or edi, FFFFFFFFh
  loc_004119E9: cmp [esi+0000006Ch], di
  loc_004119ED: jnz 004119F8h
  loc_004119EF: mov eax, [esi]
  loc_004119F1: push esi
  loc_004119F2: call [eax+00000774h]
  loc_004119F8: cmp [esi+0000006Eh], di
  loc_004119FC: jnz 00411A07h
  loc_004119FE: mov ecx, [esi]
  loc_00411A00: push esi
  loc_00411A01: call [ecx+00000790h]
  loc_00411A07: cmp [esi+00000072h], di
  loc_00411A0B: jnz 00411A16h
  loc_00411A0D: mov edx, [esi]
  loc_00411A0F: push esi
  loc_00411A10: call [edx+00000794h]
  loc_00411A16: cmp [esi+00000070h], di
  loc_00411A1A: jnz 00411A25h
  loc_00411A1C: mov eax, [esi]
  loc_00411A1E: push esi
  loc_00411A1F: call [eax+00000798h]
  loc_00411A25: pop edi
  loc_00411A26: xor eax, eax
  loc_00411A28: pop esi
  loc_00411A29: retn 0004h
End Sub

Private Sub Proc_3_40_411A30
  loc_00411A30: push ebp
  loc_00411A31: mov ebp, esp
  loc_00411A33: sub esp, 00000008h
  loc_00411A36: push 00401546h ; __vbaExceptHandler
  loc_00411A3B: mov eax, fs:[00000000h]
  loc_00411A41: push eax
  loc_00411A42: mov fs:[00000000h], esp
  loc_00411A49: sub esp, 0000004Ch
  loc_00411A4C: push ebx
  loc_00411A4D: push esi
  loc_00411A4E: push edi
  loc_00411A4F: mov var_8, esp
  loc_00411A52: mov var_4, 00401428h
  loc_00411A59: mov eax, [0041502Ch]
  loc_00411A5E: xor edi, edi
  loc_00411A60: push eax
  loc_00411A61: mov var_14, edi
  loc_00411A64: mov var_24, edi
  loc_00411A67: mov var_34, edi
  loc_00411A6A: mov var_44, edi
  loc_00411A6D: mov var_48, edi
  loc_00411A70: mov var_58, edi
  loc_00411A73: call [004010A0h] ; __vbaR4Str
  loc_00411A79: mov esi, Me
  loc_00411A7C: fstp real4 ptr var_14
  loc_00411A7F: mov ecx, [esi]
  loc_00411A81: push esi
  loc_00411A82: call [ecx+0000030Ch]
  loc_00411A88: lea edx, var_58
  loc_00411A8B: push eax
  loc_00411A8C: push edx
  loc_00411A8D: call [0040105Ch] ; __vbaObjSet
  loc_00411A93: mov eax, var_58
  loc_00411A96: push FFFFFFFFh
  loc_00411A98: push eax
  loc_00411A99: mov ecx, [eax]
  loc_00411A9B: call [ecx+0000008Ch]
  loc_00411AA1: cmp eax, edi
  loc_00411AA3: fnclex
  loc_00411AA5: jge 00411ABCh
  loc_00411AA7: mov edx, var_58
  loc_00411AAA: push 0000008Ch
  loc_00411AAF: push 00403C28h
  loc_00411AB4: push edx
  loc_00411AB5: push eax
  loc_00411AB6: call [00401040h] ; __vbaHresultCheckObj
  loc_00411ABC: mov eax, [esi]
  loc_00411ABE: lea ecx, var_24
  loc_00411AC1: lea edx, var_14
  loc_00411AC4: push ecx
  loc_00411AC5: push edx
  loc_00411AC6: push esi
  loc_00411AC7: call [eax+00000778h]
  loc_00411ACD: mov eax, var_58
  loc_00411AD0: mov edi, [00401144h] ; __vbaVarAdd
  loc_00411AD6: mov var_3C, 00000014h
  loc_00411ADD: mov var_44, 00000002h
  loc_00411AE4: mov ebx, [eax]
  loc_00411AE6: lea ecx, var_24
  loc_00411AE9: lea edx, var_44
  loc_00411AEC: push ecx
  loc_00411AED: lea eax, var_34
  loc_00411AF0: push edx
  loc_00411AF1: push eax
  loc_00411AF2: call edi
  loc_00411AF4: push eax
  loc_00411AF5: call [004010B0h] ; __vbaR4Var
  loc_00411AFB: push ecx
  loc_00411AFC: mov ecx, var_58
  loc_00411AFF: fstp real4 ptr [esp]
  loc_00411B02: push ecx
  loc_00411B03: call [ebx+0000006Ch]
  loc_00411B06: test eax, eax
  loc_00411B08: fnclex
  loc_00411B0A: jge 00411B1Eh
  loc_00411B0C: mov edx, var_58
  loc_00411B0F: push 0000006Ch
  loc_00411B11: push 00403C28h
  loc_00411B16: push edx
  loc_00411B17: push eax
  loc_00411B18: call [00401040h] ; __vbaHresultCheckObj
  loc_00411B1E: mov ebx, [00401024h] ; __vbaFreeVarList
  loc_00411B24: lea eax, var_34
  loc_00411B27: lea ecx, var_24
  loc_00411B2A: push eax
  loc_00411B2B: push ecx
  loc_00411B2C: push 00000002h
  loc_00411B2E: call ebx
  loc_00411B30: add esp, 0000000Ch
  loc_00411B33: call 0040B600h
  loc_00411B38: fstp real4 ptr var_4C
  loc_00411B3B: call 0040B5E0h
  loc_00411B40: mov edx, [0041502Ch]
  loc_00411B46: fstp real4 ptr var_50
  loc_00411B49: push edx
  loc_00411B4A: call [00401118h] ; __vbaR8Str
  loc_00411B50: fmul st0, real4 ptr var_50
  loc_00411B53: fadd st0, real4 ptr var_4C
  loc_00411B56: lea ecx, var_24
  loc_00411B59: lea edx, var_48
  loc_00411B5C: fstp real4 ptr var_48
  loc_00411B5F: fnstsw ax
  loc_00411B61: test al, 0Dh
  loc_00411B63: jnz 00411C15h
  loc_00411B69: mov eax, [esi]
  loc_00411B6B: push ecx
  loc_00411B6C: push edx
  loc_00411B6D: push esi
  loc_00411B6E: call [eax+0000077Ch]
  loc_00411B74: mov eax, var_58
  loc_00411B77: mov var_3C, 00000014h
  loc_00411B7E: mov var_44, 00000002h
  loc_00411B85: lea ecx, var_24
  loc_00411B88: mov esi, [eax]
  loc_00411B8A: lea edx, var_44
  loc_00411B8D: push ecx
  loc_00411B8E: lea eax, var_34
  loc_00411B91: push edx
  loc_00411B92: push eax
  loc_00411B93: call edi
  loc_00411B95: push eax
  loc_00411B96: call [004010B0h] ; __vbaR4Var
  loc_00411B9C: push ecx
  loc_00411B9D: mov ecx, var_58
  loc_00411BA0: fstp real4 ptr [esp]
  loc_00411BA3: push ecx
  loc_00411BA4: call [esi+00000074h]
  loc_00411BA7: test eax, eax
  loc_00411BA9: fnclex
  loc_00411BAB: jge 00411BBFh
  loc_00411BAD: mov edx, var_58
  loc_00411BB0: push 00000074h
  loc_00411BB2: push 00403C28h
  loc_00411BB7: push edx
  loc_00411BB8: push eax
  loc_00411BB9: call [00401040h] ; __vbaHresultCheckObj
  loc_00411BBF: lea eax, var_34
  loc_00411BC2: lea ecx, var_24
  loc_00411BC5: push eax
  loc_00411BC6: push ecx
  loc_00411BC7: push 00000002h
  loc_00411BC9: call ebx
  loc_00411BCB: add esp, 0000000Ch
  loc_00411BCE: lea edx, var_58
  loc_00411BD1: push 00000000h
  loc_00411BD3: push edx
  loc_00411BD4: call [0040106Ch] ; __vbaObjSetAddref
  loc_00411BDA: fwait
  loc_00411BDB: push 00411C00h
  loc_00411BE0: jmp 00411BF6h
  loc_00411BE2: lea eax, var_34
  loc_00411BE5: lea ecx, var_24
  loc_00411BE8: push eax
  loc_00411BE9: push ecx
  loc_00411BEA: push 00000002h
  loc_00411BEC: call [00401024h] ; __vbaFreeVarList
  loc_00411BF2: add esp, 0000000Ch
  loc_00411BF5: ret
  loc_00411BF6: lea ecx, var_58
  loc_00411BF9: call [00401180h] ; __vbaFreeObj
  loc_00411BFF: ret
  loc_00411C00: mov ecx, var_10
  loc_00411C03: pop edi
  loc_00411C04: pop esi
  loc_00411C05: xor eax, eax
  loc_00411C07: mov fs:[00000000h], ecx
  loc_00411C0E: pop ebx
  loc_00411C0F: mov esp, ebp
  loc_00411C11: pop ebp
  loc_00411C12: retn 0004h
End Sub

Private Sub Proc_3_41_411C20
  loc_00411C20: push ebp
  loc_00411C21: mov ebp, esp
  loc_00411C23: sub esp, 00000008h
  loc_00411C26: push 00401546h ; __vbaExceptHandler
  loc_00411C2B: mov eax, fs:[00000000h]
  loc_00411C31: push eax
  loc_00411C32: mov fs:[00000000h], esp
  loc_00411C39: sub esp, 0000004Ch
  loc_00411C3C: push ebx
  loc_00411C3D: push esi
  loc_00411C3E: push edi
  loc_00411C3F: mov var_8, esp
  loc_00411C42: mov var_4, 00401438h
  loc_00411C49: mov eax, [0041502Ch]
  loc_00411C4E: xor edi, edi
  loc_00411C50: push eax
  loc_00411C51: mov var_14, edi
  loc_00411C54: mov var_24, edi
  loc_00411C57: mov var_34, edi
  loc_00411C5A: mov var_44, edi
  loc_00411C5D: mov var_48, edi
  loc_00411C60: mov var_58, edi
  loc_00411C63: call [004010A0h] ; __vbaR4Str
  loc_00411C69: mov esi, Me
  loc_00411C6C: fstp real4 ptr var_14
  loc_00411C6F: mov ecx, [esi]
  loc_00411C71: push esi
  loc_00411C72: call [ecx+00000308h]
  loc_00411C78: lea edx, var_58
  loc_00411C7B: push eax
  loc_00411C7C: push edx
  loc_00411C7D: call [0040105Ch] ; __vbaObjSet
  loc_00411C83: mov eax, var_58
  loc_00411C86: push FFFFFFFFh
  loc_00411C88: push eax
  loc_00411C89: mov ecx, [eax]
  loc_00411C8B: call [ecx+00000084h]
  loc_00411C91: cmp eax, edi
  loc_00411C93: fnclex
  loc_00411C95: jge 00411CACh
  loc_00411C97: mov edx, var_58
  loc_00411C9A: push 00000084h
  loc_00411C9F: push 00403C48h
  loc_00411CA4: push edx
  loc_00411CA5: push eax
  loc_00411CA6: call [00401040h] ; __vbaHresultCheckObj
  loc_00411CAC: mov eax, [esi]
  loc_00411CAE: lea ecx, var_24
  loc_00411CB1: lea edx, var_14
  loc_00411CB4: push ecx
  loc_00411CB5: push edx
  loc_00411CB6: push esi
  loc_00411CB7: call [eax+00000778h]
  loc_00411CBD: mov eax, var_58
  loc_00411CC0: mov edi, [00401144h] ; __vbaVarAdd
  loc_00411CC6: mov var_3C, 0000004Bh
  loc_00411CCD: mov var_44, 00000002h
  loc_00411CD4: mov ebx, [eax]
  loc_00411CD6: lea ecx, var_24
  loc_00411CD9: lea edx, var_44
  loc_00411CDC: push ecx
  loc_00411CDD: lea eax, var_34
  loc_00411CE0: push edx
  loc_00411CE1: push eax
  loc_00411CE2: call edi
  loc_00411CE4: push eax
  loc_00411CE5: call [004010B0h] ; __vbaR4Var
  loc_00411CEB: push ecx
  loc_00411CEC: mov ecx, var_58
  loc_00411CEF: fstp real4 ptr [esp]
  loc_00411CF2: push ecx
  loc_00411CF3: call [ebx+00000064h]
  loc_00411CF6: test eax, eax
  loc_00411CF8: fnclex
  loc_00411CFA: jge 00411D12h
  loc_00411CFC: mov edx, var_58
  loc_00411CFF: mov ebx, [00401040h] ; __vbaHresultCheckObj
  loc_00411D05: push 00000064h
  loc_00411D07: push 00403C48h
  loc_00411D0C: push edx
  loc_00411D0D: push eax
  loc_00411D0E: call ebx
  loc_00411D10: jmp 00411D18h
  loc_00411D12: mov ebx, [00401040h] ; __vbaHresultCheckObj
  loc_00411D18: lea eax, var_34
  loc_00411D1B: lea ecx, var_24
  loc_00411D1E: push eax
  loc_00411D1F: push ecx
  loc_00411D20: push 00000002h
  loc_00411D22: call [00401024h] ; __vbaFreeVarList
  loc_00411D28: mov eax, var_58
  loc_00411D2B: add esp, 0000000Ch
  loc_00411D2E: lea ecx, var_48
  loc_00411D31: mov edx, [eax]
  loc_00411D33: push ecx
  loc_00411D34: push eax
  loc_00411D35: call [edx+00000060h]
  loc_00411D38: test eax, eax
  loc_00411D3A: fnclex
  loc_00411D3C: jge 00411D4Ch
  loc_00411D3E: mov edx, var_58
  loc_00411D41: push 00000060h
  loc_00411D43: push 00403C48h
  loc_00411D48: push edx
  loc_00411D49: push eax
  loc_00411D4A: call ebx
  loc_00411D4C: mov eax, var_58
  loc_00411D4F: mov edx, var_48
  loc_00411D52: push edx
  loc_00411D53: push eax
  loc_00411D54: mov ecx, [eax]
  loc_00411D56: call [ecx+00000074h]
  loc_00411D59: test eax, eax
  loc_00411D5B: fnclex
  loc_00411D5D: jge 00411D6Dh
  loc_00411D5F: mov ecx, var_58
  loc_00411D62: push 00000074h
  loc_00411D64: push 00403C48h
  loc_00411D69: push ecx
  loc_00411D6A: push eax
  loc_00411D6B: call ebx
  loc_00411D6D: call 0040B720h
  loc_00411D72: fstp real4 ptr var_4C
  loc_00411D75: mov edx, var_4C
  loc_00411D78: mov eax, [esi]
  loc_00411D7A: mov var_48, edx
  loc_00411D7D: lea ecx, var_24
  loc_00411D80: lea edx, var_48
  loc_00411D83: push ecx
  loc_00411D84: push edx
  loc_00411D85: push esi
  loc_00411D86: call [eax+0000077Ch]
  loc_00411D8C: mov eax, var_58
  loc_00411D8F: mov var_3C, 0000004Bh
  loc_00411D96: mov var_44, 00000002h
  loc_00411D9D: lea ecx, var_24
  loc_00411DA0: mov ebx, [eax]
  loc_00411DA2: lea edx, var_44
  loc_00411DA5: push ecx
  loc_00411DA6: lea eax, var_34
  loc_00411DA9: push edx
  loc_00411DAA: push eax
  loc_00411DAB: call edi
  loc_00411DAD: push eax
  loc_00411DAE: call [004010B0h] ; __vbaR4Var
  loc_00411DB4: push ecx
  loc_00411DB5: mov ecx, var_58
  loc_00411DB8: fstp real4 ptr [esp]
  loc_00411DBB: push ecx
  loc_00411DBC: call [ebx+0000006Ch]
  loc_00411DBF: test eax, eax
  loc_00411DC1: fnclex
  loc_00411DC3: jge 00411DD7h
  loc_00411DC5: mov edx, var_58
  loc_00411DC8: push 0000006Ch
  loc_00411DCA: push 00403C48h
  loc_00411DCF: push edx
  loc_00411DD0: push eax
  loc_00411DD1: call [00401040h] ; __vbaHresultCheckObj
  loc_00411DD7: mov ebx, [00401024h] ; __vbaFreeVarList
  loc_00411DDD: lea eax, var_34
  loc_00411DE0: lea ecx, var_24
  loc_00411DE3: push eax
  loc_00411DE4: push ecx
  loc_00411DE5: push 00000002h
  loc_00411DE7: call ebx
  loc_00411DE9: add esp, 0000000Ch
  loc_00411DEC: call 0040B700h
  loc_00411DF1: fstp real4 ptr var_4C
  loc_00411DF4: mov edx, var_4C
  loc_00411DF7: mov eax, [esi]
  loc_00411DF9: mov var_48, edx
  loc_00411DFC: lea ecx, var_24
  loc_00411DFF: lea edx, var_48
  loc_00411E02: push ecx
  loc_00411E03: push edx
  loc_00411E04: push esi
  loc_00411E05: call [eax+0000077Ch]
  loc_00411E0B: mov eax, var_58
  loc_00411E0E: mov var_3C, 0000004Bh
  loc_00411E15: mov var_44, 00000002h
  loc_00411E1C: lea ecx, var_24
  loc_00411E1F: mov esi, [eax]
  loc_00411E21: lea edx, var_44
  loc_00411E24: push ecx
  loc_00411E25: lea eax, var_34
  loc_00411E28: push edx
  loc_00411E29: push eax
  loc_00411E2A: call edi
  loc_00411E2C: push eax
  loc_00411E2D: call [004010B0h] ; __vbaR4Var
  loc_00411E33: push ecx
  loc_00411E34: mov ecx, var_58
  loc_00411E37: fstp real4 ptr [esp]
  loc_00411E3A: push ecx
  loc_00411E3B: call [esi+0000007Ch]
  loc_00411E3E: test eax, eax
  loc_00411E40: fnclex
  loc_00411E42: jge 00411E56h
  loc_00411E44: mov edx, var_58
  loc_00411E47: push 0000007Ch
  loc_00411E49: push 00403C48h
  loc_00411E4E: push edx
  loc_00411E4F: push eax
  loc_00411E50: call [00401040h] ; __vbaHresultCheckObj
  loc_00411E56: lea eax, var_34
  loc_00411E59: lea ecx, var_24
  loc_00411E5C: push eax
  loc_00411E5D: push ecx
  loc_00411E5E: push 00000002h
  loc_00411E60: call ebx
  loc_00411E62: add esp, 0000000Ch
  loc_00411E65: lea edx, var_58
  loc_00411E68: push 00000000h
  loc_00411E6A: push edx
  loc_00411E6B: call [0040106Ch] ; __vbaObjSetAddref
  loc_00411E71: fwait
  loc_00411E72: push 00411E97h
  loc_00411E77: jmp 00411E8Dh
  loc_00411E79: lea eax, var_34
  loc_00411E7C: lea ecx, var_24
  loc_00411E7F: push eax
  loc_00411E80: push ecx
  loc_00411E81: push 00000002h
  loc_00411E83: call [00401024h] ; __vbaFreeVarList
  loc_00411E89: add esp, 0000000Ch
  loc_00411E8C: ret
  loc_00411E8D: lea ecx, var_58
  loc_00411E90: call [00401180h] ; __vbaFreeObj
  loc_00411E96: ret
  loc_00411E97: mov ecx, var_10
  loc_00411E9A: pop edi
  loc_00411E9B: pop esi
  loc_00411E9C: xor eax, eax
  loc_00411E9E: mov fs:[00000000h], ecx
  loc_00411EA5: pop ebx
  loc_00411EA6: mov esp, ebp
  loc_00411EA8: pop ebp
  loc_00411EA9: retn 0004h
End Sub

Private Sub Proc_3_42_411EB0
  loc_00411EB0: push ebp
  loc_00411EB1: mov ebp, esp
  loc_00411EB3: sub esp, 00000008h
  loc_00411EB6: push 00401546h ; __vbaExceptHandler
  loc_00411EBB: mov eax, fs:[00000000h]
  loc_00411EC1: push eax
  loc_00411EC2: mov fs:[00000000h], esp
  loc_00411EC9: sub esp, 0000004Ch
  loc_00411ECC: push ebx
  loc_00411ECD: push esi
  loc_00411ECE: push edi
  loc_00411ECF: mov var_8, esp
  loc_00411ED2: mov var_4, 00401448h
  loc_00411ED9: mov eax, [0041502Ch]
  loc_00411EDE: xor edi, edi
  loc_00411EE0: push eax
  loc_00411EE1: mov var_14, edi
  loc_00411EE4: mov var_24, edi
  loc_00411EE7: mov var_34, edi
  loc_00411EEA: mov var_44, edi
  loc_00411EED: mov var_48, edi
  loc_00411EF0: mov var_58, edi
  loc_00411EF3: call [004010A0h] ; __vbaR4Str
  loc_00411EF9: mov esi, Me
  loc_00411EFC: fstp real4 ptr var_14
  loc_00411EFF: mov ecx, [esi]
  loc_00411F01: push esi
  loc_00411F02: call [ecx+00000304h]
  loc_00411F08: lea edx, var_58
  loc_00411F0B: push eax
  loc_00411F0C: push edx
  loc_00411F0D: call [0040105Ch] ; __vbaObjSet
  loc_00411F13: mov eax, var_58
  loc_00411F16: push FFFFFFFFh
  loc_00411F18: push eax
  loc_00411F19: mov ecx, [eax]
  loc_00411F1B: call [ecx+00000084h]
  loc_00411F21: cmp eax, edi
  loc_00411F23: fnclex
  loc_00411F25: jge 00411F3Ch
  loc_00411F27: mov edx, var_58
  loc_00411F2A: push 00000084h
  loc_00411F2F: push 00403C48h
  loc_00411F34: push edx
  loc_00411F35: push eax
  loc_00411F36: call [00401040h] ; __vbaHresultCheckObj
  loc_00411F3C: mov eax, [esi]
  loc_00411F3E: lea ecx, var_24
  loc_00411F41: lea edx, var_14
  loc_00411F44: push ecx
  loc_00411F45: push edx
  loc_00411F46: push esi
  loc_00411F47: call [eax+00000778h]
  loc_00411F4D: mov eax, var_58
  loc_00411F50: mov edi, [00401144h] ; __vbaVarAdd
  loc_00411F56: mov var_3C, 0000004Bh
  loc_00411F5D: mov var_44, 00000002h
  loc_00411F64: mov ebx, [eax]
  loc_00411F66: lea ecx, var_24
  loc_00411F69: lea edx, var_44
  loc_00411F6C: push ecx
  loc_00411F6D: lea eax, var_34
  loc_00411F70: push edx
  loc_00411F71: push eax
  loc_00411F72: call edi
  loc_00411F74: push eax
  loc_00411F75: call [004010B0h] ; __vbaR4Var
  loc_00411F7B: push ecx
  loc_00411F7C: mov ecx, var_58
  loc_00411F7F: fstp real4 ptr [esp]
  loc_00411F82: push ecx
  loc_00411F83: call [ebx+00000064h]
  loc_00411F86: test eax, eax
  loc_00411F88: fnclex
  loc_00411F8A: jge 00411FA2h
  loc_00411F8C: mov edx, var_58
  loc_00411F8F: mov ebx, [00401040h] ; __vbaHresultCheckObj
  loc_00411F95: push 00000064h
  loc_00411F97: push 00403C48h
  loc_00411F9C: push edx
  loc_00411F9D: push eax
  loc_00411F9E: call ebx
  loc_00411FA0: jmp 00411FA8h
  loc_00411FA2: mov ebx, [00401040h] ; __vbaHresultCheckObj
  loc_00411FA8: lea eax, var_34
  loc_00411FAB: lea ecx, var_24
  loc_00411FAE: push eax
  loc_00411FAF: push ecx
  loc_00411FB0: push 00000002h
  loc_00411FB2: call [00401024h] ; __vbaFreeVarList
  loc_00411FB8: mov eax, var_58
  loc_00411FBB: add esp, 0000000Ch
  loc_00411FBE: lea ecx, var_48
  loc_00411FC1: mov edx, [eax]
  loc_00411FC3: push ecx
  loc_00411FC4: push eax
  loc_00411FC5: call [edx+00000060h]
  loc_00411FC8: test eax, eax
  loc_00411FCA: fnclex
  loc_00411FCC: jge 00411FDCh
  loc_00411FCE: mov edx, var_58
  loc_00411FD1: push 00000060h
  loc_00411FD3: push 00403C48h
  loc_00411FD8: push edx
  loc_00411FD9: push eax
  loc_00411FDA: call ebx
  loc_00411FDC: mov eax, var_58
  loc_00411FDF: mov edx, var_48
  loc_00411FE2: push edx
  loc_00411FE3: push eax
  loc_00411FE4: mov ecx, [eax]
  loc_00411FE6: call [ecx+00000074h]
  loc_00411FE9: test eax, eax
  loc_00411FEB: fnclex
  loc_00411FED: jge 00411FFDh
  loc_00411FEF: mov ecx, var_58
  loc_00411FF2: push 00000074h
  loc_00411FF4: push 00403C48h
  loc_00411FF9: push ecx
  loc_00411FFA: push eax
  loc_00411FFB: call ebx
  loc_00411FFD: call 0040B6C0h
  loc_00412002: fstp real4 ptr var_4C
  loc_00412005: mov edx, var_4C
  loc_00412008: mov eax, [esi]
  loc_0041200A: mov var_48, edx
  loc_0041200D: lea ecx, var_24
  loc_00412010: lea edx, var_48
  loc_00412013: push ecx
  loc_00412014: push edx
  loc_00412015: push esi
  loc_00412016: call [eax+0000077Ch]
  loc_0041201C: mov eax, var_58
  loc_0041201F: mov var_3C, 0000004Bh
  loc_00412026: mov var_44, 00000002h
  loc_0041202D: lea ecx, var_24
  loc_00412030: mov ebx, [eax]
  loc_00412032: lea edx, var_44
  loc_00412035: push ecx
  loc_00412036: lea eax, var_34
  loc_00412039: push edx
  loc_0041203A: push eax
  loc_0041203B: call edi
  loc_0041203D: push eax
  loc_0041203E: call [004010B0h] ; __vbaR4Var
  loc_00412044: push ecx
  loc_00412045: mov ecx, var_58
  loc_00412048: fstp real4 ptr [esp]
  loc_0041204B: push ecx
  loc_0041204C: call [ebx+0000006Ch]
  loc_0041204F: test eax, eax
  loc_00412051: fnclex
  loc_00412053: jge 00412067h
  loc_00412055: mov edx, var_58
  loc_00412058: push 0000006Ch
  loc_0041205A: push 00403C48h
  loc_0041205F: push edx
  loc_00412060: push eax
  loc_00412061: call [00401040h] ; __vbaHresultCheckObj
  loc_00412067: mov ebx, [00401024h] ; __vbaFreeVarList
  loc_0041206D: lea eax, var_34
  loc_00412070: lea ecx, var_24
  loc_00412073: push eax
  loc_00412074: push ecx
  loc_00412075: push 00000002h
  loc_00412077: call ebx
  loc_00412079: add esp, 0000000Ch
  loc_0041207C: call 0040B6A0h
  loc_00412081: fstp real4 ptr var_4C
  loc_00412084: mov edx, var_4C
  loc_00412087: mov eax, [esi]
  loc_00412089: mov var_48, edx
  loc_0041208C: lea ecx, var_24
  loc_0041208F: lea edx, var_48
  loc_00412092: push ecx
  loc_00412093: push edx
  loc_00412094: push esi
  loc_00412095: call [eax+0000077Ch]
  loc_0041209B: mov eax, var_58
  loc_0041209E: mov var_3C, 0000004Bh
  loc_004120A5: mov var_44, 00000002h
  loc_004120AC: lea ecx, var_24
  loc_004120AF: mov esi, [eax]
  loc_004120B1: lea edx, var_44
  loc_004120B4: push ecx
  loc_004120B5: lea eax, var_34
  loc_004120B8: push edx
  loc_004120B9: push eax
  loc_004120BA: call edi
  loc_004120BC: push eax
  loc_004120BD: call [004010B0h] ; __vbaR4Var
  loc_004120C3: push ecx
  loc_004120C4: mov ecx, var_58
  loc_004120C7: fstp real4 ptr [esp]
  loc_004120CA: push ecx
  loc_004120CB: call [esi+0000007Ch]
  loc_004120CE: test eax, eax
  loc_004120D0: fnclex
  loc_004120D2: jge 004120E6h
  loc_004120D4: mov edx, var_58
  loc_004120D7: push 0000007Ch
  loc_004120D9: push 00403C48h
  loc_004120DE: push edx
  loc_004120DF: push eax
  loc_004120E0: call [00401040h] ; __vbaHresultCheckObj
  loc_004120E6: lea eax, var_34
  loc_004120E9: lea ecx, var_24
  loc_004120EC: push eax
  loc_004120ED: push ecx
  loc_004120EE: push 00000002h
  loc_004120F0: call ebx
  loc_004120F2: add esp, 0000000Ch
  loc_004120F5: lea edx, var_58
  loc_004120F8: push 00000000h
  loc_004120FA: push edx
  loc_004120FB: call [0040106Ch] ; __vbaObjSetAddref
  loc_00412101: fwait
  loc_00412102: push 00412127h
  loc_00412107: jmp 0041211Dh
  loc_00412109: lea eax, var_34
  loc_0041210C: lea ecx, var_24
  loc_0041210F: push eax
  loc_00412110: push ecx
  loc_00412111: push 00000002h
  loc_00412113: call [00401024h] ; __vbaFreeVarList
  loc_00412119: add esp, 0000000Ch
  loc_0041211C: ret
  loc_0041211D: lea ecx, var_58
  loc_00412120: call [00401180h] ; __vbaFreeObj
  loc_00412126: ret
  loc_00412127: mov ecx, var_10
  loc_0041212A: pop edi
  loc_0041212B: pop esi
  loc_0041212C: xor eax, eax
  loc_0041212E: mov fs:[00000000h], ecx
  loc_00412135: pop ebx
  loc_00412136: mov esp, ebp
  loc_00412138: pop ebp
  loc_00412139: retn 0004h
End Sub
