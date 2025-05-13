object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = #1052#1077#1085#1077#1076#1078#1077#1088' '#1079#1072#1076#1072#1095
  ClientHeight = 584
  ClientWidth = 733
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object pnlMain: TPanel
    Left = 0
    Top = 0
    Width = 733
    Height = 584
    Align = alClient
    ShowCaption = False
    TabOrder = 0
    ExplicitLeft = 480
    ExplicitTop = 32
    ExplicitWidth = 185
    ExplicitHeight = 41
    object spl1: TSplitter
      Left = 1
      Top = 123
      Width = 731
      Height = 3
      Cursor = crVSplit
      Align = alTop
      ExplicitTop = 129
      ExplicitWidth = 435
    end
    object spl2: TSplitter
      Left = 1
      Top = 249
      Width = 731
      Height = 3
      Cursor = crVSplit
      Align = alTop
    end
    object spl3: TSplitter
      Left = 1
      Top = 452
      Width = 731
      Height = 3
      Cursor = crVSplit
      Align = alTop
      ExplicitTop = 321
      ExplicitWidth = 243
    end
    object pnlTasks: TPanel
      Left = 1
      Top = 1
      Width = 731
      Height = 122
      Align = alTop
      Constraints.MaxHeight = 122
      ShowCaption = False
      TabOrder = 0
      DesignSize = (
        731
        122)
      object lblTaskDescription: TLabel
        Left = 419
        Top = 40
        Width = 279
        Height = 39
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        Caption = #1054#1087#1080#1089#1072#1085#1080#1077' '#1079#1072#1076#1072#1095#1080':'
        WordWrap = True
      end
      object btnLoadDLL: TButton
        Left = 5
        Top = 8
        Width = 92
        Height = 25
        Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' DLL'
        TabOrder = 0
        OnClick = btnLoadDLLClick
      end
      object lbDLLs: TListBox
        Left = 5
        Top = 39
        Width = 145
        Height = 71
        ItemHeight = 13
        TabOrder = 1
        OnClick = lbDLLsClick
      end
      object btnUnloadDLL: TButton
        Left = 103
        Top = 8
        Width = 97
        Height = 25
        Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100' DLL'
        Enabled = False
        TabOrder = 2
        OnClick = btnUnloadDLLClick
      end
      object btnStartTask: TButton
        Left = 419
        Top = 85
        Width = 38
        Height = 25
        Hint = #1047#1072#1087#1091#1089#1090#1080#1090#1100' '#1079#1072#1076#1072#1095#1091
        Caption = #1055#1091#1089#1082
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = btnStartTaskClick
      end
      object lbTasks: TListBox
        Left = 156
        Top = 39
        Width = 257
        Height = 71
        ItemHeight = 13
        TabOrder = 4
        OnClick = lbTasksClick
      end
      object btnStopTask: TButton
        Left = 463
        Top = 85
        Width = 42
        Height = 25
        Hint = #1054#1089#1090#1072#1085#1086#1074#1080#1090#1100' '#1079#1072#1076#1072#1095#1091
        Caption = #1057#1090#1086#1087
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        OnClick = btnStopTaskClick
      end
      object btnShowResults: TButton
        Left = 511
        Top = 85
        Width = 82
        Height = 25
        Hint = #1055#1086#1082#1072#1079#1072#1090#1100' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1099
        Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 6
        OnClick = btnShowResultsClick
      end
      object btnFreeTask: TButton
        Left = 632
        Top = 85
        Width = 85
        Height = 25
        Hint = #1042#1099#1075#1088#1091#1079#1080#1090#1100' '#1079#1072#1076#1072#1095#1091' '#1080#1079' '#1087#1072#1084#1103#1090#1080
        Anchors = [akTop, akRight]
        Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 7
        OnClick = btnFreeTaskClick
      end
    end
    object statBar: TStatusBar
      Left = 1
      Top = 564
      Width = 731
      Height = 19
      Panels = <
        item
          Width = 150
        end
        item
          Width = 300
        end>
      ExplicitLeft = 328
      ExplicitTop = 416
      ExplicitWidth = 0
    end
    object pnlParams: TPanel
      Left = 1
      Top = 126
      Width = 731
      Height = 123
      Align = alTop
      Caption = 'pnlParams'
      TabOrder = 2
      object pnlParams2: TScrollBox
        Left = 1
        Top = 1
        Width = 729
        Height = 121
        Align = alClient
        TabOrder = 0
        ExplicitLeft = -512
        ExplicitTop = -72
        ExplicitWidth = 697
        ExplicitHeight = 113
      end
    end
    object pnlTaskRun: TPanel
      Left = 1
      Top = 252
      Width = 731
      Height = 200
      Align = alTop
      TabOrder = 3
      object sgTasks: TStringGrid
        Left = 1
        Top = 1
        Width = 729
        Height = 198
        Align = alClient
        DefaultColWidth = 100
        FixedCols = 0
        RowCount = 1
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
        TabOrder = 0
        ExplicitLeft = 16
        ExplicitTop = 29
        ExplicitWidth = 705
        ExplicitHeight = 40
        ColWidths = (
          40
          150
          70
          100
          120)
      end
    end
    object pnlTaskLog: TPanel
      Left = 1
      Top = 455
      Width = 731
      Height = 109
      Align = alClient
      AutoSize = True
      TabOrder = 4
      ExplicitTop = 324
      ExplicitHeight = 240
      object memoLog: TMemo
        Left = 1
        Top = 1
        Width = 729
        Height = 107
        Align = alClient
        ScrollBars = ssVertical
        TabOrder = 0
        ExplicitLeft = 8
        ExplicitTop = 144
        ExplicitWidth = 705
        ExplicitHeight = 96
      end
    end
  end
  object tmrProgress: TTimer
    Interval = 500
    OnTimer = tmrProgressTimer
    Left = 688
    Top = 8
  end
  object dlgOpenDLL: TOpenDialog
    Filter = 'DLL '#1092#1072#1081#1083#1099' (*.dll)|*.dll|'#1042#1089#1077' '#1092#1072#1081#1083#1099' (*.*)|*.*'
    Title = #1042#1099#1073#1077#1088#1080#1090#1077' DLL '#1092#1072#1081#1083
    Left = 688
    Top = 56
  end
  object pmTaskMenu: TPopupMenu
    Left = 688
    Top = 112
    object miGetResult: TMenuItem
      Caption = #1055#1086#1083#1091#1095#1080#1090#1100' '#1088#1077#1079#1091#1083#1100#1090#1072#1090
      OnClick = btnGetResultClick
    end
    object miShowResults: TMenuItem
      Caption = #1055#1086#1082#1072#1079#1072#1090#1100' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1099
      OnClick = btnShowResultsClick
    end
    object miStopTask: TMenuItem
      Caption = #1054#1089#1090#1072#1085#1086#1074#1080#1090#1100' '#1079#1072#1076#1072#1095#1091
      OnClick = btnStopTaskClick
    end
    object miFreeTask: TMenuItem
      Caption = #1054#1089#1074#1086#1073#1086#1076#1080#1090#1100' '#1079#1072#1076#1072#1095#1091
      OnClick = btnFreeTaskClick
    end
  end
end
