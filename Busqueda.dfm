object LA_BUSQUEDA: TLA_BUSQUEDA
  Left = 360
  Top = 217
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = '...'
  ClientHeight = 389
  ClientWidth = 781
  Color = 16771820
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 0
    Top = 148
    Width = 781
    Height = 221
    Caption = 'Datos Existentes'
    TabOrder = 0
    object DBGrid1: TDBGrid
      Left = 2
      Top = 15
      Width = 777
      Height = 204
      Align = alClient
      DataSource = DSBUSCANDO
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnCellClick = DBGrid1CellClick
      OnTitleClick = DBGrid1TitleClick
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 0
    Width = 341
    Height = 147
    Caption = 'B'#250'squeda y Navegaci'#243'n'
    TabOrder = 1
    object Label7: TLabel
      Left = 8
      Top = 71
      Width = 67
      Height = 14
      Caption = 'Ingrese Valor:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
    object Button1: TJvBitBtn
      Left = 124
      Top = 108
      Width = 81
      Height = 33
      Hint = 'Buscar y Asignar el Cliente'
      Caption = 'Buscar'
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        1800000000000003000000000000000000000000000000000000FFFFFFFFFFFF
        FFFFFF202020404040404040202020FFFFFFFFFFFFFFFFFFFFFFFF4040404040
        40404040FFFFFFFFFFFFFFFFFFFFFFFF000000505050505050505050909090FF
        FFFFFFFFFFFFFFFF202020505050505050808080404040FFFFFFFFFFFFFFFFFF
        0000006060606F6F6F000000909090FFFFFFFFFFFFFFFFFF000000DFDFDF0000
        00505050404040FFFFFFFFFFFFFFFFFF0000006060606F6F6F00000090909040
        4040404040404040000000DFDFDF000000505050606060FFFFFFFFFFFFFFFFFF
        2020203030303030300000005050505050508080805050500000006F6F6F0000
        00505050A0A0A0606060FFFFFF3030303030300000006F6F6F00000000000000
        00002020206F6F6F000000000000000000303030808080909090003F3F9FDFDF
        6F6F6F000000606060000000000000C0C0C00000006060600000000000000000
        00306F6F5050509090903F7F7F7FFFFF7FFFFF3F3F3F30303000000000000060
        60600000003030300000000000000000007FBFBF5050509090903F7F7F7FFFFF
        7FFFFF7FFFFF3F7F7F3030300000000000003F7F7F3030300000000000007FFF
        FF7FFFFF5050509090903F7F7F7FFFFF7FFFFF7FFFFF7FFFFF3030303030303F
        7F7F7FFFFF3030303030303F7F7F7FFFFF7FFFFF5050509090903F7F7F7FFFFF
        7FFFFF7FFFFF7FFFFF3030303030303F7F7F7FFFFF3030303030303F7F7F7FFF
        FF7FFFFF5050509090903F7F7F7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7F
        FFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF5050509090903F7F7F7FFFFF
        7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFF
        FF7FFFFF4040402020203F7F7F7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF7F
        FFFF7FFFFF7FFFFF7FFFFF7FFFFF7FFFFF205F5F202020FFFFFF0000003F3F00
        7F7F007F7F007F7F007F7F007F7F000000000000000000000000000000000000
        00FFFFFFFFFFFFFFFFFFFFFFFF0000007F7F007F7F007F7F007F7F003F3F00FF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TJvBitBtn
      Left = 208
      Top = 108
      Width = 117
      Height = 33
      Caption = 'Seleccionar y Salir'
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333FF33333333333330003FF3FFFFF3333777003000003333
        300077F777773F333777E00BFBFB033333337773333F7F33333FE0BFBF000333
        330077F3337773F33377E0FBFBFBF033330077F3333FF7FFF377E0BFBF000000
        333377F3337777773F3FE0FBFBFBFBFB039977F33FFFFFFF7377E0BF00000000
        339977FF777777773377000BFB03333333337773FF733333333F333000333333
        3300333777333333337733333333333333003333333333333377333333333333
        333333333333333333FF33333333333330003333333333333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
      TabOrder = 1
      OnClick = Button2Click
    end
    object Edit1: TEdit
      Left = 8
      Top = 84
      Width = 309
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 2
      OnKeyPress = Edit1KeyPress
    end
    object GroupBox4: TGroupBox
      Left = 2
      Top = 14
      Width = 333
      Height = 54
      Caption = 'Seleccione'
      TabOrder = 3
      object Label9: TLabel
        Left = 164
        Top = 10
        Width = 111
        Height = 13
        Caption = 'Seleccione la Columna:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object R1: TRadioButton
        Left = 6
        Top = 14
        Width = 81
        Height = 17
        Caption = 'Por Defecto'
        Checked = True
        TabOrder = 0
        TabStop = True
        OnClick = R1Click
      end
      object R2: TRadioButton
        Left = 6
        Top = 31
        Width = 79
        Height = 15
        Caption = 'Por Columna'
        TabOrder = 1
        OnClick = R2Click
      end
      object COL: TComboBox
        Left = 122
        Top = 26
        Width = 195
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        Visible = False
        OnKeyPress = COLKeyPress
      end
    end
  end
  object GroupBox3: TGroupBox
    Left = 342
    Top = 0
    Width = 439
    Height = 147
    Caption = 'Datos Seleccionados'
    TabOrder = 2
    object Label1: TLabel
      Left = 9
      Top = 127
      Width = 158
      Height = 14
      Caption = 'Barra de Navegaci'#243'n: ----------->'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
    object DB1: TDBEdit
      Left = 8
      Top = 15
      Width = 423
      Height = 22
      DataSource = DSBUSCANDO
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object DB2: TDBEdit
      Left = 8
      Top = 39
      Width = 423
      Height = 22
      DataSource = DSBUSCANDO
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object DB3: TDBEdit
      Left = 8
      Top = 64
      Width = 423
      Height = 22
      DataSource = DSBUSCANDO
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
    object DB4: TDBEdit
      Left = 8
      Top = 88
      Width = 423
      Height = 22
      DataSource = DSBUSCANDO
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
    end
    object DBNavigator1: TDBNavigator
      Left = 174
      Top = 116
      Width = 256
      Height = 27
      DataSource = DSBUSCANDO
      VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
      TabOrder = 4
    end
  end
  object Barra1: TStatusBar
    Left = 0
    Top = 370
    Width = 781
    Height = 19
    Panels = <
      item
        Width = 400
      end
      item
        Width = 400
      end>
  end
  object DSBUSCANDO: TDataSource
    OnDataChange = DSBUSCANDODataChange
    Left = 18
    Top = 180
  end
end
