object Bus_CodigoPostal: TBus_CodigoPostal
  Left = 471
  Top = 160
  Caption = 'B'#250'squeda del C'#243'digo Postal'
  ClientHeight = 324
  ClientWidth = 449
  Color = 15130594
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object StatusBar1: TStatusBar
    Left = 0
    Top = 305
    Width = 449
    Height = 19
    Panels = <>
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 449
    Height = 83
    Caption = 'B'#250'squeda'
    TabOrder = 1
    object LB1: TLabel
      Left = 171
      Top = 10
      Width = 87
      Height = 13
      Caption = 'Ingrese Localidad:'
    end
    object Label1: TLabel
      Left = 6
      Top = 62
      Width = 25
      Height = 13
      Caption = 'Cod.:'
    end
    object L1: TLabel
      Left = 38
      Top = 62
      Width = 13
      Height = 13
      Caption = '...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Edit1: TEdit
      Left = 64
      Top = 26
      Width = 321
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      OnKeyPress = Edit1KeyPress
    end
    object Buscar: TJvBitBtn
      Left = 108
      Top = 54
      Width = 61
      Height = 25
      Hint = 'Buscar Informaci'#243'n'
      Caption = 'Buscar'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BuscarClick
    end
    object Salir: TJvBitBtn
      Left = 253
      Top = 54
      Width = 61
      Height = 25
      Hint = 'Sale y Cierra Ventana'
      Caption = 'Salir'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = SalirClick
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 84
    Width = 449
    Height = 221
    Caption = 'Detalle de la Localidad Buscada'
    TabOrder = 2
    object DBGrid1: TDBGrid
      Left = 2
      Top = 15
      Width = 445
      Height = 204
      Align = alClient
      DataSource = DataSource1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = [fsBold]
      OnCellClick = DBGrid1CellClick
    end
  end
  object DataSource1: TDataSource
    Left = 22
    Top = 120
  end
end
