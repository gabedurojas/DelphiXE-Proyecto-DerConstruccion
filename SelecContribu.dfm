object Seleccion_Contribu: TSeleccion_Contribu
  Left = 168
  Top = 150
  AutoScroll = False
  BorderIcons = []
  Caption = 
    'Seleccione el Contribuyente                                (Pres' +
    'ione ESC para Salir)'
  ClientHeight = 195
  ClientWidth = 744
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = [fsBold]
  OldCreateOrder = False
  Position = poDesktopCenter
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object grilla4: TStringGrid
    Left = 0
    Top = 0
    Width = 744
    Height = 195
    Hint = 'Detalle de los Medicamentos de la Receta'
    Align = alClient
    DefaultRowHeight = 15
    FixedCols = 0
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 0
    OnKeyPress = grilla4KeyPress
    ColWidths = (
      81
      389
      114
      78
      73)
  end
end
