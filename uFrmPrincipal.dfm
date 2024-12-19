object FrmPrincipal: TFrmPrincipal
  Left = 0
  Top = 0
  Caption = 'Importa'#231#227'o de Plan.Excel'
  ClientHeight = 575
  ClientWidth = 1021
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object btnImportar: TBitBtn
    Left = 8
    Top = 544
    Width = 145
    Height = 25
    Caption = 'Importa Arquivo'
    TabOrder = 0
    OnClick = btnImportarClick
  end
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 1021
    Height = 538
    Align = alTop
    FixedCols = 0
    TabOrder = 1
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel|*.xlsx'
    Left = 880
    Top = 384
  end
end
