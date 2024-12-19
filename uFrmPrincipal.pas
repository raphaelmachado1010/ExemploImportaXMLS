unit uFrmPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.Buttons, ComObj;

type
  TFrmPrincipal = class(TForm)
    btnImportar: TBitBtn;
    OpenDialog1: TOpenDialog;
    StringGrid1: TStringGrid;
    procedure btnImportarClick(Sender: TObject);
  private
    function XlsToStringGrid(XStringGrid: TStringGrid; xFileXLS: string): Boolean;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmPrincipal: TFrmPrincipal;

implementation

{$R *.dfm}

procedure TFrmPrincipal.btnImportarClick(Sender: TObject);
begin
  OpenDialog1.InitialDir := 'D:\TC\Bruder\';
  if OpenDialog1.Execute then
    XlsToStringGrid(StringGrid1,OpenDialog1.FileName);
end;
function TFrmPrincipal.XlsToStringGrid(XStringGrid: TStringGrid; xFileXLS: string): Boolean;
const xlCellTypeLastCell = $0000000B;
var XLSAplicacao, AbaXLS: OLEVariant;
    RangeMatrix: Variant;
    x, y, k, r: Integer;
begin
  Result := False;
   // Cria Excel- OLE Object
  XLSAplicacao := CreateOleObject('Excel.Application');
  try
    // Esconde Excel
    XLSAplicacao.Visible := False;
    // Abre o Workbook
    XLSAplicacao.Workbooks.Open(xFileXLS);

    {Selecione aqui a aba que você deseja abrir primeiro - 1,2,3,4....}
    XLSAplicacao.WorkSheets[1].Activate;
    {Selecione aqui a aba que você deseja ativar - começando sempre no 1 (1,2,3,4) }
    AbaXLS := XLSAplicacao.Workbooks[ExtractFileName(xFileXLS)].WorkSheets[1];

    AbaXLS.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Pegar o número da última linha
    x := XLSAplicacao.ActiveCell.Row;
    // Pegar o número da última coluna
    y := XLSAplicacao.ActiveCell.Column;
    // Seta xStringGrid linha e coluna
    XStringGrid.RowCount := x;
    XStringGrid.ColCount := y;
    // Associaca a variant WorkSheet com a variant do Delphi
    RangeMatrix := XLSAplicacao.Range['A1', XLSAplicacao.Cells.Item[x, y]].Value;
    // Cria o loop para listar os registros no TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
        XStringGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[k, r];
      Inc(k, 1);
    until k > x;
    RangeMatrix := Unassigned;
  finally
    // Fecha o Microsoft Excel
    if not VarIsEmpty(XLSAplicacao) then
    begin
      XLSAplicacao.Quit;
      XLSAplicacao := Unassigned;
      AbaXLS := Unassigned;
      Result := True;
    end;
  end;

end;

end.
