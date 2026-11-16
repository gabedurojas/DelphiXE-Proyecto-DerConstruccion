unit SelecContribu;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids;

type
  TSeleccion_Contribu = class(TForm)
    grilla4: TStringGrid;
    procedure FormActivate(Sender: TObject);
    procedure grilla4KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    fila:integer;{ Public declarations }
  end;

var
  Seleccion_Contribu: TSeleccion_Contribu;

implementation

{$R *.dfm}

procedure TSeleccion_Contribu.FormActivate(Sender: TObject);
begin
grilla4.Cells[0,0]:='N° CUIM';
grilla4.Cells[1,0]:='Contribuyente';
grilla4.Cells[2,0]:='Caracter';
grilla4.Cells[3,0]:='N° Dni';
grilla4.Cells[4,0]:='Tipo Titular';
grilla4.Focused;
end;

procedure TSeleccion_Contribu.grilla4KeyPress(Sender: TObject; var Key: Char);
begin
if key = chr(27) then begin
  fila:=grilla4.Row;
  close;
end;
end;

end.
