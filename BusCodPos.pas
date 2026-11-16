unit BusCodPos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, DBGrids, StdCtrls, ComCtrls, Mask,
  DB, ADODB, Buttons, JvExButtons, JvBitBtn;

type
  TBus_CodigoPostal = class(TForm)
    StatusBar1: TStatusBar;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    DBGrid1: TDBGrid;
    LB1: TLabel;
    DataSource1: TDataSource;
    Edit1: TEdit;
    Label1: TLabel;
    L1: TLabel;
    Buscar: TJvBitBtn;
    Salir: TJvBitBtn;
    procedure Coloca1(nro:byte);
    procedure Limpiar(nro:byte);
    procedure FormActivate(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure BuscarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SalirClick(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
  private
    DataSet1:TADODataSet;{ Private declarations }
  public
    eltpo:integer;
    muni,gh:boolean;

    { Public declarations }
  end;

var
  Bus_CodigoPostal: TBus_CodigoPostal;

implementation

uses Modulo, Unidad;


{$R *.dfm}
procedure TBus_CodigoPostal.FormCreate(Sender: TObject);
begin
gh:=false;
DataSet1:=TADODataSet.Create(nil);
DataSet1.ConnectionString:=modulo2.Sistema.ConectUsu;
DataSource1.DataSet:=DataSet1;
end;

procedure TBus_CodigoPostal.Coloca1(nro:byte);
var t:integer;
var fec,txt:string;
begin
if nro=1 then begin
  DataSet1.Close;
  DataSet1.CommandText:='SELECT Codigo as CodigoPostal,Postal FROM '
  + 'COD_POST WHERE (Postal LIKE ' + chr(39) + chr(37) + trim(Edit1.Text)
  + chr(37) + chr(39) + ') ORDER BY Postal;';
  DataSet1.Open;
end;
end;

procedure TBus_CodigoPostal.Limpiar(nro:byte);
begin
L1.Caption:='...';
end;


procedure TBus_CodigoPostal.FormActivate(Sender: TObject);
begin
if not(gh) then begin
  gh:=true;muni:=false;
  Edit1.SetFocus;
end;
end;

procedure TBus_CodigoPostal.DBGrid1CellClick(Column: TColumn);
begin
if trim(DataSet1.Fields[0].AsString) <> '' then begin
  L1.Caption:=trim(DataSet1.Fields[0].AsString);
  sALIR.SetFocus;
end;
end;

procedure TBus_CodigoPostal.BuscarClick(Sender: TObject);
begin
Limpiar(1);
Coloca1(1);
end;



procedure TBus_CodigoPostal.SalirClick(Sender: TObject);
begin
Close;
end;

procedure TBus_CodigoPostal.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
if key = chr(13) then Buscar.SetFocus;
end;

end.
