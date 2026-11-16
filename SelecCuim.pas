unit SelecCuim;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, JvExControls, JvxCheckListBox, StdCtrls, Buttons, JvExButtons,
  JvBitBtn, JvgGroupBox, ADODB;

type
  TSelec_Cuim = class(TForm)
    JvgGroupBox1: TJvgGroupBox;
    LB1: TLabel;
    CODI: TEdit;
    Buscar: TJvBitBtn;
    JvgGroupBox2: TJvgGroupBox;
    Button1: TJvBitBtn;
    LISTA: TJvxCheckListBox;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Coloca1(nro:byte);
    procedure CODIKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BuscarClick(Sender: TObject);
    procedure LISTAClickCheck(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure CODIExit(Sender: TObject);
  private
    DataSet1:TADODataSet;
    IdCombos:array of string;{ Private declarations }
  public
    Cancela,gh:boolean;
    NroCodVal:string;{ Public declarations }
  end;

var
  Selec_Cuim: TSelec_Cuim;

implementation

uses modulo;

{$R *.dfm}

procedure TSelec_Cuim.FormCreate(Sender: TObject);
begin
Modulo2.ObjConexion:=nil;

DataSet1:=TADODataSet.Create(nil);
end;

procedure TSelec_Cuim.FormActivate(Sender: TObject);
begin
if not (gh) then begin
  gh:=true;
  Cancela:=false;
  CODI.SetFocus;
end;
end;

procedure TSelec_Cuim.Coloca1(nro:byte);
var t:integer;
    txt:string;
begin
Modulo2.ObjConexion:=TADOConnection.Create(nil);
Modulo2.ObjConexion.LoginPrompt:=false;
if nro=1 then begin
  {Incorporando datos de la Localidad}
  LISTA.Clear;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectUsu;
  DataSet1.Connection:=modulo2.ObjConexion;
  DataSet1.CommandText:='SELECT dbo.ContribuyentesIdentidad.Id, dbo.ContribuyentesIdentidad.CUIM, '
  + 'dbo.Contribuyentes.Contribuyente FROM dbo.Contribuyentes INNER JOIN dbo.ContribuyentesIdentidad '
  + 'ON dbo.Contribuyentes.Id = dbo.ContribuyentesIdentidad.IdContribuyente WHERE '
  + '(dbo.ContribuyentesIdentidad.NroDni=' + CODI.Text + ') ORDER BY 2';
  DataSet1.Open;
  Setlength(IdCombos,DataSet1.RecordCount);
  for t:=1 to DataSet1.RecordCount  do begin
    DataSet1.RecNo:=t;
    LISTA.Items.Add(trim(DataSet1.Fields[1].AsString) + ' - ' + trim(DataSet1.Fields[2].AsString) );
    IdCombos[t-1]:=trim(DataSet1.Fields[0].AsString);
  end;
  DataSet1.Close;
  LISTA.ItemIndex:=0;
end;
FreeAndNil(Modulo2.ObjConexion);
end;

procedure TSelec_Cuim.CODIKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  Buscar.SetFocus;
end
else begin
  if ( StrScan('0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TSelec_Cuim.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key = VK_ESCAPE then begin
  Cancela:=true;
  Close;
end;
end;

procedure TSelec_Cuim.BuscarClick(Sender: TObject);
begin
Coloca1(1);
end;

procedure TSelec_Cuim.LISTAClickCheck(Sender: TObject);
begin
NroCodVal:=IdCombos[LISTA.ItemIndex];
Button1.Enabled:=true;
Button1.SetFocus;
end;

procedure TSelec_Cuim.Button1Click(Sender: TObject);
begin
Close;
end;

procedure TSelec_Cuim.CODIExit(Sender: TObject);
begin
if trim(CODI.Text) = '' then begin
  messagedlg('Error!! Ingrese un N° de Documento válido',mterror,[mbok],0);
  CODI.SetFocus;
end
end;

end.
