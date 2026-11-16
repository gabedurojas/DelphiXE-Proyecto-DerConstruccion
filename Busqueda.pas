unit Busqueda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, ExtCtrls, ComCtrls, DBCtrls,
  Mask, StrUtils, Buttons, JvExButtons, JvBitBtn;

type
  TLA_BUSQUEDA = class(TForm)
    DSBUSCANDO: TDataSource;
    GroupBox1: TGroupBox;
    DBGrid1: TDBGrid;
    GroupBox2: TGroupBox;
    Label7: TLabel;
    Button1: TJvBitBtn;
    Button2: TJvBitBtn;
    Edit1: TEdit;
    GroupBox3: TGroupBox;
    DB1: TDBEdit;
    DB2: TDBEdit;
    DB3: TDBEdit;
    DB4: TDBEdit;
    DBNavigator1: TDBNavigator;
    Label1: TLabel;
    GroupBox4: TGroupBox;
    R1: TRadioButton;
    R2: TRadioButton;
    Label9: TLabel;
    COL: TComboBox;
    Barra1: TStatusBar;
    procedure Coloca1();
    procedure FormActivate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure Button2Click(Sender: TObject);
    procedure R1Click(Sender: TObject);
    procedure R2Click(Sender: TObject);
    procedure COLKeyPress(Sender: TObject; var Key: Char);
    procedure DSBUSCANDODataChange(Sender: TObject; Field: TField);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure FormCreate(Sender: TObject);
  private
    DataSet1:TADODataSet;{ Private declarations }
  public
    gf:boolean;
    NroOpcion:byte;
    la_os:string;{ Public declarations }
  end;

var
  LA_BUSQUEDA: TLA_BUSQUEDA;

implementation

uses Unidad, Modulo;

{$R *.dfm}
procedure TLA_BUSQUEDA.FormCreate(Sender: TObject);
begin
gf:=false;
DataSet1:=TADODataSet.Create(nil);
DataSet1.ConnectionString:=modulo2.Sistema.ConectUsu;
DSBUSCANDO.DataSet:=DataSet1;
end;


procedure TLA_BUSQUEDA.FormActivate(Sender: TObject);
var t:integer;
begin
if not(gf) then begin
  gf:=true;
  Coloca1();
  COL.Clear;
  for t:=0 to DataSet1.FieldCount-1 do begin
    COL.Items.Insert(t,trim(DataSet1.Fields[t].FieldName));
  end;
  COL.ItemIndex:=0;
  Edit1.SetFocus;
end;
end;

procedure TLA_BUSQUEDA.Coloca1();
var t,y:integer;
begin
DataSet1.Close;
if NroOpcion < 6 then begin
  DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
  + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
  + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
  + 'FROM VDATOS_PARTIDAS WHERE (Partida=0);';
  DataSet1.Open;
end;
if NroOpcion = 6 then begin
  DataSet1.CommandText:='SELECT NroProveedor,Denominacion,TipoDoc,Documento,'
  + 'Cuit,Localidad,Domicilio,CodPost,Telefono,Fax FROM LosProveedores ORDER BY Denominacion;';
  DataSet1.Open;
end;
if NroOpcion = 7 then begin
  DataSet1.CommandText:='SELECT DISTINCT Cuim,Contribuyente as Nombre,TipoDoc,DNI FROM '
  + 'PlanesPagos.dbo.Fn_PlanPagoContribuyente (null) ORDER BY Contribuyente;';
  DataSet1.Open;
end;
if (NroOpcion >= 8) and (NroOpcion <= 12) then begin
  DataSet1.CommandText:='SELECT TOP 2 id_comercio,cuenta,th,FechaApertura,SupCubierta,'
  + 'SupDescubierta,Resolucion,Actividad,NombreTitular,FechaBaja,ResolBaja,NroSucursal,Barrio,'
  + 'Calle,Altura,DNI FROM VDatosComercios ORDER BY id_comercio';
    DataSet1.Open;
end;
if NroOpcion = 13 then begin
  DataSet1.CommandText:='SELECT top 5 Id, IdComercio, TH, FechaBaja, UsuBaja, Provisorio, Observaciones'
  + ' FROM LosComerciosTasa (0) ORDER BY TH;';
    DataSet1.Open;
end;
end;

procedure TLA_BUSQUEDA.Button1Click(Sender: TObject);
var txt:widestring;
begin
DataSet1.Close;
if R1.Checked = true then begin
  if NroOpcion = 1 then begin
    DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
    + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
    + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
    + 'FROM VDATOS_PARTIDAS WHERE (Titular LIKE ' + chr(39) + chr(37)
    + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(37) + chr(39) + ') ORDER BY Titular';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 2 then begin
    DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
    + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
    + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
    + 'FROM VDATOS_PARTIDAS WHERE (NroDni = ' + trim(Edit1.Text) + ') ORDER BY NroDni,Titular';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 3 then begin
    DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
    + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
    + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
    + 'FROM VDATOS_PARTIDAS WHERE (Nomenclatura LIKE ' + chr(39) + chr(37)
    + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(37) + chr(39) + ') ORDER BY Nomenclatura';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 4 then begin
    DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
    + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
    + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
    + 'FROM VDATOS_PARTIDAS WHERE (RespoTasa LIKE ' + chr(39) + chr(37)
    + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(37) + chr(39) + ') ORDER BY Titular';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 5 then  begin
    DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
    + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
    + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
    + 'FROM VDATOS_PARTIDAS WHERE (RespoInmo LIKE ' + chr(39) + chr(37)
    + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(37) + chr(39) + ') ORDER BY Titular';
     DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 6 then begin
    DataSet1.CommandText:='SELECT Cuim,Denominacion,TipoDoc,Documento,'
    + 'Cuit,Localidad,Domicilio,CodPost,Telefono,Fax FROM LosProveedores WHERE '
    + '(Denominacion LIKE ' + chr(39) + chr(37) + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33))
    + chr(37) + chr(39) + ') ORDER BY Denominacion';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 7 then begin
     DataSet1.CommandText:='SELECT Cuim,Contribuyente as Nombre,TipoDoc,DNI FROM '
    + 'PlanesPagos.dbo.Fn_PlanPagoContribuyente (null)  WHERE (Contribuyente LIKE ' + chr(39) + chr(37)
    + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(37) + chr(39) + ') ORDER BY Contribuyente';
    //DataSet1.CommandText:='SELECT DISTINCT NroCuim,Nombre,TipoDoc,DNI FROM '
    //+ 'PlanesPagos.dbo.LosContribuyentesPlanes WHERE (Nombre LIKE ' + chr(39) + chr(37)
    //+ AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(37) + chr(39) + ') ORDER BY Nombre';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if (NroOpcion >= 8) and (NroOpcion <= 12) then begin
    txt:='SELECT id_comercio,cuenta,th,FechaApertura,SupCubierta,SupDescubierta,'
         + 'Resolucion,Actividad,NombreTitular,FechaBaja,ResolBaja,NroSucursal,'
         + 'Barrio,Calle,Altura,DNI FROM VDatosComercios WHERE';
    if (NroOpcion = 8) then txt:=txt + ' (id_comercio = ' + trim(Edit1.Text) + ') ORDER BY id_comercio';
    if (NroOpcion = 9) then txt:=txt + ' (NombreTitular LIKE ' + chr(39) + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33))
                                     + chr(37) + chr(39) + ') ORDER BY id_comercio';
    if (NroOpcion = 10) then txt:=txt + ' (Cuenta = ' + chr(39) + trim(Edit1.Text) + chr(39) + ') ORDER BY id_comercio';
    if (NroOpcion = 11) then txt:=txt + ' (Resolucion = ' + trim(Edit1.Text) + ') ORDER BY id_comercio';
    if (NroOpcion = 12) then txt:=txt + ' (th = ' + chr(39) + trim(Edit1.Text) + chr(39) + ') ORDER BY id_comercio';
    DataSet1.CommandText:=txt;
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 13 then begin
    DataSet1.CommandText:='SELECT Id,IdComercio,TH,FechaBaja,UsuBaja,Provisorio,Observaciones FROM LosComerciosTasa (500)'
    + ' WHERE (th LIKE ' + chr(39) + chr(37) + trim(Edit1.Text) + chr(37) + chr(39) + ') ORDER BY th';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;  
end
else begin
  txt:=trim(COL.Text) + ' LIKE ' + chr (39) + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33))  + chr(37) + chr(39);
  if NroOpcion < 6 then begin
    DataSet1.CommandText:='SELECT DISTINCT Partida,Nomenclatura,TipoParcela,'
    + 'FechaAlta,FechaBaja,FechaAdjudicacion,Titular,NroDni,RespoTasa,RespoInmo,'
    + 'Barrio,Calle,Puerta,Sector,Torre,Piso,Depto,Manzana,Casa,Parcela,LoteRural '
    + 'FROM VDATOS_PARTIDAS WHERE (' + txt + ') ORDER BY Titular';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 6 then begin
    DataSet1.CommandText:='SELECT NroProveedor,Denominacion,TipoDoc,Documento,'
    + 'Cuit,Localidad,Domicilio,CodPost,Telefono,Fax FROM LosProveedores WHERE (' + txt + ') ORDER BY Denominacion;';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 7 then begin
    DataSet1.CommandText:='SELECT DISTINCT Cuim,Contribuyente as Nombre,TipoDoc,DNI FROM '
    + 'PlanesPagos.dbo.Fn_PlanPagoContribuyente (null) ORDER BY Contribuyente;';
    //DataSet1.CommandText:='SELECT DISTINCT NroCuim,Nombre,TipoDoc,DNI FROM '
    //+ 'PlanesPagos.dbo.LosContribuyentesPlanes WHERE (' + txt + ') ORDER BY Nombre;';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if (NroOpcion >= 8) and (NroOpcion <= 12) then begin
    DataSet1.CommandText:='SELECT id_comercio,cuenta,th,FechaApertura,SupCubierta,'
    + 'SupDescubierta,Resolucion,Actividad,FechaBaja,ResolBaja,NroSucursal,Barrio,'
    + 'Calle,Altura,DNI FROM VDatosComercios WHERE (' + txt + ') ORDER BY id_comercio';
     DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
  if NroOpcion = 13 then begin
  DataSet1.CommandText:='select Id,IdComercio,TH,FechaBaja,UsuBaja,Provisorio,Observaciones'
   +' FROM LosComerciosTasa (500) WHERE (' + txt + ')  ORDER BY TH;';
    DataSet1.Open;
    Barra1.Panels[0].Text:='Cantidad de Registros: ' + inttostr(DataSet1.RecordCount);
    Button2.SetFocus;
  end;
end;
end;

procedure TLA_BUSQUEDA.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
if key = chr(13) then Button1.SetFocus;
end;

procedure TLA_BUSQUEDA.Button2Click(Sender: TObject);
begin
close;
end;

procedure TLA_BUSQUEDA.R1Click(Sender: TObject);
begin
Label9.Visible:=false;
COL.Visible:=false;
Edit1.SetFocus;
end;

procedure TLA_BUSQUEDA.R2Click(Sender: TObject);
begin
Label9.Visible:=true;
COL.Visible:=true;
COL.SetFocus;
end;

procedure TLA_BUSQUEDA.COLKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then Edit1.SetFocus;
end;

procedure TLA_BUSQUEDA.DSBUSCANDODataChange(Sender: TObject;
  Field: TField);
var td:string;
    opcion:integer;
begin
opcion:=nroopcion;
DB1.DataField:=DataSet1.FieldDefs.Items[0].DisplayName; //Label4.Caption
DB2.DataField:=DataSet1.FieldDefs.Items[1].DisplayName; //Label4.Caption
if (NroOpcion <> 7) and (nroopcion <> 13)  then begin
  DB3.DataField:=DataSet1.FieldDefs.Items[6].DisplayName; //Label5.Caption
  DB4.DataField:=DataSet1.FieldDefs.Items[7].DisplayName; //modulo2.el_detalle
end;
if (NroOpcion = 7) then begin
  DB1.DataField:=DataSet1.FieldDefs.Items[0].DisplayName; //Label5.Caption
  DB2.DataField:=DataSet1.FieldDefs.Items[1].DisplayName; //modulo2.el_detalle
end;
if (NroOpcion = 8) then begin
  DB3.DataField:=DataSet1.FieldDefs.Items[2].DisplayName; //Label5.Caption
  DB4.DataField:=DataSet1.FieldDefs.Items[6].DisplayName; //modulo2.el_detalle
end;
if NroOpcion = 13 then begin
  DB1.DataField:=DataSet1.FieldDefs.Items[0].DisplayName; //Label4.Caption
  DB2.DataField:=DataSet1.FieldDefs.Items[1].DisplayName; //Label4.Caption
  DB3.DataField:=DataSet1.FieldDefs.Items[2].DisplayName; //Label5.Caption
  //DB4.DataField:=DataSet1.FieldDefs.Items[3].DisplayName; //modulo2.el_detalle
end;
Barra1.Panels[1].Text:='N° de Registro: ' + inttostr(DataSet1.RecNo);
Button2.SetFocus;
end;

procedure TLA_BUSQUEDA.DBGrid1TitleClick(Column: TColumn);
begin
DataSet1.IndexFieldNames:=Column.FieldName;
end;

procedure TLA_BUSQUEDA.DBGrid1CellClick(Column: TColumn);
begin
Button2.SetFocus;
end;



end.
