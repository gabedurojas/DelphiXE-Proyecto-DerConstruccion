unit AM_Contrib;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DB, ADODB, Mask, Buttons, Grids, DBGrids,
  ExtCtrls, StrUtils, JvExDBGrids, JvDBGrid, JvgGroupBox, JvExButtons, JvBitBtn,
  ToolWin, JvExGrids, JvStringGrid, Menus, ImgList, JvDataEmbedded, Objetos,
  JvComponentBase, JvEmbeddedForms, JvExStdCtrls, JvGroupBox,frxDBSet, DBClient,frxClass,
  JvGIF, Vcl.Imaging.jpeg;

type THackStringGrid = class(TJvStringGrid);
type
  TAM_Contribuyentes = class(TForm)
    Barra1: TStatusBar;
    DataSource1: TDataSource;
    PC1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    BORRA: TJvBitBtn;
    ToolButton2: TToolButton;
    GUARDA: TJvBitBtn;
    ToolButton3: TToolButton;
    SALE: TJvBitBtn;
    JvgGroupBox3: TJvgGroupBox;
    JvDBGrid2: TJvDBGrid;
    JvgGroupBox1: TJvgGroupBox;
    LB1: TLabel;
    Label9: TLabel;
    Label28: TLabel;
    ONOMB: TEdit;
    ODNI: TEdit;
    Buscar: TJvBitBtn;
    OCUIT: TEdit;
    JvgGroupBox2: TJvgGroupBox;
    JvgGroupBox4: TJvgGroupBox;
    Label7: TLabel;
    H3: TComboBox;
    H2: TEdit;
    Label22: TLabel;
    LB3: TLabel;
    H4: TEdit;
    Label8: TLabel;
    H1: TEdit;
    Agrega2: TJvBitBtn;
    Modifica2: TJvBitBtn;
    Borra2: TJvBitBtn;
    Quita2: TJvBitBtn;
    JvgGroupBox5: TJvgGroupBox;
    Label27: TLabel;
    L1: TEdit;
    Label4: TLabel;
    L2: TEdit;
    Label6: TLabel;
    L3: TMaskEdit;
    Label5: TLabel;
    L4: TComboBox;
    L5: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    L7: TEdit;
    Label25: TLabel;
    L8: TComboBox;
    L10: TCheckBox;
    JvgGroupBox6: TJvgGroupBox;
    Label26: TLabel;
    D15: TEdit;
    Label19: TLabel;
    D10: TEdit;
    D6: TEdit;
    Label15: TLabel;
    Label14: TLabel;
    D5: TEdit;
    Label13: TLabel;
    D3: TComboBox;
    Label12: TLabel;
    D1: TEdit;
    Label10: TLabel;
    D2: TEdit;
    Label11: TLabel;
    D4: TComboBox;
    D8: TEdit;
    Label17: TLabel;
    D7: TEdit;
    Label16: TLabel;
    Button4: TButton;
    Label24: TLabel;
    Label20: TLabel;
    D11: TEdit;
    Label21: TLabel;
    D12: TEdit;
    Label1: TLabel;
    Label18: TLabel;
    D9: TEdit;
    D13: TEdit;
    D14: TEdit;
    Label23: TLabel;
    Label29: TLabel;
    L6: TComboBox;
    Label30: TLabel;
    L9: TEdit;
    PopupMenu1: TPopupMenu;
    Activar1: TMenuItem;
    N1: TMenuItem;
    AsociarCUIM1: TMenuItem;
    Label31: TLabel;
    L11: TComboBox;
    GroupBox2: TGroupBox;
    C6: TCheckBox;
    C3: TCheckBox;
    C4: TCheckBox;
    C5: TCheckBox;
    ImageList1: TImageList;
    FormularioEmb: TJvEmbeddedFormLink;
    grillaidentidad: TStringGrid;
    TabSheet5: TTabSheet;
    JvGroupBox1: TJvGroupBox;
    JvGroupBox2: TJvGroupBox;
    Label32: TLabel;
    banco: TComboBox;
    Label33: TLabel;
    tipoctabanco: TComboBox;
    Label34: TLabel;
    nrosucursal: TEdit;
    Label35: TLabel;
    nrocuenta: TEdit;
    escredito: TCheckBox;
    grillactabancos: TStringGrid;
    agrega4: TJvBitBtn;
    modifica4: TJvBitBtn;
    C7: TCheckBox;
    Label36: TLabel;
    elrubro: TComboBox;
    elcodobjeto: TEdit;
    Label37: TLabel;
    Label38: TLabel;
    grilladebitos: TStringGrid;
    modifica5: TJvBitBtn;
    agrega5: TJvBitBtn;
    C8: TCheckBox;
    Agrega3: TJvBitBtn;
    Modifica3: TJvBitBtn;
    Borra3: TJvBitBtn;
    Quita3: TJvBitBtn;
    JvgGroupBox7: TJvgGroupBox;
    grilladomicilio: TStringGrid;
    JvGroupBox3: TJvGroupBox;
    Label39: TLabel;
    Shape2: TShape;
    Shape3: TShape;
    Label40: TLabel;
    Label41: TLabel;
    elsubrubro: TComboBox;
    PopupMenu2: TPopupMenu;
    PopupMenu3: TPopupMenu;
    QuitarSeleccion1: TMenuItem;
    QuitarDebitoSeleccionado1: TMenuItem;
    Label42: TLabel;
    frxReport1: TfrxReport;
    Label43: TLabel;
    borra5: TJvBitBtn;
    Label44: TLabel;
    cbu: TEdit;
    BusPlanes: TButton;
    TabSheet6: TTabSheet;
    JvgGroupBox8: TJvgGroupBox;
    JvgGroupBox9: TJvgGroupBox;
    GroupBox1: TGroupBox;
    agregacarnet: TJvBitBtn;
    print: TJvBitBtn;
    GroupBox3: TGroupBox;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    Image1: TImage;
    Label50: TLabel;
    LC: TLabel;
    LCUIT: TLabel;
    LD: TLabel;
    LL: TLabel;
    GroupBox4: TGroupBox;
    Label51: TLabel;
    Image2: TImage;
    Image3: TImage;
    GroupBox5: TGroupBox;
    Label52: TLabel;
    LNRO: TLabel;
    DBGrid1: TDBGrid;
    DataSource2: TDataSource;
    lcodbar: TLabel;
    Label53: TLabel;
    nombfantasia: TEdit;
    Label54: TLabel;
    domcomercio: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CargaCombos(nro:byte);
    procedure Consultas(nro:byte);
    procedure BORRAClick(Sender: TObject);
    procedure ODNIKeyPress(Sender: TObject; var Key: Char);
    procedure ONOMBKeyPress(Sender: TObject; var Key: Char);
    procedure OCUITKeyPress(Sender: TObject; var Key: Char);
    procedure BuscarClick(Sender: TObject);
    procedure L2KeyPress(Sender: TObject; var Key: Char);
    procedure L2Exit(Sender: TObject);
    procedure L3KeyPress(Sender: TObject; var Key: Char);
    procedure L3Exit(Sender: TObject);
    procedure L4KeyPress(Sender: TObject; var Key: Char);
    procedure L4Exit(Sender: TObject);
    procedure L5KeyPress(Sender: TObject; var Key: Char);
    procedure L5Exit(Sender: TObject);
    procedure L6KeyPress(Sender: TObject; var Key: Char);
    procedure L6Exit(Sender: TObject);
    procedure L7KeyPress(Sender: TObject; var Key: Char);
    procedure L7Exit(Sender: TObject);
    procedure L8KeyPress(Sender: TObject; var Key: Char);
    procedure L8Exit(Sender: TObject);
    procedure L9KeyPress(Sender: TObject; var Key: Char);
    procedure L9Exit(Sender: TObject);
    procedure L10KeyPress(Sender: TObject; var Key: Char);
    procedure H3KeyPress(Sender: TObject; var Key: Char);
    procedure H3Exit(Sender: TObject);
    procedure H4KeyPress(Sender: TObject; var Key: Char);
    procedure H4Exit(Sender: TObject);
    procedure H4Change(Sender: TObject);
    procedure Agrega2Click(Sender: TObject);
    procedure Modifica2Click(Sender: TObject);
    procedure Borra2Click(Sender: TObject);
    procedure Quita2Click(Sender: TObject);
    procedure D3KeyPress(Sender: TObject; var Key: Char);
    procedure D3Exit(Sender: TObject);
    procedure D4KeyPress(Sender: TObject; var Key: Char);
    procedure D4Exit(Sender: TObject);
    procedure D5KeyPress(Sender: TObject; var Key: Char);
    procedure D5Exit(Sender: TObject);
    procedure D6KeyPress(Sender: TObject; var Key: Char);
    procedure D6Exit(Sender: TObject);
    procedure D7KeyPress(Sender: TObject; var Key: Char);
    procedure D8KeyPress(Sender: TObject; var Key: Char);
    procedure D10KeyPress(Sender: TObject; var Key: Char);
    procedure D11KeyPress(Sender: TObject; var Key: Char);
    procedure D12KeyPress(Sender: TObject; var Key: Char);
    procedure D13KeyPress(Sender: TObject; var Key: Char);
    procedure D9KeyPress(Sender: TObject; var Key: Char);
    procedure D14KeyPress(Sender: TObject; var Key: Char);
    procedure D15KeyPress(Sender: TObject; var Key: Char);
    procedure Agrega3Click(Sender: TObject);
    procedure Modifica3Click(Sender: TObject);
    procedure Borra3Click(Sender: TObject);
    procedure Quita3Click(Sender: TObject);
    procedure GUARDAClick(Sender: TObject);
    Function Verificacion(nro:byte):boolean;
    procedure SALEClick(Sender: TObject);
    procedure JvDBGrid2CellClick(Column: TColumn);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure Activar1Click(Sender: TObject);
    procedure L2Change(Sender: TObject);
    procedure AsociarCUIM1Click(Sender: TObject);
    procedure L11KeyPress(Sender: TObject; var Key: Char);
    procedure L11Exit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure grillaidentidadClick(Sender: TObject);
    procedure grillaidentidadDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure grilladomicilioClick(Sender: TObject);
    procedure bancoExit(Sender: TObject);
    procedure tipoctabancoExit(Sender: TObject);
    procedure bancoKeyPress(Sender: TObject; var Key: Char);
    procedure tipoctabancoKeyPress(Sender: TObject; var Key: Char);
    procedure nrosucursalKeyPress(Sender: TObject; var Key: Char);
    procedure nrocuentaKeyPress(Sender: TObject; var Key: Char);
    procedure nrosucursalExit(Sender: TObject);
    procedure nrocuentaExit(Sender: TObject);
    procedure escreditoKeyPress(Sender: TObject; var Key: Char);
    procedure agrega4Click(Sender: TObject);
    procedure grillactabancosClick(Sender: TObject);
    procedure modifica4Click(Sender: TObject);
    procedure grillactabancosDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure elcodobjetoExit(Sender: TObject);
    procedure elcodobjetoKeyPress(Sender: TObject; var Key: Char);
    procedure elsubrubroExit(Sender: TObject);
    procedure elsubrubroKeyPress(Sender: TObject; var Key: Char);
    procedure agrega5Click(Sender: TObject);
    procedure grilladebitosClick(Sender: TObject);
    function VerificaGrilla(grilla:string):boolean;
    procedure modifica5Click(Sender: TObject);
    procedure QuitarSeleccion1Click(Sender: TObject);
    procedure QuitarDebitoSeleccionado1Click(Sender: TObject);
    procedure VerArea(nro:byte);
    procedure ImprimeConfirmacion();
    procedure borra5Click(Sender: TObject);
    procedure cbuKeyPress(Sender: TObject; var Key: Char);
    procedure Button1Click(Sender: TObject);
    procedure cbuExit(Sender: TObject);
    procedure agregacarnetClick(Sender: TObject);
    procedure printClick(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure PC1Change(Sender: TObject);
    procedure nombfantasiaKeyPress(Sender: TObject; var Key: Char);
    procedure domcomercioKeyPress(Sender: TObject; var Key: Char);
  private
    Lista:array of string;
    CONTRIBUYENTE,CONTRIBCTASBANCOS,OBJETOS,DataSet1,DataSet2,CARNET:TADODataSet;
    Com1:TADOCommand;
    { Private declarations }
  public
    Recibo:TRecibo;
    NroOpcion:integer;
    gh,yaexiste:boolean;
    nomb,laidcuenta,dato,FechaAct,eldni,ctadeb,lasucbanco,nrocarnet,NroCodBar,dmi,elcuim:string;
    { Public declarations }
  end;

var
  AM_Contribuyentes: TAM_Contribuyentes;

implementation

uses Unidad, modulo, BusCodPos, SelecCuim;

{$R *.dfm}
procedure TAM_Contribuyentes.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
FreeAndNil(CONTRIBUYENTE);
FreeAndNil(DataSet1);
FreeAndNil(Recibo);
FreeAndNil(DataSet2);
FreeAndNil(CONTRIBCTASBANCOS);
FreeAndNil(OBJETOS);
freeAndNil(CARNET);
FreeAndNil(Self);
end;

procedure TAM_Contribuyentes.FormCreate(Sender: TObject);
begin
gh:=false;
Recibo:=TRecibo.Create;
Modulo2.ObjConexion:=nil;
CONTRIBUYENTE:=TADODataSet.Create(Self);
CONTRIBCTASBANCOS:=TADODataSet.Create(Self);
OBJETOS:=TADODataSet.Create(Self);
CARNET:=TADODataSet.Create(Self);
DataSet1:=TADODataSet.Create(Self);
DataSet2:=TADODataSet.Create(Self);
pc1.Pages[1].TabVisible:=false;
pc1.Pages[2].TabVisible:=false;
pc1.Pages[3].TabVisible:=false;
pc1.Pages[4].TabVisible:=false;
pc1.Pages[5].TabVisible:=false;
//if Modulo2.ObtenerNivelAcceso('151') <= 3 then pc1.Pages[4].TabVisible:=true
//  else pc1.Pages[4].TabVisible:=false;
end;

procedure TAM_Contribuyentes.VerArea(nro:byte);
var Cnx2:TADOConnection;
begin
Cnx2:=TADOConnection.Create(nil);
Cnx2.LoginPrompt:=false;
Cnx2.ConnectionString:=modulo2.Sistema.ConectUsu;
DataSet1.Connection:=Cnx2;
if nro=1 then begin
  DataSet1.CommandText:='SELECT nivel FROM usumodulos WHERE (modulo = 151) AND (codusu = ' + modulo2.Usuario.UsuNro + ');';
  DataSet1.Open;
  dato:=DataSet1.Fields[0].AsString;
    //if dato = '1' then begin
      //PopupMenu2.Items[1].Enabled:=true;
    //end
    //else begin
      //if PopupMenu2.Items[1].Enabled then PopupMenu2.Items[1].Enabled:=false;
    //end;
  DataSet1.Close;
end;
Cnx2.Close;
FreeAndNil(Cnx2);
end;

procedure TAM_Contribuyentes.FormActivate(Sender: TObject);
begin
if not (gh) then begin
  gh:=true;
  PC1.ActivePage:=PC1.Pages[0];

  grillaidentidad.Cells[0,0]:='Orden';                    grillaidentidad.Cells[1,0]:='Id.';
  grillaidentidad.Cells[2,0]:='Tipo Documento';           grillaidentidad.Cells[3,0]:='N° Documento';
  grillaidentidad.Cells[4,0]:='N° CUIM';                  grillaidentidad.Cells[5,0]:='Id Tipo';
  grillaidentidad.Cells[6,0]:='Activo';                   grillaidentidad.Cells[7,0]:='MDF';

  grilladomicilio.Cells[0,0]:='Orden';                    grilladomicilio.Cells[1,0]:='Id.';
  grilladomicilio.Cells[2,0]:='Barrio';                   grilladomicilio.Cells[3,0]:='Calle';
  grilladomicilio.Cells[4,0]:='Cod.Postal';               grilladomicilio.Cells[5,0]:='Altura';
  grilladomicilio.Cells[6,0]:='Piso';                     grilladomicilio.Cells[7,0]:='Depto';
  grilladomicilio.Cells[8,0]:='Manzana';                  grilladomicilio.Cells[9,0]:='Parcela';
  grilladomicilio.Cells[10,0]:='Casa';                    grilladomicilio.Cells[11,0]:='Sector';
  grilladomicilio.Cells[12,0]:='Torre';                   grilladomicilio.Cells[13,0]:='Lt.Rural';
  grilladomicilio.Cells[14,0]:='Domicilio Textual';       grilladomicilio.Cells[15,0]:='Id Barrio';
  grilladomicilio.Cells[16,0]:='Id Calle';                grilladomicilio.Cells[17,0]:='MDF';

  grillactabancos.Cells[0,0]:='Orden';
  grillactabancos.Cells[1,0]:='IdCuenta';                 grillactabancos.Cells[2,0]:='IdContrib';
  grillactabancos.Cells[3,0]:='Banco';                    grillactabancos.Cells[4,0]:='TipoCuenta';
  grillactabancos.Cells[5,0]:='Nro.Sucursal';             grillactabancos.Cells[6,0]:='Nro.Cuenta';
  grillactabancos.Cells[7,0]:='Credito';                  grillactabancos.Cells[8,0]:='Confirmado';
  grillactabancos.Cells[9,0]:='Baja';                     grillactabancos.Cells[10,0]:='UsuModif';
  grillactabancos.Cells[11,0]:='FechaModif';              grillactabancos.Cells[12,0]:='IdBanco';
  grillactabancos.Cells[13,0]:='IdTipoCuenta';            grillactabancos.Cells[14,0]:='MDF';
  grillactabancos.Cells[15,0]:='NroCBU';

  grilladebitos.Cells[0,0]:='Orden';
  grilladebitos.Cells[1,0]:='Id';                         grilladebitos.Cells[2,0]:='IdCuenta';
  grilladebitos.Cells[3,0]:='Rubro';                      grilladebitos.Cells[4,0]:='SubRubro';
  grilladebitos.Cells[5,0]:='CodigoObjeto';               grilladebitos.Cells[6,0]:='Identificador';
  grilladebitos.Cells[7,0]:='Confirmado';                 grilladebitos.Cells[8,0]:='Baja';
  grilladebitos.Cells[9,0]:='UsuModif';                   grilladebitos.Cells[10,0]:='FechaModif';
  grilladebitos.Cells[11,0]:='idRubro';                   grilladebitos.Cells[12,0]:='IdSubRubro';
  grilladebitos.Cells[13,0]:='MDF';                       grilladebitos.Cells[14,0]:='PlanPago';

  CargaCombos(1);         CargaCombos(2);
  CargaCombos(3);         CargaCombos(4);
  CargaCombos(5);         CargaCombos(6);
  CargaCombos(7);         CargaCombos(8);
  CargaCombos(9);         CargaCombos(11);
  VerArea(1);
  FechaAct:=LeftStr(Fec_System(),10);
  L3.Text:=strcut(fec_system(),10);
  JvGroupBox2.Enabled:=false;
  GUARDA.Enabled:=false;
  ODNI.SetFocus;
end;
end;

procedure TAM_Contribuyentes.CargaCombos(nro:byte);
begin
  if nro=1 then Modulo2.CargaCombo(L4,'SELECT DetalleSexo,Sexo FROM TipoSexo',modulo2.Sistema.ConectSeg,false);
  if nro=2 then Modulo2.CargaCombo(L6,'SELECT Detalle, Id FROM TipoPersona',modulo2.Sistema.ConectSeg,false);
  if nro=3 then Modulo2.CargaCombo(L8,'SELECT desc_est_civil,est_civil FROM TipoEstadoCivil'
     + ' order by desc_est_civil',modulo2.Sistema.ConectSeg,false);
  if nro=4 then Modulo2.CargaCombo(H3,'SELECT Abreviatura,Tpo_Doc FROM TipoDocumento '
     + 'order by Abreviatura',modulo2.Sistema.ConectSeg,false);
  if nro=5 then Modulo2.CargaCombo(D3,'SELECT Detalle,Id FROM Barrios ORDER BY Detalle',modulo2.Sistema.ConectSeg,false);
  if nro=6 then Modulo2.CargaCombo(D4,'SELECT Descrip,Calle FROM CALLES WHERE (Dpto_Polit=1) AND'
     + ' (LOCALIDAD=1) ORDER BY DESCRIP',modulo2.Sistema.ConectUsu,false);
  if nro=7 then Modulo2.CargaCombo(L11,'SELECT Descripcion,Id FROM TipoNacionalidad ORDER BY Descripcion',modulo2.Sistema.ConectSeg,false);
  if nro=8 then Modulo2.CargaCombo(banco,'SELECT Banco,Id,Nrosucursal FROM Bancos',modulo2.Sistema.ConectSeg,false);
  if nro=9 then Modulo2.CargaCombo(tipoctabanco,'SELECT TipoCuenta,Id FROM TipoCuentaBancaria ORDER BY 1',modulo2.Sistema.ConectSeg,false);
  if nro=10 then Modulo2.CargaCombo(tipoctabanco,'SELECT TipoCuenta,Id FROM TipoCuentaBancaria WHERE ID=2',modulo2.Sistema.ConectSeg,false);
  if nro=11 then Modulo2.CargaCombo(elsubrubro,'SELECT * FROM [SeguridadWEB].[dbo].[Fn_ObtenerRubros] ()',modulo2.Sistema.ConectUsu,false);
  //if nro=11 then Modulo2.CargaCombo(elsubrubro,'SELECT descripcion,idSubRubro from RubrosSub '
  //+ ' where idSubRubro in (1120501,1111201,1111601,1111104) order by 1',modulo2.Sistema.ConectUsu,false);
end;

procedure TAM_Contribuyentes.cbuExit(Sender: TObject);
begin
if banco.ItemIndex >= 1 then begin
   if  Length(cbu.Text) < 22 then begin
     messagedlg('ERROR!! Este Campo CBU esta incompleto, por favor verificar que contenga los 22 caracteres exigidos...',mterror,[mbok],0);
     cbu.SetFocus;
   end
   else begin
       agrega4.Enabled:=true;
       agrega4.SetFocus;
   end;
end
else begin
       agrega4.Enabled:=true;
       agrega4.SetFocus;
end;
end;

procedure TAM_Contribuyentes.cbuKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    if trim(grillactabancos.Cells[14,grillactabancos.Row])='X' then  begin
      agrega4.Enabled:=false;
      modifica4.SetFocus;
    end
    else begin
     agrega4.Enabled:=true;
     //escredito.SetFocus;
     agrega4.SetFocus;
    end;
  end
  else begin
    if ( StrScan('0123456789',Key) <> nil ) or
        ( Key = Char(VK_BACK) ) then
     else Key := #0;
  end;
end;

procedure TAM_Contribuyentes.Consultas(nro:byte);
var t,gf,h:integer;
    txt,valor:string;
    texto:widestring;
    Cnx2:TADOConnection;
begin
Cnx2:=TADOConnection.Create(nil);
Cnx2.LoginPrompt:=false;
if nro=0 then begin
  try
    DataSet1.Close;
    Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
    DataSet1.Connection:=Cnx2;
    DataSet1.CommandText:='EXEC darmaximonro NroCarnetEstMedido ';
    DataSet1.Open;
    nrocarnet:=trim(DataSet1.Fields[0].AsString);
    DataSet1.Close;
  except
    on E: Exception do EnviarLog(Self.Caption, E.Message);
  end;
end;
if nro=1 then  begin
  CARNET.Close;
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  CARNET.Connection:=Cnx2;
  DataSource2.DataSet:=CARNET;
  CARNET.CommandText:='exec SPP_ContribuyentesCarnetEstMedido ' + trim(l1.Text);
  CARNET.Open;
  if CARNET.IsEmpty then agregacarnet.Enabled:=true
    else agregacarnet.Enabled:=false;
  CARNET.Connection:=nil;
end;
if nro=2 then  begin
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.Connection:=Cnx2;
  DataSet1.CommandText:='SPP_ContribuyentesDebitos ' + laidcuenta + ' , 2';
  DataSet1.Open;
  for t:=1 to DataSet1.RecordCount do begin
    grilladebitos.RowCount:=t+1;
    DataSet1.RecNo:=t;
    grilladebitos.Cells[0,t]:=StrFill(IntToStr(t),2,'0','');
    grilladebitos.Cells[1,t]:=DataSet1.Fields[0].AsString;
    grilladebitos.Cells[2,t]:=DataSet1.Fields[1].AsString;
    grilladebitos.Cells[3,t]:=DataSet1.Fields[2].AsString;
    grilladebitos.Cells[4,t]:=DataSet1.Fields[3].AsString;
    grilladebitos.Cells[5,t]:=DataSet1.Fields[4].AsString;
    grilladebitos.Cells[6,t]:=DataSet1.Fields[5].AsString;
    grilladebitos.Cells[7,t]:=DataSet1.Fields[6].AsString;
    grilladebitos.Cells[8,t]:=DataSet1.Fields[7].AsString;
    grilladebitos.Cells[9,t]:=DataSet1.Fields[8].AsString;
    grilladebitos.Cells[10,t]:=DataSet1.Fields[9].AsString;
    grilladebitos.Cells[11,t]:=DataSet1.Fields[10].AsString;
    grilladebitos.Cells[12,t]:=DataSet1.Fields[11].AsString;
    grilladebitos.Cells[14,t]:=DataSet1.Fields[12].AsString;
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;
if nro=3 then begin
  OBJETOS.Close;
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  OBJETOS.Connection:=Cnx2;
  if TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111104' then
     texto:='SPP_ContribuyentesDebitosObjetos 0 , ' + trim(elcodobjeto.Text)+ ', null';
  if TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111201' then
    texto:='SPP_ContribuyentesDebitosObjetos 11112,' + chr(39) + trim(elcodobjeto.Text) + chr(39) + ',null';
  if (TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111601') or
     (TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1120501') then
     texto:='SPP_ContribuyentesDebitosObjetos 11205, null ,' + trim(elcodobjeto.Text);
  OBJETOS.CommandText:=texto;
  OBJETOS.Open;
   if strtoint(OBJETOS.Fields[0].AsString) > 0 then begin
         label38.Caption:=OBJETOS.Fields[0].AsString;
   end
   else begin
     if TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111104' then begin
        messagedlg('ERROR!! Este Nro Plan de Pago No Existe  o No Presenta Deuda, y/o esta Dado de Baja. Imposible su Carga..',mterror,[mbok],0);
        valor:='No';
        label38.Caption:='Objeto inexistente';
        elcodobjeto.Clear;
        elcodobjeto.SetFocus;
     end;
     if TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111201' then begin
         messagedlg('ERROR!! Esta Patente No Existe o ya posee un Debito asociado.. Imposible su Carga..',mterror,[mbok],0);
         valor:='No';
         label38.Caption:='Objeto inexistente';
         elcodobjeto.Clear;
         elcodobjeto.SetFocus;
     end;
     if (TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111601') or
        (TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1120501') then begin
           messagedlg('ERROR!! Este Inmueble NO esta Habilitado. Imposible su Carga..',mterror,[mbok],0);
           valor:='No';
           label38.Caption:='Objeto inexistente';
           elcodobjeto.Clear;
           elcodobjeto.SetFocus;
        end;
    {  if OBJETOS.Fields[0].AsString = '0' then t:=0;
      if OBJETOS.Fields[0].AsString = '1' then t:=1;
      if OBJETOS.Fields[0].AsString = '2' then t:=2;
      if OBJETOS.Fields[0].AsString = '3' then t:=3;
      case t of
        0:begin
           messagedlg('ERROR!! Esta Patente ya posee un Debito asociado.. Imposible su Carga..',mterror,[mbok],0);
           valor:='No';
           label38.Caption:='...';
           elcodobjeto.Clear;
           elcodobjeto.SetFocus;
        end;
        1:begin
           messagedlg('ERROR!! Esta Patente esta dada de BAJA y/o NO existe en nuestra Base de Datos.. Imposible su Carga..',mterror,[mbok],0);
           valor:='No';
           label38.Caption:='...';
           elcodobjeto.Clear;
           elcodobjeto.SetFocus;
        end;
        2:begin
           messagedlg('ERROR!! Este Inmueble NO esta Habilitado. Imposible su Carga..',mterror,[mbok],0);
           valor:='No';
           label38.Caption:='...';
           elcodobjeto.Clear;
           elcodobjeto.SetFocus;
        end;
        3:begin
           messagedlg('ERROR!! Este Nro Plan de Pago No Presenta Deuda y/o esta Dado de Baja. Imposible su Carga..',mterror,[mbok],0);
           valor:='No';
           label38.Caption:='...';
           elcodobjeto.Clear;
           elcodobjeto.SetFocus;
        end;
      end;}
   end;
  OBJETOS.Connection:=nil;
end;
if nro=4 then  begin
  if laidcuenta <> '?' then begin
    Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
    DataSet1.Connection:=Cnx2;
    DataSet1.CommandText:='SPP_ContribuyentesDebitos ' + laidcuenta + ' , 2';
    DataSet1.Open;
    for t:=1 to DataSet1.RecordCount do begin
      grilladebitos.RowCount:=t+1;
      DataSet1.RecNo:=t;
      grilladebitos.Cells[0,t]:=StrFill(IntToStr(t),2,'0','');
      grilladebitos.Cells[1,t]:=DataSet1.Fields[0].AsString;
      grilladebitos.Cells[2,t]:=DataSet1.Fields[1].AsString;
      grilladebitos.Cells[3,t]:=DataSet1.Fields[2].AsString;
      grilladebitos.Cells[4,t]:=DataSet1.Fields[3].AsString;
      grilladebitos.Cells[5,t]:=DataSet1.Fields[4].AsString;
      grilladebitos.Cells[6,t]:=DataSet1.Fields[5].AsString;
      grilladebitos.Cells[7,t]:=DataSet1.Fields[6].AsString;
      grilladebitos.Cells[8,t]:=DataSet1.Fields[7].AsString;
      grilladebitos.Cells[9,t]:=DataSet1.Fields[8].AsString;
      grilladebitos.Cells[10,t]:=DataSet1.Fields[9].AsString;
      grilladebitos.Cells[11,t]:=DataSet1.Fields[10].AsString;
      grilladebitos.Cells[12,t]:=DataSet1.Fields[11].AsString;
      grilladebitos.Cells[14,t]:=DataSet1.Fields[12].AsString;
    end;
    DataSet1.Close;
    DataSet1.Connection:=nil;
  end
  else begin
    elsubrubro.SetFocus;
  end;
end;
if nro= 5 then  begin
  CONTRIBCTASBANCOS.Close;
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  CONTRIBCTASBANCOS.Connection:=Cnx2;
  CONTRIBCTASBANCOS.CommandText:='SPP_ContribuyentesCuentasBanco ' + trim(l1.Text) + ' , 2';
  CONTRIBCTASBANCOS.Open;
  if not CONTRIBCTASBANCOS.IsEmpty then BEGIN
    for t:=1 to CONTRIBCTASBANCOS.RecordCount do begin
     grillactabancos.RowCount:=t+1;
     CONTRIBCTASBANCOS.RecNo:=t;
     grillactabancos.Cells[0,t]:=StrFill(IntToStr(t),2,'0','');
     grillactabancos.Cells[1,t]:=CONTRIBCTASBANCOS.Fields[0].AsString;
     grillactabancos.Cells[2,t]:=CONTRIBCTASBANCOS.Fields[1].AsString;
     grillactabancos.Cells[3,t]:=CONTRIBCTASBANCOS.Fields[2].AsString;
     grillactabancos.Cells[4,t]:=CONTRIBCTASBANCOS.Fields[3].AsString;
     grillactabancos.Cells[5,t]:=CONTRIBCTASBANCOS.Fields[4].AsString;
     grillactabancos.Cells[6,t]:=CONTRIBCTASBANCOS.Fields[5].AsString;
     grillactabancos.Cells[7,t]:=CONTRIBCTASBANCOS.Fields[6].AsString;
     grillactabancos.Cells[8,t]:=CONTRIBCTASBANCOS.Fields[7].AsString;
     grillactabancos.Cells[9,t]:=CONTRIBCTASBANCOS.Fields[8].AsString;
     grillactabancos.Cells[10,t]:=CONTRIBCTASBANCOS.Fields[9].AsString;
     grillactabancos.Cells[11,t]:=CONTRIBCTASBANCOS.Fields[10].AsString;
     grillactabancos.Cells[12,t]:=CONTRIBCTASBANCOS.Fields[11].AsString;
     grillactabancos.Cells[13,t]:=CONTRIBCTASBANCOS.Fields[12].AsString;
     grillactabancos.Cells[15,t]:=CONTRIBCTASBANCOS.Fields[13].AsString;
    end;
  END;
  CONTRIBCTASBANCOS.Connection:=nil;
end;
if nro = 6 then begin
  yaexiste:=false;
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.Connection:=Cnx2;
  DataSet1.CommandText:='SELECT top 1 Contribuyentes.Contribuyente FROM ContribuyentesIdentidad '
  +' inner join Contribuyentes on Contribuyentes.Id = ContribuyentesIdentidad.IdContribuyente'
  +' where (ContribuyentesIdentidad.nrodni='+ trim(H4.Text) +') and '
  +' (ContribuyentesIdentidad.tipodocumento='+TComboClase(H3.Items.Objects[H3.ItemIndex]).Param1 +') order by 1 desc';
  DataSet1.Open;
  if not DataSet1.IsEmpty then begin
    yaexiste:=true;
    nomb:=DataSet1.Fields[0].AsString;;
  end
  else yaexiste:=false;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;
if nro = 7 then begin
  if trim(ONOMB.Text) = '' then txt:='NULL'
    else txt:=chr(39) + trim(ONOMB.Text) + chr(39);
  if trim(ODNI.Text) = '' then txt:=txt + ' , NULL'
    else txt:=txt + ' , ' + chr(39) + trim(ODNI.Text) + chr(39);
  if trim(OCUIT.Text) = '' then txt:=txt + ' , NULL'
    else txt:=txt + ' , ' + chr(39) + trim(OCUIT.Text) + chr(39);
  CONTRIBUYENTE.Close;
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  CONTRIBUYENTE.Connection:=Cnx2;
  DataSource1.DataSet:=CONTRIBUYENTE;
  CONTRIBUYENTE.CommandText:='EXEC SPP_Contribuyentes ' + txt;
  CONTRIBUYENTE.Open;
  if CONTRIBUYENTE.IsEmpty then BEGIN
    gf:=messagedlg('Contribuyente Inexistente...' + chr(13) + 'Desea darlo de Alta?',mtWarning,[mbok,mbCancel],0);
    if gf = 1 then begin
      pc1.Pages[1].TabVisible:=true;
      pc1.Pages[2].TabVisible:=true;
      pc1.Pages[3].TabVisible:=true;
      PC1.ActivePage:=PC1.Pages[1];
      if trim(odni.Text) <> '' then  h4.Text:=odni.Text;
      l2.SetFocus;
    end
    else begin
      PC1.ActivePage:=PC1.Pages[0];
      odni.SetFocus;
    end;
  END;
  CONTRIBUYENTE.Connection:=nil;
end;
if nro = 8 then begin
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.Connection:=Cnx2;
  DataSet1.CommandText:='SELECT Id, TipoDocumento, NroDni, CUIM, IdTipoDocumento, '
   + 'Activo FROM dbo.Fn_ContribuyentesIdentidad(' + L1.Text + ')';
  DataSet1.Open;
  for t:=1 to DataSet1.RecordCount do begin
    grillaidentidad.RowCount:=t+1;
    DataSet1.RecNo:=t;
    grillaidentidad.Cells[0,t]:=StrFill(IntToStr(t),2,'0','');
    grillaidentidad.Cells[1,t]:=DataSet1.Fields[0].AsString;
    grillaidentidad.Cells[2,t]:=DataSet1.Fields[1].AsString;
    grillaidentidad.Cells[3,t]:=DataSet1.Fields[2].AsString;
    grillaidentidad.Cells[4,t]:=DataSet1.Fields[3].AsString;
    grillaidentidad.Cells[5,t]:=DataSet1.Fields[4].AsString;
    if DataSet1.Fields[5].AsBoolean then grillaidentidad.Cells[6,t]:='1'
    else grillaidentidad.Cells[6,t]:='0';
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
  elcuim:=grillaidentidad.Cells[4,1];
end;
if nro = 9 then begin
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.Connection:=Cnx2;
  DataSet1.CommandText:='SELECT Id,Barrio,Calle,CodPostal,Altura,Piso,Dpto,Mz,Parcela,Casa,Sector,'
  + 'Torre,LoteRural,Domicilio,CodBar,CodCal FROM dbo.Fn_ContribuyentesDomicilios(' + L1.Text + ')';
  DataSet1.Open;
  for t:=1 to DataSet1.RecordCount do begin
    grilladomicilio.RowCount:=t+1;
    DataSet1.RecNo:=t;
    grilladomicilio.Cells[0,t]:=StrFill(IntToStr(t),2,'0','');
    grilladomicilio.Cells[1,t]:=DataSet1.Fields[0].AsString;
    grilladomicilio.Cells[2,t]:=DataSet1.Fields[1].AsString;
    grilladomicilio.Cells[3,t]:=DataSet1.Fields[2].AsString;
    grilladomicilio.Cells[4,t]:=DataSet1.Fields[3].AsString;
    grilladomicilio.Cells[5,t]:=DataSet1.Fields[4].AsString;
    grilladomicilio.Cells[6,t]:=DataSet1.Fields[5].AsString;
    grilladomicilio.Cells[7,t]:=DataSet1.Fields[6].AsString;
    grilladomicilio.Cells[8,t]:=DataSet1.Fields[7].AsString;
    grilladomicilio.Cells[9,t]:=DataSet1.Fields[8].AsString;
    grilladomicilio.Cells[10,t]:=DataSet1.Fields[9].AsString;
    grilladomicilio.Cells[11,t]:=DataSet1.Fields[10].AsString;
    grilladomicilio.Cells[12,t]:=DataSet1.Fields[11].AsString;
    grilladomicilio.Cells[13,t]:=DataSet1.Fields[12].AsString;
    grilladomicilio.Cells[14,t]:=DataSet1.Fields[13].AsString;
    grilladomicilio.Cells[15,t]:=DataSet1.Fields[14].AsString;
    grilladomicilio.Cells[16,t]:=DataSet1.Fields[15].AsString;
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;
if nro = 10 then begin
  Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.Connection:=Cnx2;
  DataSet1.CommandText:='SELECT detalle FROM Fn_ElDomicilioContribuyente (' + L1.Text + ')';
  DataSet1.Open;
  if not DataSet1.IsEmpty  then begin
    ld.Caption:=trim(DataSet1.Fields[0].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;

Cnx2.Close;
FreeAndNil(Cnx2);
end;

function TAM_Contribuyentes.VerificaGrilla(grilla:string):boolean;
var j,a:integer;
begin
Result:=true;
for j:=1 to grilladebitos.RowCount-1 do begin
    if trim(grilladebitos.Cells[1,j])='' then begin
      Result:=false;
      break;
    end;
end;
end;

procedure TAM_Contribuyentes.ImprimeConfirmacion();
var ReporteRst: TfrxReport;
    frxDB1,frxDB2: TfrxDBDataset;
    DtsObjetos, DtsGrilla2:TClientDataSet;
    t,y,gf:integer;
begin
  try
    DtsObjetos:=TClientDataSet.Create(nil);
    DtsObjetos.FieldDefs.Clear;
    DtsObjetos.Active:=False;
    with DtsObjetos.FieldDefs.AddFieldDef do begin
      Name:='Orden';
      DataType:=ftInteger;
    end;
    with DtsObjetos.FieldDefs.AddFieldDef do begin
      Name:='Impuesto';
      DataType:=ftString;Size:=60;
    end;
    with DtsObjetos.FieldDefs.AddFieldDef do begin
      Name:='Objeto';
      DataType:=ftString;Size:=60;
    end;
    with DtsObjetos.FieldDefs.AddFieldDef do begin
      Name:='Detalle';
      DataType:=ftString;Size:=60;
    end;
    DtsObjetos.CreateDataSet;
    DtsObjetos.Active:=True;
    y:=1;
     for t:=1 to grilladebitos.RowCount-1 do begin
      if grilladebitos.Cells[1,t] = '?' then begin
        DtsObjetos.Append;
        DtsObjetos.FieldValues['Orden']:=y;
        DtsObjetos.FieldValues['Impuesto']:=grilladebitos.Cells[3,t];
        DtsObjetos.FieldValues['Objeto']:=grilladebitos.Cells[4,t];
        DtsObjetos.FieldValues['Detalle']:=grilladebitos.Cells[5,t] + ' / ' + grilladebitos.Cells[6,t];
        DtsObjetos.Post;
        y:=y+1;
      end;
      end;
  except
    on E: Exception do EnviarLog(Self.Caption, E.Message);
  end;

//Imprimiendo
try
  ReporteRst:=TfrxReport.Create(nil);
  frxDB1:=TfrxDBDataset.Create(nil);
  frxDB1.DataSet:=DtsObjetos;
  frxDB1.UserName:='frxDBDataset1';
  ReporteRst.DataSets.Add(frxDB1);

  ReporteRst.LoadFromFile(Modulo2.Sistema.Path_todo + '\Reportes\Alta_DebitosNew.fr3');
  TfrxPictureView(ReporteRst.FindObject('Logo1')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
  ReporteRst.Variables['ElqueSuscribe'] := QuotedStr(l2.Text);
  ReporteRst.Variables['NomTitular'] := QuotedStr(l5.Text);
  ReporteRst.Variables['NroDni'] := QuotedStr(eldni);
  ReporteRst.Variables['IdCuenta'] := QuotedStr(ctadeb);
  if laidcuenta <> '?' then begin
    ReporteRst.Variables['NroPlan'] := QuotedStr(trim(nrocuenta.Text));
    ReporteRst.Variables['NomRespo'] := QuotedStr(trim(tipoctabanco.Text));
    ReporteRst.Variables['DomTitular'] := QuotedStr(trim(banco.Text));
    ReporteRst.Variables['CuimTitular'] := QuotedStr(trim(nrosucursal.Text));
  end
  else begin
   for gf:=1 to grillactabancos.RowCount-1 do begin
    if grillactabancos.Cells[1,gf] = '?' then begin
      ReporteRst.Variables['NroPlan'] := QuotedStr(trim(grillactabancos.Cells[6,gf]));
      ReporteRst.Variables['NomRespo'] := QuotedStr(trim(grillactabancos.Cells[4,gf]));
      ReporteRst.Variables['DomTitular'] := QuotedStr(trim(grillactabancos.Cells[3,gf]));
      ReporteRst.Variables['CuimTitular'] := QuotedStr(trim(grillactabancos.Cells[5,gf]));
    end;
   end;
  end;
  ReporteRst.Variables['LugarFecha'] := QuotedStr('FORMOSA, ' + FechaToStr(FechaAct,'DDMMAAAA'));
  ReporteRst.Variables['NroUsuario'] := QuotedStr(Modulo2.Usuario.UsuNro);
  ReporteRst.ShowReport;
except
  on E: Exception do EnviarLog(Self.Caption, 'Error en Reporte ' + E.Message);
end;
FreeAndNil(frxDB1);
FreeAndNil(ReporteRst);
FreeAndNil(DtsObjetos);
end;

procedure TAM_Contribuyentes.BORRAClick(Sender: TObject);
var p,u:integer;
begin
  for p:=0 to grillaidentidad.ColCount-1 do
    for u:=1 to grillaidentidad.RowCount-1 do grillaidentidad.Cells[p,u]:='';
  grillaidentidad.RowCount:=2;

  for p:=0 to grilladomicilio.ColCount-1 do
    for u:=1 to grilladomicilio.RowCount-1 do grilladomicilio.Cells[p,u]:='';
  grilladomicilio.RowCount:=2;

  for p:=0 to grillactabancos.ColCount-1 do
    for u:=1 to grillactabancos.RowCount-1 do grillactabancos.Cells[p,u]:='';
  grillactabancos.RowCount:=2;

  for p:=0 to grilladebitos.ColCount-1 do
    for u:=1 to grilladebitos.RowCount-1 do grilladebitos.Cells[p,u]:='';
  grilladebitos.RowCount:=2;

  L1.Clear;           L2.Clear;
  L3.Clear;           L4.ItemIndex:=0;
  L5.Clear;           L6.ItemIndex:=0;
  L7.Clear;           L8.ItemIndex:=0;
  L9.Clear;           L10.Checked:=false;

  H1.Clear;           H2.Text:='00-00000000-0';
  H3.ItemIndex:=0;    H4.Clear;

  D3.ItemIndex:=0;    D4.ItemIndex:=0;

  D5.Text:='3600';
  D6.Clear;           D7.Clear;
  D8.Clear;           D9.Clear;
  D10.Clear;          D11.Clear;
  D12.Clear;          D13.Clear;
  D14.Clear;          D15.Clear;

  banco.ItemIndex:=0;
  tipoctabanco.ItemIndex:=0;
  nrosucursal.Clear;
  cbu.Clear;
  nrocuenta.Clear;
  escredito.Checked:=false;

  elsubrubro.ItemIndex:=0;
  elcodobjeto.Clear;
  label38.Caption:='...';

  C3.Checked:=false;  C4.Checked:=false;
  C5.Checked:=false;  C6.Checked:=false;
  c7.Checked:=false;  c8.Checked:=false;

  GUARDA.Enabled:=false;
  pc1.Pages[1].TabVisible:=false;
  pc1.Pages[2].TabVisible:=false;
  pc1.Pages[3].TabVisible:=false;
  pc1.Pages[4].TabVisible:=false;
  pc1.pages[5].TabVisible:=false;
  PC1.ActivePage:=PC1.Pages[0];
end;

procedure TAM_Contribuyentes.ODNIKeyPress(Sender: TObject; var Key: Char);
begin
  if key =#13 then begin
    Key := #0;
    ONOMB.SetFocus;
  end
  else begin
    if ( StrScan('0123456789',Key) <> nil ) or
        ( Key = Char(VK_BACK) ) then
     else Key := #0;
  end;
end;

procedure TAM_Contribuyentes.ONOMBKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  OCUIT.SetFocus;
end
else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ-/ °',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.OCUITKeyPress(Sender: TObject; var Key: Char);
begin
if key =#13 then begin
  Key := #0;
  BUSCAR.SetFocus;
end
else begin
  if ( StrScan('0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.BuscarClick(Sender: TObject);
begin
BORRA.Click;
Consultas(7);
end;

procedure TAM_Contribuyentes.Button1Click(Sender: TObject);
begin
//
end;

procedure TAM_Contribuyentes.L2KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L3.SetFocus;
end
else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ-/ °',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.L2Exit(Sender: TObject);
begin
if trim(L2.Text) = '' then begin
  messagedlg('Error!! No Existe Contribuyente sin nombre.',mterror,[mbok],0);
  L2.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L3KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L4.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L3Exit(Sender: TObject);
var fec:Tdatetime;
begin
try
 fec:=strtodate(trim(L3.Text));
except
  messagedlg('La fecha es Erronea!!',mterror,[mbok],0);
  L3.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L4KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L5.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L4Exit(Sender: TObject);
begin
if L4.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una opción válida',mterror,[mbok],0);
  L4.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L5KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L6.SetFocus;
end
else begin
  if ( StrScan('0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.L5Exit(Sender: TObject);
var ps:byte;
    vr:string;
    Cnx2:TADOConnection;
begin
  ps:=0;
  if (Length(trim(L5.Text)) = 11) then ps:=1;
  if (ps = 1) and (codigo_ver1(LeftStr(L5.Text,10),false) = L5.Text) then ps:=1
  else ps:=0;
  if (ps = 1) then begin
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
    DataSet1.Connection:=Cnx2;
    DataSet1.CommandText:=' select Contribuyente from Contribuyentes where (cuit='+ trim(l5.Text)+')';
    DataSet1.Open;
    vr:=trim(DataSet1.Fields[0].AsString);
    DataSet1.Close;
    DataSet1.Connection:=nil;
    Cnx2.Close;
    FreeAndNil(Cnx2);
    if vr <> '' then begin
      if (vr = trim(l2.Text)) and (c6.Checked) then begin
        l6.SetFocus;
      end
      else begin
       messagedlg('ATENCION!! El CUIT ingresado hace referencia a la siguiente persona:'
       + chr(13) + vr + chr(13) + 'Por favor Ingrese otro valor..' ,mtWarning,[mbok],0);
       ps:=2;
       l5.SetFocus;
      end;
    end;
  end;
  if (trim(trim(L5.Text)) <> '') then begin
   if (Length(trim(L5.Text)) < 11) then begin
     messagedlg('ERROR!! El CUIT ingresado no es Válido',mterror,[mbok],0);
     L5.SetFocus;
   end;
  end;
end;

procedure TAM_Contribuyentes.L7KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L8.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789@.-_',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.L6KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L7.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L6Exit(Sender: TObject);
begin
if L6.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una opción válida',mterror,[mbok],0);
  L6.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L7Exit(Sender: TObject);
begin
//if not emailValido(L7.Text) then begin
//  messagedlg('Debe ingresar un MAIL válido!!',mterror,[mbok],0);
//  L7.SetFocus;
//end;
end;

procedure TAM_Contribuyentes.L8KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L9.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L8Exit(Sender: TObject);
begin
if L8.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Estado Civil válido',mterror,[mbok],0);
  L8.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L9KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L11.SetFocus;
end
else begin
  if ( StrScan('()0123456789-/',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.L9Exit(Sender: TObject);
begin
//if trim(L9.Text) = '' then begin
//  messagedlg('Debe ingresar un N° de Teléfono válido!!',mterror,[mbok],0);
//  L9.SetFocus;
//end;
end;

procedure TAM_Contribuyentes.L10KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[2];
  H3.SetFocus;
end;
end;

procedure TAM_Contribuyentes.H3KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    H4.SetFocus;
  end;
  if key = #27 then begin
    PC1.ActivePage:=PC1.Pages[3];
    D3.SetFocus;
  end;
end;

procedure TAM_Contribuyentes.H3Exit(Sender: TObject);
begin
  if H3.ItemIndex < 0 then begin
    messagedlg('Error!! Seleccione un Tipo de Documento válido',mterror,[mbok],0);
    H3.SetFocus;
  end
  else H4.SetFocus;
end;

procedure TAM_Contribuyentes.H4KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    if Modifica2.Enabled then Modifica2.SetFocus
    else if Agrega2.Enabled then Agrega2.SetFocus;
  end
  else begin
    if ( StrScan('0123456789',Key) <> nil ) or
        ( Key = Char(VK_BACK) ) then
     else Key := #0;
  end;
end;

procedure TAM_Contribuyentes.H4Exit(Sender: TObject);
var Cnx2:TADOConnection;
begin
if (trim(H4.Text) = '') or (trim(H4.Text) = '0') then begin
  messagedlg('Error!! Ingrese un N° de Documento válido',mterror,[mbok],0);
  H4.SetFocus;
end
else begin
  consultas(6);
  if yaexiste then begin
    if (nomb = trim(l2.Text)) and (c6.Checked) then begin
       modifica2.SetFocus;
    end
    else begin
       messagedlg('Error!! Tipo y Nro Documento ya Existente para el Contribuyente'
       + chr(13)+ nomb + chr(13) + 'Ingrese Otro Valor de Documento',mterror,[mbok],0);
       H4.SetFocus;
    end;
  end
  else begin
    if trim(H1.Text) = '' then begin
      Cnx2:=TADOConnection.Create(nil);
      Cnx2.LoginPrompt:=false;
      Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
      DataSet1.Connection:=Cnx2;
      DataSet1.CommandText:='SELECT Cuim FROM Fn_TodosCUIM (' + H4.Text + ' , ' + chr(39)
      + TComboClase(L4.Items.Objects[L4.ItemIndex]).Param1 + chr(39) + ')';
      DataSet1.Open;
      H2.Text:=DataSet1.Fields[0].AsString;
      DataSet1.Close;
      Cnx2.Close;
      FreeAndNil(Cnx2);
    end;
  end;
end;
end;

procedure TAM_Contribuyentes.H4Change(Sender: TObject);
begin
if trim(H4.Text) <> '' then Agrega2.Enabled:=not Modifica2.Enabled
else Agrega2.Enabled:=false;
end;

procedure TAM_Contribuyentes.Agrega2Click(Sender: TObject);
var t:integer;

  function busca():boolean;
  var p:integer;
  begin
  Result:=true;
  if grillaidentidad.Cells[0,1] <> '' then begin
    for p:=1 to grillaidentidad.RowCount-1 do
      if (grillaidentidad.Cells[2,p]=trim(H3.Text))
      and (grillaidentidad.Cells[3,p]=trim(H4.Text)) then Result:=false;
  end;
  end;

  function buscaActivo():boolean;
  var p:integer;
  begin
  Result:=true;
  if grillaidentidad.Cells[0,1] <> '' then begin
    for p:=1 to grillaidentidad.RowCount-1 do begin
      if (grillaidentidad.Cells[6,p] = '1') then begin
        Result:=false;
        break;
      end;
    end;
  end;
  end;

begin
  if trim(H2.Text)='00-00000000-0' then begin
   messagedlg('ERROR!! CUIM Erroneo. Imposible la Carga de la identidad.',mterror,[mbok],0);
  end
  else begin
   if busca then begin
    if (grillaidentidad.RowCount > 1) and (trim(grillaidentidad.Cells[0,1])<>'') then begin
      t:=grillaidentidad.RowCount;
      grillaidentidad.RowCount:=grillaidentidad.RowCount+1;
    end
    else t:=1;
    grillaidentidad.Cells[0,t]:=completa1(inttostr(t),2,'0');
    grillaidentidad.Cells[1,t]:='?';
    grillaidentidad.Cells[2,t]:=trim(H3.Text);
    grillaidentidad.Cells[3,t]:=trim(H4.Text);
    grillaidentidad.Cells[4,t]:=trim(H2.Text);
    grillaidentidad.Cells[5,t]:=TComboClase(H3.Items.Objects[H3.ItemIndex]).Param1 ;
    if buscaActivo() then grillaidentidad.Cells[6,t]:='1'
     else grillaidentidad.Cells[6,t]:='0';
    grillaidentidad.Cells[7,t]:='X';
    if modifica2.Enabled = false then Agrega2.Enabled:=true;
    C4.Checked:=true;
    GUARDA.Enabled:=true;
   end
   else messagedlg('ERROR!! No Puede Existir dos N° DNI del mismo Tipo y Número!!',mterror,[mbok],0);
   H3.SetFocus;
   end;
end;

procedure TAM_Contribuyentes.Modifica2Click(Sender: TObject);
var t:integer;

  function busca(ht:integer):boolean;
    var p:integer;
  begin
    busca:=true;
  if grillaidentidad.Cells[0,1] <> '' then begin
    for p:=1 to grillaidentidad.RowCount-1 do
      if (grillaidentidad.Cells[2,p]=trim(H3.Text))
      and (grillaidentidad.Cells[3,p]=trim(H4.Text))
      and (ht<>p) then busca:=false;
  end;
  end;

begin
  if busca(grillaidentidad.Row) then begin
    if grillaidentidad.Cells[1,grillaidentidad.Row] <> '?' then begin
      grillaidentidad.Cells[7,grillaidentidad.Row]:='X';
    end;
    grillaidentidad.Cells[2,grillaidentidad.Row]:=trim(H3.Text);
    grillaidentidad.Cells[3,grillaidentidad.Row]:=trim(H4.Text);
    grillaidentidad.Cells[4,grillaidentidad.Row]:=trim(H2.Text);
    grillaidentidad.Cells[5,grillaidentidad.Row]:=TComboClase(H3.Items.Objects[H3.ItemIndex]).Param1 ;
    grillaidentidad.Cells[7,grillaidentidad.Row]:='X';
    Agrega2.Enabled:=false;
    Borra2.Enabled:=false;
    Modifica2.Enabled:=false;
    C4.Checked:=true;
    GUARDA.Enabled:=true;
  end
  else messagedlg('ERROR!! No Puede Existir dos N° DNI del mismo Tipo y Número!!',mterror,[mbok],0);
  H3.SetFocus;
end;

procedure TAM_Contribuyentes.Borra2Click(Sender: TObject);
var t,p,gf:integer;
var bt:boolean;
    Cnx2:TADOConnection;
begin
if trim(grillaidentidad.Cells[0,1])<>'' then begin
  gf:=1;t:=0;
  if grillaidentidad.Cells[1,grillaidentidad.Row] = '?' then t:=1
  else gf:=messagedlg('Desea Borrar la Identidad Seleccionado??' + chr(13)
       + 'Se eliminará definitivamente de la Base de Datos',mtWarning,[mbok,mbCancel],0);
  if gf = 1 then begin
    if t=0 then begin
      try
        Cnx2:=TADOConnection.Create(nil);
        Cnx2.LoginPrompt:=false;
        Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
        Com1:=TADOCommand.Create(nil);
        Com1.Connection:=Cnx2;
        Com1.CommandText:='SPB_ContribuyentesIdentidad ' + grillaidentidad.Cells[1,grillaidentidad.Row]
        + ' , ' + L1.Text
        + ' , ' + modulo2.Usuario.UsuNro
        + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
        + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
        Com1.Execute;
        messagedlg('Identidad Eliminada correctamente!',mtConfirmation,[mbok],0);
      except
        EnviarLog(self.Caption,Com1.CommandText);
        messagedlg('ERROR al Eliminar los datos',mterror,[mbok],0);
      end;
      FreeAndNil(Com1);
      Cnx2.Close;
      FreeAndNil(Cnx2);
    end;
    t:=grillaidentidad.Row;
    if (t >= 1) and (grillaidentidad.RowCount > 2) then begin
      THackStringGrid(grillaidentidad).DeleteRow(t);
    end
    else begin
      for t:=0 to grillaidentidad.ColCount-1 do grillaidentidad.Cells[t,1]:='';
    end;
  end;
  Borra2.Enabled:=false;
  Modifica2.Enabled:=false;
  H3.SetFocus;
end;
end;

procedure TAM_Contribuyentes.Quita2Click(Sender: TObject);
begin
H3.ItemIndex:=0;
H4.Clear;
H1.Clear;
H2.Text:='00-00000000-0';
Modifica2.Enabled:=false;
Borra2.Enabled:=false;
Agrega2.Enabled:=true;
Quita2.Enabled:=false;
end;

procedure TAM_Contribuyentes.D3KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D4.SetFocus;
end;
end;

procedure TAM_Contribuyentes.D3Exit(Sender: TObject);
begin
if D3.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Barrio válido',mterror,[mbok],0);
  D3.SetFocus;
end;
end;

procedure TAM_Contribuyentes.D4KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D5.SetFocus;
end;
end;

procedure TAM_Contribuyentes.D4Exit(Sender: TObject);
begin
if D4.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Barrio válido',mterror,[mbok],0);
  D4.SetFocus;
end;
end;

procedure TAM_Contribuyentes.D5KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    D6.SetFocus;
  end
  else begin
    if ( StrScan('0123456789',Key) <> nil ) or
        ( Key = Char(VK_BACK) ) then
     else Key := #0;
  end;
end;


procedure TAM_Contribuyentes.D5Exit(Sender: TObject);
var Cnx2:TADOConnection;
begin
  if Button4.Focused then
  else begin
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.Sistema.ConectUsu;
    DataSet1.Connection:=Cnx2;
    DataSet1.CommandText:='SELECT Postal FROM COD_POST WHERE (Codigo=' + D5.Text + ')';
    DataSet1.Open;
    if not DataSet1.IsEmpty then
      Label24.Caption:=DataSet1.Fields[0].AsString
    else begin
      messagedlg('Error!! Seleccione un Código Postal válido',mterror,[mbok],0);
      D5.SetFocus;
    end;
    DataSet1.Close;
    Cnx2.Close;
    FreeAndNil(Cnx2);
  end;
end;

procedure TAM_Contribuyentes.D6KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D7.SetFocus;
end
else begin
  if ( StrScan('0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D6Exit(Sender: TObject);
begin
if trim(D6.Text) = '' then begin
  messagedlg('Error!! Ingrese una Altura válida, o un 0 (cero) si no posee',mterror,[mbok],0);
  D6.SetFocus;
end;
end;

procedure TAM_Contribuyentes.D7KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D8.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D8KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D9.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D9KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D10.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.DBGrid1CellClick(Column: TColumn);
begin
  agregacarnet.Enabled:=false;
  print.Enabled:=true;
  Lc.Caption :=CARNET.Fields[2].AsString;
  Lcuit.Caption:=CARNET.Fields[3].AsString;
  ll.Caption:='Formosa';
  Lnro.Caption:=CARNET.Fields[5].AsString;
  lcodbar.Caption:=CARNET.Fields[4].AsString;
  nombfantasia.Text:=CARNET.Fields[6].AsString;
  domcomercio.Text:=CARNET.Fields[7].AsString;
  Consultas(10);
end;

procedure TAM_Contribuyentes.domcomercioKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  agregacarnet.SetFocus;
end;
end;

procedure TAM_Contribuyentes.elcodobjetoExit(Sender: TObject);
begin
//if (trim(elcodobjeto.Text) = '') or (trim(elcodobjeto.Text) = '0') then begin
//  messagedlg('Error!! Ingrese un Objeto..',mterror,[mbok],0);
//  elcodobjeto.SetFocus;
//end;
end;

procedure TAM_Contribuyentes.elcodobjetoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    consultas(3);
    if trim(label38.Caption)='Objeto inexistente' then begin
      //messagedlg('Error!! Este Objeto NO Existe para el Rubro Seleccionado...',mterror,[mbok],0);
      elcodobjeto.SetFocus;
      agrega5.Enabled:=false;
    end
    else begin
      if trim(grilladebitos.Cells[13,grilladebitos.Row])='X' then  begin
        agrega5.Enabled:=false;
        modifica5.SetFocus;
      end
      else begin
       agrega5.Enabled:=true;
       agrega5.SetFocus;
      end;
    end;
  end
  else begin
    if (elsubrubro.ItemIndex=0) or (elsubrubro.ItemIndex=1) or (elsubrubro.ItemIndex=3) then begin
        if ( StrScan('0123456789',Key) <> nil ) or
          ( Key = Char(VK_BACK) ) then
        else Key := #0;
    end
    else begin
        if ( StrScan('0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ',Key) <> nil ) or
          ( Key = Char(VK_BACK) ) then
        else Key := #0;
    end;
  end;
end;

procedure TAM_Contribuyentes.elsubrubroExit(Sender: TObject);
begin
if elsubrubro.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un SubRubro válido',mterror,[mbok],0);
  elsubrubro.SetFocus;
end;
end;

procedure TAM_Contribuyentes.elsubrubroKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  label37.Caption:='Nro. Plan Pago:';
//  if elsubrubro.ItemIndex=0 then label37.Caption:='Partida Nro:';
//  if elsubrubro.ItemIndex=1 then label37.Caption:='Nro. Plan Pago:';
//  if elsubrubro.ItemIndex=2 then label37.Caption:='Dominio:';
//  if elsubrubro.ItemIndex=3 then label37.Caption:='Partida Nro:';
  elcodobjeto.SetFocus;
end;
end;

procedure TAM_Contribuyentes.escreditoKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  if Agrega4.Enabled then Agrega4.SetFocus
    else modifica4.SetFocus;
end;
end;

procedure TAM_Contribuyentes.D10KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D11.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D11KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D12.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D12KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D13.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D13KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D14.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D14KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  D15.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.D15KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  if Modifica3.Enabled then Modifica3.SetFocus
  else
  begin
     Agrega3.Enabled:=true;
     if Agrega3.Enabled then Agrega3.SetFocus;
  end;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789.°',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_Contribuyentes.Agrega3Click(Sender: TObject);
var t:integer;

  function busca():boolean;
    var p:integer;
  begin
    busca:=true;
  if grilladomicilio.Cells[0,1] <> '' then begin
    for p:=1 to grilladomicilio.RowCount-1 do
      if (grilladomicilio.Cells[2,p]=trim(D3.Text))
      and (grilladomicilio.Cells[3,p]=trim(D4.Text))
      and (grilladomicilio.Cells[5,p]=trim(D6.Text))
      and (grilladomicilio.Cells[6,p]=trim(D7.Text))
      and (grilladomicilio.Cells[7,p]=trim(D8.Text))
      and (grilladomicilio.Cells[8,p]=trim(D9.Text))
      and (grilladomicilio.Cells[9,p]=trim(D10.Text))
      and (grilladomicilio.Cells[10,p]=trim(D11.Text))
      and (grilladomicilio.Cells[11,p]=trim(D12.Text))
      and (grilladomicilio.Cells[13,p]=trim(D14.Text)) then busca:=false;
  end;
  end;

begin
  if busca then begin
    if (grilladomicilio.RowCount > 1) and (trim(grilladomicilio.Cells[0,1])<>'') then begin
      t:=grilladomicilio.RowCount;
      grilladomicilio.RowCount:=grilladomicilio.RowCount+1;
    end
    else t:=1;
    grilladomicilio.Cells[0,t]:=completa1(inttostr(t),2,'0');
    grilladomicilio.Cells[1,t]:='?';
    grilladomicilio.Cells[2,t]:=trim(D3.Text);
    grilladomicilio.Cells[3,t]:=trim(D4.Text);
    grilladomicilio.Cells[4,t]:=trim(D5.Text);
    grilladomicilio.Cells[5,t]:=trim(D6.Text);
    grilladomicilio.Cells[6,t]:=trim(D7.Text);
    grilladomicilio.Cells[7,t]:=trim(D8.Text);
    grilladomicilio.Cells[8,t]:=trim(D9.Text);
    grilladomicilio.Cells[9,t]:=trim(D10.Text);
    grilladomicilio.Cells[10,t]:=trim(D11.Text);
    grilladomicilio.Cells[11,t]:=trim(D12.Text);
    grilladomicilio.Cells[12,t]:=trim(D13.Text);
    grilladomicilio.Cells[13,t]:=trim(D14.Text);
    grilladomicilio.Cells[14,t]:=trim(D15.Text);
    grilladomicilio.Cells[15,t]:=TComboClase(D3.Items.Objects[D3.ItemIndex]).Param1 ;
    grilladomicilio.Cells[16,t]:=TComboClase(D4.Items.Objects[D4.ItemIndex]).Param1;
    grilladomicilio.Cells[17,t]:='X';
    if modifica3.Enabled = false then Agrega3.Enabled:=false;
    C5.Checked:=true;
    GUARDA.Enabled:=true;
  end
  else messagedlg('ERROR!! No Puede Existir dos Domicilios iguales!!',mterror,[mbok],0);
  D3.SetFocus;
end;

procedure TAM_Contribuyentes.agrega4Click(Sender: TObject);
var t:integer;

  function busca():boolean;
    var p:integer;
  begin
    busca:=true;
  if grillactabancos.Cells[0,1] <> '' then begin
    for p:=1 to grillactabancos.RowCount-1 do
      if (grillactabancos.Cells[2,p]=trim(banco.Text))
      and (grillactabancos.Cells[3,p]=trim(tipoctabanco.Text))
      and (grillactabancos.Cells[5,p]=trim(nrosucursal.Text))
      and (grillactabancos.Cells[6,p]=trim(nrocuenta.Text))  then busca:=false;
  end;
  end;

begin
  if Verificacion(4) then begin
   if Verificacion(2) then begin
     if busca then begin
      if (grillactabancos.RowCount > 1) and (trim(grillactabancos.Cells[0,1])<>'') then begin
        t:=grillactabancos.RowCount;
        grillactabancos.RowCount:=grillactabancos.RowCount+1;
      end
      else t:=1;
        grillactabancos.Cells[0,t]:=completa1(inttostr(t),2,'0');
        grillactabancos.Cells[1,t]:='?';
        grillactabancos.Cells[2,t]:=trim(L1.Text);
        grillactabancos.Cells[3,t]:=trim(banco.Text);
        grillactabancos.Cells[4,t]:=trim(tipoctabanco.Text);
        grillactabancos.Cells[5,t]:=trim(nrosucursal.Text);
        grillactabancos.Cells[6,t]:=trim(nrocuenta.Text);
        if escredito.Checked  then  grillactabancos.Cells[7,t]:='True'
          else grillactabancos.Cells[7,t]:='False';
        grillactabancos.Cells[8,t]:='True';
        grillactabancos.Cells[9,t]:='null';
        grillactabancos.Cells[10,t]:='null';
        grillactabancos.Cells[11,t]:='null';
        grillactabancos.Cells[12,t]:=TComboClase(banco.Items.Objects[banco.ItemIndex]).Param1;
        grillactabancos.Cells[13,t]:=TComboClase(tipoctabanco.Items.Objects[tipoctabanco.ItemIndex]).Param1;
        grillactabancos.Cells[14,t]:='X';
        grillactabancos.Cells[15,t]:=trim(cbu.Text);
        if modifica4.Enabled = false then Agrega4.Enabled:=false;
        C7.Checked:=true;
        JvGroupBox2.Enabled:=true;
        GUARDA.Enabled:=true;
        //BusPlanes.SetFocus;
        elsubrubro.SetFocus;
     end
     else begin
       messagedlg('ERROR!! No Puede Existir dos Cuntas Bancarias iguales!!',mterror,[mbok],0);
       banco.SetFocus;
     end;
   end;
  end;
end;

procedure TAM_Contribuyentes.agrega5Click(Sender: TObject);
var t:integer;

  function busca():boolean;
    var p:integer;
  begin
    busca:=true;
  if grilladebitos.Cells[0,1] <> '' then begin
    for p:=1 to grilladebitos.RowCount-1 do
      //if (grilladebitos.Cells[3,p]=trim(elrubro.Text))
      if (grilladebitos.Cells[4,p]=trim(elsubrubro.Text))
      and (grilladebitos.Cells[5,p]=trim(elcodobjeto.Text)) then busca:=false;
  end;
  end;

begin
 if Verificacion(3) then begin
   if busca then begin
    if (grilladebitos.RowCount > 1) and (trim(grilladebitos.Cells[0,1])<>'') then begin
      t:=grilladebitos.RowCount;
      grilladebitos.RowCount:=grilladebitos.RowCount+1;
    end
    else t:=1;
    grilladebitos.Cells[0,t]:=completa1(inttostr(t),2,'0');
    grilladebitos.Cells[1,t]:='?';
    if (laidcuenta='?') or (laidcuenta='') then grilladebitos.Cells[2,t]:='?'
     else grilladebitos.Cells[2,t]:=laidcuenta;
    grilladebitos.Cells[3,t]:=trim(elsubrubro.Text);
    grilladebitos.Cells[4,t]:=trim(elsubrubro.Text);
    grilladebitos.Cells[5,t]:=trim(elcodobjeto.Text);
    grilladebitos.Cells[6,t]:=trim(label38.Caption);
    grilladebitos.Cells[7,t]:='True';
    grilladebitos.Cells[8,t]:='False';
    grilladebitos.Cells[11,t]:=TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1;//TComboClase(elrubro.Items.Objects[elrubro.ItemIndex]).Param1;
    grilladebitos.Cells[12,t]:=TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1;
    grilladebitos.Cells[13,t]:='X';
    if TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1 = '1111104' then grilladebitos.Cells[14,t]:='True'
      else grilladebitos.Cells[14,t]:='False';
    if modifica5.Enabled = false then Agrega5.Enabled:=false;
    C8.Checked:=true;
    GUARDA.Enabled:=true;
    elsubrubro.ItemIndex:=0;
    elcodobjeto.Clear;
    label38.Caption:='...';
    elsubrubro.SetFocus;
  end
  else messagedlg('ERROR!! No Puede Existir dos Cuntas Debitos iguales!!',mterror,[mbok],0);
     elsubrubro.SetFocus;
  end;
end;

procedure TAM_Contribuyentes.Modifica3Click(Sender: TObject);
var t:integer;

  function busca(ht:integer):boolean;
  var p:integer;
  begin
  busca:=true;
  if grilladomicilio.Cells[0,1] <> '' then begin
    for p:=1 to grilladomicilio.RowCount-1 do
      if (grilladomicilio.Cells[2,p]=trim(D3.Text))
      and (grilladomicilio.Cells[3,p]=trim(D4.Text))
      and (grilladomicilio.Cells[5,p]=trim(D6.Text))
      and (grilladomicilio.Cells[6,p]=trim(D7.Text))
      and (grilladomicilio.Cells[7,p]=trim(D8.Text))
      and (grilladomicilio.Cells[8,p]=trim(D9.Text))
      and (grilladomicilio.Cells[9,p]=trim(D10.Text))
      and (grilladomicilio.Cells[10,p]=trim(D11.Text))
      and (grilladomicilio.Cells[11,p]=trim(D12.Text))
      and (grilladomicilio.Cells[13,p]=trim(D14.Text))
      and (ht<>p) then busca:=false;
  end;
  end;

begin
if busca(grilladomicilio.Row) then begin
  if grilladomicilio.Cells[1,grilladomicilio.Row] <> '?' then begin
    grilladomicilio.Cells[17,grilladomicilio.Row]:='X';
  end;
  grilladomicilio.Cells[2,grilladomicilio.Row]:=trim(D3.Text);
  grilladomicilio.Cells[3,grilladomicilio.Row]:=trim(D4.Text);
  grilladomicilio.Cells[4,grilladomicilio.Row]:=trim(D5.Text);
  grilladomicilio.Cells[5,grilladomicilio.Row]:=trim(D6.Text);
  grilladomicilio.Cells[6,grilladomicilio.Row]:=trim(D7.Text);
  grilladomicilio.Cells[7,grilladomicilio.Row]:=trim(D8.Text);
  grilladomicilio.Cells[8,grilladomicilio.Row]:=trim(D9.Text);
  grilladomicilio.Cells[9,grilladomicilio.Row]:=trim(D10.Text);
  grilladomicilio.Cells[10,grilladomicilio.Row]:=trim(D11.Text);
  grilladomicilio.Cells[11,grilladomicilio.Row]:=trim(D12.Text);
  grilladomicilio.Cells[12,grilladomicilio.Row]:=trim(D13.Text);
  grilladomicilio.Cells[13,grilladomicilio.Row]:=trim(D14.Text);
  grilladomicilio.Cells[14,grilladomicilio.Row]:=trim(D15.Text);
  grilladomicilio.Cells[15,grilladomicilio.Row]:=TComboClase(D3.Items.Objects[D3.ItemIndex]).Param1;
  grilladomicilio.Cells[16,grilladomicilio.Row]:=TComboClase(D4.Items.Objects[D4.ItemIndex]).Param1;
  grilladomicilio.Cells[17,grilladomicilio.Row]:='X';
  Agrega3.Enabled:=false;
  Borra3.Enabled:=false;
  Modifica3.Enabled:=false;
  C5.Checked:=true;
  GUARDA.Enabled:=true;
end
else messagedlg('ERROR!! No Puede Existir dos N° DNI del mismo Tipo y Número!!',mterror,[mbok],0);
D3.SetFocus;
end;

procedure TAM_Contribuyentes.modifica4Click(Sender: TObject);
var t:integer;
  function busca(ht:integer):boolean;
  var p:integer;
  begin
  busca:=true;
  if grillactabancos.Cells[0,1] <> '' then begin
    for p:=1 to grillactabancos.RowCount-1 do
      if (grillactabancos.Cells[3,p]=trim(banco.Text))
      and (grillactabancos.Cells[4,p]=trim(tipoctabanco.Text))
      and (grillactabancos.Cells[5,p]=trim(nrosucursal.Text))
      and (grillactabancos.Cells[6,p]=trim(nrocuenta.Text))  and (ht<>p) then busca:=false;
  end;
  end;

begin
  if busca(grillactabancos.Row) then begin
    if grillactabancos.Cells[1,grillactabancos.Row] <> '?' then begin
      grillactabancos.Cells[14,grillactabancos.Row]:='X';
    end;
    grillactabancos.Cells[3,grillactabancos.Row]:=trim(banco.Text);
    grillactabancos.Cells[4,grillactabancos.Row]:=trim(tipoctabanco.Text);
    grillactabancos.Cells[5,grillactabancos.Row]:=trim(nrosucursal.Text);
    grillactabancos.Cells[6,grillactabancos.Row]:=trim(nrocuenta.Text);
    if escredito.Checked  then grillactabancos.Cells[7,grillactabancos.Row]:='True'
     else grillactabancos.Cells[7,grillactabancos.Row]:='False';
    grillactabancos.Cells[10,grillactabancos.Row]:=modulo2.Usuario.UsuNro;
    grillactabancos.Cells[12,grillactabancos.Row]:=TComboClase(banco.Items.Objects[banco.ItemIndex]).Param1;
    grillactabancos.Cells[13,grillactabancos.Row]:=TComboClase(tipoctabanco.Items.Objects[tipoctabanco.ItemIndex]).Param1;
    grillactabancos.Cells[14,grillactabancos.Row]:='X';
    grillactabancos.Cells[15,grillactabancos.Row]:=trim(cbu.Text);
    Agrega4.Enabled:=false;
    Modifica4.Enabled:=false;
    C7.Checked:=true;
    GUARDA.Enabled:=true;
  end
  else messagedlg('ERROR!! No Puede Existir dos N° DNI del mismo Tipo y Número!!',mterror,[mbok],0);
  banco.SetFocus;
end;

procedure TAM_Contribuyentes.modifica5Click(Sender: TObject);
var t:integer;
  function busca(ht:integer):boolean;
  var p:integer;
  begin
  busca:=true;
  if grilladebitos.Cells[0,1] <> '' then begin
    for p:=1 to grilladebitos.RowCount-1 do
      if (grilladebitos.Cells[4,p]=trim(elsubrubro.Text))
      and (grilladebitos.Cells[5,p]=trim(elcodobjeto.Text))     and (ht<>p) then busca:=false;
  end;
  end;

begin
  if busca(grilladebitos.Row) then begin
    if grilladebitos.Cells[1,grilladebitos.Row] <> '?' then begin
      grilladebitos.Cells[13,grilladebitos.Row]:='X';
    end;
    grilladebitos.Cells[3,grilladebitos.Row]:=trim(elsubrubro.Text);//trim(elrubro.Text);
    grilladebitos.Cells[4,grilladebitos.Row]:=trim(elsubrubro.Text);
    grilladebitos.Cells[5,grilladebitos.Row]:=trim(elcodobjeto.Text);
    grilladebitos.Cells[6,grilladebitos.Row]:=trim(label38.Caption);
    grilladebitos.Cells[9,grilladebitos.Row]:=modulo2.Usuario.UsuNro;
    grilladebitos.Cells[11,grilladebitos.Row]:=TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1;//TComboClase(elrubro.Items.Objects[elrubro.ItemIndex]).Param1;
    grilladebitos.Cells[12,grilladebitos.Row]:=TComboClase(elsubrubro.Items.Objects[elsubrubro.ItemIndex]).Param1;
    grilladebitos.Cells[13,grilladebitos.Row]:='X';
    Agrega5.Enabled:=false;
    Modifica5.Enabled:=false;
    C8.Checked:=true;
    GUARDA.Enabled:=true;
    elsubrubro.ItemIndex:=0;
    elcodobjeto.Clear;
    label38.Caption:='...';
  end
  else messagedlg('ERROR!! No Puede Existir dos Debitos del mismo Tipo y Número!!',mterror,[mbok],0);
  elsubrubro.SetFocus;
end;

procedure TAM_Contribuyentes.nombfantasiaKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  domcomercio.SetFocus;
end;
end;

procedure TAM_Contribuyentes.nrocuentaExit(Sender: TObject);
begin
if (trim(nrocuenta.Text) = '') or (trim(nrocuenta.Text) = '0') then begin
  messagedlg('Error!! Ingrese un N° de Cuanta Bancaria válida',mterror,[mbok],0);
  nrocuenta.SetFocus;
end;
end;

procedure TAM_Contribuyentes.nrocuentaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    cbu.SetFocus;
  end
  else begin
    if ( StrScan('0123456789',Key) <> nil ) or
        ( Key = Char(VK_BACK) ) then
     else Key := #0;
  end;
end;

procedure TAM_Contribuyentes.nrosucursalExit(Sender: TObject);
begin
if (trim(nrosucursal.Text) = '') or (trim(nrosucursal.Text) = '0') then begin
  messagedlg('Error!! Ingrese un N° de Sucursal Bancario válido',mterror,[mbok],0);
  nrosucursal.SetFocus;
end;
end;

procedure TAM_Contribuyentes.nrosucursalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then begin
    Key := #0;
    nrocuenta.SetFocus;
  end
  else begin
    if ( StrScan('0123456789',Key) <> nil ) or
        ( Key = Char(VK_BACK) ) then
     else Key := #0;
  end;
end;

procedure TAM_Contribuyentes.Borra3Click(Sender: TObject);
var t,p,gf:integer;
var bt:boolean;
    Cnx2:TADOConnection;
begin
  if trim(grilladomicilio.Cells[0,1])<>'' then begin
    gf:=1;t:=0;
    if grilladomicilio.Cells[1,grilladomicilio.Row] = '?' then t:=1
    else gf:=messagedlg('Desea Borrar el Domicilio Seleccionado??' + chr(13)
         + 'Se eliminará definitivamente de la Base de Datos',mtWarning,[mbok,mbCancel],0);
    if gf = 1 then begin
      if t=0 then begin
        try
          Cnx2:=TADOConnection.Create(nil);
          Cnx2.LoginPrompt:=false;
          Cnx2.ConnectionString:=modulo2.Sistema.ConectUsu;
          Com1:=TADOCommand.Create(nil);
          Com1.Connection:=Cnx2;
          Com1.CommandText:='SPB_ContribuyentesDomicilios ' + grilladomicilio.Cells[1,grilladomicilio.Row]
          + ' , ' + modulo2.Usuario.UsuNro
          + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
          + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
          Com1.Execute;
          messagedlg('Identidad Eliminada correctamente!',mtWarning,[mbok],0);
        except
          EnviarLog(self.Caption,Com1.CommandText);
          messagedlg('ERROR al Eliminar los datos',mterror,[mbok],0);
        end;
        FreeAndNil(Com1);
        Cnx2.Close;
        FreeAndNil(Cnx2);
      end;
      t:=grilladomicilio.Row;
      if (t >= 1) and (grilladomicilio.RowCount > 2) then begin
        THackStringGrid(grilladomicilio).DeleteRow(t);
      end
      else begin
        for t:=0 to grilladomicilio.ColCount-1 do grilladomicilio.Cells[t,1]:='';
      end;
    end;
    Borra3.Enabled:=false;
    Modifica3.Enabled:=false;
    D3.SetFocus;
  end;
end;

procedure TAM_Contribuyentes.borra5Click(Sender: TObject);
var t,p,gf:integer;
var bt:boolean;
begin
if trim(grilladebitos.Cells[0,1])<>'' then begin
  gf:=1;
  t:=0;
  if grilladebitos.Cells[1,grilladebitos.Row] = '?' then t:=1
  else gf:=messagedlg('Desea Borrar este Debito Seleccionado?',mtWarning,[mbok,mbCancel],0);
  if gf = 1 then begin
    if t=0 then begin
    //      try
    //        Com1:=TADOCommand.Create(nil);
    //        Com1.ConnectionString:=modulo2.Sistema.ConectSeg;
    //        Com1.CommandText:='SPB_ContribuyentesIdentidad ' + grillaidentidad.Cells[1,grillaidentidad.Row]
    //        + ' , ' + L1.Text
    //        + ' , ' + modulo2.Usuario.UsuNro
    //        + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    //        + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
    //        Com1.Execute;
    //        messagedlg('Identidad Eliminada correctamente!',mtWarning,[mbok],0);
    //      except
    //        EnviarLog(self.Caption,Com1.CommandText);
    //        messagedlg('ERROR al Eliminar los datos',mterror,[mbok],0);
    //      end;
    //      Com1.Free;
    end;
    t:=grilladebitos.Row;
    if (t >= 1) and (grilladebitos.RowCount > 2) then begin
      THackStringGrid(grilladebitos).DeleteRow(t);
    end
    else begin
      for t:=0 to grilladebitos.ColCount-1 do grilladebitos.Cells[t,1]:='';
    end;
  end;
  Borra5.Enabled:=false;
  Modifica5.Enabled:=false;
  elsubrubro.SetFocus;
end;
end;

procedure TAM_Contribuyentes.Quita3Click(Sender: TObject);
begin
D3.ItemIndex:=0;D4.ItemIndex:=0;
D5.Text:='3600';
D6.Clear;          D7.Clear;
D8.Clear;          D9.Clear;
D10.Clear;         D11.Clear;
D12.Clear;         D13.Clear;
D14.Clear;         D15.Clear;
Modifica3.Enabled:=false;
Borra3.Enabled:=false;
Agrega3.Enabled:=true;
Quita3.Enabled:=false;
end;

procedure TAM_Contribuyentes.QuitarDebitoSeleccionado1Click(Sender: TObject);
begin
elsubrubro.ItemIndex:=0;
CargaCombos(11);
elsubrubro.ItemIndex:=0;
elcodobjeto.Clear;
label38.Caption:='...';
Modifica5.Enabled:=false;
Agrega5.Enabled:=true;
end;

procedure TAM_Contribuyentes.QuitarSeleccion1Click(Sender: TObject);
var p,u:integer;
begin
banco.ItemIndex:=0;
tipoctabanco.ItemIndex:=0;
escredito.Checked:=false;
nrosucursal.Clear;
nrocuenta.Clear;
Modifica4.Enabled:=false;
Agrega4.Enabled:=true;

for p:=0 to grilladebitos.ColCount-1 do
  for u:=1 to grilladebitos.RowCount-1 do grilladebitos.Cells[p,u]:='';
grilladebitos.RowCount:=2;

dataset1.Close;
elsubrubro.ItemIndex:=0;
elcodobjeto.Clear;
label38.Caption:='...';
end;

procedure TAM_Contribuyentes.GUARDAClick(Sender: TObject);
var gf:integer;
    tx,valcuit:string;
    txt1:widestring;
    Cnx2:TADOConnection;
begin
gf:=messagedlg('Confirma GUARDAR los datos del Contribuyente Ingresados?',mtWarning,[mbok,mbCancel],0);
if gf = 1 then begin
  if Verificacion(1) then begin
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
    Com1:=TADOCommand.Create(nil);
    Com1.Connection:=Cnx2;
    if L10.Checked then tx:='1'
    else tx:='0';
    if trim(L5.text)='' then valcuit:= 'null'
    else valcuit:=trim(L5.text);
    if not C6.Checked then begin   ///alta
      try
        DataSet1.Connection:=Cnx2;
        DataSet1.CommandText:='SPA_Contribuyentes '  + chr(39) + trim(L2.Text) + chr(39)
        + ' , ' + chr(39) + trim(L3.Text) + chr(39)
        + ' , ' + TComboClase(L4.Items.Objects[L4.ItemIndex]).Param1
        + ' , ' + valcuit
        + ' , ' + TComboClase(L6.Items.Objects[L6.ItemIndex]).Param1
        + ' , ' + chr(39) + trim(L7.Text) + chr(39)
        + ' , ' + TComboClase(L8.Items.Objects[L8.ItemIndex]).Param1
        + ' , ' + chr(39) + trim(L9.Text) + chr(39)
        + ' , ' + tx + ' , 0 '
        + ' , ' + TComboClase(L11.Items.Objects[L11.ItemIndex]).Param1
        + ' , ' + modulo2.Usuario.UsuNro
        + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
        + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39)
        + ' , 20 ';
        //+ ' , ' + modulo2.Usuario.UsuPuntoV;
        DataSet1.Open;
        tx:=DataSet1.Fields[0].AsString;
        DataSet1.Close;
        DataSet1.Connection:=nil;

        for gf:=1 to grillaidentidad.RowCount-1 do begin
          Com1.CommandText:='SPA_ContribuyentesIdentidad '  + tx
          + ' , ' + chr(39) + trim(grillaidentidad.Cells[4,gf]) + chr(39)
          + ' , ' + trim(grillaidentidad.Cells[5,gf])
          + ' , ' + trim(grillaidentidad.Cells[3,gf])
          + ' , ' + modulo2.Usuario.UsuNro
          + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
          + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
          Com1.Execute;
        end;
        for gf:=1 to grilladomicilio.RowCount-1 do begin
          Com1.CommandText:='SPA_ContribuyentesDomicilios ' + tx
          + ' , 38 , ' + trim(grilladomicilio.Cells[15,gf])
          + ' , ' + trim(grilladomicilio.Cells[16,gf])
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[14,gf]) + chr(39)
          + ' , ' + trim(grilladomicilio.Cells[5,gf])
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[6,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[7,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[8,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[12,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[10,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[11,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[9,gf]) + chr(39)
          + ' , ' + chr(39) + trim(grilladomicilio.Cells[13,gf]) + chr(39)
          + ' , ' + trim(grilladomicilio.Cells[4,gf])
          + ' , ' + modulo2.Usuario.UsuNro
          + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
          + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
          Com1.Execute;
        end;
        if c7.Checked  then begin
           if VerificaGrilla('grilladebitos') = false then begin
              messagedlg('ERROR Imposible Guardar esta Cuenta Bancaria, por NO poseer Debitos Asigandos...'+ chr(13)
              +'Por favor realice las Carga de al menos un Débito, para luego poder Guardar esta Cuenta Bancaria',mterror,[mbok],0);
              elsubrubro.SetFocus;
           end
           else begin
              for gf:=1 to grillactabancos.RowCount-1 do begin
                if (grillactabancos.Cells[14,gf] = 'X') then begin
                 if grillactabancos.Cells[1,gf] = '?' then begin
                  DataSet2.Connection:=Cnx2;
                  txt1:= tx   //IdContribuyente
                  + ' , ' + trim(grillactabancos.Cells[12,gf])  //idbanco
                  + ' , ' + trim(grillactabancos.Cells[13,gf])  //idtipocuenta
                  + ' , ' + trim(grillactabancos.Cells[5,gf])   //sucursal
                  + ' , ' + trim(grillactabancos.Cells[6,gf]);   // cuenta
                  if trim(grillactabancos.Cells[15,gf])= '' then txt1:= txt1  + ' , null '
                    else txt1:= txt1  + ' , ' + chr(39) + grillactabancos.Cells[15,gf] + chr(39);
                  if trim(grillactabancos.Cells[7,gf])= 'True' then txt1:=txt1 + ', 1 '  // credito
                      else txt1:= txt1 + ', 0 ';
                  txt1:= txt1  + ' , ' + modulo2.Usuario.UsuNro
                  + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                  + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                  DataSet2.CommandText:='SPA_ContribuyentesCuentasBanco '  + txt1;
                  DataSet2.Open;
                  ctadeb:=DataSet2.Fields[0].AsString;
                  DataSet2.Close;
                  DataSet2.Connection:=nil;
                 end;
                end;
              end;
           end;
        end;
        if c8.Checked  then begin
         for gf:=1 to grilladebitos.RowCount-1 do begin
           if (grilladebitos.Cells[13,gf] = 'X') then begin
             if grilladebitos.Cells[1,gf] = '?' then begin
               if grilladebitos.Cells[2,gf] = '?' then begin
                txt1:= ctadeb
                + ' , ' + trim(grilladebitos.Cells[12,gf])  //idCUENTA Y subrubro
                + ' , ' + chr(39) + trim(grilladebitos.Cells[5,gf]) + chr(39)       //codjeto
                + ' , ' + trim(grilladebitos.Cells[6,gf])   // identidificador
                + ' , ' + trim(grilladebitos.Cells[14,gf])   // planpago
                + ' , ' + modulo2.Usuario.UsuNro   // nrousario
                + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                Com1.CommandText:='SPA_ContribuyentesDebitos ' + txt1;
                Com1.Execute;
               end;
             end;
           end;
         end;
        end;
        messagedlg('Datos Guardados Correctamente!!',mtConfirmation,[mbok],0);
        if c8.Checked  then imprimeconfirmacion();
        BORRA.Click;
      except
       messagedlg('ERROR al Guardar los Datos',mtError,[mbok],0);
      end;
    end
    else begin  /// modificacion
      try
        if C3.Checked then begin
          Com1.CommandText:='SPM_Contribuyentes ' + L1.Text                   //id
          + ' , ' + chr(39) + trim(L2.Text) + chr(39)                         //contrib
          + ' , ' + chr(39) + trim(L3.Text) + chr(39)                         //fecnacim
          + ' , ' + TComboClase(L4.Items.Objects[L4.ItemIndex]).Param1  //tiposex
          + ' , ' + valcuit                                                   //cuit
          + ' , ' + TComboClase(L6.Items.Objects[L6.ItemIndex]).Param1 //tipopersona
          + ' , ' + chr(39) + trim(L7.Text) + chr(39)                         //mail
          + ' , ' + TComboClase(L8.Items.Objects[L8.ItemIndex]).Param1 //estcivil
          + ' , ' + chr(39) + trim(L9.Text) + chr(39)                         //telf
          + ' , ' + tx                                                        //fallecido
          + ' , ' + TComboClase(L11.Items.Objects[L11.ItemIndex]).param1//nacio
          + ' , ' + modulo2.Usuario.UsuNro
          + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
          + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
          Com1.Execute;
        end; //3
        if C4.Checked then begin
          for gf:=1 to grillaidentidad.RowCount-1 do begin
            if (grillaidentidad.Cells[7,gf] = 'X') then begin
              if grillaidentidad.Cells[1,gf] = '?' then begin
                Com1.CommandText:='SPA_ContribuyentesIdentidad ' + L1.Text
                + ' , ' + chr(39) + trim(grillaidentidad.Cells[4,gf]) + chr(39)
                + ' , ' + trim(grillaidentidad.Cells[5,gf])
                + ' , ' + trim(grillaidentidad.Cells[3,gf])
                + ' , ' + modulo2.Usuario.UsuNro
                + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                Com1.Execute;
              end
              else begin
                Com1.CommandText:='SPM_ContribuyentesIdentidad '  + grillaidentidad.Cells[1,gf]
                + ' , ' + L1.Text
                + ' , ' + chr(39) + trim(grillaidentidad.Cells[4,gf]) + chr(39)
                + ' , ' + trim(grillaidentidad.Cells[5,gf])
                + ' , ' + trim(grillaidentidad.Cells[3,gf])
                + ' , ' + modulo2.Usuario.UsuNro
                + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                Com1.Execute;
              end;
            end;
          end;
        end;
        if C5.Checked then begin
          for gf:=1 to grilladomicilio.RowCount-1 do begin
            if (grilladomicilio.Cells[17,gf] = 'X') then begin
              if grilladomicilio.Cells[1,gf] = '?' then begin
                Com1.CommandText:='SPA_ContribuyentesDomicilios ' + L1.Text
                + ' , 38 , '  + trim(grilladomicilio.Cells[15,gf])
                + ' , ' + trim(grilladomicilio.Cells[16,gf])
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[14,gf]) + chr(39)
                + ' , ' + trim(grilladomicilio.Cells[5,gf])
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[6,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[7,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[8,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[12,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[10,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[11,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[9,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[13,gf]) + chr(39)
                + ' , ' + trim(grilladomicilio.Cells[4,gf])
                + ' , ' + modulo2.Usuario.UsuNro
                + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                Com1.Execute;
              end
              else begin
                Com1.CommandText:='SPM_ContribuyentesDomicilios '  + trim(grilladomicilio.Cells[1,gf])
                + ' , ' + L1.Text + ' , 38 , '  + trim(grilladomicilio.Cells[15,gf])
                + ' , ' + trim(grilladomicilio.Cells[16,gf])
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[14,gf]) + chr(39)
                + ' , ' + trim(grilladomicilio.Cells[5,gf])
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[6,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[7,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[8,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[12,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[10,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[11,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[9,gf]) + chr(39)
                + ' , ' + chr(39) + trim(grilladomicilio.Cells[13,gf]) + chr(39)
                + ' , ' + trim(grilladomicilio.Cells[4,gf])
                + ' , ' + modulo2.Usuario.UsuNro
                + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                Com1.Execute;
              end;
            end;
          end;
        end;
        if c7.Checked  then begin
           if VerificaGrilla('grilladebitos') = false then begin
              messagedlg('ERROR Imposible Guardar esta Cuenta Bancaria, por NO poseer Debitos Asigandos...'+ chr(13)
              +'Por favor realice las Carga de al menos un Débito, para luego poder Guardar esta Cuenta Bancaria',mterror,[mbok],0);
              elsubrubro.SetFocus;
           end
           else begin
              for gf:=1 to grillactabancos.RowCount-1 do begin
                if (grillactabancos.Cells[14,gf] = 'X') then begin
                  if grillactabancos.Cells[1,gf] = '?' then begin
                    DataSet2.Connection:=Cnx2;
                    txt1:= L1.Text   //IdContribuyente
                    + ' , ' + trim(grillactabancos.Cells[12,gf])  //idbanco
                    + ' , ' + trim(grillactabancos.Cells[13,gf])  //idtipocuenta
                    + ' , ' + trim(grillactabancos.Cells[5,gf])   //sucursal
                    + ' , ' + trim(grillactabancos.Cells[6,gf]);   // cuenta
                    if trim(grillactabancos.Cells[15,gf])= '' then txt1:= txt1  + ' , null '
                      else txt1:= txt1  + ' , ' + chr(39) + grillactabancos.Cells[15,gf] + chr(39);
                    if trim(grillactabancos.Cells[7,gf])= 'True' then txt1:=txt1 + ', 1 '  // credito
                        else txt1:= txt1 + ', 0 ';
                    txt1:= txt1  + ' , ' + modulo2.Usuario.UsuNro
                    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                    DataSet2.CommandText:='SPA_ContribuyentesCuentasBanco '  + txt1;
                    DataSet2.Open;
                    ctadeb:=DataSet2.Fields[0].AsString;
                    DataSet2.Close;
                    DataSet2.Connection:=nil;
                  end
                  else begin
                    txt1:= trim(grillactabancos.Cells[1,gf])  //idcuenta
                    + ' , ' + L1.Text //IdContribuyente
                    + ' , ' + trim(grillactabancos.Cells[12,gf])  //idbanco
                    + ' , ' + trim(grillactabancos.Cells[13,gf])  //idtipocuenta
                    + ' , ' + trim(grillactabancos.Cells[5,gf])   //sucursal
                    + ' , ' + trim(grillactabancos.Cells[6,gf]);   // cuenta
                    if trim(grillactabancos.Cells[15,gf])= '' then txt1:= txt1  + ' , null '
                      else txt1:= txt1  + ' , ' + chr(39) + grillactabancos.Cells[15,gf] + chr(39);
                    if trim(grillactabancos.Cells[7,gf])= 'True' then txt1:=txt1 + ' , 1 '  // credito
                        else txt1:= txt1 + ' , 0 ';
                    txt1:= txt1  + ' , ' + modulo2.Usuario.UsuNro
                    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                    Com1.CommandText:='SPM_ContribuyentesCuentasBanco ' + txt1;
                    Com1.Execute;
                  end;
                end;
              end;
              messagedlg('Datos Guardados Correctamente!!',mtConfirmation,[mbok],0);
           end;
        end;
        if c8.Checked  then begin
           for gf:=1 to grilladebitos.RowCount-1 do begin
             if (grilladebitos.Cells[13,gf] = 'X') then begin
               if grilladebitos.Cells[1,gf] = '?' then begin
                 if grilladebitos.Cells[2,gf] = '?' then begin
                    txt1:= ctadeb
                    + ' , ' + trim(grilladebitos.Cells[12,gf])  //idCUENTA Y subrubro
                    + ' , ' + chr(39) + trim(grilladebitos.Cells[5,gf]) + chr(39)       //codjeto
                    + ' , ' + trim(grilladebitos.Cells[6,gf])   // identidificador
                    + ' , ' + trim(grilladebitos.Cells[14,gf])   // planpago
                    + ' , ' + modulo2.Usuario.UsuNro   // nrousario
                    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                    Com1.CommandText:='SPA_ContribuyentesDebitos ' + txt1;
                    Com1.Execute;
                 end
                 ELSE BEGIN
                    txt1:= trim(grilladebitos.Cells[2,gf])
                    + ' , ' + trim(grilladebitos.Cells[12,gf])  //idbanco  y subrubro
                    + ' , ' + chr(39) + trim(grilladebitos.Cells[5,gf]) + chr(39)        //codjeto
                    + ' , ' + trim(grilladebitos.Cells[6,gf])   // identif
                    + ' , ' + trim(grilladebitos.Cells[14,gf])  // planpago
                    + ' , ' + modulo2.Usuario.UsuNro   //nrousario
                    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                    Com1.CommandText:='SPA_ContribuyentesDebitos ' + txt1;
                    Com1.Execute;
                 END;
               end
               else begin
                  txt1:= trim(grilladebitos.Cells[1,gf])  //idc
                  + ' , ' + trim(grilladebitos.Cells[2,gf])  //idcuenta
                  + ' , ' + trim(grilladebitos.Cells[11,gf])  //subrubro
                  + ' , ' + trim(grilladebitos.Cells[5,gf])  //codobjeto
                  + ' , ' + trim(grilladebitos.Cells[6,gf])   // identificador
                  + ' , ' + trim(grilladebitos.Cells[14,gf])   //planpago
                  + ' , ' + modulo2.Usuario.UsuNro
                  + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
                  + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
                  Com1.CommandText:='SPM_ContribuyentesDebitos ' + txt1;
                  Com1.Execute;
               end;
             end;
           end;
        end;
        messagedlg('Datos Guardados Correctamente!!',mtConfirmation,[mbok],0);
        if c8.Checked  then imprimeconfirmacion();
        BORRA.Click;
      except
       messagedlg('ERROR al Guardar los Datos',mtError,[mbok],0);
      end;
    end;
    FreeAndNil(Com1);
    Cnx2.Close;
    FreeAndNil(Cnx2);
  end;
end;
end;

Function TAM_Contribuyentes.Verificacion(nro:byte):boolean;
var Form1:TForm;
    E1:TRichEdit;
    txt:string;
    r:extended;
    d:integer;
begin
try
  Form1:=TForm.Create(nil);
  Form1.Width:=530;
  Form1.Height:=320;
  Form1.Caption:='Errores Encontrados en Ventana';
  Form1.Position:=poDeskTopCenter;
  E1:=TRichEdit.Create(nil);
  E1.Parent:=Form1;
  E1.Color:=clNavy;
  E1.Font.Color:=clYellow;
  E1.Font.Size:=10;
  E1.Top:=0;
  E1.Left:=0;
  E1.Align:=alClient;
  E1.Clear;
  E1.ReadOnly:=true;
  if nro = 1 then begin //Altas
    if trim(L2.Text) = '' then E1.Lines.Add('Campo Nombre del Contribuyente vacío');
    if trim(L3.Text) = '/  /' then E1.Lines.Add('Campor Fecha nacimiento vacío');
    if L4.ItemIndex < 0 then E1.Lines.Add('Sexo inválido');
    if L6.ItemIndex < 0 then E1.Lines.Add('Tipo Persona inválido');
    if L8.ItemIndex < 0 then E1.Lines.Add('Estado Civil inválido o vacío');
    if L11.ItemIndex < 0 then E1.Lines.Add('Tipo Nacionalidad inválido o vacío');
    if trim(grillaidentidad.Cells[0,1]) = '' then E1.Lines.Add('Contribuyente sin Identidad declarada');
    if trim(grilladomicilio.Cells[0,1]) = '' then E1.Lines.Add('Contribuyente sin Domicilio declarado');
  end;
  if nro = 2  then  begin
    if trim(nrosucursal.Text) = '' then E1.Lines.Add('Debe completar el Campo Nro. de Sucursal Bancario. No puede contener CERO ni estar vacio');
    if trim(nrocuenta.Text) = '' then E1.Lines.Add('Debe completar el Campo Tipo de Cuenta. No puede contener CERO ni estar vacio');
  end;
  if nro = 3  then  begin
    if elsubrubro.ItemIndex < 0 then E1.Lines.Add('Tipo Subrubro inválido o vacío');
    if trim(elcodobjeto.Text) = '' then E1.Lines.Add('Debe completar un Campo Codigo de Objeto. No puede contener CERO ni estar vacio');
    if label38.Caption='...' then E1.Lines.Add('Identificador del Objeto inválido');
  end;
  if nro = 4  then  begin
    if trim(L9.Text) = '' then E1.Lines.Add('Debe completar Campo Nº Telefono. No puede estar vacio');
    if trim(L5.Text) = '' then E1.Lines.Add('Debe completar Campo Cuit/Cuil. No puede estar vacio');
    if trim(L7.Text) <> '' then
      if not UNidad.emailValido(trim(L7.Text))then E1.Lines.Add('El Mail del contribuyente es erroneo');
  end;
  if E1.Lines.Count > 0 then begin
    Form1.ShowModal;
    Result:=false;
  end
  else Result:=true;
finally
  Form1.Free;
  Form1:=nil;
end;
end;

procedure TAM_Contribuyentes.SALEClick(Sender: TObject);
begin
Close;
end;

procedure TAM_Contribuyentes.tipoctabancoExit(Sender: TObject);
begin
if tipoctabanco.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Tipo de Cuenta bancaria válido',mterror,[mbok],0);
  nrosucursal.SetFocus;
  //tipoctabanco.SetFocus;
end;
end;

procedure TAM_Contribuyentes.tipoctabancoKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  //nrocuenta.SetFocus;
  nrosucursal.SetFocus;
end;
end;

procedure TAM_Contribuyentes.agregacarnetClick(Sender: TObject);
var gf:integer;
var ReporteRst: TfrxReport;
    Cnx2:TADOConnection;
begin
gf:=messagedlg('Confirma GUARDAR los datos del CARNET del Contribuyente seleccionado?',mtWarning,[mbok,mbCancel],0);
if gf = 1 then begin
  if (trim(nombfantasia.Text) = '') or (trim(domcomercio.Text) = '') then begin
    messagedlg('Error!! No Existe nombre de fantasia ingresado.',mterror,[mbok],0);
    nombfantasia.SetFocus;
  end
  else begin
      Consultas(0);  //:=nrocarnet
      dmi:='3310000000006005900';
      dmi:=dmi + StrFill(AnsiReplaceStr(nrocarnet,'.',''),5,'0','');
      Recibo.NroCodBar:=dmi;
      LC.caption:=TRIM(L2.text);
      LCUIT.caption:=TRIM(L5.text);
      Consultas(10);
      LL.caption:='FORMOSA';
      lcodbar.Caption:=dmi;
      LNRO.caption:=nrocarnet;
      try
        Cnx2:=TADOConnection.Create(nil);
        Cnx2.LoginPrompt:=false;
        Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
        Com1:=TADOCommand.Create(nil);
        Com1.Connection:=Cnx2;
        Com1.CommandText:='SPA_ContribuyentesCarnetEstMedido ' + trim(l1.Text)
        + ' , ' + chr(39) + elcuim + chr(39)
        + ' , ' + nrocarnet
        + ' , ' + chr(39) + dmi + chr(39)
        + ' , ' + chr(39) + trim(nombfantasia.Text) + chr(39)
        + ' , ' + chr(39) + trim(domcomercio.Text) + chr(39)
        + ' , ' + modulo2.Usuario.UsuNro
        + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
        + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
        Com1.Execute;
        FreeAndNil(Com1);
        Cnx2.Close;
        FreeAndNil(Cnx2);
        messagedlg('Carnet Guardado Correctamente!!',mtConfirmation,[mbok],0);
        agregacarnet.Enabled:=false;
        try
          ReporteRst:=TfrxReport.Create(nil);
          ReporteRst.LoadFromFile(Modulo2.Sistema.Path_todo + '\Reportes\Imp_CarnetEstMedido.fr3');
          TfrxPictureView(ReporteRst.FindObject('Logo1')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
          TfrxPictureView(ReporteRst.FindObject('Logo2')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\logoformosa.jpg');
          ReporteRst.Variables['NombreTit'] := QuotedStr(lc.Caption);
          ReporteRst.Variables['CuimTit'] := QuotedStr(lcuit.Caption);
          ReporteRst.Variables['DomicilioTit'] := QuotedStr(ld.Caption);
          ReporteRst.Variables['NombreRespo'] := QuotedStr(ll.Caption);
          ReporteRst.Variables['DetalleRecibo'] := QuotedStr(trim(nombfantasia.Text));
          ReporteRst.Variables['DomicilioRespo'] := QuotedStr(trim(domcomercio.Text));
          ReporteRst.Variables['NroPlan'] := QuotedStr(nrocarnet);
          Recibo.CodBar:=codigo_bar(Recibo.NroCodBar);
          ReporteRst.Variables['CodBar'] := QuotedStr(Recibo.CodBar);
          ReporteRst.Variables['NroCodBar'] := QuotedStr(Recibo.NroCodBar);
          ReporteRst.ShowReport;
        except
          on E: Exception do EnviarLog(Self.Caption, 'Error en Reporte ' + E.Message);
        end;
        FreeAndNil(ReporteRst);
      except
        messagedlg('ERROR al Guardar el carnet...',mtError,[mbok],0);
      end;
  end;
end;
end;

procedure TAM_Contribuyentes.JvDBGrid2CellClick(Column: TColumn);
begin
if not CONTRIBUYENTE.IsEmpty then begin
 pc1.Pages[1].TabVisible:=true;
 pc1.Pages[2].TabVisible:=true;
 pc1.Pages[3].TabVisible:=true;
 {if NroOpcion=1 then begin
   pc1.Pages[4].TabVisible:=true;
   pc1.Pages[5].TabVisible:=false;
 end
 else begin
   pc1.Pages[4].TabVisible:=false;
   pc1.Pages[5].TabVisible:=true;
 end; }
 if CONTRIBUYENTE.Fields[0].AsInteger <> 0 then begin
  PC1.Pages[1].Enabled:=true;
  PC1.Pages[2].Enabled:=true;
  PC1.Pages[3].Enabled:=true;
  L1.Clear;                   L2.Clear;
  L3.Clear;                   L4.ItemIndex:=0;
  L5.Clear;                   L6.ItemIndex:=0;
  L7.Clear;                   L8.ItemIndex:=0;
  L9.Clear;                   L10.Checked:=false;
  L1.Text:=CONTRIBUYENTE.Fields[0].AsString;
  L2.Text:=CONTRIBUYENTE.Fields[1].AsString;
  L3.Text:=CONTRIBUYENTE.Fields[2].AsString;
  L4.ItemIndex:=L4.Items.IndexOf(trim(CONTRIBUYENTE.Fields[3].AsString));
  L5.Text:=CONTRIBUYENTE.Fields[4].AsString;
  eldni:=strcut_midle(trim(CONTRIBUYENTE.Fields[4].AsString),3,10);
  L6.ItemIndex:=L6.Items.IndexOf(trim(CONTRIBUYENTE.Fields[5].AsString));
  L7.Text:=CONTRIBUYENTE.Fields[6].AsString;
  L8.ItemIndex:=L8.Items.IndexOf(trim(CONTRIBUYENTE.Fields[7].AsString));
  L9.Text:=CONTRIBUYENTE.Fields[8].AsString;
  L10.Checked:=CONTRIBUYENTE.Fields[9].AsBoolean;
  L11.ItemIndex:=L11.Items.IndexOf(trim(CONTRIBUYENTE.Fields[12].AsString));
  H1.Clear;H2.Text:='00-00000000-0';
  H3.ItemIndex:=0;
  H4.Clear;
  Consultas(8);
  Consultas(5);
  D3.ItemIndex:=0;            D4.ItemIndex:=0;
  D5.Text:='3600';            D6.Clear;                   D7.Clear;
  D8.Clear;                   D9.Clear;
  D10.Clear;                  D11.Clear;
  D12.Clear;                  D13.Clear;
  D14.Clear;                  D15.Clear;
  Consultas(9);
  C3.Checked:=false;          C4.Checked:=false;
  C5.Checked:=false;          C6.Checked:=true;
  GUARDA.Enabled:=false;
  PC1.ActivePage:=PC1.Pages[1];
  Consultas(1);
  L2.SetFocus;
 end
 else begin
   messagedlg('No se puede editar el contribuyente DESCONOCIDO',mterror,[mbok],0);
   PC1.ActivePage:=PC1.Pages[0];
   PC1.Pages[1].Enabled:=false;
   PC1.Pages[2].Enabled:=false;
   PC1.Pages[3].Enabled:=false;
 end;
end;
end;

procedure TAM_Contribuyentes.grillactabancosClick(Sender: TObject);
begin
if grillactabancos.Cells[0,grillactabancos.Row] <> '' then begin
  if grillactabancos.Cells[9,grillactabancos.Row] = 'True' then begin
    laidcuenta:=trim(grillactabancos.Cells[1,grillactabancos.Row]) ;
    consultas(2);
    C7.Checked:=false;
    Agrega4.Enabled:=false;
    Modifica4.Enabled:=false;
    PopupMenu2.Items[0].Enabled:=false;
    JvGroupBox2.Enabled:=false;
    messagedlg('ATENCION esta Cuenta Bancaria se encuentra de BAJA, imposible su modificación..',mtInformation,[mbok],0);
  end
  else begin
    JvGroupBox2.Enabled:=true;
    C7.Checked:=true;
    Agrega4.Enabled:=false;
    Modifica4.Enabled:=true;
    grillactabancos.Cells[14,grillactabancos.Row]:='X';
    laidcuenta:=trim(grillactabancos.Cells[1,grillactabancos.Row]) ;
    banco.ItemIndex:=banco.Items.IndexOf(trim(grillactabancos.Cells[3,grillactabancos.Row]));
    tipoctabanco.ItemIndex:=tipoctabanco.Items.IndexOf(trim(grillactabancos.Cells[4,grillactabancos.Row]));
    nrosucursal.Text:=trim(grillactabancos.Cells[5,grillactabancos.Row]);
    nrocuenta.Text:=trim(grillactabancos.Cells[6,grillactabancos.Row]);
    if trim(grillactabancos.Cells[7,grillactabancos.Row])='True' then escredito.Checked:=true
     else  escredito.Checked:=false;
    cbu.Text:=trim(grillactabancos.Cells[15,grillactabancos.Row]);
    banco.SetFocus;
    consultas(4);
    PopupMenu2.Items[0].Enabled:=true;
  end;
end;
end;

procedure TAM_Contribuyentes.grillactabancosDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
   With Sender as TStringGrid do begin
     if trim(grillactabancos.Cells[9,Arow]) = 'True' then begin
       Canvas.Brush.Color := clYellow;
       Canvas.TextRect( Rect, Rect.Left+2, Rect.Top+2, Cells[acol, arow]);
     end;
   end;
end;

procedure TAM_Contribuyentes.PC1Change(Sender: TObject);
begin
if PC1.ActivePageIndex = 5  then begin
      nombfantasia.SetFocus;
end;
end;

procedure TAM_Contribuyentes.PopupMenu1Popup(Sender: TObject);
begin
Activar1.Enabled:=false;
if grillaidentidad.Row > 0 then begin
  if grillaidentidad.Cells[6,grillaidentidad.Row] = '0' then Activar1.Enabled:=true;
end;
end;

procedure TAM_Contribuyentes.printClick(Sender: TObject);
var ReporteRst: TfrxReport;
begin
    try
      ReporteRst:=TfrxReport.Create(nil);
      ReporteRst.LoadFromFile(Modulo2.Sistema.Path_todo + '\Reportes\Imp_CarnetEstMedido.fr3');
      TfrxPictureView(ReporteRst.FindObject('Logo1')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
      TfrxPictureView(ReporteRst.FindObject('Logo2')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\logoformosa.jpg');
      ReporteRst.Variables['NombreTit'] := QuotedStr(lc.Caption);
      ReporteRst.Variables['CuimTit'] := QuotedStr(lcuit.Caption);
      ReporteRst.Variables['DomicilioTit'] := QuotedStr(ld.Caption);
      ReporteRst.Variables['NombreRespo'] := QuotedStr(ll.Caption);
      ReporteRst.Variables['DetalleRecibo'] := QuotedStr(trim(nombfantasia.Text));
      ReporteRst.Variables['DomicilioRespo'] := QuotedStr(trim(domcomercio.Text));
      ReporteRst.Variables['NroPlan'] := QuotedStr(lnro.Caption);
      Recibo.CodBar:=codigo_bar(lcodbar.Caption);
      ReporteRst.Variables['CodBar'] := QuotedStr(Recibo.CodBar);
      ReporteRst.Variables['NroCodBar'] := QuotedStr(lcodbar.Caption);
      ReporteRst.ShowReport;
    except
      on E: Exception do EnviarLog(Self.Caption, 'Error en Reporte ' + E.Message);
    end;
    FreeAndNil(ReporteRst);
end;

procedure TAM_Contribuyentes.Activar1Click(Sender: TObject);
var t,p,gf:integer;
var bt:boolean;
    Cnx2:TADOConnection;
begin
gf:=messagedlg('Desea ACTIVAR la Identidad Seleccionado??',mtWarning,[mbok,mbCancel],0);
if gf = 1 then begin
  try
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
    Com1:=TADOCommand.Create(nil);
    Com1.Connection:=Cnx2;
    Com1.CommandText:='SPM_ContribuyentesIdentidadActivar ' + grillaidentidad.Cells[1,grillaidentidad.Row]
    + ' , ' + L1.Text
    + ' , ' + modulo2.Usuario.UsuNro
    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
    Com1.Execute;
    messagedlg('Identidad Activada correctamente!',mtWarning,[mbok],0);
  except
    EnviarLog(self.Caption,Com1.CommandText);
    messagedlg('ERROR al Guardar los datos',mterror,[mbok],0);
  end;
  FreeAndNil(Com1);
  Cnx2.Close;
  FreeAndNil(Cnx2);
  Consultas(8);
end;
end;

procedure TAM_Contribuyentes.grilladebitosClick(Sender: TObject);
begin
if grilladebitos.Cells[0,grilladebitos.Row] <> '' then begin
  if grilladebitos.Cells[8,grilladebitos.Row] = 'True' then begin
    messagedlg('ATENCION este Debito Bancario se encuentra de BAJA, imposible su modificación..',mtInformation,[mbok],0);
    C8.Checked:=false;
    Agrega5.Enabled:=false;
    PopupMenu3.Items[0].Enabled:=false;
    Modifica5.Enabled:=false;
  end
  else begin
    PopupMenu3.Items[0].Enabled:=true;
    C8.Checked:=true;
    Agrega5.Enabled:=false;
    Modifica5.Enabled:=true;
    borra5.Enabled:=true;
    grilladebitos.Cells[13,grilladebitos.Row]:='X';
    //elrubro.ItemIndex:=elrubro.Items.IndexOf(trim(grilladebitos.Cells[3,grilladebitos.Row]));
    elsubrubro.ItemIndex:=elsubrubro.Items.IndexOf(trim(grilladebitos.Cells[4,grilladebitos.Row]));
    elcodobjeto.Text:=trim(grilladebitos.Cells[5,grilladebitos.Row]);
    label38.Caption:=trim(grilladebitos.Cells[6,grilladebitos.Row]);
    elsubrubro.SetFocus;
  end;
end;
end;

procedure TAM_Contribuyentes.grilladomicilioClick(Sender: TObject);
begin
if grilladomicilio.Cells[0,grilladomicilio.Row] <> '' then begin
  D3.ItemIndex:=D3.Items.IndexOf(trim(grilladomicilio.Cells[2,grilladomicilio.Row]));
  D4.ItemIndex:=D4.Items.IndexOf(trim(grilladomicilio.Cells[3,grilladomicilio.Row]));
  D5.Text:=trim(grilladomicilio.Cells[4,grilladomicilio.Row]);
  D6.Text:=trim(grilladomicilio.Cells[5,grilladomicilio.Row]);
  D7.Text:=trim(grilladomicilio.Cells[6,grilladomicilio.Row]);
  D8.Text:=trim(grilladomicilio.Cells[7,grilladomicilio.Row]);
  D9.Text:=trim(grilladomicilio.Cells[8,grilladomicilio.Row]);
  D10.Text:=trim(grilladomicilio.Cells[9,grilladomicilio.Row]);
  D11.Text:=trim(grilladomicilio.Cells[10,grilladomicilio.Row]);
  D12.Text:=trim(grilladomicilio.Cells[11,grilladomicilio.Row]);
  D13.Text:=trim(grilladomicilio.Cells[12,grilladomicilio.Row]);
  D14.Text:=trim(grilladomicilio.Cells[13,grilladomicilio.Row]);
  D15.Text:=trim(grilladomicilio.Cells[14,grilladomicilio.Row]);
  Agrega3.Enabled:=false;
  Borra3.Enabled:=true;
  Quita3.Enabled:=true;
  Modifica3.Enabled:=true;
  D3.SetFocus;
end;
end;

procedure TAM_Contribuyentes.grillaidentidadClick(Sender: TObject);
begin
if grillaidentidad.Cells[0,grillaidentidad.Row] <> '' then begin
  C4.Checked:=true;
  Agrega2.Enabled:=false;
  Borra2.Enabled:=true;
  Quita2.Enabled:=true;
  Modifica2.Enabled:=true;
  H1.Text:=trim(grillaidentidad.Cells[1,grillaidentidad.Row]);
  H3.ItemIndex:=H3.Items.IndexOf(trim(grillaidentidad.Cells[2,grillaidentidad.Row]));
  H4.Text:=trim(grillaidentidad.Cells[3,grillaidentidad.Row]);
  H2.Text:=trim(grillaidentidad.Cells[4,grillaidentidad.Row]);
  H3.SetFocus;
end;
end;

procedure TAM_Contribuyentes.grillaidentidadDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
   With Sender as TStringGrid do begin
     if trim(grillaidentidad.Cells[6,Arow]) = '1' then begin
       Canvas.Brush.Color := clyellow;
       Canvas.TextRect( Rect, Rect.Left+2, Rect.Top+2, Cells[acol, arow]);
     end;
   end;
end;

procedure TAM_Contribuyentes.L2Change(Sender: TObject);
begin
C3.Checked:=true;
GUARDA.Enabled:=true;
end;

procedure TAM_Contribuyentes.AsociarCUIM1Click(Sender: TObject);
var Selec_Cuim:TSelec_Cuim;
    gf,j:integer;
    Cnx2:TADOConnection;
begin
Selec_Cuim:= TSelec_Cuim.create(nil);
try
  Selec_Cuim.ShowModal;
  if not Selec_Cuim.Cancela then begin
    gf:=messagedlg('CONFIRMA cambiar el CUIM seleccionado a este contribuyente??',mtWarning,[mbok,mbCancel],0);
    if gf = 1 then begin
      try
        Cnx2:=TADOConnection.Create(nil);
        Cnx2.LoginPrompt:=false;
        Cnx2.ConnectionString:=modulo2.Sistema.ConectSeg;
        Com1:=TADOCommand.Create(nil);
        Com1.Connection:=Cnx2;
        Com1.CommandText:='SPM_ContribuyentesIdentidadCUIM ' + Selec_Cuim.NroCodVal
        + ' , ' + L1.Text
        + ' , ' + modulo2.Usuario.UsuNro
        + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
        + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
        Com1.Execute;
        messagedlg('CUIM Cambiado correctamente!',mtWarning,[mbok],0);
      except
        EnviarLog(self.Caption,Com1.CommandText);
        messagedlg('ERROR al Guardar los datos',mterror,[mbok],0);
      end;
      FreeAndNil(Com1);
      Cnx2.Close;
      FreeAndNil(Cnx2);
      for gf:=0 to grillaidentidad.ColCount-1 do
        for j:=1 to grillaidentidad.RowCount-1 do grillaidentidad.Cells[gf,j]:='';
      grillaidentidad.RowCount:=2;
      Consultas(8);
    end;
  end;
finally
  Selec_Cuim.Free;
end;
end;

procedure TAM_Contribuyentes.bancoExit(Sender: TObject);
begin
if banco.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Banco válido',mterror,[mbok],0);
  banco.SetFocus;
end
else begin
 CargaCombos(9);
 //if BANCO.ItemIndex =0  then CargaCombos(9)
 // else CargaCombos(10);
 nrosucursal.Text :=TComboClase(banco.Items.Objects[banco.ItemIndex]).Param2;
end;
end;

procedure TAM_Contribuyentes.bancoKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  tipoctabanco.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L11KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  L10.SetFocus;
end;
end;

procedure TAM_Contribuyentes.L11Exit(Sender: TObject);
begin
if L11.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una opción válida',mterror,[mbok],0);
  L11.SetFocus;
end;
end;

end.
