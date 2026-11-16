unit AM_DerConstr;


interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DB, ADODB, Mask, Buttons, Grids, DBGrids,
  ExtCtrls, StrUtils, JvExDBGrids, JvDBGrid, JvgGroupBox, JvExButtons, JvBitBtn,
  ToolWin, JvExGrids, JvStringGrid, Menus, ImgList, JvSimLogic,
   jpeg, JvComponentBase, JvEmbeddedForms,Math,Objetos,frxClass, frxDBSet,frxBarcode ;

type THackStringGrid = class(TJvStringGrid);
type
  TAM_ExptesDerConst = class(TForm)
    Barra1: TStatusBar;
    DataSource1: TDataSource;
    PC1: TPageControl;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    BORRA: TJvBitBtn;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    SALE: TJvBitBtn;
    ImageList1: TImageList;
    DataSource2: TDataSource;
    TabSheet2: TTabSheet;
    TabSheet6: TTabSheet;
    PC4: TPageControl;
    TabSheet7: TTabSheet;
    Label10: TLabel;
    Label15: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label11: TLabel;
    Label37: TLabel;
    Label14: TLabel;
    Label1: TLabel;
    N2: TEdit;
    N6: TEdit;
    N7: TEdit;
    N8: TEdit;
    N4: TEdit;
    N3: TEdit;
    N9: TEdit;
    OPARTIDA: TMaskEdit;
    BuscarP: TButton;
    GroupBox5: TGroupBox;
    Label30: TLabel;
    Label53: TLabel;
    Label54: TLabel;
    Label55: TLabel;
    Label56: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    Label62: TLabel;
    Label64: TLabel;
    BARRIOINMU: TComboBox;
    CALLEINMU: TComboBox;
    DOMIINMU: TEdit;
    ALTURAINMU: TEdit;
    PISOINMU: TEdit;
    DEPTOINMU: TEdit;
    MZINMU: TEdit;
    PARINMU: TEdit;
    TORREINMU: TEdit;
    CASAINMU: TEdit;
    SECTORINMU: TEdit;
    LRINMU: TEdit;
    Label65: TLabel;
    DOMOBRA: TEdit;
    guarda: TJvBitBtn;
    TabSheet8: TTabSheet;
    JvgGroupBox6: TJvgGroupBox;
    JvgGroupBox7: TJvgGroupBox;
    JvgGroupBox0: TJvgGroupBox;
    grilla2: TJvStringGrid;
    Label68: TLabel;
    TOTAL: TEdit;
    EXEPCION: TCheckBox;
    CARGAMANUAL: TCheckBox;
    ORDENANZA: TCheckBox;
    GroupBox7: TGroupBox;
    PLANO: TEdit;
    RESOLUCION: TEdit;
    FINALOBRA: TEdit;
    Label71: TLabel;
    Label70: TLabel;
    Label69: TLabel;
    COMENTARIOS: TMemo;
    FormularioEmb: TJvEmbeddedFormLink;
    TabSheet1: TTabSheet;
    JvgGroupBox3: TJvgGroupBox;
    JvDBGrid1: TJvDBGrid;
    GroupBox6: TGroupBox;
    C6: TCheckBox;
    C3: TCheckBox;
    C4: TCheckBox;
    C5: TCheckBox;
    C7: TCheckBox;
    C8: TCheckBox;
    JvgGroupBox1: TJvgGroupBox;
    LB1: TLabel;
    Label9: TLabel;
    Edit1: TEdit;
    BusExpte: TJvBitBtn;
    DataSource3: TDataSource;
    SECTOR: TComboBox;
    NUEVO: TJvBitBtn;
    TabSheet4: TTabSheet;
    E1: TRichEdit;
    GroupBox8: TGroupBox;
    Label5: TLabel;
    ExisDoc: TEdit;
    PopupMenu1: TPopupMenu;
    ImprimirFormulario1: TMenuItem;
    ToolButton5: TToolButton;
    frxReport1: TfrxReport;
    TabSheet9: TTabSheet;
    PageControl2: TPageControl;
    TabSheet5: TTabSheet;
    JvgGroupBox4: TJvgGroupBox;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    BUSCAR: TJvBitBtn;
    ODNI: TEdit;
    ONOMB: TEdit;
    OCUIT: TEdit;
    JvDBGrid2: TJvDBGrid;
    GroupBox4: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    Label6: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    BARRIORESPO: TComboBox;
    CALLERESPO: TComboBox;
    DOMRESPO: TEdit;
    ALTURARESPO: TEdit;
    PISORESPO: TEdit;
    DEPTORESPO: TEdit;
    MZRESPO: TEdit;
    PARRESPO: TEdit;
    TORRERESPO: TEdit;
    CASARESPO: TEdit;
    SECTORRESPO: TEdit;
    LRRESPO: TEdit;
    TabSheet10: TTabSheet;
    Label12: TLabel;
    TPOTIT: TComboBox;
    Label13: TLabel;
    tpopartida: TComboBox;
    Label16: TLabel;
    emailRESPO: TEdit;
    Label17: TLabel;
    telefonorespo: TEdit;
    PageControl3: TPageControl;
    TabSheet11: TTabSheet;
    JvgGroupBox2: TJvgGroupBox;
    Label18: TLabel;
    Label19: TLabel;
    BUSCARPROF: TJvBitBtn;
    ODNIP: TEdit;
    ONOMBP: TEdit;
    JvDBGrid3: TJvDBGrid;
    DataSource4: TDataSource;
    c2: TCheckBox;
    GroupBox10: TGroupBox;
    Shape1: TShape;
    Label73: TLabel;
    Label76: TLabel;
    Shape2: TShape;
    Label77: TLabel;
    Shape3: TShape;
    Label78: TLabel;
    Shape4: TShape;
    GroupBox1: TGroupBox;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    NROE: TEdit;
    AÑO: TEdit;
    LETRA: TEdit;
    GroupBox2: TGroupBox;
    Label36: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    AÑO1: TEdit;
    NRO1: TEdit;
    LETRA1: TEdit;
    TabSheet12: TTabSheet;
    GroupBox3: TGroupBox;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label66: TLabel;
    Label4: TLabel;
    tiponotif: TComboBox;
    RELIQUI: TCheckBox;
    VISADO: TEdit;
    AÑOVISADO: TEdit;
    REVOCADO: TEdit;
    REVOCADOAÑO: TEdit;
    OPUBLICA: TCheckBox;
    DESTINOOBRA: TEdit;
    empresa: TEdit;
    PageControl1: TPageControl;
    TabSheet3: TTabSheet;
    TabSheet13: TTabSheet;
    GroupBox9: TGroupBox;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label63: TLabel;
    Label67: TLabel;
    Label72: TLabel;
    Label74: TLabel;
    Label75: TLabel;
    Label20: TLabel;
    barriop: TComboBox;
    callep: TComboBox;
    domp: TEdit;
    alturap: TEdit;
    pisop: TEdit;
    deptop: TEdit;
    mzp: TEdit;
    parp: TEdit;
    torrep: TEdit;
    casap: TEdit;
    sectorp: TEdit;
    lrp: TEdit;
    emailp: TEdit;
    telefonop: TEdit;
    matricula: TEdit;
    agrega: TJvBitBtn;
    grillaprofes: TStringGrid;
    Edit2: TEdit;
    Label79: TLabel;
    L3: TMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CargaCombos(nro:byte);
    procedure Consultas(nro:byte);
    procedure BORRAClick(Sender: TObject);
    procedure ODNIKeyPress(Sender: TObject; var Key: Char);
    procedure ONOMBKeyPress(Sender: TObject; var Key: Char);
    procedure BuscarClick(Sender: TObject);
    procedure SALEClick(Sender: TObject);
    procedure JvDBGrid2CellClick(Column: TColumn);
    procedure OPARTIDAChange(Sender: TObject);
    procedure OPARTIDAEnter(Sender: TObject);
    procedure OPARTIDAExit(Sender: TObject);
    procedure BuscarPClick(Sender: TObject);
    procedure guardaClick(Sender: TObject);
    procedure OPARTIDAKeyPress(Sender: TObject; var Key: Char);
    procedure NROEKeyPress(Sender: TObject; var Key: Char);
    procedure LETRAKeyPress(Sender: TObject; var Key: Char);
    procedure LETRAExit(Sender: TObject);
    procedure NROEExit(Sender: TObject);
    procedure AÑOExit(Sender: TObject);
    procedure LETRA1KeyPress(Sender: TObject; var Key: Char);
    procedure LETRA1Exit(Sender: TObject);
    procedure NRO1KeyPress(Sender: TObject; var Key: Char);
    procedure NRO1Exit(Sender: TObject);
    procedure AÑOKeyPress(Sender: TObject; var Key: Char);
    procedure AÑO1Exit(Sender: TObject);
    procedure AÑO1KeyPress(Sender: TObject; var Key: Char);
    procedure tiponotifExit(Sender: TObject);
    procedure tiponotifKeyPress(Sender: TObject; var Key: Char);
    procedure DESTINOOBRAExit(Sender: TObject);
    procedure DESTINOOBRAKeyPress(Sender: TObject; var Key: Char);
    procedure VISADOKeyPress(Sender: TObject; var Key: Char);
    procedure AÑOVISADOKeyPress(Sender: TObject; var Key: Char);
    procedure REVOCADOKeyPress(Sender: TObject; var Key: Char);
    procedure REVOCADOAÑOKeyPress(Sender: TObject; var Key: Char);
    procedure BARRIORESPOExit(Sender: TObject);
    procedure CALLERESPOExit(Sender: TObject);
    procedure BARRIORESPOKeyPress(Sender: TObject; var Key: Char);
    procedure CALLERESPOKeyPress(Sender: TObject; var Key: Char);
    procedure DOMRESPOKeyPress(Sender: TObject; var Key: Char);
    procedure ALTURARESPOKeyPress(Sender: TObject; var Key: Char);
    procedure ALTURARESPOExit(Sender: TObject);
    procedure PISORESPOKeyPress(Sender: TObject; var Key: Char);
    procedure DEPTORESPOKeyPress(Sender: TObject; var Key: Char);
    procedure MZRESPOKeyPress(Sender: TObject; var Key: Char);
    procedure TORRERESPOKeyPress(Sender: TObject; var Key: Char);
    procedure CASARESPOKeyPress(Sender: TObject; var Key: Char);
    procedure SECTORRESPOKeyPress(Sender: TObject; var Key: Char);
    procedure PARRESPOKeyPress(Sender: TObject; var Key: Char);
    procedure LRRESPOKeyPress(Sender: TObject; var Key: Char);
    procedure BARRIOINMUExit(Sender: TObject);
    procedure CALLEINMUExit(Sender: TObject);
    procedure BARRIOINMUKeyPress(Sender: TObject; var Key: Char);
    procedure CALLEINMUKeyPress(Sender: TObject; var Key: Char);
    procedure DOMIINMUKeyPress(Sender: TObject; var Key: Char);
    procedure ALTURAINMUExit(Sender: TObject);
    procedure ALTURAINMUKeyPress(Sender: TObject; var Key: Char);
    procedure PISOINMUKeyPress(Sender: TObject; var Key: Char);
    procedure DEPTOINMUKeyPress(Sender: TObject; var Key: Char);
    procedure MZINMUKeyPress(Sender: TObject; var Key: Char);
    procedure TORREINMUKeyPress(Sender: TObject; var Key: Char);
    procedure CASAINMUKeyPress(Sender: TObject; var Key: Char);
    procedure SECTORINMUKeyPress(Sender: TObject; var Key: Char);
    procedure PARINMUKeyPress(Sender: TObject; var Key: Char);
    procedure LRINMUKeyPress(Sender: TObject; var Key: Char);
    procedure DOMOBRAKeyPress(Sender: TObject; var Key: Char);
    procedure grilla2DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure grilla2ExitCell(Sender: TJvStringGrid; AColumn,
      ARow: Integer; const EditText: string);
    procedure grilla2KeyPress(Sender: TObject; var Key: Char);
    procedure grilla2SelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EXEPCIONKeyPress(Sender: TObject; var Key: Char);
    procedure ORDENANZAKeyPress(Sender: TObject; var Key: Char);
    procedure PLANOKeyPress(Sender: TObject; var Key: Char);
    procedure RESOLUCIONKeyPress(Sender: TObject; var Key: Char);
    procedure FINALOBRAKeyPress(Sender: TObject; var Key: Char);
    procedure empresaKeyPress(Sender: TObject; var Key: Char);
    procedure SECTORKeyPress(Sender: TObject; var Key: Char);
    procedure SECTORExit(Sender: TObject);
    procedure C6Click(Sender: TObject);
    procedure JvDBGrid1CellClick(Column: TColumn);
    procedure LETRAChange(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure BARRIORESPOChange(Sender: TObject);
    procedure COMENTARIOSChange(Sender: TObject);
    procedure TOTALChange(Sender: TObject);
    procedure NUEVOClick(Sender: TObject);
    procedure BusExpteClick(Sender: TObject);
    procedure VISADOExit(Sender: TObject);
    procedure AÑOVISADOExit(Sender: TObject);
    procedure REVOCADOExit(Sender: TObject);
    procedure REVOCADOAÑOExit(Sender: TObject);
    procedure PLANOExit(Sender: TObject);
    procedure RESOLUCIONExit(Sender: TObject);
    procedure FINALOBRAExit(Sender: TObject);
    Function Verificacion(nro:byte):boolean;
    procedure ExisDocExit(Sender: TObject);
    procedure ExisDocChange(Sender: TObject);
    procedure DOMRESPOExit(Sender: TObject);
    procedure TPOTITKeyPress(Sender: TObject; var Key: Char);
    procedure TPOTITExit(Sender: TObject);
    procedure tpopartidaExit(Sender: TObject);
    procedure tpopartidaKeyPress(Sender: TObject; var Key: Char);
    procedure emailRESPOKeyPress(Sender: TObject; var Key: Char);
    procedure telefonorespoKeyPress(Sender: TObject; var Key: Char);
    procedure BUSCARPROFClick(Sender: TObject);
    procedure JvDBGrid3CellClick(Column: TColumn);
    procedure emailRESPOChange(Sender: TObject);
    procedure telefonorespoChange(Sender: TObject);
    procedure emailpChange(Sender: TObject);
    procedure telefonopChange(Sender: TObject);
    procedure barriopExit(Sender: TObject);
    procedure callepExit(Sender: TObject);
    procedure barriopKeyPress(Sender: TObject; var Key: Char);
    procedure callepKeyPress(Sender: TObject; var Key: Char);
    procedure dompKeyPress(Sender: TObject; var Key: Char);
    procedure alturapExit(Sender: TObject);
    procedure alturapKeyPress(Sender: TObject; var Key: Char);
    procedure pisopKeyPress(Sender: TObject; var Key: Char);
    procedure deptopKeyPress(Sender: TObject; var Key: Char);
    procedure mzpKeyPress(Sender: TObject; var Key: Char);
    procedure torrepKeyPress(Sender: TObject; var Key: Char);
    procedure casapKeyPress(Sender: TObject; var Key: Char);
    procedure sectorpKeyPress(Sender: TObject; var Key: Char);
    procedure parpKeyPress(Sender: TObject; var Key: Char);
    procedure lrpKeyPress(Sender: TObject; var Key: Char);
    procedure matriculaKeyPress(Sender: TObject; var Key: Char);
    procedure emailpKeyPress(Sender: TObject; var Key: Char);
    procedure ODNIPKeyPress(Sender: TObject; var Key: Char);
    procedure ONOMBPKeyPress(Sender: TObject; var Key: Char);
    procedure telefonopKeyPress(Sender: TObject; var Key: Char);
    procedure JvDBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure agregaClick(Sender: TObject);
    procedure Edit2Exit(Sender: TObject);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit1Exit(Sender: TObject);
    procedure L3Exit(Sender: TObject);
    procedure L3KeyPress(Sender: TObject; var Key: Char);
  private
    Lista:array of string;
    CONTRIBUYENTE,EXPEDIENTES,DataSet1,DataSet2,OBRAS,PROFESIONAL:TADODataSet;
    Com1:TADOCommand;
    Recibo:TRecibo;
    gh:boolean;
    { Private declarations }
  public
    sumdoc:real;
    yaexiste,Pos4,selec:boolean;
    nomb,losperi,elcuim,cuim,losdetalles,Contr,losid,elidentif,idcontrib,idcontribp,cuimp,eltituP,
    fechas,lares,idexpte,lapartida,eltitu,FechaAct,idexptecont,cuimexptecont,nomexptecont,elexpte:string;
    seleccontrib:boolean;
    { Public declarations }
  end;

var
  AM_ExptesDerConst: TAM_ExptesDerConst;

implementation

uses Unidad, modulo;

{$R *.dfm}
procedure TAM_ExptesDerConst.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  FreeAndNil(CONTRIBUYENTE);
  freeAndNil(PROFESIONAL);
  FreeAndNil(EXPEDIENTES);
  FreeAndNil(OBRAS);
  FreeAndNil(DataSet2);
  FreeAndNil(DataSet1);
  FreeAndNil(Self);
end;

procedure TAM_ExptesDerConst.FormCreate(Sender: TObject);
begin
  gh:=false;
  Modulo2.ObjConexion:=nil;
  CONTRIBUYENTE:=TADODataSet.Create(Self);
  EXPEDIENTES:=TADODataSet.Create(Self);
  OBRAS:=TADODataSet.Create(Self);
  PROFESIONAL:=TADODataSet.Create(Self);
  DataSet1:=TADODataSet.Create(Self);
  DataSet2:=TADODataSet.Create(Self);
  pc1.Pages[1].TabVisible:=false;    pc1.Pages[2].TabVisible:=false;
  pc1.Pages[3].TabVisible:=false;    pc1.Pages[4].TabVisible:=false;
  pc1.Pages[5].TabVisible:=false;    pc1.Pages[6].TabVisible:=false;
  pc1.Pages[7].TabVisible:=false;
end;

procedure TAM_ExptesDerConst.FormActivate(Sender: TObject);
begin
if not (gh) then begin
  gh:=true;
  selec:=false;
  PageControl1.ActivePage:=PageControl1.Pages[0];
  PC1.ActivePage:=PC1.Pages[0];
  grilla2.Cells[0,0]:='Id';                 grilla2.Cells[1,0]:='Descripcion';
  grilla2.Cells[2,0]:='Metros';             grilla2.Cells[3,0]:='Costo';
  grilla2.Cells[4,0]:='Porcentaje';         grilla2.Cells[5,0]:='Total';
  grilla2.Cells[6,0]:='IdExp';              grilla2.Cells[7,0]:='Id';

  grillaprofes.Cells[0,0]:='IdExpediente';     grillaprofes.Cells[1,0]:='Cuim';
  grillaprofes.Cells[2,0]:='CodProv';          grillaprofes.Cells[3,0]:='CodBar';
  grillaprofes.Cells[4,0]:='CodCal';           grillaprofes.Cells[5,0]:='Domicilio';
  grillaprofes.Cells[6,0]:='Altura';           grillaprofes.Cells[7,0]:='Piso';
  grillaprofes.Cells[8,0]:='Dpto';             grillaprofes.Cells[9,0]:='Mz';
  grillaprofes.Cells[10,0]:='Torre';           grillaprofes.Cells[11,0]:='Casa';
  grillaprofes.Cells[12,0]:='Sector';          grillaprofes.Cells[13,0]:='Parcela';
  grillaprofes.Cells[14,0]:='LoteRural';       grillaprofes.Cells[15,0]:='Observaciones';
  grillaprofes.Cells[16,0]:='Email';           grillaprofes.Cells[17,0]:='Telefono';
  grillaprofes.Cells[18,0]:='NroMatricula	';
  CargaCombos(1);
  seleccontrib:=false;
  FechaAct:=LeftStr(Fec_System(),10);
  l3.Text :=FechaAct;
  guarda.Enabled:=false;
  nuevo.Enabled:=false;
  C3.Checked:=false;
  Pos4:=false;
  Consultas(1);
  CargaCombos(2);
  CargaCombos(3);      CargaCombos(4);
  CargaCombos(5);      //CargaCombos(6);
  CargaCombos(7);      CargaCombos(8);
  AÑO.Text:=RightStr(strcut(fec_system(),10),4);
  CARGAMANUAL.Enabled:=false;
  Consultas(3);
  EDIT2.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.CargaCombos(nro:byte);
begin
 if nro=1 then Modulo2.CargaCombo(TIPONOTIF,'SELECT Detalle,iD FROM TipoNotificacion',modulo2.Sistema.ConectDerConst,false);
 if nro=2 then Modulo2.CargaCombo(BARRIORESPO,'SELECT Detalle,Id FROM Barrios ORDER BY Detalle',modulo2.Sistema.ConectSeg,false);
 if nro=3 then Modulo2.CargaCombo(CALLERESPO,'SELECT Descrip,Calle FROM CALLES WHERE (Dpto_Polit=1) AND'
 + ' (LOCALIDAD=1) ORDER BY DESCRIP',modulo2.Sistema.ConectUsu,false);
 if nro=4 then Modulo2.CargaCombo(BARRIOINMU,'SELECT Detalle,Id FROM Barrios ORDER BY Detalle',modulo2.Sistema.ConectSeg,false);
 if nro=5 then Modulo2.CargaCombo(CALLEINMU,'SELECT Descrip,Calle FROM CALLES WHERE (Dpto_Polit=1) AND'
 + ' (LOCALIDAD=1) ORDER BY DESCRIP',modulo2.Sistema.ConectUsu,false);
 if nro=6 then Modulo2.CargaCombo(sector,'SELECT distinct substring([NroExpediente],1,1) '
 +' FROM [DerConst].[dbo].[Expediente] where substring([NroExpediente],1,1) <> '''' order by 1',modulo2.Sistema.ConectDerConst,false);
 if nro=7 then Modulo2.CargaCombo(BARRIOP,'SELECT Detalle,Id FROM Barrios ORDER BY Detalle',modulo2.Sistema.ConectSeg,false);
 if nro=8 then Modulo2.CargaCombo(CALLEP,'SELECT Descrip,Calle FROM CALLES WHERE (Dpto_Polit=1) AND'
 + ' (LOCALIDAD=1) ORDER BY DESCRIP',modulo2.Sistema.ConectUsu,false);
end;

Function TAM_ExptesDerConst.Verificacion(nro:byte):boolean;
begin
 E1.Clear;
 if nro = 1 then begin //Altas
   if NOT C6.Checked  then BEGIN
      if seleccontrib= false then E1.Lines.Add('NO existe Ningun Contribuyente Seleccionado');
   END;
   if n2.Text = '' then E1.Lines.Add('NO existe Ninguna PARTIDA Seleccionada');
   //if total.Text = '' then E1.Lines.Add('NO existe Ningun IMPORTE del Expediente Calculado');
 end;
 if E1.Lines.Count > 0 then begin
  pc1.Pages[7].TabVisible:=true;
  Result:=false;
  PC1.ActivePage:=PC1.Pages[7];
 end
 else Result:=true;
end;

procedure TAM_ExptesDerConst.CASAINMUKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  sectorinmu.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.casapKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  sectorp.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.CASARESPOKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  sectorrespo.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.COMENTARIOSChange(Sender: TObject);
begin
c8.Checked :=true;
end;

procedure TAM_ExptesDerConst.Consultas(nro:byte);
var t,y,gf,a:integer;
    eltotal:real;
    txt,dni,lapar,dnip:string;
    CLIENTE:TADODataSet;
    Comando:widestring;
    Cnx2:TADOConnection;
begin
Modulo2.ObjConexion:=TADOConnection.Create(nil);
Modulo2.ObjConexion.LoginPrompt:=false;
if nro = 0 then begin
  DataSet1.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.CommandText:='SELECT detalle FROM [SeguridadWEB].[dbo].[Fn_ElDomicilioContribuyente] (' + idcontrib + ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   //barriorespo.ItemIndex:=barriorespo.Items.IndexOf(trim(DataSet1.Fields[1].AsString));
   //callerespo.ItemIndex:=callerespo.Items.IndexOf(trim(DataSet1.Fields[2].AsString));
   domrespo.text:= trim(DataSet1.Fields[0].AsString);
   //alturarespo.text:=  trim(DataSet1.Fields[4].AsString);
   //pisorespo.text:=  trim(DataSet1.Fields[5].AsString);
   //deptorespo.text:=  trim(DataSet1.Fields[6].AsString);
   //mzrespo.text:=    trim(DataSet1.Fields[7].AsString);
   //torrerespo.text:=  trim(DataSet1.Fields[11].AsString);
   //casarespo.text:=  trim(DataSet1.Fields[9].AsString);
   //sectorrespo.text:=    trim(DataSet1.Fields[10].AsString);
   //parrespo.text:=   trim(DataSet1.Fields[8].AsString);
   //lrrespo.text:=   trim(DataSet1.Fields[12].AsString);
   //domrespo.setfocus;
  end;
  DataSet1.Close;

  DataSet1.CommandText:='SELECT Id,Barrio,Calle,CodPostal,Altura,Piso,Dpto,Mz,Parcela,Casa,Sector,'  //10
  + 'Torre,LoteRural,Domicilio,CodBar,CodCal FROM dbo.Fn_ContribuyentesDomicilios(' + idcontrib + ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   barriorespo.ItemIndex:=barriorespo.Items.IndexOf(trim(DataSet1.Fields[1].AsString));
   callerespo.ItemIndex:=callerespo.Items.IndexOf(trim(DataSet1.Fields[2].AsString));
   //domrespo.text:= trim(DataSet1.Fields[0].AsString);
   //alturarespo.text:=  trim(DataSet1.Fields[4].AsString);
   //pisorespo.text:=  trim(DataSet1.Fields[5].AsString);
   //deptorespo.text:=  trim(DataSet1.Fields[6].AsString);
   //mzrespo.text:=    trim(DataSet1.Fields[7].AsString);
   //torrerespo.text:=  trim(DataSet1.Fields[11].AsString);
   //casarespo.text:=  trim(DataSet1.Fields[9].AsString);
   //sectorrespo.text:=    trim(DataSet1.Fields[10].AsString);
   //parrespo.text:=   trim(DataSet1.Fields[8].AsString);
   //lrrespo.text:=   trim(DataSet1.Fields[12].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
  //domrespo.setfocus;
end;
if nro = 1 then begin
  for t:=1 to Grilla2.RowCount-1 do
  for y:=0 to Grilla2.ColCount-1 do
  Grilla2.Cells[y,t]:='';
  Grilla2.RowCount:=2;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectDerConst;
  CLIENTE:=TADODataSet.Create(nil);
  CLIENTE.Connection:=modulo2.ObjConexion;
  CLIENTE.CommandText:='SELECT Id, Detalle, Porcentaje FROM TipoObra';
  CLIENTE.Open;
  for t:=1 to CLIENTE.RecordCount do begin
   CLIENTE.Recno:=t;
   Grilla2.RowCount:=t+1;
   Grilla2.Cells[0,t]:=CLIENTE.FieldByName('ID').AsString;
   Grilla2.Cells[1,t]:=CLIENTE.FieldByName('Detalle').AsString;
   Grilla2.Cells[2,t]:='0.00';
   Grilla2.Cells[3,t]:='0.00';
   Grilla2.Cells[4,t]:=AnsiReplaceStr(CLIENTE.FieldByName('Porcentaje').AsString,',','.');
   Grilla2.Cells[5,t]:='0.00';
  end;
  CLIENTE.Connection:=nil;
end;
if nro = 2  then begin
  DataSet1.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.CommandText:='SELECT TOP (1) CUIM FROM ContribuyentesIdentidad  '
  + ' WHERE  (IdContribuyente = ' + idcontrib + ') AND (Activo = 1) ORDER BY Id';
  DataSet1.Open;
  if not DataSet1.isempty then begin
    elcuim:= trim(DataSet1.Fields[0].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;
if nro = 3 then begin
  EXPEDIENTES.Close;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectDerConst;
  EXPEDIENTES.Connection:=modulo2.ObjConexion;
  DataSource3.DataSet:=EXPEDIENTES;
  EXPEDIENTES.CommandText:='SELECT TOP 100 ID,nroexpediente,nroexpedienteantecedente,TipoNotificacion,reliquidacion,revocarnro,'  //5
  +'resolucionanio,recibonro,esobrapublica,empresaconstructora,adjuntosplanos,adjuntoresoluciones,adjuntofinalobra,' //12
  +'anulado,idtiponotif,left(nroexpediente,1),RIGHT(nroexpediente,4),SUBSTRING(NroExpediente,2,8),' //17
  +'left(nroexpedienteantecedente,1),'
  +'case when len(nroexpedienteantecedente)> 11 then RIGHT(nroexpedienteantecedente,4) else RIGHT(nroexpedienteantecedente,1) end'
  +',SUBSTRING(nroexpedienteantecedente,2,8),DestinoObra,RevocarAnio,ResolucionNro,partida,MetrosDocumentados,'
  +'IdContribuyente,cuimExpte,TituExpte,Confirmado' //29
  //EXPEDIENTES.CommandText:='SELECT ID,nroexpediente,nroexpedienteantecedente,TipoNotificacion,reliquidacion,revocarnro,'
  //+'resolucionanio,recibonro,esobrapublica,empresaconstructora,adjuntosplanos,adjuntoresoluciones,adjuntofinalobra,'
  //+'anulado,idtiponotif, '
  //+'left(nroexpediente,1),RIGHT(nroexpediente,4),SUBSTRING(NroExpediente,2,8),' //17
  //+'left(nroexpedienteantecedente,1) ,RIGHT(nroexpedienteantecedente,1),SUBSTRING(nroexpedienteantecedente,2,8),DestinoObra '
  +' FROM [DerConst].[dbo].[FN_ObtenerExpedientes] (null) order by id desc';
  EXPEDIENTES.Open;
  if EXPEDIENTES.IsEmpty then BEGIN
   messagedlg('Expediente de Der. de Construccion Inexistente...' + chr(13) + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
   PC1.ActivePage:=PC1.Pages[0];
   guarda.Caption :='GUARDAR';
   c6.Checked :=false;
   EDIT2.SetFocus;
  END;
  EXPEDIENTES.Connection:=nil;
end;
if nro = 4 then begin
  EXPEDIENTES.Close;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectDerConst;
  EXPEDIENTES.Connection:=modulo2.ObjConexion;
  DataSource3.DataSet:=EXPEDIENTES;
  EXPEDIENTES.CommandText:='SELECT ID,nroexpediente,nroexpedienteantecedente,TipoNotificacion,reliquidacion,revocarnro,'  //5
  +'resolucionanio,recibonro,esobrapublica,empresaconstructora,adjuntosplanos,adjuntoresoluciones,adjuntofinalobra,' //12
  +'anulado,idtiponotif,left(nroexpediente,1),RIGHT(nroexpediente,4),SUBSTRING(NroExpediente,2,8),' //17
  +'left(nroexpedienteantecedente,1)'
  +',case when len(nroexpedienteantecedente)> 11 then RIGHT(nroexpedienteantecedente,4) else RIGHT(nroexpedienteantecedente,1) end'
  +',SUBSTRING(nroexpedienteantecedente,2,8),DestinoObra,RevocarAnio,ResolucionNro,partida,'
  +'MetrosDocumentados,IdContribuyente,cuimExpte,TituExpte,Confirmado' //29
  +' FROM FN_ObtenerExpedientes (null) where substring(NroExpediente,1,1)=' + chr(39) + trim(EDIT2.Text) + chr(39)
  + ' and substring(NroExpediente,2,8) LIKE ' + chr(39) + chr(37) + AnsiReplaceStr(trim(Edit1.Text),chr(39),chr(33)) + chr(39);
  EXPEDIENTES.Open;
  if EXPEDIENTES.IsEmpty then BEGIN
   nuevo.Enabled:=true;
   nuevo.SetFocus;
  END
  else begin
   nuevo.Enabled:=false;
  end;
  EXPEDIENTES.Connection:=nil;
end;
if nro = 5 then begin /// traemos el contrib de un expediente
  DataSet1.ConnectionString:=modulo2.Sistema.ConectDerConst;
  DataSet1.CommandText:='SELECT IdExpte,Cuim , Contribuyente ,NroDni ,barrio, calle ,Domicilio,Altura , '
	+'Piso,Dpto,Mz,Torre,Casa ,Sector ,Parcela,LoteRural ,idcontrib FROM FN_ObtenerExpedientesContrib ( '+ idexpte+ ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   dni:= trim(DataSet1.Fields[3].AsString);
   barriorespo.ItemIndex:=barriorespo.Items.IndexOf(trim(DataSet1.Fields[4].AsString));
   callerespo.ItemIndex:=callerespo.Items.IndexOf(trim(DataSet1.Fields[5].AsString));
   domrespo.text:= trim(DataSet1.Fields[6].AsString);
   alturarespo.text:=trim(DataSet1.Fields[7].AsString);
   pisorespo.text:= trim(DataSet1.Fields[8].AsString);
   deptorespo.text:=trim(DataSet1.Fields[9].AsString);
   mzrespo.text:=   trim(DataSet1.Fields[10].AsString);
   torrerespo.text:=trim(DataSet1.Fields[11].AsString);
   casarespo.text:= trim(DataSet1.Fields[12].AsString);
   sectorrespo.text:=trim(DataSet1.Fields[13].AsString);
   parrespo.text:=  trim(DataSet1.Fields[14].AsString);
   lrrespo.text:=   trim(DataSet1.Fields[15].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;

  // lleno la grilla
  CONTRIBUYENTE.Close;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectSeg;
  CONTRIBUYENTE.Connection:=modulo2.ObjConexion;
  DataSource1.DataSet:=CONTRIBUYENTE;
  CONTRIBUYENTE.CommandText:='EXEC SPP_ContribuyentesDerConst 1,null,' + dni + ',' + idexpte;
  CONTRIBUYENTE.Open;
  if CONTRIBUYENTE.IsEmpty then BEGIN
   messagedlg('Contribuyente Inexistente...' + chr(13) + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
   PC1.ActivePage:=PC1.Pages[0];
   odni.SetFocus;
  END;
  CONTRIBUYENTE.Connection:=nil;
end;
if nro = 6  then  begin   // traemos la partida del expte
 DataSet1.ConnectionString:=modulo2.Sistema.ConectDerConst;
  DataSet1.CommandText:='SELECT IdExpte,Partida ,barrio, calle ,Domicilio,Altura , '
	+'Piso,Dpto ,Mz,Torre,Casa ,Sector ,Parcela,LoteRural,DomObra FROM FN_ObtenerExpedientesInmuebles ( '+ idexpte+ ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   lapar:= trim(DataSet1.Fields[2].AsString);
   barrioinmu.ItemIndex:=barrioinmu.Items.IndexOf(trim(DataSet1.Fields[2].AsString));
   calleinmu.ItemIndex:=calleinmu.Items.IndexOf(trim(DataSet1.Fields[3].AsString));
   domiinmu.text:= trim(DataSet1.Fields[4].AsString);
   alturainmu.text:=trim(DataSet1.Fields[5].AsString);
   pisoinmu.text:= trim(DataSet1.Fields[6].AsString);
   deptoinmu.text:=trim(DataSet1.Fields[7].AsString);
   mzinmu.text:=   trim(DataSet1.Fields[8].AsString);
   torreinmu.text:=trim(DataSet1.Fields[9].AsString);
   casainmu.text:= trim(DataSet1.Fields[10].AsString);
   sectorinmu.text:=trim(DataSet1.Fields[11].AsString);
   parinmu.text:=  trim(DataSet1.Fields[12].AsString);
   lrinmu.text:=   trim(DataSet1.Fields[13].AsString);
   domobra.text:=  trim(DataSet1.Fields[14].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;
if nro = 7 then begin // buscamos un contrib
 if trim(ONOMB.Text) = '' then txt:='NULL'
   else txt:=chr(39) + trim(ONOMB.Text) + chr(39);
 if trim(ODNI.Text) = '' then txt:=txt + ' , NULL'
   else txt:=txt + ' , ' + chr(39) + trim(ODNI.Text) + chr(39);
 if trim(OCUIT.Text) = '' then txt:=txt + ' , NULL'
   else txt:=txt + ' , ' + chr(39) + trim(OCUIT.Text) + chr(39);
 CONTRIBUYENTE.Close;
 Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectSeg;
 CONTRIBUYENTE.Connection:=modulo2.ObjConexion;
 DataSource1.DataSet:=CONTRIBUYENTE;
 CONTRIBUYENTE.CommandText:='EXEC SPP_ContribuyentesDerConst 0,' + txt ;
 CONTRIBUYENTE.Open;
 if CONTRIBUYENTE.IsEmpty then BEGIN
  messagedlg('Contribuyente Inexistente...' + chr(13) + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
  PC1.ActivePage:=PC1.Pages[3];
  odni.SetFocus;
 END;
 CONTRIBUYENTE.Connection:=nil;
end;
if nro = 8  then  begin  /// traemos la grilla de importes
 eltotal:=0;
 DataSet1.ConnectionString:=modulo2.Sistema.ConectDerConst;
 DataSet1.CommandText:='SELECT IdTipoObra,Metros,Costo,Porcentaje,Importe,idexpediente,id FROM ExpedienteCalculo '
 +' WHERE   (IdExpediente = ' + idexpte + ') ORDER BY IdTipoObra';
 DataSet1.Open;
 for t:=1 to DataSet1.RecordCount do begin
  grilla2.RowCount:=t+1;
  DataSet1.RecNo:=t;
  grilla2.Cells[2,t]:=DataSet1.Fields[1].AsString;
  grilla2.Cells[3,t]:=DataSet1.Fields[2].AsString;
  grilla2.Cells[4,t]:=DataSet1.Fields[3].AsString;
  grilla2.Cells[5,t]:=DataSet1.Fields[4].AsString;
  grilla2.Cells[6,t]:=DataSet1.Fields[5].AsString;
  grilla2.Cells[7,t]:=DataSet1.Fields[6].AsString;
 end;
 DataSet1.Close;
 DataSet1.Connection:=nil;
  for a:=1 to grilla2.RowCount-1 do begin
   eltotal := eltotal + strtofloat(Grilla2.Cells[5,a]);
  end;
 total.Text := floattostr(eltotal);
end;
if nro = 9 then begin
  DataSet1.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.CommandText:='SELECT detalle FROM [SeguridadWEB].[dbo].[Fn_ElDomicilioContribuyente ] (' + idexptecont + ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   domrespo.text:= trim(DataSet1.Fields[0].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
end;
if nro = 10 then begin // buscamos un contrib
  if trim(ONOMBp.Text) = '' then txt:='NULL'
    else txt:=chr(39) + trim(ONOMBp.Text) + chr(39);
  if trim(ODNIp.Text) = '' then txt:=txt + ' , NULL'
    else txt:=txt + ' , ' + chr(39) + trim(ODNIp.Text) + chr(39);
  if trim(OCUIT.Text) = '' then txt:=txt + ' , NULL'
    else txt:=txt + ' , ' + chr(39) + trim(OCUIT.Text) + chr(39);
  PROFESIONAL.Close;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectSeg;
  PROFESIONAL.Connection:=modulo2.ObjConexion;
  DataSource4.DataSet:=PROFESIONAL;
  PROFESIONAL.CommandText:='EXEC SPP_ContribuyentesDerConst 0, '+ txt;
  PROFESIONAL.Open;
  if PROFESIONAL.IsEmpty then BEGIN
   messagedlg('Contribuyente Inexistente...' + chr(13) + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
   PC1.ActivePage:=PC1.Pages[5];
   odnip.SetFocus;
  END;
  PROFESIONAL.Connection:=nil;
end;
if nro = 11 then begin
  DataSet1.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet1.CommandText:='SELECT detalle FROM [SeguridadWEB].[dbo].[Fn_ElDomicilioContribuyente] (' + idcontribp + ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   //barriorespo.ItemIndex:=barriorespo.Items.IndexOf(trim(DataSet1.Fields[1].AsString));
   //callerespo.ItemIndex:=callerespo.Items.IndexOf(trim(DataSet1.Fields[2].AsString));
   domp.text:= trim(DataSet1.Fields[0].AsString);
   //alturarespo.text:=  trim(DataSet1.Fields[4].AsString);
   //pisorespo.text:=  trim(DataSet1.Fields[5].AsString);
   //deptorespo.text:=  trim(DataSet1.Fields[6].AsString);
   //mzrespo.text:=    trim(DataSet1.Fields[7].AsString);
   //torrerespo.text:=  trim(DataSet1.Fields[11].AsString);
   //casarespo.text:=  trim(DataSet1.Fields[9].AsString);
   //sectorrespo.text:=    trim(DataSet1.Fields[10].AsString);
   //parrespo.text:=   trim(DataSet1.Fields[8].AsString);
   //lrrespo.text:=   trim(DataSet1.Fields[12].AsString);
   //domrespo.setfocus;
  end;
  DataSet1.Close;

  DataSet1.CommandText:='SELECT Id,Barrio,Calle,CodPostal,Altura,Piso,Dpto,Mz,Parcela,Casa,Sector,'  //10
  + 'Torre,LoteRural,Domicilio,CodBar,CodCal FROM dbo.Fn_ContribuyentesDomicilios(' + idcontribp + ')';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   barriop.ItemIndex:=barriop.Items.IndexOf(trim(DataSet1.Fields[1].AsString));
   callep.ItemIndex:=callep.Items.IndexOf(trim(DataSet1.Fields[2].AsString));
   //domrespo.text:= trim(DataSet1.Fields[0].AsString);
   //alturarespo.text:=  trim(DataSet1.Fields[4].AsString);
   //pisorespo.text:=  trim(DataSet1.Fields[5].AsString);
   //deptorespo.text:=  trim(DataSet1.Fields[6].AsString);
   //mzrespo.text:=    trim(DataSet1.Fields[7].AsString);
   //torrerespo.text:=  trim(DataSet1.Fields[11].AsString);
   //casarespo.text:=  trim(DataSet1.Fields[9].AsString);
   //sectorrespo.text:=    trim(DataSet1.Fields[10].AsString);
   //parrespo.text:=   trim(DataSet1.Fields[8].AsString);
   //lrrespo.text:=   trim(DataSet1.Fields[12].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;
  //domrespo.setfocus;
end;
if nro = 12 then begin /// traemos el contrib de un expediente
  DataSet1.ConnectionString:=modulo2.Sistema.ConectDerConst;
  DataSet1.CommandText :='SELECT dat.IdExpte,dat.Cuim , dat.Contribuyente ,dat.NroDni ,dat.barrio, dat.calle ,'
  +'dat.Domicilio,dat.Altura ,	dat.Piso,dat.Dpto ,dat.Mz,dat.Torre,dat.Casa ,dat.Sector ,dat.Parcela,'
  +'dat.LoteRural ,dat.idcontrib, prof.nromatricula '
	+ ' FROM FN_ObtenerExpedientesProfesional  ('+ idexpte+ ') dat '
  + ' inner join ExpedienteProfesional prof on prof.Cuim = dat.cuim';
  DataSet1.Open;
  if not DataSet1.isempty then begin
   dnip:= trim(DataSet1.Fields[3].AsString);
   barriop.ItemIndex:=barriop.Items.IndexOf(trim(DataSet1.Fields[4].AsString));
   callep.ItemIndex:=callep.Items.IndexOf(trim(DataSet1.Fields[5].AsString));
   domp.text:= trim(DataSet1.Fields[6].AsString);
   alturap.text:=trim(DataSet1.Fields[7].AsString);
   pisop.text:= trim(DataSet1.Fields[8].AsString);
   deptop.text:=trim(DataSet1.Fields[9].AsString);
   mzp.text:=   trim(DataSet1.Fields[10].AsString);
   torrep.text:=trim(DataSet1.Fields[11].AsString);
   casap.text:= trim(DataSet1.Fields[12].AsString);
   sectorp.text:=trim(DataSet1.Fields[13].AsString);
   parp.text:=  trim(DataSet1.Fields[14].AsString);
   lrp.text:=   trim(DataSet1.Fields[15].AsString);
   matricula.Text:= trim(DataSet1.Fields[17].AsString);
  end;
  DataSet1.Close;
  DataSet1.Connection:=nil;

  // lleno la grilla
  PROFESIONAL.Close;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectSeg;
  PROFESIONAL.Connection:=modulo2.ObjConexion;
  DataSource4.DataSet:=PROFESIONAL;
  PROFESIONAL.CommandText:='EXEC SPP_ContribuyentesDerConst 1,null,' + dnip + ',null';
  PROFESIONAL.Open;
  if PROFESIONAL.IsEmpty then BEGIN
   messagedlg('Contribuyente Inexistente...' + chr(13) + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
   PC1.ActivePage:=PC1.Pages[0];
   odnip.SetFocus;
  END;
  PROFESIONAL.Connection:=nil;
end;
Modulo2.ObjConexion.Close;
Modulo2.ObjConexion.Free;
end;

procedure TAM_ExptesDerConst.DEPTOINMUKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  mzinmu.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.deptopKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  mzp.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.DEPTORESPOKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  mzrespo.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.DESTINOOBRAExit(Sender: TObject);
begin
if trim(DESTINOOBRA.Text) = '' then begin
  messagedlg('Error!! No Existe un DESTINO de OBRA ingresado.',mterror,[mbok],0);
  DESTINOOBRA.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.DESTINOOBRAKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  empresa.SetFocus;
 end
 else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ-/ °',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.DOMIINMUKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  alturainmu.SetFocus;
end
else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ-/ °',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.DOMOBRAKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[5];
  odnip.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789 /º-.:',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.dompKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  alturap.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.DOMRESPOExit(Sender: TObject);
begin
 if (trim(DOMRESPO.Text) = '0')  then begin
  messagedlg('Error!! No Existe ningun Domicilio del Contribuyente ingresado.',mterror,[mbok],0);
  DOMRESPO.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.DOMRESPOKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  tpotit.SetFocus;
  //alturarespo.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.Edit1Exit(Sender: TObject);
begin
 if trim(Edit1.Text) = '' then begin
  messagedlg('Error!! No Existe ningun Expediente ingresado.',mterror,[mbok],0);
  Edit1.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  busexpte.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.Edit2Exit(Sender: TObject);
begin
if trim(Edit2.Text) = '' then begin
//  messagedlg('Error!! No Existe ninguna Letra de Expediente ingresado.',mterror,[mbok],0);
//  Edit2.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.Edit2KeyPress(Sender: TObject; var Key: Char);
begin
if key =#13 then begin
  Key := #0;
  edit1.SetFocus;
end
else begin
  if ( StrScan('ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwyxz',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.ALTURAINMUExit(Sender: TObject);
begin
 if trim(ALTURAINMU.Text) = '' then begin
  messagedlg('Error!! Ingrese una Altura válida, o un 0 (cero) si no posee',mterror,[mbok],0);
  ALTURAINMU.text:='0';
  ALTURAINMU.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.ALTURAINMUKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  pisoinmu.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.alturapExit(Sender: TObject);
begin
 if trim(ALTURAp.Text) = '' then begin
  messagedlg('Error!! Ingrese una Altura válida, o un 0 (cero) si no posee',mterror,[mbok],0);
  ALTURAp.text:='0';
  ALTURAp.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.alturapKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  pisop.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.ALTURARESPOExit(Sender: TObject);
begin
 if trim(ALTURARESPO.Text) = '' then begin
  messagedlg('Error!! Ingrese una Altura válida, o un 0 (cero) si no posee',mterror,[mbok],0);
  ALTURARESPO.text:='0';
  ALTURARESPO.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.ALTURARESPOKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  pisorespo.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.AÑO1Exit(Sender: TObject);
begin
 if trim(año1.Text) = ''  then begin
  messagedlg('Error!! No Existe ningun Año de Expediente ingresado.',mterror,[mbok],0);
  AÑO1.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.AÑO1KeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[2];
  tiponotif.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.AÑOExit(Sender: TObject);
begin
 if (trim(AÑO.Text) = '0') or (trim(año.Text) = '' ) then begin
  messagedlg('Error!! No Existe ningun Año de Expediente ingresado.',mterror,[mbok],0);
  AÑO.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.AÑOKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  l3.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.AÑOVISADOExit(Sender: TObject);
begin
 if trim(AÑOVISADO.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  AÑOVISADO.Text := '0' ;
  AÑOVISADO.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.AÑOVISADOKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  revocado.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.BARRIOINMUExit(Sender: TObject);
begin
if BARRIOINMU.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Barrio válido',mterror,[mbok],0);
  BARRIOINMU.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.BARRIOINMUKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  calleinmu.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.barriopExit(Sender: TObject);
begin
if BARRIOp.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Barrio válido',mterror,[mbok],0);
  BARRIOp.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.barriopKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  callep.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.BARRIORESPOChange(Sender: TObject);
begin
//C2.Checked:=true;
end;

procedure TAM_ExptesDerConst.BARRIORESPOExit(Sender: TObject);
begin
if BARRIORESPO.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Barrio válido',mterror,[mbok],0);
  BARRIORESPO.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.BARRIORESPOKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  callerespo.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.BORRAClick(Sender: TObject);
var p,u:integer;
begin
  edit1.Clear;        edit2.Clear;
  letra1.Clear;       letra.Clear;
  nro1.Clear;         nroe.Clear;
  año1.Clear;         año.Clear;

  tiponotif.ItemIndex :=0;        destinoobra.Clear;
  reliqui.Checked :=false;        opublica.Checked :=false;
  empresa.Clear;
  visado.Text:='0';               añovisado.Text:='0';
  revocado.Text:='0';             revocadoaño.Text:='0';

  odni.Clear;                     onomb.Clear;
  emailrespo.Clear;               telefonorespo.Clear;
  barriorespo.ItemIndex :=0;      callerespo.ItemIndex :=0;
  domrespo.Clear;
  alturarespo.Text:='0';          pisorespo.Clear;
  deptorespo.Clear;               mzrespo.Clear;
  torrerespo.Clear;               casarespo.Clear;
  sectorrespo.Clear;              parrespo.Clear;
  lrrespo.Clear;                 tpotit.ItemIndex :=0;

  opartida.Text:='000000';
  n9.Clear;              n2.Clear;
  n3.Clear;              n4.Clear;
  n6.Clear;              n7.Clear;
  n8.Clear;              tpopartida.ItemIndex :=0;
  barrioinmu.ItemIndex :=0;      calleinmu.ItemIndex :=0;
  domiinmu.Clear;
  alturainmu.Text:='0';          pisoinmu.Clear;
  deptoinmu.Clear;               mzinmu.Clear;
  torreinmu.Clear;               casainmu.Clear;
  sectorinmu.Clear;              parinmu.Clear;
  lrinmu.Clear;                  domobra.Clear;

  odnip.Clear;                onombp.Clear;
  matricula.Clear;            emailp.Clear;
  telefonop.Clear;
  barriop.ItemIndex :=0;      callep.ItemIndex :=0;
  for p:=0 to grillaprofes.ColCount-1 do
    for u:=1 to grillaprofes.RowCount-1 do grillaprofes.Cells[p,u]:='';
  grillaprofes.RowCount:=2;
  alturap.Text:='0';          pisop.Clear;
  deptop.Clear;               mzp.Clear;
  torrep.Clear;               casap.Clear;
  sectorp.Clear;              parp.Clear;
  lrp.Clear;                  domp.Clear;

  exisdoc.Text:='0';
  Consultas(1);
  plano.Text:='0';
  exepcion.Checked :=false;
  cargamanual.Checked :=false;
  ordenanza.Checked :=false;
  resolucion.Text:='0';
  finalobra.Text:='0';
  comentarios.Clear;
  PC1.ActivePage:=PC1.Pages[0];
end;

procedure TAM_ExptesDerConst.ODNIKeyPress(Sender: TObject; var Key: Char);
begin
 if key =#13 then begin
  Key := #0;
  ONOMB.SetFocus;
 end
 else begin
  if ( StrScan('0123456789',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.ODNIPKeyPress(Sender: TObject; var Key: Char);
begin
 if key =#13 then begin
  Key := #0;
  ONOMBp.SetFocus;
 end
 else begin
  if ( StrScan('0123456789',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.ONOMBKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  buscar.SetFocus;
 end
 else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ ',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.ONOMBPKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  buscarprof.SetFocus;
 end
 else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ ',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.BuscarClick(Sender: TObject);
begin
 if (trim(odni.Text)= '') and (trim(onomb.Text)='') then begin
  messagedlg('Error!! No Existe ningun Dato del Titular a Buscar.',mterror,[mbok],0);
  odni.SetFocus;
 end
 else Consultas(7); //
end;

procedure TAM_ExptesDerConst.L3Exit(Sender: TObject);
var fec:Tdatetime;
begin
try
 fec:=strtodate(trim(L3.Text));
except
  messagedlg('La fecha es Erronea!!',mterror,[mbok],0);
  L3.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.L3KeyPress(Sender: TObject; var Key: Char);
begin
begin
if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[2];
  tiponotif.SetFocus;
  //Letra1.SetFocus;
end;
end;
end;

procedure TAM_ExptesDerConst.LETRA1Exit(Sender: TObject);
begin
 if trim(LETRA1.Text) = '' then begin
  messagedlg('Error!! No Existe ninguna Letra de Antecedente Expediente ingresado.',mterror,[mbok],0);
  LETRA1.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.LETRA1KeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  NRO1.SetFocus;
 end
 else begin
  if ( StrScan('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.LETRAChange(Sender: TObject);
begin
 C3.Checked:=true;
end;

procedure TAM_ExptesDerConst.LETRAExit(Sender: TObject);
begin
 if trim(LETRA.Text) = '' then begin
  messagedlg('Error!! No Existe ninguna Letra de Expediente ingresado.',mterror,[mbok],0);
  LETRA.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.LETRAKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  NROE.SetFocus;
 end
 else begin
  if ( StrScan('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnñopqrstuvwyxz',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.LRINMUKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  GUARDA.SetFocus;
  //PC1.ActivePage:=PC1.Pages[5];
  //odnip.SetFocus;
  //domobra.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.lrpKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  agrega.SetFocus;
  //PC1.ActivePage:=PC1.Pages[6];
  //exisdoc.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.LRRESPOKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  TPOTIT.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.matriculaKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  emailp.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.MZINMUKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  torreinmu.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.mzpKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  torrep.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.MZRESPOKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  torrerespo.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.SALEClick(Sender: TObject);
begin
Close;
end;

procedure TAM_ExptesDerConst.SECTORExit(Sender: TObject);
begin
if SECTOR.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una Letra de Expediente válido',mterror,[mbok],0);
  SECTOR.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.SECTORINMUKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  parinmu.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.SECTORKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  edit1.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.sectorpKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  parp.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.SECTORRESPOKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  parrespo.SetFocus;
 end
 else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.telefonopChange(Sender: TObject);
begin
//C2.Checked:=true;
end;

procedure TAM_ExptesDerConst.telefonopKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  barriop.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.telefonorespoChange(Sender: TObject);
begin
//C4.Checked:=true;
end;

procedure TAM_ExptesDerConst.telefonorespoKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  barriorespo.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.tiponotifExit(Sender: TObject);
begin
if tiponotif.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un tipo de notificacion válido',mterror,[mbok],0);
  tiponotif.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.tiponotifKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  destinoobra.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.TORREINMUKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  casainmu.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.torrepKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  casap.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.TORRERESPOKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  casarespo.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.TOTALChange(Sender: TObject);
begin
c7.Checked:=true;
end;

procedure TAM_ExptesDerConst.tpopartidaExit(Sender: TObject);
begin
if tpopartida.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Tipo de Partida válido',mterror,[mbok],0);
  tpopartida.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.tpopartidaKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  barrioinmu.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.TPOTITExit(Sender: TObject);
begin
if TPOTIT.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un Tipo de Titular válido',mterror,[mbok],0);
  TPOTIT.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.TPOTITKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[4];
  Opartida.SetFocus;
end
end;

procedure TAM_ExptesDerConst.VISADOExit(Sender: TObject);
begin
if trim(VISADO.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  VISADO.Text := '0' ;
  VISADO.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.VISADOKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  añovisado.SetFocus;
end
end;

procedure TAM_ExptesDerConst.agregaClick(Sender: TObject);
var t:integer;

  function busca():boolean;
  var p:integer;
  begin
  Result:=true;
  if grillaprofes.Cells[0,1] <> '' then begin
    for p:=1 to grillaprofes.RowCount-1 do
      if (grillaprofes.Cells[1,p]=trim(cuimP)) then Result:=false;
  end;
  end;

begin
 if busca then begin
  if (grillaprofes.RowCount > 1) and (trim(grillaprofes.Cells[0,1])<>'') then begin
   t:=grillaprofes.RowCount;
   grillaprofes.RowCount:=grillaprofes.RowCount+1;
  end
  else t:=1;
   grillaprofes.Cells[0,t]:='?';
   grillaprofes.Cells[1,t]:=cuimP;
   grillaprofes.Cells[2,t]:='38';
   grillaprofes.Cells[3,t]:=TComboClase(barriop.Items.Objects[barriop.ItemIndex]).Param1;
   grillaprofes.Cells[4,t]:=TComboClase(callep.Items.Objects[callep.ItemIndex]).Param1;
   grillaprofes.Cells[5,t]:=AnsiReplaceStr(trim(DOMp.Text),chr(39),chr(33));
   grillaprofes.Cells[6,t]:=trim(alturap.text);
   grillaprofes.Cells[7,t]:= trim(pisop.text) ;
   grillaprofes.Cells[8,t]:= trim(deptop.text) ;
   grillaprofes.Cells[9,t]:= trim(mzp.text);
   grillaprofes.Cells[10,t]:=trim(torrep.text);
   grillaprofes.Cells[11,t]:=trim(casap.text) ;
   grillaprofes.Cells[12,t]:=trim(sectorp.text);
   grillaprofes.Cells[13,t]:=trim(parp.text);
   grillaprofes.Cells[14,t]:=trim(lrp.text);
   grillaprofes.Cells[15,t]:='null' ;
   grillaprofes.Cells[16,t]:=trim(emailp.text);
   grillaprofes.Cells[17,t]:=trim(telefonop.text);
   grillaprofes.Cells[18,t]:=trim(Matricula.text) ;
   odnip.Clear;           onombp.Clear;
   matricula.Clear;       emailp.Clear;
   telefonop.Clear;       barriop.ItemIndex :=0;
   callep.ItemIndex :=0;  domp.Clear;
   alturap.Clear;         pisop.Clear;
   deptop.Clear;          mzp.Clear;
   torrep.Clear;          casap.Clear;
   sectorp.Clear;         parp.Clear;
   lrp.Clear;
   PageControl3.ActivePage:=PageControl3.Pages[0];
   odnip.SetFocus;
 end
 else messagedlg('ERROR!! No Puede Existir dos Profesionales Iguales!!',mterror,[mbok],0);
  //matricula.SetFocus;
 PageControl3.ActivePage:=PageControl3.Pages[0];
 odnip.SetFocus;
end;


procedure TAM_ExptesDerConst.JvDBGrid1CellClick(Column: TColumn);
begin
if not EXPEDIENTES.IsEmpty then begin
 if EXPEDIENTES.Fields[0].AsInteger <> 0 then begin
  if trim(EXPEDIENTES.Fields[13].AsString)='True' then begin
   messagedlg('ATENCION!! Este Expediente se encuentra ANULADO, es imposible la modificación de los datos.',mtWarning,[mbok],0);
   PC1.ActivePage:=PC1.Pages[0];
  end
  else begin
   if trim(EXPEDIENTES.Fields[29].AsString)='True' then begin
    messagedlg('ATENCION!! Este Expediente se encuentra CONFIRMADO, es imposible la modificación de los datos.',mtWarning,[mbok],0);
    PC1.ActivePage:=PC1.Pages[0];
   end
   else begin
    selec:=true;
    pc1.Pages[1].TabVisible:=true;
    pc1.Pages[2].TabVisible:=true;
    pc1.Pages[3].TabVisible:=true;
    pc1.Pages[4].TabVisible:=true;
    pc1.Pages[5].TabVisible:=false;
    pc1.Pages[6].TabVisible:=false;
    idexpte := trim(EXPEDIENTES.Fields[0].AsString);
    elexpte:= trim(EXPEDIENTES.Fields[1].AsString);
    letra.Text := trim(EXPEDIENTES.Fields[15].AsString);
    nroe.Text :=  trim(EXPEDIENTES.Fields[17].AsString);
    año.Text :=  trim(EXPEDIENTES.Fields[16].AsString);
    letra1.Text := trim(EXPEDIENTES.Fields[18].AsString);
    nro1.Text :=  trim(EXPEDIENTES.Fields[20].AsString);
    año1.Text :=  trim(EXPEDIENTES.Fields[19].AsString);
    tiponotif.ItemIndex:=tiponotif.Items.IndexOf(trim(EXPEDIENTES.Fields[3].AsString));
    destinoobra.Text :=  trim(EXPEDIENTES.Fields[21].AsString);
    empresa.Text :=  trim(EXPEDIENTES.Fields[9].AsString);
    reliqui.Checked:=EXPEDIENTES.Fields[4].AsBoolean;
    opublica.Checked:=EXPEDIENTES.Fields[8].AsBoolean;
    visado.Text := trim(EXPEDIENTES.Fields[23].AsString);
    añovisado.Text := trim(EXPEDIENTES.Fields[6].AsString);
    revocado.text := trim(EXPEDIENTES.Fields[5].AsString);
    revocadoaño.text := trim(EXPEDIENTES.Fields[22].AsString);
    opartida.Text  := trim(EXPEDIENTES.Fields[24].AsString);
    plano.Text  := trim(EXPEDIENTES.Fields[10].AsString);
    resolucion.Text := trim(EXPEDIENTES.Fields[11].AsString);
    finalobra.Text  := trim(EXPEDIENTES.Fields[12].AsString);
    exisdoc.Text := trim(EXPEDIENTES.Fields[25].AsString);
    idexptecont := trim(EXPEDIENTES.Fields[26].AsString);
    cuimexptecont :=  trim(EXPEDIENTES.Fields[27].AsString);
    nomexptecont :=  trim(EXPEDIENTES.Fields[28].AsString);
    buscarp.Click;
    c6.Checked:=true;
    //consultas(5);
    //consultas(6);
    //consultas(8);
    //consultas(9);
    Consultas(12);
    guarda.Enabled:=true;
    elidentif:=EXPEDIENTES.Fields[0].AsString;
    PC1.ActivePage:=PC1.Pages[1];
    letra.SetFocus;
   end;
  end;
 end;
end;
end;

procedure TAM_ExptesDerConst.JvDBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var t:integer;
begin
 if trim(JvDBGrid1.Fields[29].Text) = 'True' then t:=1 ;
 If trim(JvDBGrid1.Fields[13].Text) = 'True' then t:=2 ;
 case t of
   1:begin
    JvDBGrid1.Canvas.Brush.Color:=clLime;
    JvDBGrid1.Canvas.FillRect(Rect);
    JvDBGrid1.Canvas.Font.Color:= clBlack;
   end;
   2:begin
    JvDBGrid1.Canvas.Brush.Color:=clYellow;
    JvDBGrid1.Canvas.FillRect(Rect);
    JvDBGrid1.Canvas.Font.Color:= clBlack;
   end;
 end;
 JvDBGrid1.DefaultDrawColumnCell(Rect,Datacol,Column,State);
end;

procedure TAM_ExptesDerConst.JvDBGrid2CellClick(Column: TColumn);
begin
if not CONTRIBUYENTE.IsEmpty then begin
 if CONTRIBUYENTE.Fields[0].AsInteger <> 0 then begin
  seleccontrib:=true;
  idcontrib:=CONTRIBUYENTE.Fields[0].AsString;
  cuim:=trim(CONTRIBUYENTE.Fields[1].AsString);
  eltitu:=trim(CONTRIBUYENTE.Fields[2].AsString);
  Consultas(0);
  emailrespo.Text := trim(CONTRIBUYENTE.Fields[16].AsString);
  telefonorespo.Text :=  trim(CONTRIBUYENTE.Fields[17].AsString);
  emailRESPO.SetFocus;
  C4.Checked:=true;
 end
 else begin
   messagedlg('No se puede editar el contribuyente DESCONOCIDO',mterror,[mbok],0);
   PC1.ActivePage:=PC1.Pages[0];
   PC1.Pages[1].Enabled:=false;
 end;
end;
end;

procedure TAM_ExptesDerConst.JvDBGrid3CellClick(Column: TColumn);
begin
if not PROFESIONAL.IsEmpty then begin
 if PROFESIONAL.Fields[0].AsInteger <> 0 then begin
  idcontribP:=PROFESIONAL.Fields[0].AsString;
  cuimP:=trim(PROFESIONAL.Fields[1].AsString);
  eltituP:=trim(PROFESIONAL.Fields[2].AsString);
  Consultas(11);
  C2.Checked:=true;
  emailp.Text := trim(PROFESIONAL.Fields[7].AsString);
  telefonop.Text :=  trim(PROFESIONAL.Fields[9].AsString);
  PageControl1.ActivePage:=PageControl1.Pages[0];
  matricula.SetFocus;
 end
 else begin
   messagedlg('No se puede editar el contribuyente DESCONOCIDO',mterror,[mbok],0);
   odnip.SetFocus;
 end;
end;
end;

procedure TAM_ExptesDerConst.grilla2DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
with TJvStringGrid(Sender) do begin
  if not (gdFixed in State) then begin
    if (gdRowSelected in State) then
      Canvas.Brush.Color:=clWindow
    else
      Canvas.Brush.Color:=clInfoBk; //$00E6C4CD;
  end;
  Canvas.TextRect( Rect, Rect.Left+2, Rect.Top+2, Cells[acol, arow]);
end;
end;

procedure TAM_ExptesDerConst.grilla2ExitCell(Sender: TJvStringGrid;
  AColumn, ARow: Integer; const EditText: string);
var r1,r2,eltotal:real;
    ps:boolean;
    t,y:integer;
begin
grilla2.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goEditing];
ps:=true;
eltotal:=0;
sumdoc:=0;
 if AColumn <= 4 then begin
  try
    r1:=StrToFloat(Grilla2.Cells[AColumn,ARow]);
  except
   try
    r1:=StrToFloat(AnsiReplaceStr(Grilla2.Cells[AColumn,ARow],'.',','));
   except
    ps:=false;
    messagedlg('El valor ingresado en una fila no es numérico',mtError,[mbok],0);
   end;
  end;
  if ps then begin
    sumdoc:= sumdoc + StrToFloat(trim(exisdoc.Text))
    + StrToFloat(Grilla2.Cells[2,1])
    + StrToFloat(Grilla2.Cells[2,2])
    + StrToFloat(Grilla2.Cells[2,3]) ;

    if sumdoc > 100 then begin
     Grilla2.Cells[4,1]:='0.006';
     Grilla2.Cells[4,2]:='0.003';
    end
    else begin
     Grilla2.Cells[4,1]:='0.002';
     Grilla2.Cells[4,2]:='0.001';
    end;
    //Aca hay que calcular el total
    //if StrToFloat(Grilla2.Cells[2,1]) > 100 then Grilla2.Cells[4,1]:='0.006'
    //else  Grilla2.Cells[4,1]:='0.002';
    //if StrToFloat(Grilla2.Cells[2,2]) < 100 then Grilla2.Cells[4,2]:='0.001'
    //else  Grilla2.Cells[4,2]:='0.003';
    try
     r1:=StrToFloat(Grilla2.Cells[3,ARow]);
     r2:=StrToFloat(Grilla2.Cells[4,ARow]);
    except
     try
      r1:=StrToFloat(AnsiReplaceStr(Grilla2.Cells[3,ARow],'.',','));
      r2:=StrToFloat(AnsiReplaceStr(Grilla2.Cells[4,ARow],'.',','));
     except
     end;
    end;
    grilla2.Cells[5,ARow]:=FloatToStr(r1 * r2);
    for t:=1 to grilla2.RowCount-1 do begin
     eltotal := eltotal + strtofloat(Grilla2.Cells[5,t]);
    end;
    total.Text := floattostr(eltotal);
   end;
 end;
end;

procedure TAM_ExptesDerConst.grilla2KeyPress(Sender: TObject;
  var Key: Char);
begin
if (key = #13) or (key = #9) then begin
  if (Grilla2.Row = Grilla2.RowCount-1) and (Grilla2.Col = 4) then begin
    //
  end
  else begin
   Grilla2.Col:=Grilla2.Col + 1;
   if (Grilla2.Col = 4) and not(Pos4) then begin
    Grilla2.Col:=2;
    if not (Grilla2.Row = Grilla2.RowCount-1) then Grilla2.Row:=Grilla2.Row + 1;
   end;
   if (Grilla2.Col = 5) then begin
    Grilla2.Col:=2;
    if not (Grilla2.Row = Grilla2.RowCount-1) then Grilla2.Row:=Grilla2.Row + 1;
   end;
  end;
end
else begin
  if ( StrScan('0123456789.',Key) <> nil )
  or ( Key = Char(VK_BACK) ) then begin
    if (Grilla2.Col = 4) and not Pos4 then Key := #0
    else if (Grilla2.Col = 5) then Key := #0
    else if (Grilla2.Col < 4) then Key:=UpCase(Key)
  end
  else
    Key := #0;
end;
end;

procedure TAM_ExptesDerConst.PARINMUKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  lrinmu.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.parpKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  lrp.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.PARRESPOKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  lrrespo.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.PISOINMUKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  deptoinmu.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.pisopKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  deptop.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.PISORESPOKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  deptorespo.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.PLANOExit(Sender: TObject);
begin
if trim(PLANO.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  PLANO.Text:= '0';
  PLANO.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.PLANOKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  resolucion.SetFocus;
 end
 else begin
  if ( StrScan('0123456789',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.OPARTIDAChange(Sender: TObject);
begin
c5.Checked:=true;
end;

procedure TAM_ExptesDerConst.OPARTIDAEnter(Sender: TObject);
begin
BuscarP.Enabled:=true;
end;

procedure TAM_ExptesDerConst.OPARTIDAExit(Sender: TObject);
begin
 if (trim(OPARTIDA.Text) = '')  or (trim(OPARTIDA.Text) = '0')
 or (trim(OPARTIDA.Text) = '00') or (trim(OPARTIDA.Text) = '000')  then begin
  messagedlg('Error!! Ingrese una Partida válida..',mterror,[mbok],0);
  OPARTIDA.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.BuscarPClick(Sender: TObject);
var t,j,p:integer;
r,bloqueo:boolean;
txt:string;
TxtW:widestring;

begin
  N2.text:='...';
  N3.text:='...';N4.text:='...';N6.text:='...';
  N7.text:='...';N8.text:='...';N9.text:='...';

  Barra1.Panels[1].Text:='Buscando Datos del Inmueble. Por Favor Espere!!';
  DataSet1.ConnectionString:=modulo2.Sistema.ConectUsu;
  DataSet1.CommandText:='SELECT PARTIDA,rtrim(ltrim(NOMRAZ)) + '' ('' + rtrim(ltrim(tipotitular)) + '')'' ,'
  + 'BARRIO,CALLE,PUERTA,RESPOPAG,' //5
  + 'FECHA_ALTA,CUOTA_TASA,CUOTA_INMO1,CUOTA_INMO2,CUOTA_INMO3,CUOTA_INMO4,CUOTA_INMO5,'  //12
  + 'VAL_TOTAL,NOMENC,NRO_DOC,CUIM,'//16
  + 'MANZANA,SECTOR,PARCELA,TORRE,PISO,DEPTO,CASA,sup_mens,'  //24
  + 'anio,fecha_baja,fecha_adj,Habilitada FROM DatosPartidasActualesBis (' + trim(OPARTIDA.Text) + ') order by nomraz;';
  DataSet1.Open;
  Barra1.Panels[1].Text:='...';
  if DataSet1.RecordCount = 0 then begin
    messagedlg('No Existe el N° de Partida Ingresado',mterror,[mbok],0);
    DataSet1.Close;
    OPARTIDA.SetFocus;
  end
  else begin
   if DataSet1.Fields[28].AsBoolean then begin
    DataSet1.RecNo:=DataSet1.RecordCount;
    N9.Text:= trim(DataSet1.Fields[1].AsString);
    n2.Text :=trim(DataSet1.Fields[0].AsString);
    n3.Text :=trim(DataSet1.Fields[24].AsString);
    n4.Text :=trim(DataSet1.Fields[14].AsString);
    n6.Text :=trim(DataSet1.Fields[6].AsString);
    n7.Text :=trim(DataSet1.Fields[26].AsString);
    n8.Text :=trim(DataSet1.Fields[27].AsString);
   if not selec then tpopartida.setfocus; ///barrioinmu.SetFocus;
   end
   else begin
    messagedlg('Partida NO HABILITADA O NO PUBLICA!!... Posible Espacio VERDE...',mterror,[mbok],0);
    r:=false;
   end;
  end;
  DataSet1.Close;
  Barra1.Panels[1].Text:='Datos Encontrados y en Grilla!!';
end;

procedure TAM_ExptesDerConst.BUSCARPROFClick(Sender: TObject);
begin
Consultas(10); //
end;

procedure TAM_ExptesDerConst.BusExpteClick(Sender: TObject);
begin
 if (edit2.Text) = '' then begin
   messagedlg('Error!! Ingrese una Letra de Expediente válido',mterror,[mbok],0);
   edit2.SetFocus;
 end else Consultas(4);
end;

procedure TAM_ExptesDerConst.C6Click(Sender: TObject);
begin
 if c6.Checked then guarda.Caption:='MODIFICAR'
  else guarda.Caption:='GUARDAR'
end;

procedure TAM_ExptesDerConst.CALLEINMUExit(Sender: TObject);
begin
if CALLEINMU.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una Calle válida',mterror,[mbok],0);
  CALLEINMU.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.CALLEINMUKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  domiinmu.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.callepExit(Sender: TObject);
begin
if callep.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una Calle válida',mterror,[mbok],0);
  callep.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.callepKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  domp.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.CALLERESPOExit(Sender: TObject);
begin
if CALLERESPO.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione una Calle válida',mterror,[mbok],0);
  CALLERESPO.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.CALLERESPOKeyPress(Sender: TObject;
  var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  domrespo.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.emailpChange(Sender: TObject);
begin
//C2.Checked:=true;
end;

procedure TAM_ExptesDerConst.emailpKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  telefonop.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.emailRESPOChange(Sender: TObject);
begin
//C4.Checked:=true;
end;

procedure TAM_ExptesDerConst.emailRESPOKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  telefonorespo.SetFocus;
 end;
end;

procedure TAM_ExptesDerConst.empresaKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[3];
  odni.SetFocus;
  //visado.SetFocus;
end
else begin
  if ( StrScan('0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ-/ °',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.EXEPCIONKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  if cargamanual.Enabled then cargamanual.SetFocus
    else ordenanza.SetFocus;
end
end;

procedure TAM_ExptesDerConst.ExisDocChange(Sender: TObject);
begin
 sumdoc:= sumdoc + StrToFloat(trim(exisdoc.Text))
 + StrToFloat(Grilla2.Cells[2,1])
 + StrToFloat(Grilla2.Cells[2,2])
 + StrToFloat(Grilla2.Cells[2,3]) ;
end;

procedure TAM_ExptesDerConst.ExisDocExit(Sender: TObject);
begin
if trim(ExisDoc.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  ExisDoc.Text := '0';
  ExisDoc.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.NRO1Exit(Sender: TObject);
begin
if (trim(nro1.Text) = '' ) then begin
  messagedlg('Error!! No Existe ningun Número de Expediente ingresado.',mterror,[mbok],0);
  NRO1.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.NRO1KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  AÑO1.SetFocus;
end
else begin
  if ( StrScan('1234567890',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TAM_ExptesDerConst.NROEExit(Sender: TObject);
begin
if (trim(NROE.Text) = '0') or (trim(nroE.Text) = '' ) then begin
  messagedlg('Error!! No Existe ningun Número de Expediente ingresado.',mterror,[mbok],0);
  NROE.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.NROEKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  AÑO.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or   ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.NUEVOClick(Sender: TObject);
 var gf:integer;
begin
 gf:=messagedlg('Expediente de Der. de Construccion Inexistente...' + chr(13)
   + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
 if gf = 1 then begin
  guarda.Enabled:=true;
  pc1.Pages[1].TabVisible:=true;
  pc1.Pages[2].TabVisible:=true;
  pc1.Pages[3].TabVisible:=true;
  pc1.Pages[4].TabVisible:=true;
  pc1.Pages[5].TabVisible:=false;
  pc1.Pages[6].TabVisible:=false;
  PC1.ActivePage:=PC1.Pages[1];
  letra.Text := TRIM(edit2.Text);//trim(sector.Text);
  nroE.Text:= trim(edit1.Text);
  letra.SetFocus;
 end
 else begin
  PC1.ActivePage:=PC1.Pages[0];
  edit1.SetFocus;
  nuevo.Enabled:=false;
 end;
end;

procedure TAM_ExptesDerConst.guardaClick(Sender: TObject);
var gf:integer;
    veri:boolean;
    Cnx2:TADOConnection;
    tx,txt,ELIdexpte:string;
    txtm:widestring;
    comando:widestring;
begin
if not c6.Checked then begin   /// Alta de Datos
 gf:=messagedlg('Confirma GUARDAR los datos de este Expediente..?',mtWarning,[mbok,mbCancel],0);
 if gf = 1 then begin
  if Verificacion(1) then begin
   try
    Cnx2:=TADOConnection.Create(nil);
    Cnx2.LoginPrompt:=false;
    Cnx2.ConnectionString:=modulo2.Sistema.ConectDerConst;
    DataSet1.Connection:=Cnx2;

    Com1:=TADOCommand.Create(nil);
    Com1.Connection:=Cnx2;
    tx:=tx + chr(39) + trim(letra.Text) + StrFill(trim(nroe.Text),8,'0','I') + '/'+ trim(año.Text) + chr(39)
    + ' , ' + chr(39) + trim(letra1.Text) +  StrFill(trim(nro1.Text),8,'0','I') + '/'+ trim(año1.Text) + chr(39)
    + ' , ' + TComboClase(tiponotif.Items.Objects[tiponotif.ItemIndex]).Param1;
    if reliqui.Checked  then tx:= tx + ' ,  1  ,'
     else tx:= tx + ' , 0  ,';
    tx := tx + trim(revocado.Text)
    + ' , ' + trim(revocadoaño.Text)
    + ' , ' + trim(visado.Text)
    + ' , ' + trim(añovisado.Text)
     + ' , null ';
    if opublica.Checked  then tx:= tx + ' , 1 , '
     else tx:= tx + ' , 0 , ' ;
    tx := tx + chr(39)+ AnsiReplaceStr(trim(EMPRESA.Text),chr(39),chr(33)) + chr(39)
    + ' , ' + chr(39) + AnsiReplaceStr(trim(DESTINOOBRA.Text),chr(39),chr(33)) + chr(39)
    + ' , ' + trim(plano.Text)
    + ' , ' + trim(resolucion.Text)
    + ' , ' + trim(finalobra.Text)
    + ' , ' + trim(exisdoc.Text)
    + ' , ' + modulo2.Usuario.UsuNro
    + ' , ' + chr(39)+ trim(l3.Text) + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);

    DataSet1.CommandText:='SPA_Expediente '  + tx ;
    DataSet1.Open;
    Idexpte:=DataSet1.Fields[0].AsString;
    DataSet1.Close;
    DataSet1.Connection:=nil;

    Com1.CommandText:='SPA_ExpedienteContribuyente '  + Idexpte
    + ' , ' + chr(39) + cuim + chr(39)
    + ' , 38 , ' + TComboClase(barriorespo.Items.Objects[barriorespo.ItemIndex]).Param1
    + ' , ' + TComboClase(callerespo.Items.Objects[callerespo.ItemIndex]).Param1
    + ' , ' + chr(39) + AnsiReplaceStr(trim(DOMRESPO.Text),chr(39),chr(33))+ chr(39)
    + ' , ' + trim(alturarespo.text)
    + ' , ' + chr(39) + trim(pisorespo.text) + chr(39)
    + ' , ' + chr(39) + trim(deptorespo.text) + chr(39)
    + ' , ' + chr(39) + trim(mzrespo.text) + chr(39)
    + ' , ' + chr(39) + trim(torrerespo.text) + chr(39)
    + ' , ' + chr(39) + trim(casarespo.text) + chr(39)
    + ' , ' + chr(39) + trim(sectorrespo.text) + chr(39)
    + ' , ' + chr(39) + trim(parrespo.text) + chr(39)
    + ' , ' + chr(39) + trim(lrrespo.text) + chr(39)
    + ' , null , ' + chr(39) + trim(emailrespo.text) + chr(39)
    + ' , ' + chr(39) + trim(telefonorespo.text) + chr(39)
    + ' , ' + chr(39) + trim(tpotit.text) + chr(39)
    + ' , ' + modulo2.Usuario.UsuNro
    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
    Com1.Execute;

    for gf:=1 to grillaprofes.RowCount-1 do begin
     if grillaprofes.Cells[0,gf] = '?' then begin
      Com1.CommandText:='SPA_ExpedienteProfesional ' + Idexpte
      + ' , ' + chr(39) + trim(grillaprofes.Cells[1,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[2,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[3,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[4,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[5,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[6,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[7,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[8,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[9,gf])  + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[10,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[11,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[12,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[13,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[14,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[15,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[16,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[17,gf]) + chr(39)
      + ' , ' + chr(39) + trim(grillaprofes.Cells[18,gf]) + chr(39)
      + ' , ' + modulo2.Usuario.UsuNro
      + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
      + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
      Com1.Execute;
     end;
    end;

    Com1.CommandText:='SPA_ExpedienteInmueble '  + Idexpte
    + ' , ' + trim(opartida.text)
    + ' , ' + TComboClase(barrioinmu.Items.Objects[barrioinmu.ItemIndex]).Param1
    + ' , ' + TComboClase(calleinmu.Items.Objects[calleinmu.ItemIndex]).Param1
    + ' , ' + chr(39) + AnsiReplaceStr(trim(domiinmu.Text),chr(39),chr(33)) + chr(39)
    + ' , ' + chr(39) + AnsiReplaceStr(trim(domobra.Text),chr(39),chr(33)) + chr(39)
    + ' , ' + trim(alturainmu.text)
    + ' , ' + chr(39) + trim(pisoinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(deptoinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(mzinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(torreinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(casainmu.text) + chr(39)
    + ' , ' + chr(39) + trim(sectorinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(parinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(lrinmu.text) + chr(39)
    + ' , ' + chr(39) + trim(tpopartida.text) + chr(39)
    + ' , ' + modulo2.Usuario.UsuNro
    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
    Com1.Execute;

    if exepcion.Checked  then txt:= txt + ' , 1 '
     else  txt:= txt + ' , 0 ';
    if ordenanza.Checked then txt:= txt + ' , 1 '
     else  txt:= txt + ' , 0 ';
    for gf:=1 to grilla2.RowCount-1 do begin
     Com1.CommandText:='SPA_ExpedienteCalculo ' + Idexpte
     + ' , ' + trim(grilla2.Cells[0,gf])
     + ' , ' + trim(grilla2.Cells[2,gf])
     + ' , ' + trim(grilla2.Cells[3,gf])
     + ' , ' + trim(grilla2.Cells[4,gf])
     + ' , ' + trim(grilla2.Cells[5,gf])  + txt
     + ' , ' + modulo2.Usuario.UsuNro
     + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
     + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
     Com1.Execute;
    end;

    Com1.CommandText:='SPA_ExpedienteComentario ' + Idexpte
    + ' , ' + chr(39) + AnsiReplaceStr(trim(comentarios.Text),chr(39),chr(33)) + chr(39)
    + ' , ' + modulo2.Usuario.UsuNro
    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
    Com1.Execute;

    Com1.CommandText:='SPA_ExpedienteEstado ' + Idexpte
    + ' , 1  , null , ' + modulo2.Usuario.UsuNro
    + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
    + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
    Com1.Execute;

    messagedlg('Datos Guardados del Expediente Correctamente!!',mtConfirmation,[mbok],0);
    PC1.ActivePage:=PC1.Pages[0];
    borra.Click;
   except
    messagedlg('ERROR al Guardar los Datos',mtError,[mbok],0);
   end;
   FreeAndNil(DataSet1);
   FreeAndNil(Com1);
   Cnx2.Close;
  end
  else begin
  //
  end;
 end;
end
else begin  /// Modificacion de datos
 gf:=messagedlg('Confirma Modificar los datos de este Expediente..?',mtWarning,[mbok,mbCancel],0);
 if gf = 1 then begin
  try
   Cnx2:=TADOConnection.Create(nil);
   Cnx2.LoginPrompt:=false;
   Cnx2.ConnectionString:=modulo2.Sistema.ConectDerConst;
   Com1:=TADOCommand.Create(nil);
   Com1.Connection:=Cnx2;
   if Verificacion(1) then begin
    ELIdexpte:= Idexpte;
    if c3.Checked then begin  // DATOS EXPTES
     comando:='SPM_Expediente ' + ELIdexpte
     + ' , ' + chr(39) + trim(letra.Text) + StrFill(trim(nroe.Text),8,'0','I') + '/' + trim(año.Text) + chr(39)
     + ' , ' + chr(39) + trim(letra1.Text) + StrFill(trim(nro1.Text),8,'0','I') + '/' + trim(año1.Text) + chr(39)
     + ' , ' + TComboClase(tiponotif.Items.Objects[tiponotif.ItemIndex]).Param1;
     if reliqui.Checked  then comando:= comando + ' ,  1  ,'
       else comando:= comando + ' , 0  ,';
     comando := comando + revocado.Text
     + ' , ' + revocadoaño.Text
     + ' , ' + visado.Text
     + ' , ' + añovisado.Text
     + ' , null ';
     if opublica.Checked  then comando:= comando + ' , 1 , '
       else comando:= comando + ' , 0 , ' ;
     comando := comando + chr(39)+ AnsiReplaceStr(trim(EMPRESA.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + chr(39) + AnsiReplaceStr(trim(DESTINOOBRA.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + plano.Text
     + ' , ' + resolucion.Text
     + ' , ' + finalobra.Text
     + ' , ' + modulo2.Usuario.UsuNro
     + ' , ' + exisdoc.Text
     + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
     + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
     Com1.CommandText:=comando;
    end;
    if c4.Checked  then begin  /// titular
     comando:='SPM_ExpedienteContribuyente ' + Idexpte
     + ' , ' + chr(39) + cuim  + chr(39)
     + ' , 38 , ' + TComboClase(barriorespo.Items.Objects[barriorespo.ItemIndex]).Param1
     + ' , ' + TComboClase(callerespo.Items.Objects[callerespo.ItemIndex]).Param1
     + ' , ' + chr(39) + AnsiReplaceStr(trim(domrespo.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + trim(alturarespo.text)
     + ' , ' + trim(pisorespo.text)
     + ' , ' + trim(deptorespo.text)
     + ' , ' + trim(mzrespo.text)
     + ' , ' + trim(torrerespo.text)
     + ' , ' + trim(casarespo.text)
     + ' , ' + trim(sectorrespo.text)
     + ' , ' + trim(parrespo.text)
     + ' , ' + trim(lrrespo.text)
     + ' , nuul , ' + chr(39) + trim(emailrespo.text) + chr(39)
     + ' , ' + chr(39) + trim(telefonorespo.text) + chr(39)
     + ' , ' + trim(tpotit.text)
     + ' , ' + modulo2.Usuario.UsuNro
     + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
     + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
     Com1.CommandText:=comando;
     Com1.Execute;
    end;
    if c5.Checked then begin  // DATOS INMUEBLE
     comando:='SPM_ExpedienteInmueble '  + Idexpte
     + ' , ' + trim(opartida.text)
     + ' , ' + TComboClase(barrioinmu.Items.Objects[barrioinmu.ItemIndex]).Param1
     + ' , ' + TComboClase(calleinmu.Items.Objects[calleinmu.ItemIndex]).Param1
     + ' , ' + chr(39) + AnsiReplaceStr(trim(domiinmu.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + chr(39) + AnsiReplaceStr(trim(domobra.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + trim(alturainmu.text)
     + ' , ' + trim(pisoinmu.text)
     + ' , ' + trim(deptoinmu.text)
     + ' , ' + trim(mzinmu.text)
     + ' , ' + trim(torreinmu.text)
     + ' , ' + trim(casainmu.text)
     + ' , ' + trim(sectorinmu.text)
     + ' , ' + trim(parinmu.text)
     + ' , ' + trim(lrinmu.text)
     + ' , ' + chr(39) + trim(TPOPARTIDA.text) + chr(39)
     + ' , ' + modulo2.Usuario.UsuNro
     + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
     + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
     Com1.CommandText:=comando;
     Com1.Execute;
    end;
    if c2.Checked  then begin  /// profesional
     comando:='SPM_ExpedienteProfesional ' + Idexpte
     + ' , ' + chr(39) + cuimp + chr(39)
     + ' , 38 , ' + TComboClase(barriop.Items.Objects[barriop.ItemIndex]).Param1
     + ' , ' + TComboClase(callep.Items.Objects[callep.ItemIndex]).Param1
     + ' , ' + chr(39) + AnsiReplaceStr(trim(domp.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + trim(alturap.text)
     + ' , ' + trim(pisop.text)
     + ' , ' + trim(deptop.text)
     + ' , ' + trim(mzp.text)
     + ' , ' + trim(torrep.text)
     + ' , ' + trim(casap.text)
     + ' , ' + trim(sectorp.text)
     + ' , ' + trim(parp.text)
     + ' , ' + trim(lrp.text)
     + ' , nuul , ' + chr(39) + trim(emailp.text) + chr(39)
     + ' , ' + chr(39) + trim(telefonop.text) + chr(39)
     + ' , ' + chr(39) + trim(matricula.text) + chr(39)
     + ' , ' + modulo2.Usuario.UsuNro
     + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
     + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
     Com1.CommandText:=comando;
     Com1.Execute;
    end;
    if c7.Checked then begin  // DATOS IMPORTES
     for gf:=1 to grilla2.RowCount-1 do begin
      comando:='SPM_ExpedienteCalculo '
      + trim(grilla2.Cells[7,gf])
      + ' , ' + trim(grilla2.Cells[6,gf])
      + ' , ' + trim(grilla2.Cells[0,gf])
      + ' , ' + trim(grilla2.Cells[2,gf])
      + ' , ' + trim(grilla2.Cells[3,gf])
      + ' , ' + trim(grilla2.Cells[4,gf])
      + ' , ' + trim(grilla2.Cells[5,gf])
      + ' , ' + modulo2.Usuario.UsuNro
      + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
      + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
      Com1.CommandText:=comando;
      Com1.Execute;
     end;
    end;
    if c8.Checked then begin  // DATOS COMENTARIOS
     Com1.CommandText:='SPM_ExpedienteComentario ' + Idexpte
     + ' , ' + chr(39) + AnsiReplaceStr(trim(comentarios.Text),chr(39),chr(33)) + chr(39)
     + ' , ' + modulo2.Usuario.UsuNro
     + ' , ' + chr(39) + modulo2.Sistema.MyHost + chr(39)
     + ' , ' + chr(39) + modulo2.Sistema.MyIP + chr(39);
     Com1.Execute;
    end;
    messagedlg('Datos Modificados Correctamente!!',mtConfirmation,[mbok],0);
    borra.Click;
    FreeAndNil(DataSet1);
    FreeAndNil(Com1);
    Cnx2.Close;
   end
   else begin
    // pc1.Pages[6].TabVisible:=true;
    // PC1.ActivePage:=PC1.Pages[6];
   end;
  except
   messagedlg('ERROR al Modificar los Datos',mtError,[mbok],0);
  end;
   FreeAndNil(DataSet1);
   FreeAndNil(Com1);
   Cnx2.Close;
 end;
end;
end;

procedure TAM_ExptesDerConst.OPARTIDAKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  if buscarp.Enabled then buscarp.SetFocus;
 end
 else begin
  if ( StrScan('0123456789',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.ORDENANZAKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  plano.SetFocus;
end
end;

procedure TAM_ExptesDerConst.RESOLUCIONExit(Sender: TObject);
begin
if trim(RESOLUCION.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  RESOLUCION.Text:= '0';
  RESOLUCION.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.RESOLUCIONKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  finalobra.SetFocus;
 end
 else begin
  if ( StrScan('0123456789',Key) <> nil ) or ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.REVOCADOAÑOExit(Sender: TObject);
begin
if trim(REVOCADOAÑO.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  REVOCADOAÑO.Text :='0';
  REVOCADOAÑO.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.REVOCADOAÑOKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  PC1.ActivePage:=PC1.Pages[3];
  odni.SetFocus;
 end
 else begin
  if ( StrScan('1234567890',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.REVOCADOExit(Sender: TObject);
begin
if trim(REVOCADO.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  REVOCADO.Text :='0';
  REVOCADO.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.REVOCADOKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  revocadoaño.SetFocus;
 end
end;


procedure TAM_ExptesDerConst.FINALOBRAExit(Sender: TObject);
begin
if trim(FINALOBRA.Text) = '' then begin
  messagedlg('Error!! Este campo no puede ser vacio.',mterror,[mbok],0);
  PLANO.Text:= '0';
  FINALOBRA.SetFocus;
end;
end;

procedure TAM_ExptesDerConst.FINALOBRAKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then begin
  Key := #0;
  comentarios.SetFocus;
 end
 else begin
  if ( StrScan('0123456789',Key) <> nil ) or  ( Key = Char(VK_BACK) ) then
   else Key := #0;
 end;
end;

procedure TAM_ExptesDerConst.grilla2SelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  if ((ACol = 4) and not(Pos4)) or (ACol = 5) then grilla2.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine]
  else grilla2.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goEditing];
  if (grilla2.Cells[0,4] = '4') and (Acol = 2) then grilla2.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine]
  else grilla2.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goEditing];
end;

end.
