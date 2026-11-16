unit LiquidarDerecho;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, JvExButtons,
  JvBitBtn, Vcl.ComCtrls, Vcl.ToolWin,ADODB,Objetos, Vcl.Grids, JvExGrids,
  JvStringGrid, JvExComCtrls, JvComCtrls, JvComponentBase, JvEmbeddedForms,
  Vcl.ImgList,System.Math,frxClass, frxDBSet,frxBarcode,StrUtils, Data.DB,
  Vcl.ExtCtrls, Vcl.DBGrids, JvExDBGrids, JvDBGrid, JvgGroupBox, Vcl.Mask,
  Vcl.Menus;

type THackStringGrid = class(TJvStringGrid);
type
  TCalcuDerecho = class(TForm)
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    JvBitBtn1: TJvBitBtn;
    PorNoPresup: TGroupBox;
    SC: TEdit;
    SSC: TEdit;
    GroupBox1: TGroupBox;
    TS: TEdit;
    Label1: TLabel;
    TipoConst: TComboBox;
    Memo1: TMemo;
    Label70: TLabel;
    Label69: TLabel;
    PC1: TJvPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    GroupBox2: TGroupBox;
    GrillaCalculo: TJvStringGrid;
    agrega: TJvBitBtn;
    GrillaEntidad: TJvStringGrid;
    FormularioEmb: TJvEmbeddedFormLink;
    ImageList1: TImageList;
    TabSheet3: TTabSheet;
    GroupBox3: TGroupBox;
    losdatos: TJvBitBtn;
    TotLiqui: TEdit;
    GroupBox4: TGroupBox;
    Label2: TLabel;
    ENTIDAD: TComboBox;
    cargaliqui: TJvBitBtn;
    Label3: TLabel;
    grillaImportes: TJvStringGrid;
    frxReport1: TfrxReport;
    TabSheet4: TTabSheet;
    GroupBox5: TGroupBox;
    impliqui: TJvBitBtn;
    JvgGroupBox2: TJvgGroupBox;
    Label18: TLabel;
    Label19: TLabel;
    BUSCARPROF: TJvBitBtn;
    ODNIP: TEdit;
    ONOMBP: TEdit;
    JvDBGrid3: TJvDBGrid;
    Panel1: TPanel;
    DataSource1: TDataSource;
    Label4: TLabel;
    OPARTIDA: TMaskEdit;
    BuscarP: TButton;
    N9: TEdit;
    Label14: TLabel;
    Label37: TLabel;
    N2: TEdit;
    Label15: TLabel;
    N3: TEdit;
    Label10: TLabel;
    N4: TEdit;
    N8: TEdit;
    Label7: TLabel;
    N7: TEdit;
    Label8: TLabel;
    N6: TEdit;
    Label11: TLabel;
    OCUIT: TEdit;
    PopupMenu1: TPopupMenu;
    BorrarCalculo1: TMenuItem;
    Label5: TLabel;
    Edit1: TEdit;
    Label6: TLabel;
    Label9: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    procedure JvBitBtn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CargaCombos(nro:byte);
    procedure Consultas(nro:byte);
    procedure ControlaSelec(nro:byte);
    procedure TipoConstExit(Sender: TObject);
    procedure TipoConstKeyPress(Sender: TObject; var Key: Char);
    procedure SSCExit(Sender: TObject);
    procedure TSChange(Sender: TObject);
    procedure SCKeyPress(Sender: TObject; var Key: Char);
    procedure SSCKeyPress(Sender: TObject; var Key: Char);
    procedure SCExit(Sender: TObject);
    procedure SCChange(Sender: TObject);
    procedure SSCChange(Sender: TObject);
    procedure GrillaCalculoClick(Sender: TObject);
    procedure ENTIDADExit(Sender: TObject);
    procedure ENTIDADKeyPress(Sender: TObject; var Key: Char);
    procedure agregaClick(Sender: TObject);
    procedure GrillaEntidadClick(Sender: TObject);
    procedure TipoConstChange(Sender: TObject);
    procedure cargaliquiClick(Sender: TObject);
    procedure losdatosClick(Sender: TObject);
    procedure TotLiquiChange(Sender: TObject);
    procedure ImprimeRecibo();
    procedure GuardarRecibo(nro:byte);
    procedure impliquiClick(Sender: TObject);
    procedure ODNIPKeyPress(Sender: TObject; var Key: Char);
    procedure ONOMBPKeyPress(Sender: TObject; var Key: Char);
    procedure BUSCARPROFClick(Sender: TObject);
    procedure JvDBGrid3CellClick(Column: TColumn);
    procedure OPARTIDAEnter(Sender: TObject);
    procedure OPARTIDAExit(Sender: TObject);
    procedure OPARTIDAKeyPress(Sender: TObject; var Key: Char);
    procedure BuscarPClick(Sender: TObject);
    procedure BorrarCalculo1Click(Sender: TObject);
  private
      gh:boolean;
      Recibo:TRecibo;
      Com1:TADOCommand;
      descri:string;
      eltotal:real;
      cants,cantis,cant:integer;
      alicuota,valor,FechaAct,idcontribp,cuimp,eltituP,eldomicilio:string;
      DataSet1,DataSet2,OBRAS,PROFESIONAL:TADODataSet;
    { Private declarations }
  public

    { Public declarations }
  end;

var
  CalcuDerecho: TCalcuDerecho;

implementation

uses Unidad, modulo;

{$R *.dfm}

procedure TCalcuDerecho.FormCreate(Sender: TObject);
begin
 gh:=false;
 DataSet1:=TADODataSet.Create(Self);
 DataSet2:=TADODataSet.Create(Self);
 PROFESIONAL:=TADODataSet.Create(Self);
 Modulo2.ObjConexion:=nil;
end;

procedure TCalcuDerecho.FormActivate(Sender: TObject);
begin
 if not (gh) then begin
  gh:=true;
  FechaAct:=LeftStr(Fec_System(),10);
  CargaCombos(1);
  CargaCombos(2);
  PC1.ActivePage :=PC1.Pages[0];
  grillacalculo.Cells[0,0]:='Id';                 grillacalculo.Cells[1,0]:='Descripcion';
  grillacalculo.Cells[2,0]:='Alicuota';           grillacalculo.Cells[3,0]:='Selecc.';

  GrillaEntidad.Cells[0,0]:='Id';                 GrillaEntidad.Cells[1,0]:='Descripcion';
  GrillaEntidad.Cells[2,0]:='Valor';              GrillaEntidad.Cells[3,0]:='Selecc.';

  GrillaImportes.Cells[0,0]:='Descripcion';       GrillaImportes.Cells[1,0]:='MetrosCuadrados';
  GrillaImportes.Cells[2,0]:='Valor';             GrillaImportes.Cells[3,0]:='Alicuota';
  GrillaImportes.Cells[4,0]:='';                  GrillaImportes.Cells[5,0]:='Importes';
  cargaliqui.Enabled :=false;
  impliqui.Enabled :=false;
  idcontribP:='';
  impliqui.Enabled :=false;
  tipoconst.SetFocus;
 end;
end;

procedure TCalcuDerecho.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(DataSet1);
  FreeAndNil(DataSet2);
  freeAndNil(PROFESIONAL);
  FreeAndNil(Self);
end;

procedure TCalcuDerecho.GrillaCalculoClick(Sender: TObject);
begin
 if (trim(GrillaCalculo.Cells[3,GrillaCalculo.Row])='') then GrillaCalculo.Cells[3,GrillaCalculo.Row]:='x'
 else GrillaCalculo.Cells[3,GrillaCalculo.Row]:='';
 ControlaSelec(2);
 ControlaSelec(4);
 agrega.Enabled :=true;
end;

procedure TCalcuDerecho.agregaClick(Sender: TObject);
var eldato:real;
    a,p,u:integer;
begin
     if (trim(ts.Text)='0') and (TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4') then begin
       messagedlg('ERROR.. El Importe Calculado para la Liquidación NO puede ser Cero(0)!',mtError,[mbok],0);
       sc.SetFocus;
     end
     else begin
      if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 = '4' then begin
       ControlaSelec(2);
       if cant = 0 then begin
         messagedlg('Atencion.. No Existe ningún Importe seleccionado!',mtWarning,[mbok],0);
       end
       else begin
        for p:=0 to grillaimportes.ColCount-1 do
          for u:=1 to grillaimportes.RowCount-1 do grillaimportes.Cells[p,u]:='';
        grillaimportes.RowCount:=2;

       PC1.ActivePage :=PC1.Pages[2];
       ControlaSelec(4);
       if trim(grillaimportes.Cells[0,1])='' then begin
        eldato:=0;
        grillaimportes.Cells[0,1]:=trim(tipoconst.Text);
        grillaimportes.Cells[1,1]:='0';
        grillaimportes.Cells[2,1]:=trim(sc.Text);
        grillaimportes.Cells[3,1]:=alicuota;
        eldato:=strtofloat(grillaimportes.Cells[2,1]) * strtofloat(grillaimportes.Cells[3,1]);
        grillaimportes.Cells[5,1]:=floattostr(RoundTo(eldato,-2));
       end
       else begin
        eldato:=0;
        grillaimportes.RowCount:=grillaimportes.RowCount+1;
        a:=grillaimportes.RowCount-1;
        grillaimportes.Cells[0,a]:=trim(tipoconst.Text);
        grillaimportes.Cells[1,a]:='0';
        grillaimportes.Cells[2,a]:=trim(sc.Text);
        grillaimportes.Cells[3,a]:=alicuota;
        eldato:=strtofloat(grillaimportes.Cells[2,a]) * strtofloat(grillaimportes.Cells[3,a]);
        grillaimportes.Cells[5,a]:=floattostr(RoundTo(eldato,-2));
       end;
       end;
       ControlaSelec(5);
      end
      else begin
       ControlaSelec(4);
       PC1.ActivePage :=PC1.Pages[1];
       entidad.SetFocus;
      end;
     end;
end;

procedure TCalcuDerecho.BorrarCalculo1Click(Sender: TObject);
var p,u:integer;
begin
  for p:=0 to grillaimportes.ColCount-1 do
    for u:=1 to grillaimportes.RowCount-1 do grillaimportes.Cells[p,u]:='';
  grillaimportes.RowCount:=2;
  totliqui.Text :='0';
  PC1.ActivePage:=PC1.Pages[1];
end;

procedure TCalcuDerecho.BuscarPClick(Sender: TObject);
var t,j,p:integer;
r,bloqueo:boolean;
txt:string;
TxtW:widestring;

begin
  N2.text:='...';
  N3.text:='...';N4.text:='...';N6.text:='...';
  N7.text:='...';N8.text:='...';N9.text:='...';

  DataSet1.ConnectionString:=modulo2.Sistema.ConectUsu;
  DataSet1.CommandText:='SELECT PARTIDA,rtrim(ltrim(NOMRAZ)) + '' ('' + rtrim(ltrim(tipotitular)) + '')'' ,'
  + 'BARRIO,CALLE,PUERTA,RESPOPAG,' //5
  + 'FECHA_ALTA,CUOTA_TASA,CUOTA_INMO1,CUOTA_INMO2,CUOTA_INMO3,CUOTA_INMO4,CUOTA_INMO5,'  //12
  + 'VAL_TOTAL,NOMENC,NRO_DOC,CUIM,'//16
  + 'MANZANA,SECTOR,PARCELA,TORRE,PISO,DEPTO,CASA,sup_mens,'  //24
  + 'anio,fecha_baja,fecha_adj,Habilitada FROM DatosPartidasActualesBis (' + trim(OPARTIDA.Text) + ') order by nomraz;';
  DataSet1.Open;
  if DataSet1.RecordCount = 0 then begin
    messagedlg('No Existe el N° de Partida Ingresado',mterror,[mbok],0);
    DataSet1.Close;
    OPARTIDA.SetFocus;
  end
  else begin
    if DataSet1.Fields[28].AsBoolean then begin
     DataSet1.RecNo:=DataSet1.RecordCount;
     N9.Text:= trim(DataSet1.Fields[1].AsString);
     n2.Text:= trim(DataSet1.Fields[0].AsString);
     n3.Text:= trim(DataSet1.Fields[24].AsString);
     n4.Text:= trim(DataSet1.Fields[14].AsString);
     n6.Text:= trim(DataSet1.Fields[6].AsString);
     n7.Text:= trim(DataSet1.Fields[26].AsString);
     n8.Text:= trim(DataSet1.Fields[27].AsString);
     impliqui.Enabled :=true;
     impliqui.SetFocus;
    end
    else begin
      messagedlg('Partida NO HABILITADA O NO PUBLICA!!... Posible Espacio VERDE...',mterror,[mbok],0);
      r:=false;
    end;
  end;
  DataSet1.Close;
end;

procedure TCalcuDerecho.BUSCARPROFClick(Sender: TObject);
begin
 if (trim(odnip.Text)='') and  (trim(onombp.Text)='')  then begin
   messagedlg('Error!! No Hay ningun dato para realizar la busqueda del Contribiyente..',mtWarning,[mbok],0);
   odnip.SetFocus;
 end
 else Consultas(3);
end;

procedure TCalcuDerecho.CargaCombos(nro:byte);
begin
 if nro=1 then Modulo2.CargaCombo(TipoConst,'SELECT Detalle,Id FROM TiposConstrucciones',modulo2.Sistema.ConectDerConst,false);
 if nro=2 then Modulo2.CargaCombo(Entidad,'SELECT Entidades,Id FROM TarifasDerConstEntidades order by 1 ',modulo2.Sistema.ConectDerConst,false);
end;

procedure TCalcuDerecho.ControlaSelec(nro:byte);
var t,y,gf:integer;
begin
cant:=0;
cants:=0;
cantis:=0;
eltotal:=0;
 if nro=1 then begin
  for gf:=1 to grillaentidad.RowCount-1 do begin
    if (grillaentidad.Cells[3,gf] = 'x') then begin
      cants:=cants+1;
    end;
  end;
  if cants > 1 then begin
     messagedlg('Error!! No Se Puede seleccionar mas de un importe..',mtWarning,[mbok],0);
     PC1.ActivePage :=PC1.Pages[1];
     GRILLAENTIDAD.SetFocus;
  end
  ELSE BEGIN
    PC1.ActivePage :=PC1.Pages[2];
  END;
 end;
 if nro = 2 then begin
  for t:=1 to grillacalculo.RowCount-1 do begin
    if (grillacalculo.Cells[3,t] = 'x') then begin
      cant:=cant+1;
    end;
  end;
  if cant > 1 then begin
     messagedlg('Error!! No Se Puede seleccionar mas de un importe..',mtWarning,[mbok],0);
  end;
 end;
 if nro = 3 then begin
  for y:=1 to grillaentidad.RowCount-1 do begin
    if (grillaentidad.Cells[3,y] = 'x') then begin
      valor:=trim(grillaentidad.Cells[2,y]);
    end;
  end;
 end;
 if nro = 4 then begin
  for y:=1 to grillacalculo.RowCount-1 do begin
    if (grillacalculo.Cells[3,y] = 'x') then begin
      alicuota:=trim(grillacalculo.Cells[2,y]);
    end;
  end;
 end;
 if nro = 5 then begin   // sumamos los importes de cada tramite
  for y:=1 to GrillaImportes.RowCount-1 do begin
    //if (grillaentidad.Cells[3,y] = 'x') then begin
      eltotal:=eltotal + strtofloat(GrillaImportes.Cells[5,y]);
    //end;
  end;
  totliqui.Text :=floattostr(eltotal);
 end;
end;

procedure TCalcuDerecho.Consultas(nro:byte);
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
  DataSet1.ConnectionString:=modulo2.Sistema.ConectDerConst;
  DataSet1.CommandText:='SELECT descripcion FROM TiposConstrucciones where (id = '+ TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 +') ';
  DataSet1.Open;
  memo1.Text :=  trim(DataSet1.Fields[0].AsString);
  DataSet1.Close;
  DataSet1.Connection:=nil;
  if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 = '4' then begin
   sc.Text:='0';
   label69.Caption:='Ingrese Presupuesto:';
   ssc.Visible:=false;
   label70.Visible:=false;
   groupbox1.Visible :=false;
   grillacalculo.Enabled :=true;
   grillacalculo.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goRangeSelect,goRowSelect];
  end
  else begin
   label69.Caption:='Supercie Cubierta:';
   grillacalculo.Options:=[goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goRangeSelect];
   ssc.Visible:=true;
   label70.Visible:=true;
   groupbox1.Visible :=true;
   grillacalculo.Enabled :=false;
   sc.SetFocus;
  end;
end;
if nro = 1 then begin
  DataSet2.ConnectionString:=modulo2.Sistema.ConectDerConst;
  if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 = '4' then
     DataSet2.CommandText:='SELECT ID,Descripcion,alicuota,selec FROM SeleccionTipoObra ( '
    + TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 +' ,null)'
  else
  DataSet2.CommandText:='SELECT ID,Descripcion,alicuota,selec FROM SeleccionTipoObra ( '
  + TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 +' , '+ trim(ts.Text) +')';
  DataSet2.Open;
  for t:=1 to DataSet2.RecordCount do begin
   grillacalculo.RowCount:=t+1;
   DataSet2.RecNo:=t;
   grillacalculo.Cells[0,t]:=DataSet2.Fields[0].AsString;
   grillacalculo.Cells[1,t]:=DataSet2.Fields[1].AsString;
   grillacalculo.Cells[2,t]:=DataSet2.Fields[2].AsString;
   grillacalculo.Cells[3,t]:=DataSet2.Fields[3].AsString;
  end;
  DataSet2.Close;
  DataSet2.Connection:=nil;
end;
if nro = 2 then begin
  DataSet2.ConnectionString:=modulo2.Sistema.ConectDerConst;
  DataSet2.CommandText:='SELECT IdTarifa,Descripcion,Tarifa,'''' as Selec  FROM TarifasDerConstEntidadesDetalle  '
  +' where idderconstentidad= ' + TComboClase(Entidad.Items.Objects[Entidad.ItemIndex]).Param1 +' and activo=1 order by 1';
  DataSet2.Open;
  for t:=1 to DataSet2.RecordCount do begin
   GrillaEntidad.RowCount:=t+1;
   DataSet2.RecNo:=t;
   GrillaEntidad.Cells[0,t]:=DataSet2.Fields[0].AsString;
   GrillaEntidad.Cells[1,t]:=DataSet2.Fields[1].AsString;
   GrillaEntidad.Cells[2,t]:=DataSet2.Fields[2].AsString;
   GrillaEntidad.Cells[3,t]:=DataSet2.Fields[3].AsString;
  end;
  DataSet2.Close;
  DataSet2.Connection:=nil;
end;
if nro = 3 then begin // buscamos un contrib
  if trim(ONOMBp.Text) = '' then txt:='NULL'
    else txt:=chr(39) + trim(ONOMBp.Text) + chr(39);
  if trim(ODNIp.Text) = '' then txt:=txt + ' , NULL'
    else txt:=txt + ' , ' + chr(39) + trim(ODNIp.Text) + chr(39);
  if trim(OCUIT.Text) = '' then txt:=txt + ' , NULL'
    else txt:=txt + ' , ' + chr(39) + trim(OCUIT.Text) + chr(39);
  PROFESIONAL.Close;
  Modulo2.ObjConexion.ConnectionString:=modulo2.Sistema.ConectSeg;
  PROFESIONAL.Connection:=modulo2.ObjConexion;
  DataSource1.DataSet :=PROFESIONAL;
  PROFESIONAL.CommandText:='EXEC SPP_ContribuyentesDerConst 0, ' + txt;
  PROFESIONAL.Open;
  if PROFESIONAL.IsEmpty then BEGIN
   messagedlg('Contribuyente Inexistente...' + chr(13) + 'Debe darlo de  Alta?',mtWarning,[mbok,mbCancel],0);
   PC1.ActivePage:=PC1.Pages[3];
   odnip.SetFocus;
  END;
  PROFESIONAL.Connection:=nil;
end;
if nro = 4 then begin
  DataSet2.ConnectionString:=modulo2.Sistema.ConectSeg;
  DataSet2.CommandText:='SELECT detalle FROM Fn_ElDomicilioContribuyente (' + idcontribP + ')';
  DataSet2.Open;
  eldomicilio:=DataSet2.Fields[0].AsString;
  DataSet2.Close;
  DataSet2.Connection:=nil;
end;
end;

procedure TCalcuDerecho.ENTIDADExit(Sender: TObject);
begin
begin
if ENTIDAD.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un tipo de Entidad válida',mterror,[mbok],0);
  ENTIDAD.SetFocus;
end
else Consultas(2);
end;
end;

procedure TCalcuDerecho.ENTIDADKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  grillaentidad.SetFocus;
end;
end;

procedure TCalcuDerecho.JvBitBtn1Click(Sender: TObject);
begin
CLOSE;
end;

procedure TCalcuDerecho.JvDBGrid3CellClick(Column: TColumn);
begin
if not PROFESIONAL.IsEmpty then begin
 if PROFESIONAL.Fields[0].AsInteger <> 0 then begin
  idcontribP:=PROFESIONAL.Fields[0].AsString;
  consultas(4);
  cuimP:=trim(PROFESIONAL.Fields[1].AsString);
  eltituP:=trim(PROFESIONAL.Fields[2].AsString);
  opartida.SetFocus;
 end
 else begin
   messagedlg('No se puede editar el contribuyente DESCONOCIDO',mterror,[mbok],0);
   PC1.ActivePage:=PC1.Pages[3];
 end;
end;
end;

procedure TCalcuDerecho.cargaliquiClick(Sender: TObject);
 var x,a,p,u:integer;
     dato:real;
begin
 if TRIM(TS.Text ) = '0' then begin
   messagedlg('Atencion.. No Existe ningun Importe seleccionado!',mtWarning,[mbok],0);
   cargaliqui.Enabled:=false;
 end
 else begin
  ControlaSelec(1);
  if cants = 1 THEN BEGIN
   /// limpiamos la grilla
   for p:=0 to grillaimportes.ColCount-1 do
    for u:=1 to grillaimportes.RowCount-1 do grillaimportes.Cells[p,u]:='';
    grillaimportes.RowCount:=2;

   ControlaSelec(2);
   ControlaSelec(3);
   ControlaSelec(4);
   if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4' then begin
     if trim(grillaimportes.Cells[0,1])='' then begin
      grillaimportes.Cells[0,1]:=trim(tipoconst.Text) + 'Sup. Cubierta';
      grillaimportes.Cells[1,1]:=trim(sc.Text);
      grillaimportes.Cells[2,1]:=valor;
      grillaimportes.Cells[3,1]:=alicuota;
      dato:=(strtofloat(grillaimportes.Cells[1,1])
       * StrToFloat(grillaimportes.Cells[2,1])) * StrToFloat(grillaimportes.Cells[3,1]);
      grillaimportes.Cells[5,1]:=floattostr(RoundTo(dato,-2));
      dato:=0;
      grillaimportes.RowCount:=grillaimportes.RowCount+1;
      a:=grillaimportes.RowCount-1;
      grillaimportes.Cells[0,a]:=trim(tipoconst.Text)+ 'Sup. SemiCubierta';
      grillaimportes.Cells[1,a]:=trim(ssc.Text);
      grillaimportes.Cells[2,a]:=valor;
      grillaimportes.Cells[3,a]:=alicuota;
      grillaimportes.Cells[4,a]:='0.5';
      dato:=(strtofloat(grillaimportes.Cells[1,a])    * StrToFloat(grillaimportes.Cells[2,a])
         * StrToFloat(grillaimportes.Cells[3,a])) * StrToFloat(grillaimportes.Cells[4,a]);
      grillaimportes.Cells[5,a]:=floattostr(RoundTo(dato,-2));
     end
     else begin
      dato:=0;
      grillaimportes.RowCount:=grillaimportes.RowCount+1;
      a:=grillaimportes.RowCount-1;
      grillaimportes.Cells[0,a]:=trim(tipoconst.Text) + 'Sup. Cubierta';
      grillaimportes.Cells[1,a]:=trim(sc.Text);
      grillaimportes.Cells[2,a]:=valor;
      grillaimportes.Cells[3,a]:=alicuota;
      dato:=(strtofloat(grillaimportes.Cells[1,a])
        * StrToFloat(grillaimportes.Cells[2,a])) * StrToFloat(grillaimportes.Cells[3,a]);
      grillaimportes.Cells[5,a]:=floattostr(RoundTo(dato,-2));
      dato:=0;
      grillaimportes.RowCount:=grillaimportes.RowCount+1;
      a:=grillaimportes.RowCount-1;
      grillaimportes.Cells[0,a]:=trim(tipoconst.Text)+ 'Sup. SemiCubierta';
      grillaimportes.Cells[1,a]:=trim(ssc.Text);
      grillaimportes.Cells[2,a]:=valor;
      grillaimportes.Cells[3,a]:=alicuota;
      grillaimportes.Cells[4,a]:='0.5';
      dato:=(strtofloat(grillaimportes.Cells[1,a])  * StrToFloat(grillaimportes.Cells[2,a])
        * StrToFloat(grillaimportes.Cells[3,a])) * strtofloat(grillaimportes.Cells[4,a]);
      grillaimportes.Cells[5,a]:=floattostr(RoundTo(dato,-2));
     end;
   end
   else begin
     if trim(grillaimportes.Cells[0,1])='' then begin
      dato:=0;
      grillaimportes.Cells[0,1]:=trim(tipoconst.Text);
      grillaimportes.Cells[1,1]:=trim(sc.Text);
      grillaimportes.Cells[2,1]:=valor;
      grillaimportes.Cells[3,1]:='0';
      dato:=strtofloat(grillaimportes.Cells[1,1]) * strtofloat(grillaimportes.Cells[3,1]);
      grillaimportes.Cells[5,1]:=floattostr(RoundTo(dato,-2));
     end
     else begin
      dato:=0;
      grillaimportes.RowCount:=grillaimportes.RowCount+1;
      a:=grillaimportes.RowCount-1;
      grillaimportes.Cells[0,a]:=trim(tipoconst.Text);
      grillaimportes.Cells[1,a]:=trim(sc.Text);
      grillaimportes.Cells[2,a]:=valor;
      grillaimportes.Cells[3,a]:='0';
      dato:=strtofloat(grillaimportes.Cells[1,a]) * strtofloat(grillaimportes.Cells[3,a]);
      grillaimportes.Cells[5,a]:=floattostr(RoundTo(dato,-2));
     end;
   end;
   cargaliqui.Enabled:=false;
   ControlaSelec(5);
  end;
 end;
end;

procedure TCalcuDerecho.SCChange(Sender: TObject);
var p,u:integer;
begin
{ if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4' then begin
    ts.Text := floattostr (strtofloat(ssc.Text)   + strtofloat(sc.Text));
    consultas(1);
 end
 else begin
  for p:=0 to grillacalculo.ColCount-1 do
    for u:=1 to grillacalculo.RowCount-1 do grillacalculo.Cells[p,u]:='';
     grillacalculo.RowCount:=2;
     consultas(1);
 end;  }
end;

procedure TCalcuDerecho.SCExit(Sender: TObject);
var p,u:integer;
begin
  if (trim(sc.Text)= '') or (trim(sc.Text)='0') then begin
   messagedlg('Error!! El valor ingresado es Incorrecto',mterror,[mbok],0);
   //if trim(sc.Text) = '' then begin
   //sc.Text:='0';
   sc.SetFocus;
  end
  else begin
   if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4' then begin
     ts.Text := floattostr (strtofloat(ssc.Text)   + strtofloat(sc.Text));
     consultas(1);
   end
   else begin
    for p:=0 to grillacalculo.ColCount-1 do
      for u:=1 to grillacalculo.RowCount-1 do grillacalculo.Cells[p,u]:='';
       grillacalculo.RowCount:=2;
       consultas(1);
   end;
    if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 = '4' then  grillacalculo.SetFocus
      ELSE ssc.SetFocus;
 end;
end;

procedure TCalcuDerecho.SCKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  //if trim(sc.Text) = '' then sc.Text:='0';
  if (trim(sc.Text)= '') or (trim(sc.Text)='0') then begin
    messagedlg('Error!! El valor ingresado es Incorrecto',mterror,[mbok],0);
   //if trim(sc.Text) = '' then begin
   //sc.Text:='0';
   sc.SetFocus;
  end
  else begin
    if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 = '4' then begin
     Consultas(1);
     PC1.ActivePage :=PC1.Pages[0];
     grillacalculo.SetFocus;
    end
    else ssc.SetFocus;
  end;
end
else begin
  if ( StrScan('0123456789.',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TCalcuDerecho.SSCChange(Sender: TObject);
begin
// if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4' then begin
//    ts.Text := floattostr (strtofloat(ssc.Text)   + strtofloat(sc.Text));
//    consultas(1) ;
// end
end;

procedure TCalcuDerecho.SSCExit(Sender: TObject);
begin
 if trim(ssc.Text) = '' then begin
  ssc.Text:='0';
  ts.Text := floattostr(strtofloat(sc.Text) + strtofloat(ssc.Text));
 end
 else begin
  if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4' then begin
    ts.Text := floattostr (strtofloat(ssc.Text)   + strtofloat(sc.Text));
    consultas(1) ;
  end;
  ts.Text := floattostr(strtofloat(sc.Text) + strtofloat(ssc.Text));
 end;
end;

procedure TCalcuDerecho.SSCKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  if trim(ssc.Text) = '' then ssc.Text:='0';
  ts.Text := floattostr(strtofloat(sc.Text) + strtofloat(ssc.Text));
  PC1.ActivePage :=PC1.Pages[0];
  agrega.SetFocus;
end
else begin
  if ( StrScan('0123456789.',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TCalcuDerecho.TipoConstChange(Sender: TObject);
var p,u:integer;
begin
  for p:=0 to grillaentidad.ColCount-1 do
    for u:=1 to grillaentidad.RowCount-1 do grillaentidad.Cells[p,u]:='';
  grillaentidad.RowCount:=2;

  for p:=0 to grillacalculo.ColCount-1 do
    for u:=1 to grillacalculo.RowCount-1 do grillacalculo.Cells[p,u]:='';
  grillacalculo.RowCount:=2;

  for p:=0 to grillaimportes.ColCount-1 do
    for u:=1 to grillaimportes.RowCount-1 do grillaimportes.Cells[p,u]:='';
  grillaimportes.RowCount:=2;

    PC1.ActivePage :=PC1.Pages[0];
end;

procedure TCalcuDerecho.TipoConstExit(Sender: TObject);
begin
if TipoConst.ItemIndex < 0 then begin
  messagedlg('Error!! Seleccione un tipo de Construcción válido',mterror,[mbok],0);
  TipoConst.SetFocus;
end
else Consultas(0);
end;

procedure TCalcuDerecho.TipoConstKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  sc.Text:='0';
  ssc.Text:='0';
  sc.SetFocus;
end;
end;

procedure TCalcuDerecho.TotLiquiChange(Sender: TObject);
begin
  impliqui.Enabled :=true;
  edit1.Text := trim(totliqui.Text);
end;

procedure TCalcuDerecho.TSChange(Sender: TObject);
begin
 if trim(ts.Text)='0' then begin
   messagedlg('ERROR.. El Importe Calculado para la Liquidación NO puede ser Cero(0)!',mtError,[mbok],0);
   sc.SetFocus;
 end
 else Consultas(1);
end;

procedure TCalcuDerecho.GrillaEntidadClick(Sender: TObject);
begin
 if (trim(GrillaEntidad.Cells[3,GrillaEntidad.Row])='') then GrillaEntidad.Cells[3,GrillaEntidad.Row]:='x'
 else GrillaEntidad.Cells[3,GrillaEntidad.Row]:='';
 cargaliqui.Enabled :=true;
 cargaliqui.SetFocus;
end;

procedure TCalcuDerecho.losdatosClick(Sender: TObject);
var gf:integer;
begin
 if trim(totliqui.Text ) = '0' then begin
   messagedlg('Atencion.. No Existe ningun Importe Calculado para la Liquidación!',mtWarning,[mbok],0);
   PC1.ActivePage :=PC1.Pages[2];
 end
 else begin
  impliqui.Enabled:=false;
  PC1.ActivePage :=PC1.Pages[3];
  odnip.SetFocus;
 end;
end;

procedure TCalcuDerecho.ODNIPKeyPress(Sender: TObject; var Key: Char);
begin
if key =#13 then begin
  Key := #0;
  ONOMBp.SetFocus;
end
else begin
  if ( StrScan('0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TCalcuDerecho.ONOMBPKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  buscarprof.SetFocus;
end
else begin
  if ( StrScan('abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ ',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure TCalcuDerecho.OPARTIDAEnter(Sender: TObject);
begin
BuscarP.Enabled:=true;
end;

procedure TCalcuDerecho.OPARTIDAExit(Sender: TObject);
begin
 if (trim(OPARTIDA.Text) = '')  or (trim(OPARTIDA.Text) = '0')
 or (trim(OPARTIDA.Text) = '00') or (trim(OPARTIDA.Text) = '000')  then begin
  messagedlg('Error!! Ingrese una Partida válida..',mterror,[mbok],0);
  OPARTIDA.SetFocus;
 end;
end;

procedure TCalcuDerecho.OPARTIDAKeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then begin
  Key := #0;
  if buscarp.Enabled then buscarp.SetFocus;
end
else begin
  if ( StrScan('0123456789',Key) <> nil ) or
      ( Key = Char(VK_BACK) ) then
   else Key := #0;
end;
end;

procedure  TCalcuDerecho.ImprimeRecibo();
var ReporteRst: TfrxReport;
    t,y:integer;
    dmi:widestring;
    tmp:string;
begin
 try
  ReporteRst:=TfrxReport.Create(nil);
  ReporteRst.LoadFromFile(Modulo2.Sistema.Path_todo + '\Reportes\LiquidacionDerConst.fr3');
  TfrxPictureView(ReporteRst.FindObject('Logo1')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
  TfrxPictureView(ReporteRst.FindObject('Logo2')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
  TfrxPictureView(ReporteRst.FindObject('Logo3')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
  TfrxPictureView(ReporteRst.FindObject('Logo4')).Picture.LoadFromFile(Modulo2.Sistema.Path_todo + '\Imagenes\LogoM2.jpg');
  ReporteRst.Variables['NombreTit'] := QuotedStr(Recibo.NombreTit);
  ReporteRst.Variables['CuimTit'] := QuotedStr(Recibo.CuimTit);
  ReporteRst.Variables['DomicilioTit'] := QuotedStr(eldomicilio);
  ReporteRst.Variables['NombreRespo'] := QuotedStr(Recibo.NombreRespo);
  ReporteRst.Variables['CuimRespo'] := QuotedStr(Recibo.CuimRespo);
  ReporteRst.Variables['DomicilioRespo'] := QuotedStr(Recibo.DomicilioRespo);
  ReporteRst.Variables['CuotaDe'] := QuotedStr(Recibo.CuotaDe);
  ReporteRst.Variables['Anio'] := QuotedStr(Recibo.Anio);
  ReporteRst.Variables['NroPlan'] := QuotedStr(Recibo.NroPlan);
  ReporteRst.Variables['NroCuota'] := QuotedStr(Recibo.NroCuota);
  ReporteRst.Variables['FechaIni'] := QuotedStr(Recibo.FechaIni);
  ReporteRst.Variables['CantObj'] := QuotedStr(Recibo.CantObj);
  ReporteRst.Variables['Objetos'] := QuotedStr(Recibo.Objetos);
  ReporteRst.Variables['DetalleRecibo'] := QuotedStr(Recibo.DetalleRecibo);
  ReporteRst.Variables['NroCobro'] := QuotedStr(Recibo.NroCobro);
  ReporteRst.Variables['NroLiquidacion'] := QuotedStr(Recibo.NroLiquidacion);
  ReporteRst.Variables['PrimerV'] := QuotedStr(Recibo.PrimerV);
  ReporteRst.Variables['Importe1'] := QuotedStr(Recibo.Importe1);
  ReporteRst.Variables['SegundoV'] := QuotedStr(Recibo.SegundoV);
  ReporteRst.Variables['Importe2'] := QuotedStr(Recibo.Importe2);
  ReporteRst.Variables['DenomCatas'] := QuotedStr(' Denom. Cat.:(' + trim(n4.Text) +')');
  Recibo.CodBar:=codigo_bar(Recibo.NroCodBar);
  ReporteRst.Variables['CodBar'] := QuotedStr(Recibo.CodBar);
  ReporteRst.Variables['NroCodBar'] := QuotedStr(Recibo.NroCodBar);
  ReporteRst.Variables['TituloDesc'] := QuotedStr(Recibo.TituloDesc);
  ReporteRst.Variables['NroUsuario'] := QuotedStr(Modulo2.Usuario.UsuNro);
  ReporteRst.ShowReport;
 except
    on E: Exception do EnviarLog(Self.Caption, 'Error en Reporte ' + E.Message);
 end;
 FreeAndNil(ReporteRst);
end;

procedure  TCalcuDerecho.GuardarRecibo(nro:byte);
var dmi,tmp,imp,GER:string;
    t,y:integer;
    DataSet3:TADODataSet;
    Com1:TADOCommand;
    Cnx2:TADOConnection;
begin
 dmi:='';
 if nro = 2 then begin
  Recibo.TituloDesc:='EDIF. LINEA Y NIVEL';
  Recibo.CuotaDe:='UNICA';
  recibo.PrimerV:= FechaAct;
  recibo.Anio:=RightStr(FechaAct,4);
  recibo.NroCuota:=trim(tipoconst.Text);
  if TComboClase(TipoConst.Items.Objects[TipoConst.ItemIndex]).Param1 <> '4' then
    recibo.FechaIni:='Supercie Cubierta: ' +  trim(sc.Text) + '  -  Superficia SemiCubierta: ' + trim(ssc.Text)
    + '  -  Total Superficie: ' + trim(ts.Text)//=trim(destinoobra.Text);
  else recibo.FechaIni:='Valor del Presupuesto: $' +  trim(sc.Text);
  recibo.Objetos:='';//trim(empresa.Text);
  Recibo.NombreTit:=AnsiReplaceStr(eltituP,chr(39),chr(33));    //AnsiReplaceStr(eltitu,chr(39),chr(33));
  Recibo.DomicilioTit:='';//AnsiReplaceStr('domi',chr(39),chr(33));
  Recibo.CuimTit:=cuimP;
  Recibo.NroPlan :=trim(opartida.Text);
  //GER:=trim(letra.Text) + StrFill(trim(nroe.Text),8,'0','I') + '/'+ trim(año.Text) ;
  Recibo.DetalleRecibo:='PLANO - DERECHO DE CONSTRUCCION';// + GER;//;
  Recibo.NroUsuario:=Modulo2.Usuario.UsuNro;
  try
   Cnx2:=TADOConnection.Create(nil);
   Cnx2.LoginPrompt:=false;
   Cnx2.ConnectionString:=Modulo2.Sistema.ConectUsu;
   DataSet3:=TADODataSet.Create(nil);
   DataSet3.Connection:=Cnx2;
   DataSet3.CommandText:='EXEC darmaximonro ' + chr(39) + 'TASAANUALLIQUIDACION ' + chr(39) + ' , ' + modulo2.Sistema.PuntoV ;
   DataSet3.Open;
   Recibo.NroLiquidacion:=StrFill(DataSet3.Fields[0].AsString,7,'0','');
   DataSet3.Close;
  except
    on E: Exception do EnviarLog(Self.Caption, E.Message);
  end;

  imp:='13';
  Recibo.SubRubro:='1110504';
  Recibo.Importe1:=FormatFloat('000000.00',StrToFloat(AnsiReplaceStr(trim(totliqui.Text),',','.')));
  Recibo.Importe2:=FormatFloat('000000.00',StrToFloat(AnsiReplaceStr(trim(totliqui.Text),',','.')));
  Recibo.NroCobro:=modulo2.Sistema.PuntoV + RightStr(FechaAct,2) + ' - ' + imp + '90';
  dmi:='927' + StrFill(AnsiReplaceStr(Recibo.Importe1,'.',''),8,'0','');
  tmp:=FechaAct;
  if strtodate(trim(FechaAct)) >  strtodate(FechaAct) then tmp:=trim(FechaAct);
   dmi:=dmi + RightStr(tmp,2);
   dmi:=dmi + StrFill(fecha_dif('01/01/' +  RightStr(tmp,4),FechaAct),3,'0','');
   dmi:=dmi + StrFill(modulo2.Sistema.PuntoV,2,'0','') + RightStr(tmp,2);
   dmi:=dmi + imp + '00';
   dmi:=dmi + Recibo.NroLiquidacion + '0';
   dmi:=dmi + StrFill(AnsiReplaceStr(Recibo.Importe1,'.',''),8,'0','');
   dmi:=dmi + RightStr(tmp,2);
   dmi:=dmi + StrFill(fecha_dif('01/01/' +  RightStr(tmp,4),FechaAct),3,'0','');
   dmi:=dmi + digitos(dmi, 1);
   Recibo.NroCodBar:=dmi;
  end;

  try
   Com1:=TADOCommand.Create(nil);
   Com1.Connection:=Cnx2;
   DataSet3.CommandText:='EXEC SPA_tasaanualLiquidacion '
   + chr(39) + Recibo.CuimTit + chr(39)
   + ' , ' + Recibo.NroPlan
   + ' , ' + 'NULL, ' + '0, '  + MidStr(Recibo.CuimTit,4,8)
   + ' , ' + chr(39) + LeftStr(Recibo.NombreTit,50) + chr(39)
   + ' , ' + '0, ' + Recibo.Importe1
   + ' , ' + chr(39) + FechaAct + chr(39)
   + ' , ' + Recibo.Importe1
   + ' , ' + chr(39) + FechaAct + chr(39)
   + ' , ' + RightStr(tmp,4)
   + ' , ' + imp + '90 ,' + Recibo.NroLiquidacion
   + ' , ' + chr(39) + Recibo.NroCodBar + chr(39)
   + ' , ' + Modulo2.Sistema.PuntoV
   + ' , ' + Modulo2.Usuario.UsuNro
   + ' , ' + LeftStr(Recibo.SubRubro,5)
   + ' , ' + Recibo.NroPlan + ', NULL, '  + '0, ' + '0, '   + '0, '  + '0, ' + Recibo.SubRubro;
   DataSet3.Open;
   Recibo.IdRegistro:=DataSet3.Fields[0].AsString;
   DataSet3.Close;
   DataSet3.Connection:=nil;
  except
    on E: Exception do EnviarLog(Self.Caption, E.Message);
  end;
  Cnx2.Close;
  FreeAndNil(DataSet3);
  FreeAndNil(Com1);
  FreeAndNil(Cnx2);
end;

procedure TCalcuDerecho.impliquiClick(Sender: TObject);
var gf:integer;
begin
 gf:=messagedlg('Confirma Imprimir la Liquidación del Derecho de Construcción.?',mtInformation,[mbok,mbCancel],0);
 if gf = 1 then begin
  if (idcontribP = '') OR (TRIM(N9.TEXT)='') then  BEGIN
    messagedlg('ERROR... Falta datos para confeccionar la Liquidación!',mtError,[mbok],0);
  END
  else begin
   Recibo:=TRecibo.Create;
   GuardarRecibo(2);
   ImprimeRecibo;
   FreeAndNil(Recibo);
  end;
 end
 else messagedlg('Ud., Ha cancelado la Impresion de la liquidación!',mtWarning,[mbok],0);
end;

end.
