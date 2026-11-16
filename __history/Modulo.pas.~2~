unit Modulo;

interface

uses
  SysUtils, Classes, Messages, Windows, Dialogs, StdCtrls, ComCtrls, Forms,
  JvComponentBase, JvComputerInfoEx, Encryp, ADODB, JvAppStorage,
  JvAppIniStorage, Winsock, Objetos, RzCmboBx, DCPcrypt2, DCPsha1;

type
  TModulo2 = class(TDataModule)
    DCP_sha11: TDCP_sha1;
    function getIPs: Tstrings;
    function obtenerNombrePC : Ansistring;
    procedure DataModuleCreate(Sender: TObject);
    Function CreaHash(Archivo:string):string;
    procedure CargaCombo(var Combo:TComboBox;comando,ConectDB:Widestring; IncluirTodos:boolean);
    procedure LiberarCombo(var Combo:TComboBox);
    function ObtenerNivelAcceso(NroModulo:string):integer;
  private
    Encrypt:TTomEncryption;
    { Private declarations }
  public
    ObjConexion:TADOConnection;
    Ini1: TJvAppIniFileStorage;
    LstArchivos:TStringList;
    Sistema:TSistema;
    Usuario:TUsuario;
    ConAcceso:boolean;
    Filtro:TFiltro;
    { Public declarations }
  end;


var
  Modulo2: TModulo2;

implementation


{$R *.dfm}

function TModulo2.getIPs: Tstrings;
type
  TaPInAddr = array[0..10] of PInAddr;
  PaPInAddr = ^TaPInAddr; 
var 
  phe: PHostEnt; 
  pptr: PaPInAddr; 
  Buffer: array[0..63] of AnsiChar;
  I: Integer; 
  GInitData: TWSAData; 
begin 
  WSAStartup($101, GInitData); 
  Result := TstringList.Create; 
  Result.Clear; 
  GetHostName(Buffer, SizeOf(Buffer)); 
  phe := GetHostByName(buffer);
  if phe = nil then Exit; 
  pPtr := PaPInAddr(phe^.h_addr_list); 
  I    := 0; 
  while pPtr^[I] <> nil do 
  begin 
    Result.Add(inet_ntoa(pptr^[I]^));
    Inc(I);
  end;
  WSACleanup; 
end;

function TModulo2.obtenerNombrePC():Ansistring;
var
  nombre: AnsiString;
  DatosSocket: WSAData;
begin
  WSAStartup($0101, DatosSocket);
  SetLength(nombre, MAX_PATH);
  gethostname(PAnsiChar(Nombre), MAX_PATH);
  SetLength(nombre, StrLen(PAnsiChar(nombre)));
  result := nombre;
end;


Function TModulo2.CreaHash(Archivo:string):string;
var t:integer;
var
  Hashes: TDCP_hash;
  HashDigest: array of byte;
  i, j, read: integer;
  s: string;
  buffer: array[0..16383] of byte;
  strmInput: TFileStream;
begin
  Hashes := nil;
  // make a list of all the hash algorithms to use
  Hashes := TDCP_hash(FindComponent('DCP_sha11'));
  TDCP_hash(FindComponent('DCP_sha11')).Init;

  strmInput := nil;
  try
    strmInput := TFileStream.Create(Archivo,fmOpenRead);
    repeat
      // read into the buffer
      read := strmInput.Read(buffer,Sizeof(buffer));
      // hash the buffer with each of the selected hashes
      Hashes.Update(buffer,read);
    until read <> Sizeof(buffer);
    strmInput.Free;
    // iterate through the selected hashes
    SetLength(HashDigest,Hashes.HashSize div 8);
    Hashes.Final(HashDigest[0]);  // get the output
    s := '';
    for j := 0 to Length(HashDigest) - 1 do  // convert it into a hex string
      s := s + IntToHex(HashDigest[j],2);
    Result:=s;
  except
    strmInput.Free;
    Result:='';
  end;
end;


procedure TModulo2.DataModuleCreate(Sender: TObject);
var BaseDatos,usr,key,Psw,Segu:string;
    rs:integer;
begin
Sistema:=TSistema.Create;
Usuario:=TUsuario.Create;
Filtro:=TFiltro.Create;
Sistema.Path_todo:=GetCurrentDir();
try
  //Obtenemos lops datos del host cliente

  Sistema.MyIP:=getIPs.Strings[0];
  Sistema.MyHost:=obtenerNombrePC;

  //Componente de lectura de INI
  Ini1:=TJvAppIniFileStorage.Create(nil);
  Ini1.FileName:='Planes.ini';
  Ini1.Location:=flExeFile;

  Sistema.Server:=INI1.IniFile.ReadString('SQL','Host','127.0.0.1');
  //Componente de Cifrado
  Encrypt:=TTomEncryption.Create(nil);
  Encrypt.Action:=atDecryption;

  BaseDatos:=INI1.IniFile.ReadString('SQL','Base1','');
  key:=INI1.IniFile.ReadString('SQL','Default1','');
  usr:=INI1.IniFile.ReadString('SQL','Ini1','');
  Psw:=INI1.IniFile.ReadString('SQL','Cl1','');

  Encrypt.Input:=usr;
  Encrypt.Key:=key;
  Encrypt.Execute;
  Usr:=Encrypt.Output;

  Encrypt.Input:=Psw;
  Encrypt.Key:=key;
  Encrypt.Execute;
  Psw:=Encrypt.Output;

  /// conexion a testing
  { Sistema.ConectDb:='Provider=SQLOLEDB.1;Password=$uperu$er'  //+ Psw
       + ';Persist Security Info=True;User ID=sa' //+ Usr
       + ';Initial Catalog=PlanesPagos;Data Source=192.168.0.105' //+ Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectUsu:='Provider=SQLOLEDB.1;Password=$uperu$er'  //  + Psw
       + ';Persist Security Info=True;User ID=sa' // + Usr
       + ';Initial Catalog=municipio;Data Source=192.168.0.105' //+ Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectSeg:='Provider=SQLOLEDB.1;Password=$uperu$er'  // + Psw
       + ';Persist Security Info=True;User ID=sa' // + Usr
       + ';Initial Catalog=SeguridadWeb;Data Source=192.168.0.105' //+ Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectInti:='Provider=SQLOLEDB.1;Password=$uperu$er'  //+ Psw
       + ';Persist Security Info=True;User ID=sa' //+ Usr
       + ';Initial Catalog=IntimacionesNew;Data Source=192.168.0.105' //+ Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

   Sistema.ConectDerConst:='Provider=SQLOLEDB.1;Password=$uperu$er'   // +Psw
       + ';Persist Security Info=True;User ID=sa' //+ Usr
       + ';Initial Catalog=DerConst;Data Source=192.168.0.105'  //+ Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';
  }


 /// conexion a produccion
  Sistema.ConectDb:='Provider=SQLOLEDB.1;Password='  + Psw
       + ';Persist Security Info=True;User ID=' + Usr
       + ';Initial Catalog=' + BaseDatos + ';Data Source=' + Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectUsu:='Provider=SQLOLEDB.1;Password='  + Psw
       + ';Persist Security Info=True;User ID=' + Usr
       + ';Initial Catalog=municipio;Data Source=' + Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectSeg:='Provider=SQLOLEDB.1;Password='  + Psw
       + ';Persist Security Info=True;User ID=' + Usr
       + ';Initial Catalog=SeguridadWeb;Data Source=' + Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectInti:='Provider=SQLOLEDB.1;Password='+ Psw
       + ';Persist Security Info=True;User ID=' + Usr
       + ';Initial Catalog=IntimacionesNew;Data Source=' + Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';

  Sistema.ConectDerConst:='Provider=SQLOLEDB.1;Password='+ Psw
       + ';Persist Security Info=True;User ID=' + Usr
       + ';Initial Catalog=DerConst;Data Source=' + Sistema.Server
       + ';Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'
       + 'Use Encryption for Data=False;Tag with column collation when possible=False;';


  Encrypt.Input:=INI1.IniFile.ReadString('SQL','PuntoV','20');
  Encrypt.Key:=key;
  Encrypt.Execute;
  Sistema.PuntoV:=Encrypt.Output;
  FreeAndNil(Encrypt);
except
  on Exception do
    messagedlg('Error al Intento de Conección a la Base de Datos de SQL Server',mterror,[mbok],0);
end;
end;

procedure TModulo2.CargaCombo(var Combo:TComboBox;  comando,ConectDB:Widestring; IncluirTodos:boolean);
var k,z:integer;
    DataSet1:TADODataSet;
    Cnx:TADOConnection;
    MiObjeto:TComboClase;
begin
Combo.Clear;
Cnx:=TADOConnection.Create(nil);
DataSet1:=TADODataSet.Create(nil);
DataSet1.ConnectionString:=ConectDB;
DataSet1.CommandText:=Comando;
DataSet1.Open;
if IncluirTodos then begin
  MiObjeto:=TComboClase.Create;
  MiObjeto.Param1:='NULL';
  Combo.Items.AddObject('TODOS',MiObjeto);
end;
for k:=1 to DataSet1.RecordCount do begin
  DataSet1.RecNo:=k;
  MiObjeto:=TComboClase.Create;
  for z:=1 to DataSet1.FieldCount-1 do begin
    if z = 1 then MiObjeto.Param1:=trim(DataSet1.Fields[1].AsString);
    if z = 2 then MiObjeto.Param2:=trim(DataSet1.Fields[2].AsString);
    if z = 3 then MiObjeto.Param3:=trim(DataSet1.Fields[3].AsString);
    if z = 4 then MiObjeto.Param4:=trim(DataSet1.Fields[4].AsString);
  end;
  Combo.Items.AddObject(trim(DataSet1.Fields[0].AsString),MiObjeto);
end;
DataSet1.Close;
DataSet1.Connection:=nil;
FreeAndNil(DataSet1);
Combo.ItemIndex:=0;
end;

procedure TModulo2.LiberarCombo(var Combo:TComboBox);
var q:integer;
begin
For q:= Combo.items.count-1 downto 0 do
  TComboClase(Combo.items.objects[q]).Free;
end;

function TModulo2.ObtenerNivelAcceso(NroModulo:string):integer;
var q:integer;
begin
Result:=0;
For q:=0 to Usuario.UsuNroArea.Count-1 do begin
  if Usuario.UsuNroArea.Strings[q] = NroModulo then begin
    Result:=StrToInt(Usuario.UsuNvl.Strings[q]);
    break;
  end;
end;
end;





end.
