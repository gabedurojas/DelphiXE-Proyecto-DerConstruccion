unit Acceso1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DB, ADODB, Vcl.Buttons, JvExButtons, JvBitBtn,
  Vcl.ExtCtrls;

type
  TAcceso = class(TForm)
    Edit2: TMaskEdit;
    edit1: TEdit;
    GroupBox1: TGroupBox;
    Label4: TLabel;
    Boton1: TJvBitBtn;
    Button2: TJvBitBtn;
    GroupBox2: TGroupBox;
    Label5: TLabel;
    Image1: TImage;
    Image3: TImage;
    Image2: TImage;
    Image4: TImage;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormActivate(Sender: TObject);
    procedure edit1Change(Sender: TObject);
    procedure edit1KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit2Exit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure Boton1Click(Sender: TObject);
    function GetAppVersion: string;
    procedure VerVersion(nro:byte);
  private
    DataSet1,DataSet2:TADODataSet;
    Com1:TADOCommand;
    Cnx:TADOConnection;{ Private declarations }
  public
    control:boolean;
    cuenta:integer;
    usuario,dato:string;{ Public declarations }
  end;

var
  Acceso: TAcceso;

implementation

uses modulo;

{$R *.dfm}
procedure TAcceso.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(DataSet1);
Cnx.Close;
FreeAndNil(Cnx);
end;

procedure TAcceso.FormCreate(Sender: TObject);
begin
Cnx:=TADOConnection.Create(nil);
Cnx.LoginPrompt:=false;
Cnx.ConnectionString:=Modulo2.Sistema.ConectUsu;
DataSet1:=TADODataSet.Create(nil);
DataSet1.Connection:=Cnx;
end;

procedure TAcceso.FormActivate(Sender: TObject);
begin
boton1.Enabled:=false;
control:=false;
label5.Caption :='Version - ' + GetAppVersion;
VerVersion(1);
groupbox2.Visible :=false;
Edit2.SetFocus;
cuenta:=1;
end;

function TAcceso.GetAppVersion: string;
var
    Size, Size2: DWord;
    Pt, Pt2: Pointer;
begin
  Size := GetFileVersionInfoSize(PChar(ParamStr(0)), Size2);
  if Size > 0 then
  begin
    GetMem(Pt, Size);
    try
      GetFileVersionInfo(PChar(ParamStr(0)), 0, Size, Pt);
      VerQueryValue(Pt, '\', Pt2, Size2);
      with TVSFixedFileInfo(Pt2^) do
      begin
        Result := '' +
        IntToStr(HiWord(dwFileVersionMS)) +  IntToStr(LoWord(dwFileVersionMS)) +
        IntToStr(HiWord(dwFileVersionLS)) +  IntToStr(LoWord(dwFileVersionLS)) ;
        //IntToStr(HiWord(dwFileVersionMS)) + '.' +
        //IntToStr(LoWord(dwFileVersionMS)) + '.' +
        //IntToStr(HiWord(dwFileVersionLS));
        //+ ' Build ' + IntToStr(LoWord(dwFileVersionLS));
      end;
    finally
      FreeMem(Pt);
    end;
  end;
end;

procedure TAcceso.VerVersion(nro:byte);
 var   Cnx2:TADOConnection;
begin
  Cnx2:=TADOConnection.Create(nil);
  Cnx2.LoginPrompt:=false;
  Cnx2.ConnectionString:=Modulo2.Sistema.ConectUsu;
  if nro=1 then begin
    DataSet2:=TADODataSet.Create(nil);
    DataSet2.Close;
    DataSet2.Connection:=Cnx2;
    DataSet2.CommandText:='SELECT NroVersion FROM VersionSoft where (NroAplicacion=6) and (activo=1)';
    DataSet2.Open;
    if not Dataset2.IsEmpty  then dato:=DataSet2.Fields[0].AsString
      else dato:='0';
    DataSet2.Connection:=nil;
  end;
  Cnx2.Close;
  FreeAndNil(Cnx2);
end;


procedure TAcceso.Boton1Click(Sender: TObject);
var clav,clav1:string;
    t:integer;
begin
if GetAppVersion <> dato then begin
 groupbox2.Visible :=true;
 boton1.Enabled :=false;
 button2.SetFocus;
end
else begin
 if (cuenta < 3) or (DataSet1.Fields[3].AsInteger > 3) then begin
  if not(DataSet1.Fields[3].AsInteger > 3) then begin
   clav:=trim(DataSet1.Fields[2].AsString);{2}
   DataSet2:=TADODataSet.Create(nil);
   DataSet2.Connection:=Cnx;
   DataSet2.CommandText:='todos_loguinSinmod ' + trim(Edit2.Text) + ' , ' + chr(39) + trim(Edit1.Text) + chr(39) + ';';
   DataSet2.Open;
   clav1:=DataSet2.Fields[0].AsString;
   DataSet2.Close;

   Label4.Caption:=trim(DataSet1.Fields[1].AsString);
   if clav1 <> 'OK' then begin
     messagedlg( clav1 + '!!!'+chr(13)+chr(13)+'   Acceso DENEGADO!!!',mterror,[mbok],0);
     cuenta:=cuenta+1;
     edit1.SetFocus;
   end
   else begin
     Modulo2.Usuario.UsuNro:=trim(Edit2.Text);
     Modulo2.Usuario.UsuNombre:=usuario;
     Modulo2.Usuario.UsuNroArea.Clear;
     Modulo2.Usuario.UsuNvl.Clear;
     Modulo2.Usuario.UsuDetalleNivel.Clear;
     //DataSet2.CommandText:='SELECT Nivel FROM Municipio.dbo.Usumodulos WHERE (codusu='  + trim(Edit2.Text) + ') AND (modulo IN (26));';
     DataSet2.CommandText:='SELECT Municipio.dbo.Usumodulos.Modulo, Municipio.dbo.Usumodulos.Nivel,'
     + 'Municipio.dbo.NivelAcceso.DetalleNivel FROM Municipio.dbo.Usumodulos inner join '
     + 'Municipio.dbo.NivelAcceso on Municipio.dbo.NivelAcceso.idNivel = Municipio.dbo.Usumodulos.NIVEL '
     + 'WHERE (codusu=' + trim(Edit2.Text) + ');';
     DataSet2.Open;
     for t:=1 to DataSet2.RecordCount do begin
       DataSet2.RecNo:=t;
       Modulo2.Usuario.UsuNvl.Create.Append(DataSet2.FieldByName('Nivel').AsString);
       Modulo2.Usuario.UsuNroArea.Create.Append(DataSet2.FieldByName('Modulo').AsString);
       Modulo2.Usuario.UsuDetalleNivel.Append(DataSet2.FieldByName('DetalleNivel').AsString);
       control:=true;
     end;
     DataSet2.Close;
     DataSet2.Connection:=nil;
     FreeAndNil(DataSet2);
     close;
   end;
  end
  else begin
   cuenta:=DataSet1.Fields[3].AsInteger;
   messagedlg('El Nº de Operador ha Caducado su Cuenta'+chr(13)+chr(13)
   + '   Acceso DENEGADO!!!'+chr(13)+'Consulte con el Administrador',mterror,[mbok],0);
   close;
  end;
 end
 else begin
  try
   Com1:=TADOCommand.Create(nil);
   Com1.Connection:=Cnx;
   Com1.CommandText:='UPDATE Usuarios2 SET intentos =' + chr(39) + inttostr(cuenta+1) + chr(39) +
                      'WHERE codusu =' + Modulo2.Usuario.UsuNro;
   Com1.Execute;
  except
  end;
  Com1.Connection:=nil;
 end;
end;
end;


procedure TAcceso.Button2Click(Sender: TObject);
begin
close;
end;

procedure TAcceso.edit1Change(Sender: TObject);
begin
if trim(edit1.Text)='' then begin
 boton1.Enabled:=false;
 end
else begin
 boton1.Enabled:=true;
 button2.Enabled:=true;
end;
end;

procedure TAcceso.edit1KeyPress(Sender: TObject; var Key: Char);
begin
if (key = chr(13))  then begin
 if boton1.Enabled = true then boton1.SetFocus;
end;
end;

procedure TAcceso.Edit2KeyPress(Sender: TObject; var Key: Char);
begin
if key = chr(13) then edit1.SetFocus;
end;

procedure TAcceso.Edit2Exit(Sender: TObject);
begin
if button2.Focused then button2.SetFocus
else begin
 try
   DataSet1.Close;
   DataSet1.CommandText:='SELECT * FROM Usuarios2 WHERE (codusu =' + trim(edit2.Text) + ');';
   DataSet1.Open;
   if DataSet1.RecordCount = 0 then begin
     messagedlg('No exite el Nº de Operador Ingresado',mterror,[mbok],0);
     edit2.SetFocus;
   end
   else begin
     Label4.Caption:=trim(DataSet1.Fields[1].AsString);
     usuario:=trim(DataSet1.Fields[1].AsString);
     if DataSet1.Fields[3].AsInteger < 1 then cuenta:=1
      else cuenta:=DataSet1.Fields[3].AsInteger;
     Modulo2.Usuario.UsuNro:=DataSet1.Fields[0].AsString;
   end;
 except
    control:=true;
 end;
end;
end;


end.
