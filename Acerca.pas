unit acerca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, jpeg, GR32_Image, JvExStdCtrls, JvBehaviorLabel,
  JvExExtCtrls, JvExtComponent, JvPanel;

type
  TAcerca_de = class(TForm)
    Image321: TImage32;
    Label5: TLabel;
    Label8: TLabel;
    Label7: TLabel;
    Label4: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    Label6: TLabel;
    Panel3: TPanel;
    JvPanel2: TJvPanel;
    JvBehaviorLabel2: TJvBehaviorLabel;
    Label1: TLabel;
    procedure Panel3MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Label1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Acerca_de: TAcerca_de;

implementation

uses Menu;

{$R *.dfm}

procedure TAcerca_de.Label1Click(Sender: TObject);
begin
Close;
end;

procedure TAcerca_de.Panel3MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
ReleaseCapture();
Perform( WM_SYSCOMMAND, $F012, 0 );
end;

end.
