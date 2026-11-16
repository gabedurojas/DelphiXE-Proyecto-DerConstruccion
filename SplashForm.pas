unit SplashForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, Gauges, StdCtrls, JvExControls, JvgDigits,
  JvAnimatedImage, JvGIFCtrl;

type
  TSplash = class(TForm)
    ProgressBar1: TProgressBar;
    Label2: TLabel;
    Label3: TLabel;
    Label1: TLabel;
    reloj: TJvGIFAnimator;
    procedure FormActivate(Sender: TObject);
  private

    { Private declarations }
  public
    valor:string;
    { Public declarations }
  end;

var
  Splash: TSplash;

implementation

{$R *.dfm}

procedure TSplash.FormActivate(Sender: TObject);
begin
//valor:=inttostr(progressbar1.Position);
//label3.Caption:=inttostr(progressbar1.Position);
reloj.Animate:=true;
end;

end.
