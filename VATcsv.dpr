program VATcsv;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {main},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Carbon');
  Application.CreateForm(Tmain, main);
  Application.Run;
end.
