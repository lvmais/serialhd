program DadosHardware2015;



{$R *.dres}

uses
  sharemem,
  Forms,
  uDadosHardware in 'uDadosHardware.pas' {Form1},
  magwmi in 'obj\magwmi.pas',
  WbemScripting_TLB in 'obj\WbemScripting_TLB.pas',
  magsubs1 in 'obj\magsubs1.pas',
  smartapi in 'obj\smartapi.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Dados de Hardware';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
