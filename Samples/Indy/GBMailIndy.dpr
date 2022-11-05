program GBMailIndy;

uses
  Vcl.Forms,
  FIndy in 'FIndy.pas' {Form1},
  GBMail.Interfaces in '..\..\Source\GBMail.Interfaces.pas',
  GBMail.Base.Server in '..\..\Source\GBMail.Base.Server.pas',
  GBMail.Base in '..\..\Source\GBMail.Base.pas',
  GBMail.Indy in '..\..\Source\GBMail.Indy.pas',
  GBMail.Factory in '..\..\Source\GBMail.Factory.pas',
  GBMail.Outlook in '..\..\Source\GBMail.Outlook.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
