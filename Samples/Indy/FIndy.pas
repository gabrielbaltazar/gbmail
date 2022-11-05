unit FIndy;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, GBMail.Interfaces, GBMail.Indy;

type
  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var
  mail : IGBMail;
begin
  mail := TGBMailIndy.New;
  mail.Server
        .&Begin
            .Host('smtp.gmail.com')
            .Port(587)
            .Username('email@gmail.com')
            .Password('*********')
            .UseSSL(True)
            .UseTLS(True)
            .RequireAuthentication(True)
        .&End
      .From('gabrielbaltazar.f1@gmail.com')
      .AddRecipient('gabrielbaltazar.f1@gmail.com')
      .Subject('teste')
      .Message('Oi Gabriel')
      .AddAttachment('C:\Users\gabriel.baltazar\Documents\CnWizards\Scripts.xml')
      .AddAttachment('C:\Users\gabriel.baltazar\Documents\CnWizards\FeedCfg.xml')
      .Send;
end;

end.
