unit FIndy;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  GBMail.Interfaces, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  LMail : IGBMail;
begin
  LMail := GetMailDefault;
  LMail.Server
      .Host('smtp.gmail.com')
      .Port(587)
      .Username('email@gmail.com')
      .Password('*********')
      .UseSSL(True)
      .UseTLS(True)
      .RequireAuthentication(True)
    .&End
    .From('No-Reply Test')
    .AddRecipient('recipient@email.com')
    .Subject('test mail')
    .Message('Hi Test')
    .Send;
end;

end.
