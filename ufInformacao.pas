unit ufInformacao;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons;

type
  TfInformacao = class(TForm)
    mm: TMemo;
    btnOk: TBitBtn;
    procedure btnOkClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fInformacao: TfInformacao;

implementation

{$R *.dfm}

procedure TfInformacao.btnOkClick(Sender: TObject);
begin
  Close;
end;

end.
