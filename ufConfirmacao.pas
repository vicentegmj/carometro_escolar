unit ufConfirmacao;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons;

type
  TfConfirmacao = class(TForm)
    btnSim: TBitBtn;
    btnNao: TBitBtn;
    edtRandon: TEdit;
    edtConfirma: TEdit;
    lblTexto: TLabel;
    lblConfirma: TLabel;
    lblConfirma1: TLabel;
    procedure btnSimClick(Sender: TObject);
    procedure btnNaoClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure edtConfirmaChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fConfirmacao: TfConfirmacao;
  confirmacao: Boolean;

implementation

{$R *.dfm}

function GerarNumeroAleatorio3Digitos: Integer;
begin
  Randomize; // inicializa o gerador de números aleatórios (necessário apenas uma vez)
  Result := 100 + Random(900); // gera número entre 100 e 999
end;

procedure TfConfirmacao.btnNaoClick(Sender: TObject);
begin
  confirmacao := False;
  Close;
end;

procedure TfConfirmacao.btnSimClick(Sender: TObject);
begin
  confirmacao := True;
  Close;
end;

procedure TfConfirmacao.edtConfirmaChange(Sender: TObject);
begin
  if edtRandon.Text = edtConfirma.Text then
    btnSim.Enabled := True
  else
    btnSim.Enabled := False;
end;

procedure TfConfirmacao.FormShow(Sender: TObject);
begin
  edtConfirma.Clear;
  confirmacao := False;
  btnSim.Enabled := False;
  edtRandon.Text := IntToStr(GerarNumeroAleatorio3Digitos);
  lblTexto.Width := 449;
  lblTexto.Height := 64;
end;

end.

