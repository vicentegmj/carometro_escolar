program Carometro;

uses
  Vcl.Forms,
  DSPack in 'src\DSPack.pas',
  DXSUtil in 'src\DXSUtil.pas',
  uDM in 'uDM.pas' {DM: TDataModule},
  ufNovosRegistros in 'ufNovosRegistros.pas' {fNovosRegistros},
  ufImagemAtual in 'ufImagemAtual.pas' {fImagemAtual},
  ufmain in 'ufmain.pas' {fmain},
  ufConfirmacao in 'ufConfirmacao.pas' {fConfirmacao},
  ufInformacao in 'ufInformacao.pas' {fInformacao};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TDM, DM);
  Application.CreateForm(Tfmain, fmain);
  Application.CreateForm(TfImagemAtual, fImagemAtual);
  Application.CreateForm(TfConfirmacao, fConfirmacao);
  Application.CreateForm(TfInformacao, fInformacao);
  Application.Run;
end.
