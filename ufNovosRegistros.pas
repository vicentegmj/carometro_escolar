unit ufNovosRegistros;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Buttons;

type
  TfNovosRegistros = class(TForm)
    pnlTitulo: TPanel;
    btnGravar: TBitBtn;
    mmRegistros: TMemo;
    btnExemploTurmas: TBitBtn;
    btnExemploAluno: TBitBtn;
    btnLimpar: TBitBtn;
    lbl1: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnGravarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnExemploTurmasClick(Sender: TObject);
    procedure btnExemploAlunoClick(Sender: TObject);
    procedure btnLimparClick(Sender: TObject);
    procedure RemoverLinhasEmBrancoOuEspacos(Memo: TMemo);
    procedure LimparMemoCaracteresInvalidos(Memo: TMemo);
  private
    procedure InserirDados(Memo: TMemo; IdTurma, Ano: Integer);
  public
    { Public declarations }
  end;

var
  fNovosRegistros: TfNovosRegistros;
  IdTurma, Ano: Integer;

implementation

uses
  uDM, ufmain;

{$R *.dfm}

{ TfNovosRegistros }

procedure TfNovosRegistros.btnExemploAlunoClick(Sender: TObject);
var
  lista: TStringList;
begin
  mmRegistros.Clear;
  lista := TStringList.Create;
  try
    lista.Add('ISAAC NEWTON');
    lista.Add('ALBERT EINSTEIN');
    lista.Add('GALILEU GALILEI');
    lista.Add('NIELS BOHR');
    lista.Add('MARIE CURIE');
    lista.Add('JAMES CLERK MAXWELL');
    lista.Add('RICHARD FEYNMAN');
    lista.Add('STEPHEN HAWKING');
    lista.Add('ERWIN SCHRÖDINGER');
    lista.Add('WERNER HEISENBERG');
    mmRegistros.Lines.AddStrings(lista);
  finally
    lista.Free;
  end;
//  btnGravar.Enabled := False;
end;

procedure TfNovosRegistros.btnExemploTurmasClick(Sender: TObject);
var
  lista: TStringList;
begin
  mmRegistros.Clear;
  lista := TStringList.Create;
  try
    lista.Add('5A');
    lista.Add('6B');
    lista.Add('9A');
    lista.Add('1A');
    lista.Add('1B');
    lista.Add('2A');
    lista.Add('2A (TÉCNICO)');
    lista.Add('2B');
    lista.Add('3A');
    lista.Add('3B');
    mmRegistros.Lines.AddStrings(lista);
  finally
    lista.Free;
  end;
//  btnGravar.Enabled := False;
end;

function LimparCaracteresInvalidos(const Texto: string): string;
const
  Invalidos: array[0..8] of Char = ('<', '>', ':', '"', '/', '\', '|', '?', '*');
var
  i: Integer;
  Resultado: string;
begin
  Resultado := Texto;
  for i := Low(Invalidos) to High(Invalidos) do
    Resultado := StringReplace(Resultado, Invalidos[i], '', [rfReplaceAll]);
  Result := Resultado;
end;

function MemoEstaVazio(Memo: TMemo): Boolean;
var
  i: Integer;
begin
  Result := True;
  for i := 0 to Memo.Lines.Count - 1 do
  begin
    if Trim(Memo.Lines[i]) <> '' then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

procedure TfNovosRegistros.btnGravarClick(Sender: TObject);
var
  Confirmacao: Integer;
  i: Integer;
  TextoNumerado: string;
  TextoCabecalho: string;
begin

  RemoverLinhasEmBrancoOuEspacos(mmRegistros);

  LimparMemoCaracteresInvalidos(mmRegistros);

  if MemoEstaVazio(mmRegistros) = true then
    Exit;

  TextoNumerado := '';
  for i := 0 to mmRegistros.Lines.Count - 1 do
    TextoNumerado := TextoNumerado + IntToStr(i + 1) + '. ' + mmRegistros.Lines[i] +
      sLineBreak;

  if Ano > 0 then
  begin
    TextoCabecalho := 'Você está prestes a gravar ' + IntToStr(i) + ' "Turmas Novas":'
  end;

  if IdTurma > 0 then
  begin
    TextoCabecalho := 'Você está prestes a gravar ' + IntToStr(i) + ' "Alunos Novos":'
  end;

  Confirmacao := MessageDlg(TextoCabecalho + sLineBreak + sLineBreak + TextoNumerado
    + sLineBreak +
    'ATENÇÃO: Após a gravação, não será possível alterar esses registros.' + sLineBreak
    + 'Deseja continuar?', mtWarning, [mbYes, mbNo], 0);

  if Confirmacao = mrYes then
  begin
    if Ano > 0 then
    begin
      if fmain.qrTurmasAno.Active then
        fmain.qrTurmasAno.Close;

      InserirDados(mmRegistros, 0, Ano);

      fmain.qrTurmasAno.ParamByName('ano').AsInteger := Ano;
      fmain.qrTurmasAno.Open();
    end;

    if IdTurma > 0 then
    begin
      if fmain.qrAlunosTurma.Active then
        fmain.qrAlunosTurma.Close;

      InserirDados(mmRegistros, fmain.qrTurmasAnoid.Value, 0);
      fmain.qrAlunosTurma.Open();
    end;
  end;
end;

procedure TfNovosRegistros.btnLimparClick(Sender: TObject);
begin
  mmRegistros.Clear;
  btnGravar.Enabled := True;
end;

procedure TfNovosRegistros.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
  fNovosRegistros := nil;
end;

procedure TfNovosRegistros.FormShow(Sender: TObject);
begin
  btnExemploTurmas.Caption := 'Exemplo' + sLineBreak + 'Turma(s)';
  btnExemploAluno.Caption := 'Exemplo' + sLineBreak + 'Aluno(s)';
  mmRegistros.Clear;
end;

procedure TfNovosRegistros.InserirDados(Memo: TMemo; IdTurma, Ano: Integer);
var
  i: Integer;
begin
  if (Memo = nil) or (Memo.Lines.Count = 0) then
    Exit;

  for i := 0 to Memo.Lines.Count - 1 do
  begin
    if Trim(Memo.Lines[i]) <> '' then
    begin
      DM.FDQuery1.SQL.Clear;

      if Ano > 0 then
      begin
        // Inserir na tabela TURMAS
        DM.FDQuery1.SQL.Add('INSERT INTO turmas (nome, ano) VALUES (:nome, :ano)');
        DM.FDQuery1.ParamByName('nome').AsString := Trim(Memo.Lines[i]);
        DM.FDQuery1.ParamByName('ano').AsInteger := Ano;
      end
      else if IdTurma > 0 then
      begin
        // Inserir na tabela ALUNOS
        DM.FDQuery1.SQL.Add('INSERT INTO alunos (nome, id_turma) VALUES (:nome, :id_turma)');
        DM.FDQuery1.ParamByName('nome').AsString := Trim(Memo.Lines[i]);
        DM.FDQuery1.ParamByName('id_turma').AsInteger := IdTurma;
      end
      else
        Exit; // Nenhum parâmetro válido

      DM.FDQuery1.ExecSQL;
    end;
  end;
  Close;
end;

procedure TfNovosRegistros.LimparMemoCaracteresInvalidos(Memo: TMemo);
var
  i: Integer;
begin
  for i := 0 to Memo.Lines.Count - 1 do
    Memo.Lines[i] := LimparCaracteresInvalidos(Memo.Lines[i]);
end;

procedure TfNovosRegistros.RemoverLinhasEmBrancoOuEspacos(Memo: TMemo);
var
  i: Integer;
begin
  for i := Memo.Lines.Count - 1 downto 0 do
    if Trim(Memo.Lines[i]) = '' then
      Memo.Lines.Delete(i);
end;

end.

