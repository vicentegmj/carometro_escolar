unit ufmain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  DSPack, Vcl.Menus, DXSUtil, DirectShow9, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Grids, Vcl.DBGrids, Vcl.Buttons, Vcl.DBCtrls,
  System.IOUtils, Vcl.Imaging.jpeg, ShellAPI, Vcl.ExtDlgs, ActiveX, frxClass,
  frxDBSet, Printers, Vcl.Mask, System.IniFiles, System.DateUtils, Vcl.ToolWin,
  Vcl.ComCtrls, Clipbrd;

type
  Tfmain = class(TForm)
    FilterGraph: TFilterGraph;
    Filter: TFilter;
    SampleGrabber: TSampleGrabber;
    cbbAno: TComboBox;
    lblAno: TLabel;
    qrTurmasAno: TFDQuery;
    qrAlunosTurma: TFDQuery;
    ds1: TDataSource;
    ds2: TDataSource;
    dbgTurmas: TDBGrid;
    qrTurmasAnoid: TFDAutoIncField;
    qrTurmasAnonome: TStringField;
    qrTurmasAnoano: TIntegerField;
    qrAlunosTurmaid: TFDAutoIncField;
    qrAlunosTurmaid_turma: TIntegerField;
    qrAlunosTurmanome: TStringField;
    dbgAlunos: TDBGrid;
    btnNovaTurma: TBitBtn;
    btnNovoAlunoTurma: TBitBtn;
    cbbDevices: TComboBox;
    btnLigarDevice: TBitBtn;
    lblWebCams: TLabel;
    btnGravarImagem: TBitBtn;
    Label1: TLabel;
    Image: TImage;
    qrAlunosTurmapath_imagem: TStringField;
    edtPesquisa: TEdit;
    SpeedButton1: TSpeedButton;
    btn1: TBitBtn;
    Label3: TLabel;
    btnFotometro: TBitBtn;
    Label4: TLabel;
    image3: TImage;
    chkImagem: TCheckBox;
    mm: TMainMenu;
    Cadastro1: TMenuItem;
    NovaTurma1: TMenuItem;
    NovoAluno1: TMenuItem;
    Relatrio1: TMenuItem;
    FotmetroPdf1: TMenuItem;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    N1: TMenuItem;
    Bevel2: TBevel;
    btnDesligar: TBitBtn;
    Configuraes1: TMenuItem;
    NomedaEscola1: TMenuItem;
    lblCaptura: TLabel;
    lblImagemAluno: TLabel;
    Label2: TLabel;
    Label5: TLabel;
    qrAlunosTurmaqtde: TAggregateField;
    DBText1: TDBText;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    N2: TMenuItem;
    CordoTextoNasImagens1: TMenuItem;
    dlgColor: TColorDialog;
    shape: TShape;
    lbl1: TLabel;
    btn2: TSpeedButton;
    VideoWindow: TVideoWindow;
    Bevel1: TBevel;
    N3: TMenuItem;
    CorIndicadoradeImagem1: TMenuItem;
    qrTurmasAnoturma_ano: TWideStringField;
    N4: TMenuItem;
    FotmetroPDFEscola1: TMenuItem;
    N5: TMenuItem;
    CorSombraTextoImagem1: TMenuItem;
    N6: TMenuItem;
    ImpressoraPadro1: TMenuItem;
    N7: TMenuItem;
    Informaes1: TMenuItem;
    CorMolduraCaptura1: TMenuItem;
    N8: TMenuItem;
    Informao1: TMenuItem;
    pnlTitulo: TPanel;
    lbl2: TLabel;
    lbl3: TLabel;
    N9: TMenuItem;
    N10: TMenuItem;
    CopiarAlunosdaTurma1: TMenuItem;
    N11: TMenuItem;
    Carteira1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SnapShotClick(Sender: TObject);
    procedure cbbAnoClick(Sender: TObject);
    procedure dbgTurmasCellClick(Column: TColumn);
    procedure btnLigarDeviceClick(Sender: TObject);
    procedure btnNovaTurmaClick(Sender: TObject);
    procedure btnNovoAlunoTurmaClick(Sender: TObject);
    procedure qrTurmasAnoAfterScroll(DataSet: TDataSet);
    procedure btnGravarImagemClick(Sender: TObject);
    procedure dbgTurmasDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol:
      Integer; Column: TColumn; State: TGridDrawState);
    procedure dbgAlunosDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol:
      Integer; Column: TColumn; State: TGridDrawState);
    procedure dbgAlunosCellClick(Column: TColumn);
    procedure edtPesquisaChange(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnFotometroClick(Sender: TObject);
    procedure qrAlunosTurmaAfterScroll(DataSet: TDataSet);
    procedure chkImagemClick(Sender: TObject);
    procedure NovaTurma1Click(Sender: TObject);
    procedure NovoAluno1Click(Sender: TObject);
    procedure FotmetroPdf1Click(Sender: TObject);
    procedure btnDesligarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure NomedaEscola1Click(Sender: TObject);
    procedure CordoTextoNasImagens1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure image3DblClick(Sender: TObject);
    procedure dbgAlunosDblClick(Sender: TObject);
    procedure CorIndicadoradeImagem1Click(Sender: TObject);
    procedure FotmetroPDFEscola1Click(Sender: TObject);
    procedure CorSombraTextoImagem1Click(Sender: TObject);
    procedure ImpressoraPadro1Click(Sender: TObject);
    procedure CorMolduraCaptura1Click(Sender: TObject);
    procedure Informao1Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure CopiarAlunosdaTurma1Click(Sender: TObject);
    procedure Carteira1Click(Sender: TObject);
  private
    procedure SalvarImagemJPEG(Imagem: TBitmap; const Ano, Turma, NomeAluno: string);
    procedure DesenharAreaRecorte;
    procedure DesenharAreaRecorteTVideo;
    procedure AbrirPasta(const Ano, Turma: string);
    procedure ImprimirAlunosComFotos;
    procedure ImprimirCarteirasEstudantes;
    procedure AtualizarImagemAluno;
    procedure CarregarConfiguracoes;
    procedure SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada: string;
      CorTextoImagem, CorIndicaImagem, CorSombraTexto, CorMolduraCaptura: integer);
    procedure DecodificarEAN13(const Codigo: string; out ID: Integer; out Iniciais: string);
    { Private declarations }
  public
    { Public declarations }
    procedure OnSelectDevice(sender: TObject);
    procedure CopiarCampoParaClipboard(DataSet: TDataSet; const FieldName, Objeto: string);
  end;

var
  fmain: Tfmain;
  SysDev: TSysDevEnum;
  NomeEscola: string;
  CorTextoImagem: Integer;
  CorIndicaImagem: Integer;
  CorSombraTexto: Integer;
  nomeImpressoraDesejada: string;
  CorMolduraCaptura: Integer;

implementation

uses
  uDM, ufNovosRegistros, ufImagemAtual, ufConfirmacao, ufInformacao;

{$R *.dfm}

function MmToPixels(mm: Integer): Integer;
begin
  // Multiplica mm por 10 para trabalhar com 1 casa decimal (25.4 → 254)
  Result := MulDiv(mm * 10, GetDeviceCaps(Printer.Handle, LOGPIXELSX), 254);
end;

function LimparNomeArquivo(const Nome: string): string;
const
  CaracteresInvalidos: array[0..8] of Char = ('\', '/', ':', '*', '?', '"', '<', '>', '|');
var
  C: Char;
  S: string;
begin
  S := Nome.Trim;
  for C in CaracteresInvalidos do
    S := StringReplace(S, C, '_', [rfReplaceAll]);
  Result := S;
end;

procedure Tfmain.AbrirPasta(const Ano, Turma: string);
var
  Caminho: string;
begin
  Caminho := ExtractFilePath(ParamStr(0)) + 'imagens\' + Ano + '\' + Turma;

  if DirectoryExists(Caminho) then
    ShellExecute(0, 'open', PChar(Caminho), nil, nil, SW_SHOWNORMAL)
  else
    MessageDlg('A pasta dessa turma (' + 'imagens\' + Ano + '\' + Turma +
      ') ainda não existe, ela será criada automaticamente.', mtWarning, [mbOK], 0);
end;

procedure Tfmain.AtualizarImagemAluno;
var
  caminho: string;
begin
  image3.Picture := nil;

  if not qrAlunosTurma.IsEmpty then
  begin
    caminho := qrAlunosTurma.FieldByName('path_imagem').AsString;
    if FileExists(caminho) then
      image3.Picture.LoadFromFile(caminho);
  end;
end;

procedure Tfmain.btn1Click(Sender: TObject);
begin
  if (FilterGraph.Active = True) then
  begin
    SampleGrabber.GetBitmap(Image.Picture.Bitmap);
    DesenharAreaRecorte;
  end
  else
    ShowMessage('WebCam desligada.');
end;

procedure Tfmain.btn2Click(Sender: TObject);
begin
  dlgColor.Color := CorTextoImagem; // cor atual como padrão
  if dlgColor.Execute then
  begin
    CorTextoImagem := dlgColor.Color;
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
    shape.Brush.Color := CorTextoImagem;
  end;
end;

procedure Tfmain.btnDesligarClick(Sender: TObject);
begin
  try
    FilterGraph.Stop;
    FilterGraph.ClearGraph;
    FilterGraph.Active := False;

    Image.Picture := nil;

    if (FilterGraph.Active = False) then
    begin
      btnDesligar.Enabled := False;
      btnLigarDevice.Enabled := True;
    end
  except
    on E: Exception do
      ShowMessage('Erro ao desligar o dispositivo: ' + E.Message);
  end;
end;

procedure Tfmain.btnFotometroClick(Sender: TObject);
begin
  ImprimirAlunosComFotos;
end;

procedure Tfmain.btnGravarImagemClick(Sender: TObject);
var
  nomeAluno, nomeLimpo, turma, ano: string;
begin
  if (not qrAlunosTurma.IsEmpty) and (Image.Picture.Graphic <> nil) then
  begin
    try
      nomeAluno := qrAlunosTurmanome.Text;
      nomeLimpo := LimparNomeArquivo(nomeAluno);
      turma := qrTurmasAnonome.Text;
      ano := cbbAno.Text;

      SalvarImagemJPEG(Image.Picture.Bitmap, ano, turma, nomeLimpo);

    except
      on E: Exception do
        MessageDlg('ERRO!!! Imagem NÃO foi gravada. Erro: ' + E.Message, mtError, [mbOK], 0);
    end;

    // Limpa a imagem somente se salvou ou tentou salvar
    Image.Picture := nil;
  end
  else
    ShowMessage('Nenhum aluno selecionado ou imagem não carregada.');
end;

procedure Tfmain.btnLigarDeviceClick(Sender: TObject);
begin
  if cbbDevices.ItemIndex < 0 then
  begin
    ShowMessage('Selecione uma câmera primeiro!');
    Exit;
  end;

  FilterGraph.ClearGraph;
  FilterGraph.Active := False;

  Filter.BaseFilter.Moniker := SysDev.GetMoniker(cbbDevices.ItemIndex);

  FilterGraph.Active := True;
  with FilterGraph as ICaptureGraphBuilder2 do
    RenderStream(@PIN_CATEGORY_PREVIEW, nil, Filter as IBaseFilter, SampleGrabber as
      IBaseFilter, VideoWindow as IBaseFilter);

  FilterGraph.Play;

  if (FilterGraph.Active = True) then
  begin
    btnDesligar.Enabled := True;
    btnLigarDevice.Enabled := False;
  end
end;

procedure Tfmain.btnNovaTurmaClick(Sender: TObject);
begin
  if cbbAno.Text <> '' then
  begin
    if (fNovosRegistros = nil) then
      fNovosRegistros := TfNovosRegistros.Create(Application);

    ufNovosRegistros.IdTurma := 0;
    ufNovosRegistros.Ano := StrToInt(cbbAno.Text);

    fNovosRegistros.pnlTitulo.Caption := 'Nova(s) turma(s), ano ' + cbbAno.Text;

    fNovosRegistros.ShowModal;
  end
  else
    ShowMessage('Ano não definido.');

end;

procedure Tfmain.btnNovoAlunoTurmaClick(Sender: TObject);
begin
  if not qrTurmasAno.IsEmpty then
  begin
    if (fNovosRegistros = nil) then
      fNovosRegistros := TfNovosRegistros.Create(Application);

    ufNovosRegistros.IdTurma := qrTurmasAnoid.Value;
    ufNovosRegistros.Ano := 0;

    fNovosRegistros.pnlTitulo.Caption := 'Novo(s) Aluno(s), turma ' + qrTurmasAnonome.Text
      + '/' + IntToStr(qrTurmasAnoano.Value);

    fNovosRegistros.ShowModal;
  end
  else
    ShowMessage('Turma não definida.');
end;

procedure Tfmain.CarregarConfiguracoes;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(ExtractFilePath(ParamStr(0)) + 'config.ini');
  try
    NomeEscola := Ini.ReadString('Sistema', 'NomeEscola', '');
    CorTextoImagem := Ini.ReadInteger('Sistema', 'CorTextoImagem', 16777215);
    CorIndicaImagem := Ini.ReadInteger('Sistema', 'CorIndicaImagem', 65535);
    CorSombraTexto := Ini.ReadInteger('Sistema', 'CorSombraTexto', 0);
    nomeImpressoraDesejada := Ini.ReadString('Sistema', 'ImpressoraPDF',
      'Microsoft Print to PDF');
    CorMolduraCaptura := Ini.ReadInteger('Sistema', 'CorMolduraCaptura', 32768);
  finally
    Ini.Free;
  end;

  shape.Brush.Color := CorTextoImagem;

end;

procedure Tfmain.Carteira1Click(Sender: TObject);
begin
  ImprimirCarteirasEstudantes;
end;

procedure Tfmain.cbbAnoClick(Sender: TObject);
begin
  qrTurmasAno.Close;
  qrTurmasAno.SQL.Text :=
    'SELECT id, nome || '' / '' || ano AS turma_ano, nome, ano FROM turmas WHERE ano = :ano ORDER BY nome';
  qrTurmasAno.ParamByName('ano').AsInteger := StrToInt(cbbAno.Text);
  qrTurmasAno.Open();

  if qrTurmasAno.IsEmpty then
  begin
    qrAlunosTurma.Close;
  end;
end;

procedure Tfmain.chkImagemClick(Sender: TObject);
begin
  if chkImagem.Checked then
  begin
    if FileExists(qrAlunosTurmapath_imagem.Text) then
    begin
      image3.Picture.LoadFromFile(qrAlunosTurmapath_imagem.Text);
    end
    else
    begin
      image3.Picture := nil;

    end;
  end
  else
    image3.Picture := nil; // limpa se desmarcar o checkbox
end;

procedure Tfmain.CopiarAlunosdaTurma1Click(Sender: TObject);
begin
  CopiarCampoParaClipboard(qrAlunosTurma, 'nome', 'Alunos da turma ' +
    qrTurmasAnonome.Text + '/' + cbbAno.Text + ' copiados com sucesso.');
end;

procedure Tfmain.CopiarCampoParaClipboard(DataSet: TDataSet; const FieldName, objeto: string);
var
  Lista: TStringList;
  Valor: string;
begin
  if not Assigned(DataSet) or not DataSet.Active then
    Exit;

  if DataSet.FindField(FieldName) = nil then
  begin
    ShowMessage('Campo "' + FieldName + '" não encontrado no dataset.');
    Exit;
  end;

  Lista := TStringList.Create;
  try
    DataSet.DisableControls;
    try
      DataSet.First;
      while not DataSet.Eof do
      begin
        Valor := DataSet.FieldByName(FieldName).AsString;
        Lista.Add(Valor);
        DataSet.Next;
      end;
    finally
      DataSet.EnableControls;
    end;
    DataSet.First;

    Clipboard.AsText := Lista.Text; // Copia para o Ctrl+V

    ShowMessage(objeto);
  finally
    Lista.Free;
  end;
end;

procedure Tfmain.CordoTextoNasImagens1Click(Sender: TObject);
begin
  dlgColor.Color := CorTextoImagem; // cor atual como padrão
  if dlgColor.Execute then
  begin
    CorTextoImagem := dlgColor.Color;
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
    shape.Brush.Color := CorTextoImagem;
  end;
end;

procedure Tfmain.CorIndicadoradeImagem1Click(Sender: TObject);
begin
  dlgColor.Color := CorIndicaImagem; // cor atual como padrão
  if dlgColor.Execute then
  begin
    CorIndicaImagem := dlgColor.Color;
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
  end;
end;

procedure Tfmain.CorMolduraCaptura1Click(Sender: TObject);
begin
  dlgColor.Color := CorMolduraCaptura; // cor atual como padrão
  if dlgColor.Execute then
  begin
    CorMolduraCaptura := dlgColor.Color;
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
  end;
end;

procedure Tfmain.CorSombraTextoImagem1Click(Sender: TObject);
begin
  dlgColor.Color := CorSombraTexto; // cor atual como padrão
  if dlgColor.Execute then
  begin
    CorSombraTexto := dlgColor.Color;
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
  end;
end;

procedure Tfmain.dbgAlunosCellClick(Column: TColumn);
var
  path_arquivo: string;
begin
  if (Column.FieldName = 'image') and (not qrAlunosTurma.IsEmpty) then
  begin

    if MessageDlg('Você deseja excluir este aluno ' + qrAlunosTurmanome.Text +
      ', turma ' + qrTurmasAnonome.Text + '/' + IntToStr(qrTurmasAnoano.Value) + '?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin

      fConfirmacao.lblTexto.Caption :=
        'ATENÇÃO: Isso também excluirá a imagem, caso exista, do ALUNO. PARA ATIVAR O BOTÃO SIM DIGITE O "CÓDIGO GERADO" NO CAMPO CONFIRMAÇÃO.';
      fConfirmacao.ShowModal;

      if ufConfirmacao.confirmacao then
      begin
        path_arquivo := ExtractFilePath(ParamStr(0)) + 'imagens\' + IntToStr(qrTurmasAnoano.Value)
          + '\' + qrTurmasAnonome.Text + '\' + qrAlunosTurmanome.Text + '.jpg';

        DM.FDQuery1.Close;
        DM.FDQuery1.SQL.Text := 'DELETE FROM alunos WHERE alunos.id = :id_aluno';
        DM.FDQuery1.ParamByName('id_aluno').AsInteger := qrAlunosTurmaid.Value;
        DM.FDQuery1.ExecSQL;

        qrAlunosTurma.Refresh;

        try
          if TFile.Exists(path_arquivo) then
            TFile.Delete(path_arquivo);
        finally
        end;

        if chkImagem.Checked then
          AtualizarImagemAluno;
      end;
//      else
//        ShowMessage('Cancelado.');

    end;

//    fImagemAtual.Image.Picture.LoadFromFile(qrAlunosTurmapath_imagem.Text);
//    fImagemAtual.Label1.Caption := qrAlunosTurmanome.Text + sLineBreak +
//      qrTurmasAnonome.Text + '/' + IntToStr(qrTurmasAnoano.Value);
//
//    fImagemAtual.ShowModal;

  end;
end;

procedure Tfmain.dbgAlunosDblClick(Sender: TObject);
begin

//  try
//    fImagemAtual.Image.Picture.LoadFromFile(qrAlunosTurmapath_imagem.Text);
//    fImagemAtual.Label1.Caption := qrAlunosTurmanome.Text + sLineBreak +
//      qrTurmasAnonome.Text + '/' + IntToStr(qrTurmasAnoano.Value);
//
//    fImagemAtual.ShowModal;
//  except
//    on E: Exception do
//
//
//  end;
end;

procedure Tfmain.dbgAlunosDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol:
  Integer; Column: TColumn; State: TGridDrawState);
begin

  if ((not qrAlunosTurmapath_imagem.IsNull) and (qrAlunosTurmapath_imagem.Text <> '')) then
  begin
    dbgAlunos.Canvas.Brush.Color := CorIndicaImagem;
    dbgAlunos.DefaultDrawDataCell(Rect, Column.Field, State);
  end;

  if ((Column.FieldName = 'image') and (not qrAlunosTurma.IsEmpty)) then
  begin
    dbgAlunos.Canvas.FillRect(Rect);
    Dm.ilimage3.Draw(dbgAlunos.Canvas, Rect.Left + 8, Rect.Top, 182);
  end;

end;

procedure Tfmain.dbgTurmasCellClick(Column: TColumn);
var
  path_pasta: string;
begin
  if not qrTurmasAno.IsEmpty then
  begin
    qrAlunosTurma.Close;
    qrAlunosTurma.SQL.Text :=
      'SELECT id, id_turma, nome, path_imagem FROM alunos WHERE id_turma = :id_turma ORDER BY nome';
    qrAlunosTurma.ParamByName('id_turma').AsInteger := qrTurmasAnoid.Value;
    qrAlunosTurma.Open;

    if ((Column.FieldName = 'abrir') and (not qrTurmasAno.IsEmpty)) then
    begin
      AbrirPasta(cbbAno.Text, qrTurmasAnonome.Text);
    end;
  end;

  if ((Column.FieldName = 'del') and (not qrTurmasAno.IsEmpty)) then
  begin

    if MessageDlg('Você deseja excluir esta turma ' + qrTurmasAnonome.Text + '/' +
      IntToStr(qrTurmasAnoano.Value) + '?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin

      fConfirmacao.lblTexto.Caption :=
        'ATENÇÃO: Isso também excluirá TODOS os ALUNOS dessa turma e as IMAGENS dos alunos. PARA ATIVAR O BOTÃO SIM DIGITE O "CÓDIGO GERADO" NO CAMPO CONFIRMAÇÃO.';
      fConfirmacao.ShowModal;

      if ufConfirmacao.confirmacao then
      begin
        path_pasta := ExtractFilePath(ParamStr(0)) + 'imagens\' + IntToStr(qrTurmasAnoano.Value)
          + '\' + qrTurmasAnonome.Text;

        DM.FDQuery1.Close;
        DM.FDQuery1.SQL.Text := 'DELETE FROM turmas WHERE turmas.id = :id_turma';
        DM.FDQuery1.ParamByName('id_turma').AsInteger := qrTurmasAnoid.Value;
        DM.FDQuery1.ExecSQL;

        qrTurmasAno.Refresh;
        qrAlunosTurma.Refresh;

        try
          if TDirectory.Exists(path_pasta) then
            TDirectory.Delete(path_pasta, True);
        finally
        end;

        if chkImagem.Checked then
          AtualizarImagemAluno;
      end;
//      else
//        ShowMessage('Cancelado.');

    end;

  end;

end;

procedure Tfmain.dbgTurmasDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol:
  Integer; Column: TColumn; State: TGridDrawState);
begin
  if ((Column.FieldName = 'abrir') and (not qrTurmasAno.IsEmpty)) then
  begin
    dbgTurmas.Canvas.FillRect(Rect);
    Dm.ilimage3.Draw(dbgTurmas.Canvas, Rect.Left + 8, Rect.top, 54)
  end;

  if ((Column.FieldName = 'del') and (not qrTurmasAno.IsEmpty)) then
  begin
    dbgTurmas.Canvas.FillRect(Rect);
    Dm.ilimage3.Draw(dbgTurmas.Canvas, Rect.Left + 8, Rect.top, 182)
  end;
end;

function NumeroParaLetra(Valor: string): Char;
var
  Num: Integer;
begin
  Num := StrToIntDef(Valor, 0);
  if (Num >= 1) and (Num <= 26) then
    Result := Chr(Ord('A') + Num - 1)
  else
    Result := '?'; // caractere inválido
end;

procedure Tfmain.DecodificarEAN13(const Codigo: string; out ID: Integer; out Iniciais: string);
var
  IDStr, L1, L2: string;
begin
  if Length(Codigo) < 13 then
  begin
    ID := -1;
    Iniciais := 'Inválido';
    Exit;
  end;

  IDStr := Copy(Codigo, 5, 4);      // posições 5 a 8 → ID (4 dígitos)
  L1 := Copy(Codigo, 9, 2);         // posições 9 e 10 → primeira inicial
  L2 := Copy(Codigo, 11, 2);        // posições 11 e 12 → segunda inicial

  ID := StrToIntDef(IDStr, -1);
  Iniciais := NumeroParaLetra(L1) + NumeroParaLetra(L2);
end;

procedure Tfmain.DesenharAreaRecorte;
var
  Percentual: Double;
  x1, y1, x2, y2: Integer;
  ImgWidth, ImgHeight: Integer;
begin
  Percentual := 0.5; // 60% da largura da imagem real

  if (Image.Picture.Graphic <> nil) then
  begin
    ImgWidth := Image.Picture.Width;
    ImgHeight := Image.Picture.Height;

    // Calcula as posições com base no tamanho real da imagem
    x1 := (ImgWidth - Round(ImgWidth * Percentual)) div 2;
    y1 := 0;
    x2 := x1 + Round(ImgWidth * Percentual);
    y2 := ImgHeight;

    // Configura caneta
    Image.Canvas.Pen.Color := CorMolduraCaptura;
    Image.Canvas.Pen.Width := 5;
    Image.Canvas.Brush.Style := bsClear;

    // Desenha o retângulo correto
    Image.Canvas.Rectangle(x1, y1, x2, y2);
  end;
end;

procedure Tfmain.DesenharAreaRecorteTVideo;
var
  Percentual: Double;
  x1, y1, x2, y2: Integer;
  ImgWidth, ImgHeight: Integer;
begin
  Percentual := 0.5; // 60% da largura da imagem real

  if (FilterGraph.Active <> False) then
  begin
    ImgWidth := VideoWindow.Width;
    ImgHeight := VideoWindow.Height;

    // Calcula as posições com base no tamanho real da imagem
    x1 := (ImgWidth - Round(ImgWidth * Percentual)) div 2;
    y1 := 0;
    x2 := x1 + Round(ImgWidth * Percentual);
    y2 := ImgHeight;

    // Configura caneta
    VideoWindow.Canvas.Pen.Color := clWhite;
    VideoWindow.Canvas.Pen.Width := 2;
    VideoWindow.Canvas.Brush.Style := bsClear;

    // Desenha o retângulo correto
    VideoWindow.Canvas.Rectangle(x1, y1, x2, y2);
  end;
end;

procedure Tfmain.edtPesquisaChange(Sender: TObject);
begin
  if not qrAlunosTurma.IsEmpty then
  begin
    qrAlunosTurma.Locate('nome', edtPesquisa.Text, [loPartialKey, loCaseInsensitive]);
  end;

//  if not qrAlunosTurma.IsEmpty then
//  begin
//    if edtPesquisa.Text = '' then
//      qrAlunosTurma.Filtered := False
//    else
//    begin
//      qrAlunosTurma.Filter := 'nome LIKE ' + QuotedStr('%' + edtPesquisa.Text + '%');
//      qrAlunosTurma.Filtered := True;
//    end;
//  end;
end;

procedure Tfmain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  SysDev.Free;
  FilterGraph.ClearGraph;
  FilterGraph.Active := false;
end;

procedure Tfmain.FormCreate(Sender: TObject);
var
  i: Integer;
  Ini: TIniFile;
  caminhoINI: string;
begin
  DM.CriarBancoSeNaoExistir;

  SysDev := TSysDevEnum.Create(CLSID_VideoInputDeviceCategory);

  cbbDevices.Items.Clear; // <-- seu ComboBox onde vai listar

  if SysDev.CountFilters > 0 then
  begin
    for i := 0 to SysDev.CountFilters - 1 do
    begin
      cbbDevices.Items.Add(SysDev.Filters[i].FriendlyName);
    end;

    // Opcional: já selecionar a primeira câmera
    //    cbbDevices.ItemIndex := 0;
  end;

  caminhoINI := ExtractFilePath(ParamStr(0)) + 'config.ini';

  Ini := TIniFile.Create(caminhoINI);
  try
    if not Ini.ValueExists('Sistema', 'NomeEscola') then
    begin
      Ini.WriteString('Sistema', 'NomeEscola', 'CEPI Professor Pedro Gomes');
      Ini.WriteInteger('Sistema', 'CorTextoImagem', 16777215); // clWhite
      Ini.WriteInteger('Sistema', 'CorIndicaImagem', 65535); // clYellow
      Ini.WriteInteger('Sistema', 'CorSombraImagem', 0); // clYellow
      Ini.WriteString('Sistema', 'ImpressoraPDF', 'Microsoft Print to PDF');
      Ini.WriteInteger('Sistema', 'CorMolduraCaptura', 32768);
    end;
  finally
    Ini.Free;
  end;

end;

procedure Tfmain.FormShow(Sender: TObject);
var
  anoAtual: string;
  idx: Integer;
begin
  anoAtual := IntToStr(YearOf(Now)); // Pega o ano atual
  idx := cbbAno.Items.IndexOf(anoAtual); // Procura o ano na lista

  if idx >= 0 then
  begin
    cbbAno.ItemIndex := idx;

    // Simula o clique chamando o evento
    if Assigned(cbbAno.OnClick) then
      cbbAno.OnClick(cbbAno);
  end;

  CarregarConfiguracoes;
  fmain.Caption := 'Carômetro ' + NomeEscola;
end;

procedure Tfmain.FotmetroPdf1Click(Sender: TObject);
begin
  ImprimirAlunosComFotos;
end;

procedure Tfmain.FotmetroPDFEscola1Click(Sender: TObject);
var
  Img: TJPEGImage;
  x, y, i, contadorColuna: Integer;
  caminhoImagem, nomeAluno, nomeTurma, linha1, linha2: string;
  larguraImg, alturaImg, margemEsq, margemTopo, espacamentoCol, espacamentoLin,
    textoLargura, maxLargura, posQuebra: Integer;
  nomeLinha1, nomeLinha2: string;
begin
  // === ESCOLHE A IMPRESSORA CERTA ===
  nomeImpressoraDesejada := 'Microsoft Print to PDF'; // <-- Substitua pelo nome EXATO da impressora

  for i := 0 to Printer.Printers.Count - 1 do
  begin
    if SameText(Printer.Printers[i], nomeImpressoraDesejada) then
    begin
      Printer.PrinterIndex := i;
      Break;
    end;
  end;

  if Printer.PrinterIndex = -1 then
  begin
    ShowMessage('Impressora "' + nomeImpressoraDesejada + '" não encontrada.');
    Exit;
  end;

  with DM.FDQuery1 do
  begin
    Close;
    SQL.Text := 'SELECT alunos.id, ' + '       alunos.id_turma, ' +
      '       alunos.nome, ' + '       alunos.path_imagem, ' +
      '       turmas.nome AS nome_turma ' + 'FROM alunos ' +
      'JOIN turmas ON turmas.id = alunos.id_turma ' + 'ORDER BY turmas.nome, alunos.nome';
    Open;
  end;

  if not DM.FDQuery1.Active then
    Exit;

  if DM.FDQuery1.IsEmpty then
    Exit;

  larguraImg := MmToPixels(35);  // 3,5 cm
  alturaImg := MmToPixels(45);  // 4,5 cm
  margemEsq := MmToPixels(15);
  margemTopo := MmToPixels(25);  // espaço maior para título
  espacamentoCol := MmToPixels(45);
  espacamentoLin := MmToPixels(60);  // espaço extra para nomes

  x := margemEsq;
  y := margemTopo;
  contadorColuna := 0;

//  nomeTurma := 'TURMA ' + qrTurmasAnonome.Text + ' - ' + cbbAno.Text;
  Printer.Orientation := poPortrait;
  Printer.BeginDoc;
  try
    // ======= IMPRIME TÍTULO NO TOPO =======
    Printer.Canvas.Font.Size := 14;

    Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth(NomeEscola)
      div 2), MmToPixels(10), NomeEscola);

//    Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth(nomeTurma)
//      div 2), MmToPixels(17), nomeTurma);

    Printer.Canvas.Font.Size := 10;

    // ======= IMPRESSÃO DAS IMAGENS =======
    DM.FDQuery1.First;
    while not DM.FDQuery1.Eof do
    begin
      caminhoImagem := DM.FDQuery1.FieldByName('path_imagem').AsString;
      nomeAluno := DM.FDQuery1.FieldByName('nome').AsString;

      Img := TJPEGImage.Create;
      try
        if (Trim(caminhoImagem) <> '') and FileExists(caminhoImagem) then
          Img.LoadFromFile(caminhoImagem)
        else
          Img.LoadFromFile(IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)))
            + 'no_image.jpg');

        Printer.Canvas.StretchDraw(Rect(x, y, x + larguraImg, y + alturaImg), Img);

          // ======= Nome do aluno com quebra automática =======
        if Printer.Canvas.TextWidth(nomeAluno) > larguraImg then
        begin
          posQuebra := Length(nomeAluno) div 2;
          while (posQuebra > 1) and (nomeAluno[posQuebra] <> ' ') do
            Dec(posQuebra);

          if posQuebra > 1 then
          begin
            nomeLinha1 := Trim(Copy(nomeAluno, 1, posQuebra));
            nomeLinha2 := Trim(Copy(nomeAluno, posQuebra + 1, Length(nomeAluno)));

            Printer.Canvas.TextOut(x + (larguraImg div 2) - (Printer.Canvas.TextWidth
              (nomeLinha1) div 2), y + alturaImg + MmToPixels(2), nomeLinha1);

            Printer.Canvas.TextOut(x + (larguraImg div 2) - (Printer.Canvas.TextWidth
              (nomeLinha2) div 2), y + alturaImg + MmToPixels(2) + Printer.Canvas.TextHeight
              (nomeLinha1), nomeLinha2);
          end
          else
          begin
            textoLargura := Printer.Canvas.TextWidth(nomeAluno);
            Printer.Canvas.TextOut(x + (larguraImg div 2) - (textoLargura div 2), y +
              alturaImg + MmToPixels(2), nomeAluno);
          end;
        end
        else
        begin
          textoLargura := Printer.Canvas.TextWidth(nomeAluno);
          Printer.Canvas.TextOut(x + (larguraImg div 2) - (textoLargura div 2), y +
            alturaImg + MmToPixels(2), nomeAluno);
        end;

      finally
        Img.Free;
      end;

      Inc(x, espacamentoCol);
      Inc(contadorColuna);
      if contadorColuna = 4 then
      begin
        contadorColuna := 0;
        x := margemEsq;
        Inc(y, espacamentoLin);

        if y + alturaImg + MmToPixels(15) > Printer.PageHeight - margemTopo then
        begin
          Printer.NewPage;

            // ======= REPETE TÍTULO =======
//          Printer.Canvas.Font.Size := 14;
//          textoLargura := Printer.Canvas.TextWidth(nomeTurma);
//          if textoLargura > maxLargura then
//          begin
//            if (linha1 <> '') and (linha2 <> '') then
//            begin
//              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth
//                (linha1) div 2), MmToPixels(10), linha1);
//              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth
//                (linha2) div 2), MmToPixels(17), linha2);
//            end
//            else
//              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (textoLargura div 2),
//                MmToPixels(12), nomeTurma);
//          end
//          else
//            Printer.Canvas.TextOut((Printer.PageWidth div 2) - (textoLargura div 2),
//              MmToPixels(12), nomeTurma);

          Printer.Canvas.Font.Size := 10;
          y := margemTopo;
        end;
      end;

      DM.FDQuery1.Next;
    end;
    DM.FDQuery1.First;
  finally
    Printer.EndDoc;
  end;
end;

procedure Tfmain.image3DblClick(Sender: TObject);
begin
  if not image3.Picture.Graphic.Empty then
  begin

    fImagemAtual.Image.Picture.LoadFromFile(qrAlunosTurmapath_imagem.Text);
    fImagemAtual.Label1.Caption := qrAlunosTurmanome.Text + sLineBreak +
      qrTurmasAnonome.Text + '/' + IntToStr(qrTurmasAnoano.Value);

    fImagemAtual.ShowModal;
  end;
end;

procedure Tfmain.ImpressoraPadro1Click(Sender: TObject);
var
  nome: string;
begin
  if InputQuery('Informe o nome (correto) da impressora padrão.', 'Impressora:',
    nomeImpressoraDesejada) then
  begin
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
  end;
end;

procedure Tfmain.ImprimirAlunosComFotos;
var
  Img: TJPEGImage;
  x, y, contadorColuna: Integer;
  caminhoImagem, nomeAluno, nomeTurma, linha1, linha2: string;
  larguraImg, alturaImg, margemEsq, margemTopo, espacamentoCol, espacamentoLin,
    textoLargura, maxLargura, posQuebra: Integer;
  nomeLinha1, nomeLinha2: string;
  i: Integer;
  nomeImpressoraDesejada: string;
begin
  // === ESCOLHE A IMPRESSORA CERTA ===
  nomeImpressoraDesejada := 'Microsoft Print to PDF'; // <-- Substitua pelo nome EXATO da impressora

  for i := 0 to Printer.Printers.Count - 1 do
  begin
    if SameText(Printer.Printers[i], nomeImpressoraDesejada) then
    begin
      Printer.PrinterIndex := i;
      Break;
    end;
  end;

  if Printer.PrinterIndex = -1 then
  begin
    ShowMessage('Impressora "' + nomeImpressoraDesejada + '" não encontrada.');
    Exit;
  end;

  if not qrAlunosTurma.Active then
    Exit;
  if qrAlunosTurma.IsEmpty then
    Exit;

  larguraImg := MmToPixels(35);
  alturaImg := MmToPixels(45);
  margemEsq := MmToPixels(15);
  margemTopo := MmToPixels(25);
  espacamentoCol := MmToPixels(45);
  espacamentoLin := MmToPixels(60);

  x := margemEsq;
  y := margemTopo;
  contadorColuna := 0;

  nomeTurma := 'TURMA ' + qrTurmasAnonome.Text + ' - ' + cbbAno.Text;

  Printer.Orientation := poPortrait;
  Printer.BeginDoc;
  try
    // ======= IMPRIME TÍTULO NO TOPO =======
    Printer.Canvas.Font.Size := 14;

    Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth(NomeEscola)
      div 2), MmToPixels(10), NomeEscola);

    Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth(nomeTurma)
      div 2), MmToPixels(17), nomeTurma);

    Printer.Canvas.Font.Size := 10;

    // ======= IMPRESSÃO DAS IMAGENS =======
    qrAlunosTurma.First;
    while not qrAlunosTurma.Eof do
    begin
      caminhoImagem := qrAlunosTurmapath_imagem.AsString;
      nomeAluno := qrAlunosTurmanome.AsString;

      Img := TJPEGImage.Create;
      try
        if (Trim(caminhoImagem) <> '') and FileExists(caminhoImagem) then
          Img.LoadFromFile(caminhoImagem)
        else
          Img.LoadFromFile(IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)))
            + 'no_image.jpg');

        Printer.Canvas.StretchDraw(Rect(x, y, x + larguraImg, y + alturaImg), Img);

          // ======= Nome do aluno com quebra automática =======
        if Printer.Canvas.TextWidth(nomeAluno) > larguraImg then
        begin
          posQuebra := Length(nomeAluno) div 2;
          while (posQuebra > 1) and (nomeAluno[posQuebra] <> ' ') do
            Dec(posQuebra);

          if posQuebra > 1 then
          begin
            nomeLinha1 := Trim(Copy(nomeAluno, 1, posQuebra));
            nomeLinha2 := Trim(Copy(nomeAluno, posQuebra + 1, Length(nomeAluno)));

            Printer.Canvas.TextOut(x + (larguraImg div 2) - (Printer.Canvas.TextWidth
              (nomeLinha1) div 2), y + alturaImg + MmToPixels(2), nomeLinha1);

            Printer.Canvas.TextOut(x + (larguraImg div 2) - (Printer.Canvas.TextWidth
              (nomeLinha2) div 2), y + alturaImg + MmToPixels(2) + Printer.Canvas.TextHeight
              (nomeLinha1), nomeLinha2);
          end
          else
          begin
            textoLargura := Printer.Canvas.TextWidth(nomeAluno);
            Printer.Canvas.TextOut(x + (larguraImg div 2) - (textoLargura div 2), y +
              alturaImg + MmToPixels(2), nomeAluno);
          end;
        end
        else
        begin
          textoLargura := Printer.Canvas.TextWidth(nomeAluno);
          Printer.Canvas.TextOut(x + (larguraImg div 2) - (textoLargura div 2), y +
            alturaImg + MmToPixels(2), nomeAluno);
        end;

      finally
        Img.Free;
      end;

      Inc(x, espacamentoCol);
      Inc(contadorColuna);
      if contadorColuna = 4 then
      begin
        contadorColuna := 0;
        x := margemEsq;
        Inc(y, espacamentoLin);

        if y + alturaImg + MmToPixels(15) > Printer.PageHeight - margemTopo then
        begin
          Printer.NewPage;

            // ======= REPETE TÍTULO =======
          Printer.Canvas.Font.Size := 14;
          textoLargura := Printer.Canvas.TextWidth(nomeTurma);
          if textoLargura > maxLargura then
          begin
            if (linha1 <> '') and (linha2 <> '') then
            begin
              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth
                (linha1) div 2), MmToPixels(10), linha1);
              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth
                (linha2) div 2), MmToPixels(17), linha2);
            end
            else
              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (textoLargura div 2),
                MmToPixels(12), nomeTurma);
          end
          else
            Printer.Canvas.TextOut((Printer.PageWidth div 2) - (textoLargura div 2),
              MmToPixels(12), nomeTurma);

          Printer.Canvas.Font.Size := 10;
          y := margemTopo;
        end;
      end;

      qrAlunosTurma.Next;
    end;
    qrAlunosTurma.First;
  finally
    Printer.EndDoc;
  end;
end;

function CalcularDigitoVerificadorEAN13(const Codigo12: string): Char;
var
  i, SomaImpar, SomaPar, Total: Integer;
begin
  SomaImpar := 0;
  SomaPar := 0;
  for i := 1 to 12 do
  begin
    if i mod 2 = 0 then
      Inc(SomaPar, StrToInt(Codigo12[i]))
    else
      Inc(SomaImpar, StrToInt(Codigo12[i]));
  end;
  Total := SomaImpar + SomaPar * 3;
  Result := Chr(((10 - (Total mod 10)) mod 10) + Ord('0'));
end;

function LetraParaNumero2(L: Char): string;
begin
  L := UpCase(L);
  if (L >= 'A') and (L <= 'Z') then
    Result := Format('%.2d', [Ord(L) - Ord('A') + 1])
  else
    Result := '00';
end;

function GerarEAN13(ID: Integer; Nome: string): string;
var
  Base, Iniciais, Codigo12: string;
  Palavras: TArray<string>;
begin
  Palavras := Nome.Trim.Split([' ']);
  if Length(Palavras) >= 2 then
    Iniciais := LetraParaNumero2(Palavras[0][1]) + LetraParaNumero2(Palavras[1][1])
  else
    Iniciais := LetraParaNumero2(Palavras[0][1]) + '00';

  Base := Format('%.4d', [ID]) + Iniciais;
  Codigo12 := Base.PadLeft(12, '0'); // Preenche com zeros à esquerda

  Result := Codigo12 + CalcularDigitoVerificadorEAN13(Codigo12);
end;

procedure Tfmain.ImprimirCarteirasEstudantes;
var
  Img: TJPEGImage;
  x, y, contadorColuna: Integer;
  caminhoImagem, nomeAluno, nomeTurma, linha1, linha2: string;
  larguraImg, alturaImg, margemEsq, margemTopo, espacamentoCol, espacamentoLin,
    textoLargura, maxLargura, posQuebra: Integer;
  nomeLinha1, nomeLinha2: string;
  i: Integer;
  nomeImpressoraDesejada: string;
begin
  // === ESCOLHE A IMPRESSORA CERTA ===
  nomeImpressoraDesejada := 'Microsoft Print to PDF'; // <-- Substitua pelo nome EXATO da impressora

  for i := 0 to Printer.Printers.Count - 1 do
  begin
    if SameText(Printer.Printers[i], nomeImpressoraDesejada) then
    begin
      Printer.PrinterIndex := i;
      Break;
    end;
  end;

  if Printer.PrinterIndex = -1 then
  begin
    ShowMessage('Impressora "' + nomeImpressoraDesejada + '" não encontrada.');
    Exit;
  end;

  if not qrAlunosTurma.Active then
    Exit;
  if qrAlunosTurma.IsEmpty then
    Exit;

  larguraImg := MmToPixels(22);
  alturaImg := MmToPixels(30);
  margemEsq := MmToPixels(10);
  margemTopo := MmToPixels(40);
  espacamentoCol := MmToPixels(66);
  espacamentoLin := MmToPixels(50);

  x := margemEsq;
  y := margemTopo;
  contadorColuna := 0;

  nomeTurma := 'TURMA ' + qrTurmasAnonome.Text + ' - ' + cbbAno.Text;

  Printer.Orientation := poPortrait;
  Printer.BeginDoc;
  try
    // ======= IMPRIME TÍTULO NO TOPO =======
    Printer.Canvas.Font.Size := 14;

    Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth(NomeEscola)
      div 2), MmToPixels(10), NomeEscola);

    Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth(nomeTurma)
      div 2), MmToPixels(17), nomeTurma);

    Printer.Canvas.Font.Size := 8;


    // ======= IMPRESSÃO DAS IMAGENS =======
    qrAlunosTurma.First;
    while not qrAlunosTurma.Eof do
    begin


        // Desenha o retângulo ao redor da carteira
      Printer.Canvas.Pen.Width := 1;
      Printer.Canvas.Rectangle(x - MmToPixels(2), // margem interna esquerda
        y - MmToPixels(2), // margem interna superior
        x + MmToPixels(60) + MmToPixels(2), // largura total + margem
        y + MmToPixels(35) + MmToPixels(2)  // altura total + margem
      );

      caminhoImagem := qrAlunosTurmapath_imagem.AsString;
      nomeAluno := qrAlunosTurmanome.AsString;

      Img := TJPEGImage.Create;

      try
        if (Trim(caminhoImagem) <> '') and FileExists(caminhoImagem) then
          Img.LoadFromFile(caminhoImagem)
        else
          Img.LoadFromFile(IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)))
            + 'no_image.jpg');

        Printer.Canvas.StretchDraw(Rect(x, y, x + larguraImg, y + alturaImg), Img);




          // ======= Nome do aluno com quebra automática =======
        if Printer.Canvas.TextWidth(nomeAluno) > larguraImg + MmToPixels(25) then
        begin
          posQuebra := Length(nomeAluno) div 2;
          while (posQuebra > 1) and (nomeAluno[posQuebra] <> ' ') do
            Dec(posQuebra);

          if posQuebra > 1 then
          begin
            nomeLinha1 := Trim(Copy(nomeAluno, 1, posQuebra));
            nomeLinha2 := Trim(Copy(nomeAluno, posQuebra + 1, Length(nomeAluno)));

            Printer.Canvas.TextOut(x + (larguraImg) + MmToPixels(1), y, nomeLinha1);

            Printer.Canvas.TextOut(x + (larguraImg) + MmToPixels(1), y + Printer.Canvas.TextHeight
              (nomeLinha1), nomeLinha2);
          end
          else
          begin
            textoLargura := Printer.Canvas.TextWidth(nomeAluno);
            Printer.Canvas.TextOut(x + (larguraImg) + MmToPixels(1), y, nomeAluno);
          end;
        end
        else
        begin
          textoLargura := Printer.Canvas.TextWidth(nomeAluno);
          Printer.Canvas.TextOut(x + (larguraImg) + MmToPixels(1), y, nomeAluno);
        end;

      finally
        Img.Free;
      end;

      Inc(x, espacamentoCol);
      Inc(contadorColuna);
      if contadorColuna = 3 then
      begin
        contadorColuna := 0;
        x := margemEsq;
        Inc(y, espacamentoLin);

        if y + alturaImg + MmToPixels(15) > Printer.PageHeight - margemTopo then
        begin
          Printer.NewPage;

            // ======= REPETE TÍTULO =======
          Printer.Canvas.Font.Size := 14;
          textoLargura := Printer.Canvas.TextWidth(nomeTurma);
          if textoLargura > maxLargura then
          begin
            if (linha1 <> '') and (linha2 <> '') then
            begin
              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth
                (linha1) div 2), MmToPixels(10), linha1);
              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (Printer.Canvas.TextWidth
                (linha2) div 2), MmToPixels(17), linha2);
            end
            else
              Printer.Canvas.TextOut((Printer.PageWidth div 2) - (textoLargura div 2),
                MmToPixels(12), nomeTurma);
          end
          else
            Printer.Canvas.TextOut((Printer.PageWidth div 2) - (textoLargura div 2),
              MmToPixels(12), nomeTurma);

          Printer.Canvas.Font.Size := 10;
          y := margemTopo;
        end;
      end;

      qrAlunosTurma.Next;
    end;
    qrAlunosTurma.First;
  finally
    Printer.EndDoc;
  end;
end;

procedure Tfmain.Informao1Click(Sender: TObject);
begin
  fInformacao.ShowModal;
end;

procedure Tfmain.N9Click(Sender: TObject);
begin
  CopiarCampoParaClipboard(qrTurmasAno, 'nome', 'Turmas do ano ' + cbbAno.Text +
    ' copiadas com sucesso.');
end;

procedure Tfmain.NomedaEscola1Click(Sender: TObject);
var
  nome: string;
begin
  if InputQuery('Informe o nome da escola', 'Nome:', NomeEscola) then
  begin
    SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada, CorTextoImagem,
      CorIndicaImagem, CorSombraTexto, CorMolduraCaptura);
    fmain.Caption := 'Carômetro Escolar ' + NomeEscola;
  end;
end;

procedure Tfmain.NovaTurma1Click(Sender: TObject);
begin
  btnNovaTurma.Click;
end;

procedure Tfmain.NovoAluno1Click(Sender: TObject);
begin
  btnNovoAlunoTurma.Click;
end;

procedure Tfmain.OnSelectDevice(sender: TObject);
begin
  FilterGraph.ClearGraph;
  FilterGraph.Active := false;
  Filter.BaseFilter.Moniker := SysDev.GetMoniker(TMenuItem(sender).tag);
  FilterGraph.Active := true;
  with FilterGraph as ICaptureGraphBuilder2 do
    RenderStream(@PIN_CATEGORY_PREVIEW, nil, Filter as IBaseFilter, SampleGrabber as
      IBaseFilter, VideoWindow as IbaseFilter);
  FilterGraph.Play;
end;

procedure Tfmain.qrAlunosTurmaAfterScroll(DataSet: TDataSet);
begin
  if chkImagem.Checked then
    AtualizarImagemAluno;
end;

procedure Tfmain.qrTurmasAnoAfterScroll(DataSet: TDataSet);
begin
  if not qrTurmasAno.IsEmpty then
  begin
    qrAlunosTurma.Close;
    qrAlunosTurma.SQL.Text :=
      'SELECT id, id_turma, nome, path_imagem FROM alunos WHERE id_turma = :id_turma ORDER BY nome';
    qrAlunosTurma.ParamByName('id_turma').AsInteger := qrTurmasAnoid.Value;
    qrAlunosTurma.Open;

    if chkImagem.Checked then
      AtualizarImagemAluno;
  end;
end;

procedure Tfmain.SalvarConfiguracoes(NomeEscola, nomeImpressoraDesejada: string;
  CorTextoImagem, CorIndicaImagem, CorSombraTexto, CorMolduraCaptura: integer);
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(ExtractFilePath(ParamStr(0)) + 'config.ini');
  try
    Ini.WriteString('Sistema', 'NomeEscola', NomeEscola);
    Ini.WriteInteger('Sistema', 'CorTextoImagem', CorTextoImagem);
    Ini.WriteInteger('Sistema', 'CorIndicaImagem', CorIndicaImagem);
    Ini.WriteInteger('Sistema', 'CorSombraTexto', CorSombraTexto);
    Ini.WriteString('Sistema', 'ImpressoraPDF', nomeImpressoraDesejada);
    Ini.WriteInteger('Sistema', 'CorMolduraCaptura', CorMolduraCaptura);
  finally
    Ini.Free;
  end;
end;

procedure Tfmain.SalvarImagemJPEG(Imagem: TBitmap; const Ano, Turma, NomeAluno: string);
var
  CaminhoBase, NomeArquivo, Texto, Texto1: string;
  JPEGImage: TJPEGImage;
  BmpTemp, BmpCortada, MarcaDagua: TBitmap;
  x, y, x1, y1, x2, y2: Integer;
  NovaLargura, NovaAltura: Integer;
  XInicio: Integer;
  id_aluno: Integer;
begin
  CaminhoBase := ExtractFilePath(ParamStr(0)) + 'imagens\' + Ano + '\' + Turma + '\';

  if not DirectoryExists(CaminhoBase) then
    ForceDirectories(CaminhoBase);

  NomeArquivo := CaminhoBase + NomeAluno + '.jpg';

  BmpTemp := TBitmap.Create;
  BmpCortada := TBitmap.Create;
  JPEGImage := TJPEGImage.Create;
  try
    // Copia a imagem original
    BmpTemp.Assign(Imagem);

    // Define nova largura (por exemplo, 60% da largura original) e altura total
    NovaLargura := Round(BmpTemp.Width * 0.5); // Cortar 40% (20% de cada lado)
    NovaAltura := BmpTemp.Height; // Mantém altura inteira

    // Define onde começar o corte (para centralizar)
    XInicio := (BmpTemp.Width - NovaLargura) div 2;

    // Prepara o bitmap cortado
    BmpCortada.Width := NovaLargura;
    BmpCortada.Height := NovaAltura;

    // Copia apenas a faixa central
    BmpCortada.Canvas.CopyRect(Rect(0, 0, NovaLargura, NovaAltura), BmpTemp.Canvas,
      Rect(XInicio, 0, XInicio + NovaLargura, NovaAltura));

    // Texto adicional (nome e data)
    Texto1 := NomeAluno;
    Texto := Turma + ' - ' + Ano + '    ' + FormatDateTime('dd/mm/yyyy HH:mm', Now);

    with BmpCortada.Canvas do
    begin
      Font.Name := 'Arial';
      Font.Size := 16;
      Font.Style := [fsBold];
      Brush.Style := bsClear;  // Sem fundo branco

      // Posição para o texto
      x2 := BmpCortada.Width - TextWidth(NomeEscola) - 10;
      y2 := BmpCortada.Height - TextHeight(NomeEscola) - 10;

      x := BmpCortada.Width - TextWidth(Texto) - 10;
      y := BmpCortada.Height - TextHeight(Texto) - 30;

      x1 := BmpCortada.Width - TextWidth(Texto1) - 10;
      y1 := BmpCortada.Height - TextHeight(Texto1) - 50;

      // Desenha sombra preta
      Font.Color := CorSombraTexto;
      TextOut(x + 1, y + 1, Texto);
      TextOut(x1 + 1, y1 + 1, Texto1);
      TextOut(x2 + 1, y2 + 1, NomeEscola);

      // Desenha o texto branco
      Font.Color := CorTextoImagem;
      TextOut(x, y, Texto);
      TextOut(x1, y1, Texto1);
      TextOut(x2, y2, NomeEscola);
    end;

    // Salva
    JPEGImage.Assign(BmpCortada);
    JPEGImage.SaveToFile(NomeArquivo);

    id_aluno := qrAlunosTurmaid.Value;

    if fmain.qrAlunosTurma.Active then
      fmain.qrAlunosTurma.Close;

    DM.FDQuery1.SQL.Clear;
    DM.FDQuery1.SQL.Add('UPDATE alunos SET path_imagem = :path WHERE id = :id');
    DM.FDQuery1.ParamByName('path').AsString := NomeArquivo;
    DM.FDQuery1.ParamByName('id').AsInteger := id_aluno;
    DM.FDQuery1.ExecSQL;

    fmain.qrAlunosTurma.Open;
    fmain.qrAlunosTurma.Locate('id', id_aluno, []);

  finally
    BmpTemp.Free;
    BmpCortada.Free;
    JPEGImage.Free;
  end;
end;

procedure Tfmain.SnapShotClick(Sender: TObject);
begin
  if (FilterGraph.Active = True) then
  begin
    SampleGrabber.GetBitmap(Image.Picture.Bitmap);
    DesenharAreaRecorte;
  end
  else
    ShowMessage('WebCam desligada.');
end;

procedure Tfmain.SpeedButton1Click(Sender: TObject);
begin
  edtPesquisa.Clear;
end;

procedure Tfmain.tmr1Timer(Sender: TObject);
begin
  DesenharAreaRecorteTVideo;
end;

end.

