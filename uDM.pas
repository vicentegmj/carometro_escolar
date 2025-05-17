unit uDM;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Comp.Client, FireDAC.UI.Intf,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Phys.SQLite, FireDAC.Phys.SQLiteDef, FireDAC.Stan.Def, FireDAC.Stan.Pool,
  FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Comp.UI, Data.DB, FireDAC.Stan.ExprFuncs,
  FireDAC.DApt, FireDAC.Comp.DataSet, System.ImageList, Vcl.ImgList, Vcl.Controls;

type
  TDM = class(TDataModule)
    FDConnection1: TFDConnection;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    FDQuery1: TFDQuery;
    ilimage3: TImageList;
    ImageList1: TImageList;
    procedure DataModuleCreate(Sender: TObject);
    procedure FDConnection1AfterConnect(Sender: TObject);
  private
  public
    procedure CriarBancoSeNaoExistir;
    function ListarTurmasPorAno(AAno: Integer): TFDQuery;
    function ListarAlunosPorTurma(AIdTurma: Integer): TFDQuery;
  end;

var
  DM: TDM;

implementation

uses
  System.IOUtils;

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

procedure TDM.CriarBancoSeNaoExistir;
var
  DatabasePath: string;
  BancoNovo: Boolean;
begin
  DatabasePath := TPath.Combine(ExtractFilePath(ParamStr(0)), 'database.db');
  BancoNovo := not FileExists(DatabasePath);

  // Configura conexão
  FDConnection1.DriverName := 'SQLite';
  FDConnection1.Params.Database := DatabasePath;
  FDConnection1.Params.Add('LockingMode=Normal');
  FDConnection1.Params.Add('Synchronous=Normal');
  FDConnection1.LoginPrompt := False;

  // Conecta (isso pode criar o arquivo se não existir)
  FDConnection1.Connected := True;

  // Ativa suporte a chave estrangeira
  FDConnection1.ExecSQL('PRAGMA foreign_keys = ON');

  // Se for novo, cria as tabelas
  if BancoNovo then
  begin
    FDConnection1.ExecSQL(
      'CREATE TABLE turmas (' +
      ' id INTEGER PRIMARY KEY AUTOINCREMENT,' +
      ' nome VARCHAR(100),' +
      ' ano INTEGER NOT NULL' +
      ');'
    );

    FDConnection1.ExecSQL(
      'CREATE TABLE alunos (' +
      ' id INTEGER PRIMARY KEY AUTOINCREMENT,' +
      ' id_turma INTEGER NOT NULL,' +
      ' nome VARCHAR(100),' +
      ' path_imagem VARCHAR(255),' +
      ' FOREIGN KEY (id_turma) REFERENCES turmas(id) ON DELETE CASCADE' +
      ');'
    );
  end;
end;



procedure TDM.DataModuleCreate(Sender: TObject);
begin
  DM.FDConnection1.TxOptions.AutoCommit := True;
end;

procedure TDM.FDConnection1AfterConnect(Sender: TObject);
begin
  DM.FDConnection1.ExecSQL('PRAGMA foreign_keys = ON');
end;

function TDM.ListarAlunosPorTurma(AIdTurma: Integer): TFDQuery;
begin
  Result := TFDQuery.Create(nil);
  Result.Connection := FDConnection1;
  Result.SQL.Text :=
    'SELECT id, id_turma, nome FROM alunos WHERE id_turma = :id_turma ORDER BY nome';
  Result.ParamByName('id_turma').AsInteger := AIdTurma;
  Result.Open;
end;

function TDM.ListarTurmasPorAno(AAno: Integer): TFDQuery;
begin
  Result := TFDQuery.Create(nil);
  Result.Connection := FDConnection1;
  Result.SQL.Text := 'SELECT id, nome, ano FROM turmas WHERE ano = :ano ORDER BY nome';
  Result.ParamByName('ano').AsInteger := AAno;
  Result.Open;
end;

end.

