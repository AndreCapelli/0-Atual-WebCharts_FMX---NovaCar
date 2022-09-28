unit NovaConexao;

interface

uses Tlhelp32, SysUtils, DB, Winapi.Windows, ShellApi,
  Vcl.Controls,
  Vcl.DBCtrls, Vcl.ExtCtrls, Vcl.Graphics, Vcl.Forms, System.Classes, Variants,
  FireDAC.Comp.Client, FireDAC.Phys.ADS,
  FireDAC.Phys.MSSQL, Vcl.Dialogs, FMX.Dialogs;

procedure ExecutaStrCon();
procedure BuildConnection(pSistema: String);
function ProcessoExiste(ExeFileName: string): boolean;
procedure StartConnection(pConexao: TFDConnection;
  pDriverLink: TFDPhysMSSQLDriverLink);

var
  vSistemaGlobal: String;

var
  ConexaoSQL: String;

implementation

uses unitDmDados, unitFuncoes, unitPrincipal;

function ProcessoExiste(ExeFileName: string): boolean;
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32 { declarar Uses Tlhelp32 };
begin
  result := false;
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile))
      = UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile)
      = UpperCase(ExeFileName))) then
    begin
      result := true;
      exit;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

procedure BuildConnection(pSistema: String);
var
  Caminho: string;
begin
  vSistemaGlobal := pSistema;

//  if not FileExists(ExtractFilePath(Application.ExeName) + 'strCon\strCon' +
//    vSistemaGlobal + '.ini') then
//  begin
//     if MsgYesNo('Arquivo de conexão não encontrado!' + #13 +
//      'Deseja selecionar o arquivo?') = true then
//    begin
//       unitHome.OpenDialog1 := TOpenDialog.Create(unitHome.OpenDialog1);
//      if unitHome.OpenDialog1.Execute then
//      Begin
//        Caminho := unitHome.OpenDialog1.FileName;
//        BD_NovoCaminho := Caminho;
//      End;
//      unitHome.OpenDialog1.Free;
//    end;
//
//  end;
end;

procedure ExecutaStrCon();
var
  TheWindow: HWND;
begin
  TheWindow := FindWindow(nil, 'ORIGEM - String de Conexão');
  if TheWindow <> 0 then // Chama String de Conexão se já estiver carregada
    SetForegroundWindow(TheWindow)
  else // Carrega String de Conexão se estiver fechada
    ShellExecute(0, 'open', PWIDECHAR(ExtractFilePath(Application.ExeName) +
      'strCon\StrCon.exe'), '', '', SW_SHOWNORMAL);
end;

procedure StartConnection(pConexao: TFDConnection;
  pDriverLink: TFDPhysMSSQLDriverLink);
// Z94T1259
Var
  vListaBD: TStringList;
  vParams: TStringList;

  function CriaString(): TStringList;
  begin
    vParams := TStringList.Create;
    vListaBD := TStringList.Create;
    Try
       if BD_NovoCaminho = '' then
        vListaBD.LoadFromFile(ExtractFilePath(Application.ExeName) +
          'strCon\strCon' + vSistemaGlobal + '.ini')
      else
        vListaBD.LoadFromFile(BD_NovoCaminho);
    Except
      vListaBD.Free;
      MessageDlg('O Arquivo de Conexão não foi encontrado!', mtError,
        [mbOK], 0);
      Application.Terminate;
    End;
    if vListaBD[0] = 'ADS' then // ADS
    begin
      vParams.Clear;
      vParams.Add('Alias=' + vListaBD[4]);
      vParams.Add('User_Name=' + (vListaBD[1]));
      vParams.Add('Password=' + (vListaBD[2]));
      vParams.Add('ServerTypes=' + vListaBD[3]);
      vParams.Add('DriverID=' + vListaBD[0]);
      // Seu componente TFDPhysADSDriverLink
      pDriverLink.DriverID := vListaBD[0];
    end;
    if vListaBD[0] = 'MSSQL' then // ADS
    begin
      ConexaoSQL := 'Server=' + dmdados.Descript(vListaBD[5]) + #$D#$A +
        'Database=' + vListaBD[4] + #$D#$A + 'DriverID=' + vListaBD[0] + #$D#$A
        + 'User_Name=' + dmdados.Descript(vListaBD[1]) + #$D#$A + 'Password=' +
        dmdados.Descript(vListaBD[2]);
      pDriverLink.DriverID := vListaBD[0];

      dmdados.GB_Servidor := dmdados.Descript(vListaBD[5]);
      dmdados.GB_BancoDeDados := vListaBD[4];
      dmdados.BD_Usuario := dmdados.Descript(vListaBD[1]);
      dmdados.BD_senha := dmdados.Descript(vListaBD[2]);
      dmdados.BD_servidor := dmdados.Descript(vListaBD[5]);
      dmdados.BD_Base := vListaBD[4];
    end;
    result := vParams;
  end;

begin

  pConexao.LoginPrompt := false;
  pConexao.Connected := false;
  CriaString();

  pConexao.Params.text := ConexaoSQL;

  try
    pConexao.Connected := true;
  Except
    ShowMessage('Não foi possível conectar ao banco de dados!' + #13 +
      'Verifique as credenciais de acesso e se o serviço de banco de dados está ativo no servidor!');
    pConexao.Connected := false;
    Application.Terminate;
  End;
End;

end.
