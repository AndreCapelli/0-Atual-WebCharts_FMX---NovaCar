program MaxCharts;

uses
  System.StartUpCopy,
  FMX.Forms,
  UCEFApplication,
  uCEFConstants,
  unitPrincipal in 'unitPrincipal.pas' {unitHome} ,
  NovaConexao in 'Util\NovaConexao.pas',
  unitDmDados in 'Util\unitDmDados.pas' {dmDados: TDataModule} ,
  unitFuncoes in 'Util\unitFuncoes.pas',
  unit_esmaecer_fundo in 'unit_esmaecer_fundo.pas' {form_esmaecer_fundo} ,
  uLogin in 'uLogin.pas' {frmLogin};

{$R *.res}

begin
  GlobalCEFApp := TCefApplication.Create;

  GlobalCEFApp.FrameworkDirPath := 'cefCharts';
  GlobalCEFApp.ResourcesDirPath := 'cefCharts';
  GlobalCEFApp.LocalesDirPath := 'cefCharts\locales';

  GlobalCEFApp.AllowRunningInsecureContent := true;

  GlobalCEFApp.Locale := 'pt-BR';
  GlobalCEFApp.AcceptLanguageList := 'pt-BR,pt-BR;q=0.9,en-US;q=0.8,en;q=0.7';
  // GlobalCEFApp.LogFile := '\debug.log';
  // GlobalCEFApp.LogSeverity := LOGSEVERITY_DEBUG;

  if GlobalCEFApp.StartMainProcess then
  begin
    Application.Initialize;
    Application.CreateForm(TunitHome, unitHome);
    Application.CreateForm(Tform_esmaecer_fundo, form_esmaecer_fundo);
    Application.Run;
  end;
  GlobalCEFApp.Free;
  GlobalCEFApp := nil;

end.
