; =====================================================================
; Ink Canvas 安装脚本（Inno Setup 6.3+）
; ---------------------------------------------------------------------
; 构建方式：Build\build-zips.ps1 在生成暂存目录后自动调用 ISCC 编译本脚本
; 产物：Releases\InkCanvas-v{版本}-Setup.exe（单文件安装包）
;
; 设计要点（改前请读）：
; 1. 【用户级安装】软件把 Settings.json / Log.txt / CustomShapes 等
;    运行期数据写在 exe 同目录。若装到 Program Files，普通权限账号
;    无法写入，设置将保存失败。故默认装到 {localappdata}\Programs，
;    免 UAC、免管理员权限，双击即装（与 Chrome/VS Code 用户安装同模式）。
; 2. 【升级覆盖】AppId 固定不变，重复安装 = 原地升级，无需先卸载。
; 3. 【中文单语言】面向国内教师，只注册简体中文（语言包在
;    Build\InnoLang\ChineseSimplified.isl，来源 kira-96 维护的
;    Inno-Setup-Chinese-Simplified-Translation，MIT，适配 6.5.0+）。
; 4. 【卸载保数据】卸载只删除安装时装入的文件；运行期生成的
;    Settings.json 等留在原目录，用户可自行备份或手动删除。
; =====================================================================

#define MyAppName "Ink Canvas"
#define MyAppVersion "5.2.0"
#define MyAppExeName "Ink Canvas.exe"
#define MyAppPublisher "XueRenYi0"

[Setup]
; 注意：此 GUID 是软件的永久身份标识，升级版本时切勿更改
AppId={{B6E5D2E0-4C8A-4E9B-9F3D-1A2B3C4D5E6F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/XueRenYi0/ink-canvas-dev
AppSupportURL=https://github.com/XueRenYi0/ink-canvas-dev/issues

; 用户级安装：免 UAC 提权（理由见文件头注释 1）
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}

; 安装器图标复用应用图标
SetupIconFile=..\Ink Canvas\Resources\InkCanvas.ico
UninstallDisplayName={#MyAppName} {#MyAppVersion}
UninstallDisplayIcon={app}\{#MyAppExeName}

; GPL-3.0 开源协议
LicenseFile=..\LICENSE

; 现代向导样式 + 单 LZMA2 压缩（体积最小）
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes

; 输出单文件安装包
OutputDir=..\Releases
OutputBaseFilename=InkCanvas-v{#MyAppVersion}-Setup

; 覆盖安装时若软件正在运行，提示用户关闭（走 Windows 重启管理器）
CloseApplications=yes

[Languages]
Name: "chinesesimplified"; MessagesFile: "InnoLang\ChineseSimplified.isl"

[Tasks]
; 桌面快捷方式默认勾选（教师场景刚需），开始菜单入口始终创建
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; 源自 build-zips.ps1 生成的暂存目录（已排除用户数据/调试文件）
Source: "..\Releases\stage-v5\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; —— 快捷方式分两种：带版本号（主入口，残留时一眼可辨）+ 无版本号（兼容入口）
; 旧版（v5.1.0 及之前自制安装器）可能在开始菜单/桌面留有同名快捷方式，
; PrepareToInstall 阶段会自动删除那些指向 Program Files 的残留项
Name: "{group}\{#MyAppName} {#MyAppVersion}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName} {#MyAppVersion}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName} {#MyAppVersion}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; 安装完成后提供"立即运行"勾选项
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

; ==============================================================
; 旧版残留清理（PrepareToInstall 阶段：安装前自动执行）
; --------------------------------------------------------------
; 背景：v5.1.0 及之前使用自制 csc 安装器，安装到 {pf32}\Ink Canvas
; 或 {pf}\Ink Canvas 且未正确写入卸载注册表。升级到 Inno Setup
; 后会出现"快捷方式两套并存、卸载里只显示新版"的现象（本会话
; 本轮真实遇到）。
;
; 策略：
;   1) 检测 Program Files(x86/64) 下是否有旧版 Ink Canvas.exe
;   2) 如存在，删除其在公共开始菜单和公共桌面残留的快捷方式
;   3) 旧安装目录本身可能有 Settings.json 等用户数据，且
;      用户级安装器（PrivilegesRequired=lowest）没有权限删
;      Program Files，故只清快捷方式残留，安装完成后通过 MsgBox
;      提示用户可手动删除
; ==============================================================
[Code]
// 读取 .lnk 快捷方式的目标路径（通过 Shell.Application COM，不依赖未导出的 ResolveShortcut）
function GetShortcutTarget(const LnkPath: String): String;
var
  Shell: Variant;
  Link: Variant;
begin
  Result := '';
  try
    Shell := CreateOleObject('Shell.Application');
    Link := Shell.NameSpace(0).ParseName(LnkPath);
    if not VarIsEmpty(Link) then
    begin
      try
        Result := Link.GetLink.Target.Path;
      except
        Result := '';
      end;
    end;
  except
    Result := '';
  end;
end;

// 删除公共开始菜单/公共桌面里指向旧 Program Files 路径的 Ink Canvas 快捷方式
procedure CleanupLegacyShortcuts;
var
  ProgramDir: String;
  Links: array of String;
  I: Integer;
  Target: String;
begin
  // 旧自制安装器可能装在 32 或 64 位 Program Files
  ProgramDir := ExpandConstant('{pf32}\Ink Canvas\Ink Canvas.exe');
  if not FileExists(ProgramDir) then
  begin
    ProgramDir := ExpandConstant('{pf}\Ink Canvas\Ink Canvas.exe');
    if not FileExists(ProgramDir) then
      Exit; // 无旧版：直接返回
  end;

  // 待扫描路径：公共开始菜单 + 当前用户开始菜单 + 公共桌面 + 当前用户桌面
  SetArrayLength(Links, 4);
  Links[0] := ExpandConstant('{commonprograms}\Ink Canvas.lnk');
  Links[1] := ExpandConstant('{userprograms}\Ink Canvas.lnk');
  Links[2] := ExpandConstant('{commondesktop}\Ink Canvas.lnk');
  Links[3] := ExpandConstant('{userdesktop}\Ink Canvas.lnk');

  for I := 0 to GetArrayLength(Links) - 1 do
  begin
    if FileExists(Links[I]) then
    begin
      Target := GetShortcutTarget(Links[I]);
      if (Pos('Program Files', Target) > 0) or (Pos('Program Files (x86)', Target) > 0) then
      begin
        DeleteFile(Links[I]);
        Log('[LegacyCleanup] 删除残留快捷方式: ' + Links[I] + ' -> ' + Target);
      end;
    end;
  end;
end;

// 安装前：清理旧快捷方式，并提示旧安装目录
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Old32, Old64, OldDir: String;
  Msg: String;
begin
  Result := '';
  CleanupLegacyShortcuts;

  Old32 := ExpandConstant('{pf32}\Ink Canvas\Ink Canvas.exe');
  Old64 := ExpandConstant('{pf}\Ink Canvas\Ink Canvas.exe');
  OldDir := '';
  if FileExists(Old32) then OldDir := ExpandConstant('{pf32}\Ink Canvas')
  else if FileExists(Old64) then OldDir := ExpandConstant('{pf}\Ink Canvas');

  if OldDir <> '' then
  begin
    Msg := #13#10 +
      '检测到系统中还有旧版本的 Ink Canvas，位于：' + #13#10 +
      '  ' + OldDir + #13#10 + #13#10 +
      '旧版本的开始菜单/桌面快捷方式已自动清理。' + #13#10 +
      '如果你不再使用旧版本，可以手动删除上述文件夹。' + #13#10;
    MsgBox(Msg, mbInformation, MB_OK);
  end;
end;
