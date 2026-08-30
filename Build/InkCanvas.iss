; =====================================================================
; Inkboard 安装脚本（Inno Setup 6.3+）
; ---------------------------------------------------------------------
; 构建方式：Build\build-zips.ps1 在生成暂存目录后自动调用 ISCC 编译本脚本
; 产物：Releases\Inkboard-v{版本}-Setup.exe（单文件安装包）
;
; 设计要点（改前请读）：
; 1. 【用户级安装】软件把 Settings.json / Log.txt / CustomShapes 等
;    运行期数据写在 exe 同目录。若装到 Program Files，普通权限账号
;    无法写入，设置将保存失败。故默认装到 {localappdata}\Programs，
;    免 UAC、免管理员权限，双击即装（与 Chrome/VS Code 用户安装同模式）。
; 2. 【升级覆盖】AppId 固定不变，重复安装 = 原地升级，无需先卸载；
;    [InstallDelete] 会在安装前清掉旧版本号的快捷方式残留（换版本号不残留两套）。
; 3. 【中文单语言】面向国内教师，只注册简体中文（语言包在
;    Build\InnoLang\ChineseSimplified.isl，来源 kira-96 维护的
;    Inno-Setup-Chinese-Simplified-Translation，MIT，适配 6.5.0+）。
; 4. 【卸载保数据】卸载只删除安装时装入的文件；运行期生成的
;    Settings.json 等留在原目录，用户可自行备份或手动删除。
; 5. 【v6.0.0 更名迁移】软件由 Ink Canvas 更名为 Inkboard：换用新 AppId
;    + 新安装目录，PrepareToInstall 阶段自动把旧目录（{localappdata}\Programs
;    \Ink Canvas）里的 Settings.json / custom.json / CustomShapes 图库迁移
;    到新目录，并清理旧版快捷方式、旧卸载注册表项、旧安装目录。
; =====================================================================

#define MyAppName "Inkboard"
#define MyAppVersion "6.0.0"
#define MyAppExeName "Inkboard.exe"
#define MyAppPublisher "XueRenYi0"
; 旧名（v5.x 及之前），仅在迁移清理代码中使用
#define OldAppName "Ink Canvas"

[Setup]
; 注意：此 GUID 是软件的永久身份标识，升级版本时切勿更改
; （v6.0.0 更名时换新 ID 是刻意的：与旧 Ink Canvas 安装解耦，走迁移而非原地升级）
AppId={{7C1F9A4E-2D35-4B6A-8E77-9F0A1B2C3D4E}
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
OutputBaseFilename=Inkboard-v{#MyAppVersion}-Setup

; 覆盖安装时若软件正在运行，提示用户关闭（走 Windows 重启管理器）
CloseApplications=yes

[Languages]
Name: "chinesesimplified"; MessagesFile: "InnoLang\ChineseSimplified.isl"

[Tasks]
; 桌面快捷方式默认勾选（教师场景刚需），开始菜单入口始终创建
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[InstallDelete]
; 版本迭代防残留：安装前删掉旧版本号的快捷方式（旧脚本带版本号命名，
; 升级后旧 lnk 会留着，这里按通配符统一清掉再重建）
Type: files; Name: "{group}\{#OldAppName}*.lnk"
Type: files; Name: "{group}\{#MyAppName}*.lnk"
Type: files; Name: "{autodesktop}\{#OldAppName}*.lnk"
Type: files; Name: "{autodesktop}\{#MyAppName}*.lnk"
; v6.0.0 首包事故自愈：首个 6.0.0 安装包把文件误装进了 {app}\Inkboard v6.0.0\
; 嵌套目录（快捷方式因此找不到 exe）。这里把那个嵌套目录整个清掉，
; 已装错的老用户重跑修复包后自动恢复，不残留 60MB 垃圾
Type: filesandordirs; Name: "{app}\Inkboard v6.0.0"

[Files]
; 源自 build-zips.ps1 生成的暂存目录（已排除用户数据/调试文件）。
; 注意暂存结构是 stage-v6\Inkboard v{版本}\（zip 打包需要这层文件夹），
; 安装包必须穿透这层直接取内容，否则 exe 会装到 {app}\Inkboard v{版本}\ 里，
; 快捷方式指向 {app}\Inkboard.exe 找不到文件 → "安装后打不开"（v6.0.0 首包真实踩过）
Source: "..\Releases\stage-v6\Inkboard v{#MyAppVersion}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; —— 主入口不带版本号（版本号在"应用和功能"里看），避免每次升级残留一套旧快捷方式
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; 安装完成后提供"立即运行"勾选项
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

; ==============================================================
; 安装前处理（PrepareToInstall 阶段：安装向导点击"安装"后、复制文件前执行）
; --------------------------------------------------------------
; A. v6.0.0 更名迁移（旧 Ink Canvas 用户级安装 → 新 Inkboard 目录）
; B. 旧自制安装器残留清理（v5.1.0 及之前装在 Program Files 的版本）
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
// （v5.1.0 及之前的自制 csc 安装器装的，无正规卸载项）
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

// 整目录递归复制（Inno 无内置递归 copy，简单实现：先 Files 再递归子目录）
// 注意：Inno Pascal 要求函数先定义后使用，故放在 MigrateFromOldInkCanvas 之前
function CopyDirRecursive(const SrcDir, DstDir: String): Boolean;
var
  FindRec: TFindRec;
  SrcPath, DstPath: String;
begin
  // Inno Pascal 不支持 Exit(值) 语法，用 Result 赋值 + 普通 Exit 实现"失败即返回"
  Result := ForceDirectories(DstDir);
  if not Result then Exit;
  if FindFirst(SrcDir + '\*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          SrcPath := SrcDir + '\' + FindRec.Name;
          DstPath := DstDir + '\' + FindRec.Name;
          if FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0 then
          begin
            if not CopyDirRecursive(SrcPath, DstPath) then Result := False;
          end
          else
          begin
            if not FileCopy(SrcPath, DstPath, False) then Result := False;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

// 删除当前用户开始菜单/桌面上指向 OldDir 的 "Ink Canvas" 快捷方式
procedure DeleteLegacyUserShortcuts(const OldDir: String);
var
  Links: array of String;
  I: Integer;
  Target: String;
begin
  SetArrayLength(Links, 2);
  Links[0] := ExpandConstant('{userprograms}\Ink Canvas.lnk');
  Links[1] := ExpandConstant('{userdesktop}\Ink Canvas.lnk');
  for I := 0 to GetArrayLength(Links) - 1 do
  begin
    if FileExists(Links[I]) then
    begin
      Target := GetShortcutTarget(Links[I]);
      // 目标指向旧安装目录（或读不到目标也删：名字匹配的只剩它）
      if (Pos(Lowercase(OldDir), Lowercase(Target)) > 0) or (Target = '') then
      begin
        DeleteFile(Links[I]);
        Log('[Migrate] 删除旧快捷方式: ' + Links[I]);
      end;
    end;
  end;
end;

// v6.0.0 更名迁移：把旧 Ink Canvas 用户级安装的数据搬到新 Inkboard 目录，
// 并清理旧快捷方式 / 旧卸载注册表 / 旧安装目录。
// 用户级安装（{localappdata}）本账号有完整权限，可放心整目录删除。
procedure MigrateFromOldInkCanvas;
var
  OldDir, NewDir: String;
  Files: array of String;
  I: Integer;
  OldUninstKey: String;
begin
  OldDir := ExpandConstant('{localappdata}\Programs\Ink Canvas');
  NewDir := ExpandConstant('{localappdata}\Programs\Inkboard');

  // 没装过旧版（或已迁移过）：无事可做
  if not FileExists(OldDir + '\Ink Canvas.exe') then Exit;

  // 1) 迁移用户数据：设置 + 图库 + 名单（不存在就跳过）
  SetArrayLength(Files, 3);
  Files[0] := 'Settings.json';
  Files[1] := 'custom.json';
  Files[2] := 'Versions.ini';
  for I := 0 to GetArrayLength(Files) - 1 do
  begin
    if FileExists(OldDir + '\' + Files[I]) then
    begin
      ForceDirectories(NewDir);
      if FileCopy(OldDir + '\' + Files[I], NewDir + '\' + Files[I], False) then
        Log('[Migrate] 迁移配置: ' + Files[I]);
    end;
  end;

  // 2) 迁移自定义图库目录（整目录复制，含全部 .isc 文件）
  if DirExists(OldDir + '\CustomShapes') then
  begin
    ForceDirectories(NewDir + '\CustomShapes');
    // 复制所有 .isc（自制墨迹图形）——旧库是纯数据，直接全量搬
    if not CopyDirRecursive(OldDir + '\CustomShapes', NewDir + '\CustomShapes') then
      Log('[Migrate] WARN: CustomShapes 迁移失败（旧目录保留，可手动拷贝）');
  end;

  // 3) 删除旧版快捷方式（用户开始菜单 + 用户桌面，指向旧目录的）
  DeleteLegacyUserShortcuts(OldDir);

  // 4) 删除旧版卸载注册表项（用户级安装在 HKCU；旧 AppId = v5.2.0 脚本所用 GUID）
  OldUninstKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B6E5D2E0-4C8A-4E9B-9F3D-1A2B3C4D5E6F}_is1';
  if RegKeyExists(HKEY_CURRENT_USER, OldUninstKey) then
  begin
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, OldUninstKey);
    Log('[Migrate] 删除旧版卸载注册表项');
  end;

  // 5) 删除旧安装目录（配置已迁移走；万一有没搬到的文件，放弃删除并记录日志）
  // DelTree 签名：(Path, IsDir, DeleteFiles, DeleteSubdirsAlso)
  if not DelTree(OldDir, True, True, True) then
    Log('[Migrate] WARN: 旧目录删除失败（可能有文件被占用），请手动删除: ' + OldDir);

  Log('[Migrate] Ink Canvas → Inkboard 迁移完成');
end;

// 安装前：更名迁移 + 清理旧自制安装器残留并提示
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Old32, Old64, OldDir: String;
  Msg: String;
begin
  Result := '';

  // v6.0.0 更名：先迁移旧 Ink Canvas 用户级安装
  MigrateFromOldInkCanvas;

  // 再处理 v5.1.0 及之前 Program Files 里的自制安装器残留
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
