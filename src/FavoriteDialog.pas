(*//
标题：调用IE收藏夹相关对话框
说明：目前就封装添加IE收藏夹对话框，整理收藏夹就留给ie自己了。
设计：ZswangY37
日期：2008-03-18
支持：wjhu111#21cn.com
//*)

unit FavoriteDialog;

interface

uses
  Windows, Messages, SysUtils, Classes, Dialogs, ActiveX, ShlObj;

type
  TAddFavoriteDialog = class(TCommonDialog)
  private
    { Private declarations }
    FDirectory: string;
    FDisplayName: string;
    FWindowHandle: THandle;
  protected
    { Protected declarations }
  public
    { Public declarations }
    property WindowHandle: THandle read FWindowHandle write FWindowHandle;
    function Execute: Boolean; override;
  published
    { Published declarations }
    property DisplayName: string read FDisplayName write FDisplayName;
    property Directory: string read FDirectory write FDirectory;
  end;

function DoAddToFavDlg(Wnd: HWND; szPath: PChar; cszPath: Integer;
  szTitle: PChar; cszTitle: Integer; pidlFavorites: PItemIDList): BOOL; stdcall;

function CreateUrlShortcut( // 创建链接快捷方式
  AUrl: string; // Url字符串
  AFileName: WideString // 文件名
): Boolean; // 返回创建文件是否成功

procedure Register;

implementation

uses ComObj, SHDocVw;

function DoAddToFavDlg; stdcall; external 'shdocvw.dll';

procedure Register;
begin
  RegisterComponents('Zswang', [TAddFavoriteDialog]);
end;

const
  CLSID_InternetShortcut: TGUID = (D1:$FBF23B40; D2:$E3F0; D3:$101B; D4:($84, $88, $00, $AA, $00, $3E, $56, $F8));

  IID_IUniformResourceLocatorA: TGUID = (D1:$FBF23B80; D2:$E3F0; D3:$101B; D4:($84, $88, $00, $AA, $00, $3E, $56, $F8));
  IID_IUniformResourceLocatorW: TGUID = (D1:$CABB0DA0; D2:$DA57; D3:$11CF; D4:($99, $74, $00, $20, $AF, $D7, $97, $62));

  {$IFDEF UNICODE}
  IID_IUniformResourceLocator: TGUID = (D1:$CABB0DA0; D2:$DA57; D3:$11CF; D4:($99, $74, $00, $20, $AF, $D7, $97, $62));
  {$ELSE}
  IID_IUniformResourceLocator: TGUID = (D1:$FBF23B80; D2:$E3F0; D3:$101B; D4:($84, $88, $00, $AA, $00, $3E, $56, $F8));
  {$ENDIF UNICODE}

type
  PUrlInvokeCommandInfoA = ^TUrlInvokeCommandInfoA;

  TUrlInvokeCommandInfoA = record
    dwcbSize: DWORD;
    dwFlags: DWORD;
    hwndParent: HWND;
    pcszVerb: LPCSTR;
  end;

  PUrlInvokeCommandInfoW = ^TUrlInvokeCommandInfoW;
  TUrlInvokeCommandInfoW = record
    dwcbSize: DWORD;
    dwFlags: DWORD;
    hwndParent: HWND;
    pcszVerb: LPCWSTR;
  end;

  IUniformResourceLocatorA = interface(IUnknown)
    function SetURL(pcszURL: LpcStr; dwInFlags: DWORD): HRESULT; stdcall;
    function GetURL(ppszURL: LpStr): HRESULT; stdcall;
    function InvokeCommand(purlici: PURLINVOKECOMMANDINFOA): HRESULT; stdcall;
  end;

  IUniformResourceLocatorW = interface(IUnknown)
    function SetURL(pcszURL: LpcWStr; dwInFlags: DWORD): HRESULT; stdcall;
    function GetURL(ppszURL: LpWStr): HRESULT; stdcall;
    function InvokeCommand(purlici: PURLINVOKECOMMANDINFOW): HRESULT; stdcall;
  end;
  
  IUniformResourceLocator = IUniformResourceLocatorA;

function CreateUrlShortcut( // 创建链接快捷方式
  AUrl: string; // Url字符串
  AFileName: WideString // 文件名
): Boolean; // 返回创建文件是否成功
var
  vUniformResourceLocator: IUniformResourceLocatorA;
  vPersistFile: IPersistfile;
begin
  Result := False;
  CoCreateInstance(CLSID_InternetShortcut, nil, CLSCTX_INPROC_SERVER,
    IID_IUniformResourceLocator, vUniformResourceLocator);
  if not Assigned(vUniformResourceLocator) then Exit;
  vUniformResourceLocator.SetURL(PChar(AUrl), 0);
  vPersistFile := vUniformResourceLocator as IPersistfile;
  if not Assigned(vPersistFile) then
  begin
    vUniformResourceLocator := nil;
    Exit;
  end;
  Result := Succeeded(vPersistFile.Save(PWideChar(AFileName), True));
  vUniformResourceLocator := nil;
  vPersistFile := nil;
end;

function DesktopShellFolder: IShellFolder;
begin
  OleCheck(SHGetDesktopFolder(Result));
end;

{ TAddFavoriteDialog }

type
  TDialogData = packed record
    rWindowHandle: THandle;
    rDirectory: array[0..MAX_PATH] of Char;
    rDisplayName: array[0..MAX_PATH] of Char;
    rItemIDList: PItemIDList;
  end;
  PDialogData = ^TDialogData;

function DoExecute(var ADialogData: TDialogData): Bool; stdcall;
begin
  Result := DoAddToFavDlg(ADialogData.rWindowHandle, @ADialogData.rDirectory[0],
    MAX_PATH, @ADialogData.rDisplayName[0], MAX_PATH, ADialogData.rItemIDList);
end;

function TAddFavoriteDialog.Execute: Boolean;
var
  Malloc: IMalloc;
  vAttributes: Longword;
  vEaten: Longword;
  vDialogData: TDialogData;
begin
  Result := False;
  if CoGetMalloc(1, Malloc) <> S_OK then Exit;
  FillChar(vDialogData, SizeOf(vDialogData), 0);
  vDialogData.rWindowHandle := FWindowHandle;
  if FDirectory = EmptyStr then
    SHGetSpecialFolderLocation(FWindowHandle, CSIDL_FAVORITES,
      vDialogData.rItemIDList)
  else DesktopShellFolder.ParseDisplayName(FWindowHandle, nil,
    PWideChar(WideString(FDirectory)), vEaten, vDialogData.rItemIDList,
    vAttributes);
  if not Assigned(vDialogData.rItemIDList) then Exit;
  if FDisplayName <> EmptyStr then
    Move(FDisplayName[1], vDialogData.rDisplayName[0], Length(FDisplayName));
  if FDirectory <> EmptyStr then
    Move(FDirectory[1], vDialogData.rDirectory[0], Length(FDirectory));
  Result := TaskModalDialog(@DoExecute, vDialogData);
  Malloc.Free(vDialogData.rItemIDList);
  FDirectory := ExtractFileDir(vDialogData.rDirectory);
  FDisplayName := vDialogData.rDisplayName;
  Malloc := nil;
end;

end.
