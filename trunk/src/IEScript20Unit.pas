(*//
���⣺IE�ű���������
˵����Ϊ������IE�е��Խű���ƣ�Ҳ����Firefox�е��ԣ�
��ƣ�ZswangY37
���ڣ�2008-03-10
֧�֣�wjhu111#21cn.com
//*)

//�޸���־
//2008-03-10 ZswangY37 No.1 �½�
//2008-03-13 ZswangY37 No.1 ���� 1.00.001
//2008-03-14 ZswangY37 No.1 ���� 1.02.001
//2008-03-14 ZswangY37 No.2 ���� ���ű��б��޸�Ϊ��״�����̶��ű�Ŀ¼�����ڵ�ǰĿ¼����
//2008-03-14 ZswangY37 No.3 ��� ���ƾ���ű��Ĺ��ܣ����㵽IE��FF��ַ������
//2008-03-14 ZswangY37 No.4 ��� ����ű��Ĺ���
//2008-03-19 ZswangY37 No.1 ��� ���浽IE�ղؼеĹ���
//2008-03-21 ZswangY37 No.1 ��� �������ļ��Ĺ���
//2008-03-21 ZswangY37 No.2 ��� �ļ��϶��Ĺ���
//2008-03-24 Zswangy37 No.1 ��� ɾ���ű���Ŀ¼�Ĺ���
//2008-03-25 Zswangy37 No.1 ���� �����Ŀ¼�Ĺ��ܷ�Ϊ�����Ŀ¼�����ͬ��Ŀ¼
//2008-05-22 ZswangY37 No.1 ��� �ڱ༭���в��Һ��滻�Ĺ���
//2008-06-11 ZswangY37 No.1 ���� �򵥼��ܸ��ƴ������ı���ʧЧ������
//------------------------------------------------------------------------------1.02.005
//2008-11-15 ZswangY37 No.1 ��� ����ִ�нű�
//2008-11-15 ZswangY37 No.2 ���� �滻ΪRichEdit
//------------------------------------------------------------------------------1.02.007
//2008-12-14 ZswangY37 No.1 ��� �ڵ��л�ʱ���ű��ı䱣����ʾ
//------------------------------------------------------------------------------1.02.008
//2009-12-20 ZswangY37 No.1 ��� ����IE�����Ϊ��
//2009-12-20 ZswangY37 No.2 ��� Tab���Ű洦��
//------------------------------------------------------------------------------1.02.008
//2009-12-23 ZswangY37 No.1 ��� �ļ�����Ĵ�������Unicode��Utf8����BOM��Utf8�ļ�
//------------------------------------------------------------------------------1.02.008
//2010-12-21 ZswangY37 No.1 ���� ����Ҳ��Ĭ�ϱ༭�ļ�
//2010-12-21 ZswangY37 No.2 ���� �����BOM Utf8�˵�

unit IEScript20Unit;

interface

{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, MSHTML, StdCtrls, ComCtrls, ExtCtrls, ToolWin, ImgList, Menus,
  ActnList, (*DragDrop, DropTarget, DragDropFile, DropSource,*)AppEvnts,
  Buttons, SHDocVw;

type
  TFormIEScript = class(TForm)
    ImageListDrag: TImageList;
    ToolBarSelect: TToolBar;
    ImageDrag: TImage;
    ToolButtonSelectWindow: TToolButton;
    ToolButtonJavaScript: TToolButton;
    ToolButtonSeparatorSelect: TToolButton;
    EditSelect: TEdit;
    ImageListTools: TImageList;
    ActionListTools: TActionList;
    ActionSelectWindow: TAction;
    ActionAbout: TAction;
    ActionClose: TAction;
    MainMenuOne: TMainMenu;
    MenuItemFile: TMenuItem;
    MenuItemSelectWindow: TMenuItem;
    MenuItemClose: TMenuItem;
    MenuItemHelp: TMenuItem;
    MenuItemAbout: TMenuItem;
    StatusBarOne: TStatusBar;
    LabelHttp: TLabel;
    ActionExecuteScript: TAction;
    GroupBoxScriptEditor: TGroupBox;
    GroupBoxScriptList: TGroupBox;
    SplitterA: TSplitter;
    ActionRefurbish: TAction;
    ToolButtonRefurbish: TToolButton;
    ActionBlog: TAction;
    MenuItemExecuteScript: TMenuItem;
    MenuItemRefurbish: TMenuItem;
    ImageListScriptList: TImageList;
    MenuItemBlog: TMenuItem;
    MenuItemHelpA: TMenuItem;
    ActionHelp: TAction;
    ActionViewSource: TAction;
    ActionProperties: TAction;
    MenuItemTools: TMenuItem;
    MenuItemViewSource: TMenuItem;
    MenuItemProperties: TMenuItem;
    ActionSave: TAction;
    ToolButtonSave: TToolButton;
    MenuItemSave: TMenuItem;
    ActionCopyShort: TAction;
    MenuItemCopy: TMenuItem;
    ActionOpenDir: TAction;
    ToolButtonOpenDir: TToolButton;
    ActionAddFavorite: TAction;
    MenuItemAddFavorite: TMenuItem;
    ToolButtonCopyShort: TToolButton;
    ToolButtonAddFavorite: TToolButton;
    ActionRename: TAction;
    TreeViewScriptList: TTreeView;
    MenuItemRename: TMenuItem;
    PopupMenuScriptList: TPopupMenu;
    MenuItemSaveB: TMenuItem;
    MenuItemOpenDirB: TMenuItem;
    MenuItemRenameB: TMenuItem;
    MenuItemExecuteScriptB: TMenuItem;
    MenuItemCopyShortB: TMenuItem;
    MenuItemScript: TMenuItem;
    MenuItemAddFavoriteB: TMenuItem;
    MenuItemOpenDirA: TMenuItem;
    ToolButtonRename: TToolButton;
    ActionNewVBScript: TAction;
    ActionNewJavaScript: TAction;
    MenuItemLineA: TMenuItem;
    MenuItemNewJavaScriptA: TMenuItem;
    MenuItemNewVBScriptA: TMenuItem;
    ActionNewSibling: TAction;
    MenuItemNewSiblingA: TMenuItem;
    MenuItemNewJavaScriptB: TMenuItem;
    MenuItemNewVBScriptB: TMenuItem;
    MenuItemNewSiblingB: TMenuItem;
    MenuItemLineB: TMenuItem;
    ActionCopyEval: TAction;
    MenuItemCopyEval: TMenuItem;
    ActionDelete: TAction;
    MenuItemDeleteB: TMenuItem;
    MenuItemDeleteA: TMenuItem;
    ActionNewChildren: TAction;
    MenuItemNewChildrenB: TMenuItem;
    MenuItemNewChildrenA: TMenuItem;
    ActionFindText: TAction;
    MenuItemSearch: TMenuItem;
    MenuItemFindTextA: TMenuItem;
    FindDialogOne: TFindDialog;
    ReplaceDialogOne: TReplaceDialog;
    ActionReplace: TAction;
    MenuItemReplace: TMenuItem;
    TimerOne: TTimer;
    ActionExecuteScriptTimer: TAction;
    MenuItemExecuteScriptTimer: TMenuItem;
    MenuItemExecuteScriptTimerA: TMenuItem;
    ActionNoForeground: TAction;
    MenuItemNoForeground: TMenuItem;
    PopupMenuScriptEditor: TPopupMenu;
    MenuItemEditUndo: TMenuItem;
    MenuItemEditRedo: TMenuItem;
    MenuItemLineM: TMenuItem;
    MenuItemEditCut: TMenuItem;
    MenuItemEditCopy: TMenuItem;
    MenuItemEditPaste: TMenuItem;
    MenuItemEditDelete: TMenuItem;
    MenuItemLineN: TMenuItem;
    MenuItemEditSelectAll: TMenuItem;
    ActionEditCut: TAction;
    ActionEditCopy: TAction;
    ActionEditPaste: TAction;
    ActionEditDelete: TAction;
    ActionEditUndo: TAction;
    ActionEditRedo: TAction;
    ActionEditSelectAll: TAction;
    ActionSearchAgain: TAction;
    MenuItemSearchAgain: TMenuItem;
    MenuItemEdit: TMenuItem;
    MenuItemEditCopyA: TMenuItem;
    MenuItemEditCutA: TMenuItem;
    MenuItemEditDeleteA: TMenuItem;
    MenuItemEditPasteA: TMenuItem;
    MenuItemEditRedoA: TMenuItem;
    MenuItemEditSelectAllA: TMenuItem;
    MenuItemEditUndoA: TMenuItem;
    MenuItemLineO: TMenuItem;
    MenuItemLineP: TMenuItem;
    MemoScriptEditor: TMemo;
    ActionSavePage: TAction;
    MenuItemSavePage: TMenuItem;
    MenuItemCoding: TMenuItem;
    MenuItemAnsiA: TMenuItem;
    MenuItemUtf8A: TMenuItem;
    MenuItemUnicode: TMenuItem;
    ActionCodingAnsi: TAction;
    ActionCodingUtf8: TAction;
    ActionCodingUnicode: TAction;
    ActionCodingUtf8NoneBOM: TAction;
    MenuItemUtf8NoneBOM: TMenuItem;
    MenuItemLineC: TMenuItem;
    ActionSelectRoot: TAction;
    MenuItemSelectRoot: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure ImageDragMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ImageDragMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ImageDragMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ToolBarSelectResize(Sender: TObject);
    procedure ActionSelectWindowExecute(Sender: TObject);
    procedure ActionCloseExecute(Sender: TObject);
    procedure ActionExecuteScriptExecute(Sender: TObject);
    procedure StatusBarOneResize(Sender: TObject);
    procedure LabelHttpClick(Sender: TObject);
    procedure ActionRefurbishExecute(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ActionAboutExecute(Sender: TObject);
    procedure ActionBlogExecute(Sender: TObject);
    procedure ActionHelpExecute(Sender: TObject);
    procedure ActionListToolsUpdate(Action: TBasicAction;
      var Handled: Boolean);
    procedure ActionViewSourceExecute(Sender: TObject);
    procedure ActionPropertiesExecute(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ActionSaveExecute(Sender: TObject);
    procedure MemoScriptEditorChange(Sender: TObject);
    procedure TreeViewScriptListChange(Sender: TObject; Node: TTreeNode);
    procedure ActionCopyShortExecute(Sender: TObject);
    procedure ActionOpenDirExecute(Sender: TObject);
    procedure TreeViewScriptListKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ActionAddFavoriteExecute(Sender: TObject);
    procedure TreeViewScriptListEdited(Sender: TObject; Node: TTreeNode;
      var S: String);
    procedure ActionRenameExecute(Sender: TObject);
    procedure TreeViewScriptListMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ActionNewScriptExecute(Sender: TObject);
    procedure ActionCopyEvalExecute(Sender: TObject);
    procedure ActionDeleteExecute(Sender: TObject);
    procedure MemoScriptEditorKeyPress(Sender: TObject; var Key: Char);
    procedure ActionFindTextExecute(Sender: TObject);
    procedure FindDialogOneFind(Sender: TObject);
    procedure ActionReplaceExecute(Sender: TObject);
    procedure ReplaceDialogOneReplace(Sender: TObject);
    procedure MemoScriptEditorKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ActionExecuteScriptTimerExecute(Sender: TObject);
    procedure TimerOneTimer(Sender: TObject);
    procedure ActionNoForegroundExecute(Sender: TObject);
    procedure TreeViewScriptListChanging(Sender: TObject; Node: TTreeNode;
      var AllowChange: Boolean);
    procedure ActionEditCutExecute(Sender: TObject);
    procedure ActionEditCopyExecute(Sender: TObject);
    procedure ActionEditPasteExecute(Sender: TObject);
    procedure ActionEditDeleteExecute(Sender: TObject);
    procedure ActionEditUndoExecute(Sender: TObject);
    procedure ActionEditRedoExecute(Sender: TObject);
    procedure ActionEditSelectAllExecute(Sender: TObject);
    procedure ActionSearchAgainExecute(Sender: TObject);
    procedure ActionSavePageExecute(Sender: TObject);
    procedure ActionCodingAnsiExecute(Sender: TObject);
    procedure ActionCodingUtf8Execute(Sender: TObject);
    procedure ActionCodingUnicodeExecute(Sender: TObject);
    procedure ActionCodingUtf8NoneBOMExecute(Sender: TObject);
    procedure ActionSelectRootExecute(Sender: TObject);
  private
    { Private declarations }
    FMouseDown: Boolean;
    FIEHandle: THandle;
    FDocument: IHTMLDocument2;
    FWebBrowser2: IWebBrowser2;
    FEncoding: string;
    
    FHasDocument: Boolean;
    FScriptChanging: Boolean;
    FSelectPath: string;
    FEditorShowing: Boolean;
    (*DropFileSourceOne: TDropFileSource;*)
    FLastChar: Char; // �������������ַ�
    FSearchEvent: TNotifyEvent;
    FSearchDialog: TFindDialog;
    FLastChangingNode: TTreeNode;
    function ShortScript: string;
    function EvalScript: string;
    procedure ApplicationHint(Sender: TObject);
    function GetNodePath(ATreeNode: TTreeNode): string;
  public
    { Public declarations }
  end;

var
  FormIEScript: TFormIEScript;

implementation

{$R *.dfm}
{$R CursorRes.res}

uses ShellAPI, ActiveX, IniFiles, CommCtrl, Clipbrd, Math, WindowDialog,
  FavoriteDialog, RecycleDialog, Search, MSHTMCID, FileFunctions;

const
{ VK_0 thru VK_9 are the same as ASCII '0' thru '9' ($30 - $39) }
  VK_0 = $30;  VK_1 = $31;  VK_2 = $32;  VK_3 = $33;  VK_4 = $34;
  VK_5 = $35;  VK_6 = $36;  VK_7 = $37;  VK_8 = $38;  VK_9 = $39;

{ VK_A thru VK_Z are the same as ASCII 'A' thru 'Z' ($41 - $5A) }
  VK_A = $41;  VK_B = $42;  VK_C = $43;  VK_D = $44;
  VK_E = $45;  VK_F = $46;  VK_G = $47;  VK_H = $48;
  VK_I = $49;  VK_J = $4A;  VK_K = $4B;  VK_L = $4C;  VK_M = $4D;  VK_N = $4E;
  VK_O = $4F;  VK_P = $50;  VK_Q = $51;  VK_R = $52;  VK_S = $53;  VK_T = $54;
  VK_U = $55;  VK_V = $56;  VK_W = $57;  VK_X = $58;  VK_Y = $59;  VK_Z = $5A;
  
function FirstSpace( // ȡǰ���հ��ַ�
  mStr: string; // �ַ���
  mSpaceChar: TSysCharSet = [#9, #32] // �ַ���
): string; // ����ǰ���հ��ַ�
var
  I: Integer;
begin
  Result := '';
  for I := 1 to Length(mStr) do
    if not (mStr[I] in mSpaceChar) then
    begin
      Result := Copy(mStr, 1, I - 1);
      Break;
    end;
end; { FirstSpace }

const crFindWindow = 10;

type
  TFileVersionInfomation = record // �ļ��汾��Ϣ
    rCommpanyName: string;
    rFileDescription: string;
    rFileVersion: string;
    rInternalName: string;
    rLegalCopyright: string;
    rLegalTrademarks: string;
    rOriginalFileName: string;
    rProductName: string;
    rProductVersion: string;
    rComments: string;
    rVsFixedFileInfo: VS_FIXEDFILEINFO;
    rDefineValue: string;
  end;

var
  vModuleVersionInfomation: TFileVersionInfomation;

function TrackerWindow( // ��ɫ��ʾ�����������ѡȡ������
  AHandle: THandle // ������
): Boolean; // ���ش����Ƿ�ɹ�
var
  vDC: HDC;
  vRect: TRect;
begin
  Result := False;
  if not IsWindow(AHandle) then Exit;
  GetWindowRect(AHandle, vRect);
  OffsetRect(vRect, -vRect.Left, -vRect.Top);
  if IsRectEmpty(vRect) then Exit;
  vDC := GetWindowDC(AHandle);
  try
    PatBlt(vDC, vRect.Left, vRect.Top,                                          //2008-03-05 ZswangY37 No.1
      4, vRect.Bottom - vRect.Top, DSTINVERT);
    PatBlt(vDC, vRect.Right - 4, vRect.Top,
      4, vRect.Bottom - vRect.Top, DSTINVERT);
    PatBlt(vDC, vRect.Left + 4, vRect.Top,
      (vRect.Right - vRect.Left) - 8, 4, DSTINVERT);
    PatBlt(vDC, vRect.Left + 4, vRect.Bottom - 4,
      (vRect.Right - vRect.Left) - 8, 4, DSTINVERT);
  finally
    ReleaseDC(AHandle, vDC);
  end;
  Result := True;
end; { TrackerWindow }

function ForceForegroundWindow( // ����������Ϊ��ǰ��,����ý���
  mHandle: THandle // ������
): Boolean; // ���������Ƿ�ɹ�
var
  vHandle: THandle;
  vResult: Longword;
begin
  if IsIconic(mHandle) then
    SendMessageTimeOut(mHandle, WM_SYSCOMMAND, SC_RESTORE, 0, SMTO_NORMAL,
      1000, vResult)
  else                                                                          //2006-10-13 ZswangY37 No.1
  begin
    vHandle := GetWindow(mHandle, GW_OWNER);
    if IsIconic(vHandle) then
      SendMessageTimeOut(vHandle, WM_SYSCOMMAND, SC_RESTORE, 0, SMTO_NORMAL,
        1000, vResult);
  end;
  vHandle := GetForegroundWindow;
  AttachThreadInput(GetWindowThreadProcessId(vHandle, nil),
    GetCurrentThreadId, True);
  Result := SetForegroundWindow(mHandle);
  AttachThreadInput(GetWindowThreadProcessId(vHandle, nil),
    GetCurrentThreadId, False);
end; { ForceForegroundWindow }

type
  TObjectFromLResult = function(
    LRESULT: lResult;
    const IID: TIID;
    WPARAM: wParam;
    out pObject
  ): HRESULT; stdcall;

function DocumentFromHWND( //��ô����е�IHTMLDocument2����
  mHandle: HWND; //�ô�����
  var nDocument2: IHTMLDocument2 //���ص�IHTMLDocument2����
): HRESULT; //���ش�����룬����ɹ��򷵻�0
{$J+}const WM_HTML_GETOBJECT: Integer = 0;{$J-}
var
  vLibraryHandle: HWND;
  vResult: Cardinal;
  vObjectFromLresult: TObjectFromLresult;
begin
  vLibraryHandle := LoadLibrary('Oleacc.dll');
  @vObjectFromLresult := GetProcAddress(vLibraryHandle, 'ObjectFromLresult');
  Result := S_FALSE;
  if @vObjectFromLresult <> nil then
  begin
    try
      WM_HTML_GETOBJECT := RegisterWindowMessage('WM_HTML_GETOBJECT');
      SendMessageTimeOut(mHandle,
        WM_HTML_GETOBJECT, 0, 0, SMTO_ABORTIFHUNG, 1000, vResult);
      Result := vObjectFromLresult(vResult, IHTMLDocument2, 0, nDocument2);
    finally
      FreeLibrary(vLibraryHandle);
    end;
  end;
end; { DocumentFromHWND }

function ExePath: string; // ���ص�ǰӦ�ó����·��
begin
  Result := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)));
end; { ExePath }

function IsValidFileName(AFileName: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  if Length(AFileName) <= 0 then Exit;
  for I := 1 to Length(AFileName) do
    if AFileName[I] in ['/', '\', '<', '>', '|', ':', '?', '*', '"'] then Exit;
  Result := True;
end; { IsValidFileName }

function SystemImageList( // ���ϵͳͼ���б�
  mImageList: TImageList // ͼ���б�
): Boolean; // ����ϵͳͼ�굽ͼ���б����Ƿ�ɹ�
var
  vHandle: THandle;
  vSHFileInfo: TSHFileInfo;
begin
  FillChar(vSHFileInfo, SizeOf(vSHFileInfo), 0);
  vHandle := SHGetFileInfo('', 0, vSHFileInfo, SizeOf(vSHFileInfo),
    SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
  Result := vHandle <> 0;
  mImageList.Handle := vHandle;
  mImageList.ShareImages := True;
end; { GetSystemImageList }

function PathIconIndex( // ���·����ͼ�����
  mPath: string // ·��
): Integer; // �����ļ���·������Ӧ��ͼ�����
var
  vSHFileInfo: TSHFileInfo;
begin
  FillChar(vSHFileInfo, SizeOf(vSHFileInfo), 0);
  SHGetFileInfo(PChar(mPath), 0,
    vSHFileInfo, SizeOf(vSHFileInfo), SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
  Result := vSHFileInfo.iIcon;
end; { GetIconIndex }

//type
//  TScriptInfo = packed record
//    rFileName: string;
//  end;
//  PScriptInfo = ^TScriptInfo;

procedure TFormIEScript.FormCreate(Sender: TObject);
var
  I: Integer;
begin
  ActionAbout.ImageIndex := ImageListTools.Count;
  ImageListTools.AddIcon(Application.Icon);
  
  (*DropFileSourceOne := TDropFileSource.Create(Self);*)

  Font.Assign(Screen.MenuFont);
  Screen.Cursors[crFindWindow] :=
    {//�򿪿������ļ�����}LoadCursor(HInstance, 'CursorFindWindow'); (*//
LoadCursorFromFile(PChar(ExtractFilePath(ParamStr(0)) + 'FindWindow2.cur'));//*)
  ImageListDrag.GetBitmap(Ord(FMouseDown), ImageDrag.Picture.Bitmap);

  LabelHttp.Font.Color := clBlue;
  LabelHttp.Font.Style := [fsUnderline];

  Application.OnHint := ApplicationHint;
  Application.Title := vModuleVersionInfomation.rProductName + '-' +
    vModuleVersionInfomation.rProductVersion;
  Application.HelpFile := ChangeFileExt(ParamStr(0), '.hlp');
  Caption := Application.Title;
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do try
    Font.Name := ReadString(Self.ClassName, 'Font.Name', Font.Name);
    Font.Size := ReadInteger(Self.ClassName, 'Font.Size', Font.Size);
    try
      Font.Color := StringToColor(ReadString(Self.ClassName, 'Font.Color',
        ColorToString(Font.Color)));
    except
    end;
    FSelectPath := ReadString(Self.ClassName, 'SelectPath', FSelectPath);
    Left := ReadInteger(Self.ClassName, 'Left', Left);
    Top := ReadInteger(Self.ClassName, 'Top', Top);
    Height := ReadInteger(Self.ClassName, 'Height', Height);
    Width := ReadInteger(Self.ClassName, 'Width', Width);

    ActionNoForeground.Checked := ReadBool(Self.ClassName, 'NoForeground',
      ActionNoForeground.Checked);
    for I := 0 to ComponentCount - 1 do
    begin
      if Components[I] is TOpenDialog then
        with TOpenDialog(Components[I]) do
          FilterIndex := ReadInteger(Name, 'FilterIndex', FilterIndex);
    end;
    GroupBoxScriptList.Width :=
      ReadInteger(GroupBoxScriptList.Name, 'Width', GroupBoxScriptList.Width);
  finally
    Free;
  end;
  ActionRefurbish.Execute;
  SystemImageList(ImageListScriptList);
end;

procedure TFormIEScript.ImageDragMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  vRect: TRect;
begin
  if Button <> mbLeft then Exit;
  FIEHandle := 0;
  FDocument := nil;
  FHasDocument := False;
  EditSelect.Text := '';
  FMouseDown := True;
  ImageListDrag.GetBitmap(Ord(FMouseDown), TImage(Sender).Picture.Bitmap);
  vRect := TImage(Sender).BoundsRect;
  InvalidateRect(TImage(Sender).Parent.Handle, @vRect, True);
  Screen.Cursor := crFindWindow;
end;

procedure TFormIEScript.ImageDragMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  vHandle: THandle;
  vBuffer: array[0..255] of Char;
begin
  if not FMouseDown then Exit;
  ///////Begin �õ�IE������
  vHandle := WindowFromPoint(Mouse.CursorPos);
  GetClassName(vHandle, vBuffer, SizeOf(vBuffer));
  while (vHandle <> 0) and not SameText(vBuffer, 'Internet Explorer_Server') do
  begin
    vHandle := GetParent(vHandle);
    GetClassName(vHandle, vBuffer, SizeOf(vBuffer));
  end;
  ///////End �õ�IE������
  if FIEHandle = vHandle then Exit;
  if FIEHandle <> 0 then TrackerWindow(FIEHandle);
  FIEHandle := vHandle;
  TrackerWindow(FIEHandle);
  FDocument := nil;
  FHasDocument := DocumentFromHWND(FIEHandle, FDocument) = 0;
  if FHasDocument then
    EditSelect.Text := Format('%d=%s', [FIEHandle, FDocument.url])
  else EditSelect.Text := Format('%d=<null>', [FIEHandle])
end;

procedure TFormIEScript.ImageDragMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  vRect: TRect;
begin
  if not FMouseDown then Exit;
  if FIEHandle <> 0 then TrackerWindow(FIEHandle);
  FMouseDown := False;
  ImageListDrag.GetBitmap(Ord(FMouseDown), TImage(Sender).Picture.Bitmap);
  vRect := TImage(Sender).BoundsRect;
  InvalidateRect(TImage(Sender).Parent.Handle, @vRect, True);
  Screen.Cursor := crDefault;
end;

procedure TFormIEScript.ToolBarSelectResize(Sender: TObject);
begin
  EditSelect.Width := TToolBar(Sender).ClientWidth - EditSelect.Left - 2;
end;

procedure TFormIEScript.ActionSelectWindowExecute(Sender: TObject);
var
  vHandle: THandle;
  vBuffer: array[0..255] of Char;
begin
  with TWindowDialog.Create(nil) do try
    Hide;
    if not Execute then Exit;
    ///////Begin �õ�IE������
    vHandle := WindowFromPoint(Mouse.CursorPos);
    GetClassName(vHandle, vBuffer, SizeOf(vBuffer));
    while (vHandle <> 0) and not SameText(vBuffer, 'Internet Explorer_Server') do
    begin
      vHandle := GetParent(vHandle);
      GetClassName(vHandle, vBuffer, SizeOf(vBuffer));
    end;
    ///////End �õ�IE������
    FIEHandle := vHandle;
    FDocument := nil;
    FHasDocument := DocumentFromHWND(FIEHandle, FDocument) = 0;
    if FHasDocument then
      EditSelect.Text := Format('%d=%s', [FIEHandle, FDocument.url])
    else EditSelect.Text := Format('%d=<null>', [FIEHandle]);
    ForceForegroundWindow(Self.Handle);                                         //2006-11-17 ZswangY37 No.1
  finally
    ForceForegroundWindow(Self.Handle);
    Show;
    Free;
  end;
end;

procedure TFormIEScript.ActionCloseExecute(Sender: TObject);
begin
  Close;
end;

procedure TFormIEScript.ActionExecuteScriptExecute(Sender: TObject);
var
  vHandle: THandle;
  vLanguage: string;
begin
  if MemoScriptEditor.GetTextLen <= 0 then Exit; // �ű�������
  FDocument := nil;
  FHasDocument := DocumentFromHWND(FIEHandle, FDocument) = 0;
  if not FHasDocument then Exit;
  if Assigned(TreeViewScriptList.Selected) and
    SameText(ExtractFileExt(TreeViewScriptList.Selected.Text), '.vbs') then
    vLanguage := 'VBScript'
  else vLanguage := 'JavaScript';
  vHandle := FIEHandle;
  while GetParent(vHandle) <> 0 do
    vHandle := GetParent(vHandle);
  if not ActionNoForeground.Checked then ForceForegroundWindow(vHandle);
  try
    FDocument.parentWindow.execScript(MemoScriptEditor.Text, vLanguage);
  finally
    if not ActionNoForeground.Checked then ForceForegroundWindow(Self.Handle);
  end;
end;

procedure TFormIEScript.StatusBarOneResize(Sender: TObject);
begin
  LabelHttp.Left := (LabelHttp.Parent.Width - LabelHttp.Width) - 12;
end;

procedure TFormIEScript.LabelHttpClick(Sender: TObject);
begin
  ShellExecute(Handle, nil, PChar(TLabel(Sender).Caption), nil, nil, SW_SHOW);
end;

procedure TFormIEScript.ApplicationHint(Sender: TObject);
var
  vStatusBar: TStatusBar;
begin
  if not Assigned(Screen.ActiveForm) then Exit;
  TComponent(vStatusBar) := Screen.ActiveForm.FindComponent('StatusBarOne');
  if not Assigned(vStatusBar) then Exit;
  if Length(Application.Hint) > 0 then
  begin
    vStatusBar.SimplePanel := True;
    vStatusBar.SimpleText := Application.Hint;
  end else vStatusBar.SimplePanel := False;
  LabelHttp.Visible := not vStatusBar.SimplePanel;
end;

procedure TFormIEScript.ActionRefurbishExecute(Sender: TObject);
var
  vSelectNode: TTreeNode;

  procedure pScanPath(APath: string; ATreeNode: TTreeNode);                     //2008-03-14 ZswangY37 No.2
  var
    vSearchRec: TSearchRec;
    vTreeNode: TTreeNode;
  begin
    if FindFirst(APath + '*.js', faArchive, vSearchRec) = 0 then
    begin
      repeat
        vTreeNode := TreeViewScriptList.Items.AddChild(ATreeNode, vSearchRec.Name);
        vTreeNode.ImageIndex := PathIconIndex(APath + vSearchRec.Name);
        vTreeNode.SelectedIndex := vTreeNode.ImageIndex;
        if SameText(FSelectPath, APath + vSearchRec.Name) then
          vSelectNode := vTreeNode;
      until FindNext(vSearchRec) <> 0;
      FindClose(vSearchRec);
    end;
    if FindFirst(APath + '*.vbs', faArchive, vSearchRec) = 0 then
    begin
      repeat
        vTreeNode := TreeViewScriptList.Items.AddChild(ATreeNode, vSearchRec.Name);
        vTreeNode.ImageIndex := PathIconIndex(APath + vSearchRec.Name);
        vTreeNode.SelectedIndex := vTreeNode.ImageIndex;
        if SameText(FSelectPath, APath + vSearchRec.Name) then
          vSelectNode := vTreeNode;
      until FindNext(vSearchRec) <> 0;
      FindClose(vSearchRec);
    end;
    if FindFirst(APath + '*.*', faDirectory, vSearchRec) = 0 then
    begin
      repeat
        if (vSearchRec.Attr and faDirectory = faDirectory) and
          (Pos('.', vSearchRec.Name) <> 1) then
        begin
          vTreeNode :=
            TreeViewScriptList.Items.AddChild(ATreeNode, vSearchRec.Name);
          vTreeNode.ImageIndex := PathIconIndex(APath + vSearchRec.Name);
          vTreeNode.SelectedIndex := vTreeNode.ImageIndex;
          pScanPath(APath + vSearchRec.Name + '\', vTreeNode);
          vTreeNode.Data := Pointer(1);
          if SameText(FSelectPath, APath + vSearchRec.Name) then
            vSelectNode := vTreeNode;
        end;
      until FindNext(vSearchRec) <> 0;
      FindClose(vSearchRec);
    end;
  end;

begin
  vSelectNode := nil;
  TreeViewScriptList.Items.Clear;
  pScanPath(ExePath, nil);
  TreeViewScriptList.Selected := vSelectNode;
  if Assigned(TreeViewScriptList.Selected) then
    TreeViewScriptList.Selected.Expand(False)
  else begin
    ActionCodingUtf8.Checked := True;
    TreeViewScriptListChange(TreeViewScriptList, nil);
  end;
end;

procedure TFormIEScript.FormDestroy(Sender: TObject);
var
  I: Integer;
begin
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do try
    WriteInteger(Self.ClassName, 'Left', Left);
    WriteInteger(Self.ClassName, 'Top', Top);
    WriteInteger(Self.ClassName, 'Height', Height);
    WriteInteger(Self.ClassName, 'Width', Width);
    WriteInteger(GroupBoxScriptList.Name, 'Width', GroupBoxScriptList.Width);
    WriteString(Self.ClassName, 'SelectPath', FSelectPath);
    WriteBool(Self.ClassName, 'NoForeground', ActionNoForeground.Checked);
    for I := 0 to ComponentCount - 1 do
    begin
      if Components[I] is TOpenDialog then
        with TOpenDialog(Components[I]) do
          WriteInteger(Name, 'FilterIndex', FilterIndex);
    end;
  finally
    Free;
  end;
end;

function GetFileVersionInfomation( // ��ȡ�ļ��İ汾��Ϣ
  mFileName: string; // Ŀ���ļ���
  var nFileVersionInfomation: TFileVersionInfomation; // �ļ���Ϣ�ṹ
  mDefineName: string = '' // �Զ����ֶ���
): Boolean; // �����Ƿ��ȡ�ɹ�
var
  vHandle: Cardinal;
  vInfoSize: Cardinal;
  vVersionInfo: Pointer;
  vTranslation: Pointer;
  vVersionValue: string;
  vInfoPointer: Pointer;
begin
  Result := False;
  vInfoSize := GetFileVersionInfoSize(PChar(mFileName), vHandle); // ȡ���ļ��汾��Ϣ�ռ估��Դ���
  FillChar(nFileVersionInfomation, SizeOf(nFileVersionInfomation), 0); // ��ʼ��������Ϣ
  if vInfoSize <= 0 then Exit; // ��ȫ���

  GetMem(vVersionInfo, vInfoSize); // ������Դ
  with nFileVersionInfomation do try
    if not GetFileVersionInfo(PChar(mFileName),
      vHandle, vInfoSize, vVersionInfo) then Exit;
    CloseHandle(vHandle);
    VerQueryValue(vVersionInfo, '\VarFileInfo\Translation',
      vTranslation, vInfoSize);
    if not Assigned(vTranslation) then Exit;
    vVersionValue := Format('\StringFileInfo\%.4x%.4x\',
      [LOWORD(Longint(vTranslation^)), HIWORD(Longint(vTranslation^))]);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'CompanyName'),
      vInfoPointer, vInfoSize);
    rCommpanyName := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'FileDescription'),
      vInfoPointer, vInfoSize);
    rFileDescription := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'FileVersion'),
      vInfoPointer, vInfoSize);
    rFileVersion := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'InternalName'),
      vInfoPointer, vInfoSize);
    rInternalName := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'LegalCopyright'),
      vInfoPointer, vInfoSize);
    rLegalCopyright := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'LegalTrademarks'),
      vInfoPointer, vInfoSize);
    rLegalTrademarks := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'OriginalFileName'),
      vInfoPointer, vInfoSize);
    rOriginalFileName := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'ProductName'),
      vInfoPointer, vInfoSize);
    rProductName := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'ProductVersion'),
      vInfoPointer, vInfoSize);
    rProductVersion := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, PChar(vVersionValue + 'Comments'),
      vInfoPointer, vInfoSize);
    rComments := PChar(vInfoPointer);
    VerQueryValue(vVersionInfo, '\', vInfoPointer, vInfoSize);
    rVsFixedFileInfo := TVSFixedFileInfo(vInfoPointer^);
    if mDefineName <> '' then begin
      VerQueryValue(vVersionInfo, PChar(vVersionValue + mDefineName),
        vInfoPointer, vInfoSize);
      rDefineValue := PChar(vInfoPointer);
    end else rDefineValue := '';
  finally
    FreeMem(vVersionInfo, vInfoSize);
  end;
  Result := True;
end; { GetFileVersionInfomation }

procedure TFormIEScript.ActionAboutExecute(Sender: TObject);
begin
  ShowMessage(Format(
'��������:%s'#13#10 +
'����汾:%s'#13#10 +
'�ļ��汾:%s'#13#10 +
'%s'#13#10 +
'���:ZswangY37' +
'', [
    vModuleVersionInfomation.rProductName,
    vModuleVersionInfomation.rProductVersion,
    vModuleVersionInfomation.rFileVersion,
    vModuleVersionInfomation.rLegalCopyright
  ]));
end;

procedure TFormIEScript.ActionBlogExecute(Sender: TObject);
begin
  ShellExecute(Handle, nil, PChar(TAction(Sender).Hint), nil, nil, SW_SHOW);
end;

procedure TFormIEScript.ActionHelpExecute(Sender: TObject);
begin
  ShowMessage(
'IE�ű�����'#13#10 +
''#13#10 +
'  IE��ַ��������á�javascript:xxxx�������нű���'#13#10 +
'  �����ű��Ƕ��еģ��������ݺܶ࣬�ڵ�ַ�������룬�Ͳ������ˡ�'#13#10 +
'  ����һЩҳ�汾���û�е�ַ����'#13#10 +
'  �����߾���Ϊ���ṩ��IE������ִ�д�νű��Ĺ��ܡ�'#13#10 +
''#13#10 +
'[��д�ű�]'#13#10 +
'  ������û���ṩ�ű�����Ĺ��ܣ�������ʹ������������(�磺���±�)��д�ű���'#13#10 +
'  "Script"��Ŀ¼��ר�Ŵ�Žű��ļ���'#13#10 +
'  �����ļ���չ����.js����ӦJavaScript���ԡ���.vbs����ӦVBScript���ԣ�Ŀǰ֧�������ֽű���'#13#10 +
''#13#10 +
'[ѡ��IE����]'#13#10 +
'  ͨ���϶�����׼������ͼ�����ѡ��Ŀ��IE���塣�ο���Spy++����'#13#10 +
''#13#10 +
'[ִ�нű�]'#13#10 +
'  ѡ��IE���岢���нű����ݺ��������нű�����'#13#10 +
''#13#10 +
'���ܼ򵥣�����Ҳ�򵥡�'#13#10 +
''#13#10 +
'              --ZswangY37 2008-03-10'#13#10);
  //Application.HelpCommand(HELP_CONTENTS, 0);
end;

procedure TFormIEScript.ActionListToolsUpdate(Action: TBasicAction;
  var Handled: Boolean);
var
  vPoint: TPoint;
begin
  ActionExecuteScript.Enabled :=
    FHasDocument and (MemoScriptEditor.GetTextLen > 0) and IsWindow(FIEHandle);
  ActionAddFavorite.Enabled := MemoScriptEditor.GetTextLen > 0;
  ActionRename.Enabled := Assigned(TreeViewScriptList.Selected);
  ActionNewJavaScript.Enabled := (TreeViewScriptList.Items.Count <= 0) or
      Assigned(TreeViewScriptList.Selected);
  ActionNewVBScript.Enabled := (TreeViewScriptList.Items.Count <= 0) or
      Assigned(TreeViewScriptList.Selected);
  ActionNewSibling.Enabled := (TreeViewScriptList.Items.Count <= 0) or
      Assigned(TreeViewScriptList.Selected);
  ActionNewChildren.Enabled := Assigned(TreeViewScriptList.Selected) and
    Assigned(TreeViewScriptList.Selected.Data);
  ActionViewSource.Enabled := FHasDocument and IsWindow(FIEHandle);
  ActionProperties.Enabled := FHasDocument and IsWindow(FIEHandle);
  ActionSavePage.Enabled := FHasDocument and IsWindow(FIEHandle);
  ActionSave.Enabled := FScriptChanging;
  ActionOpenDir.Enabled := True;
  ActionDelete.Enabled := Assigned(TreeViewScriptList.Selected);
  ActionCopyShort.Enabled := MemoScriptEditor.GetTextLen > 0;
  ActionCopyEval.Enabled := MemoScriptEditor.GetTextLen > 0;
  ActionExecuteScriptTimer.Enabled := 
    FHasDocument and (MemoScriptEditor.GetTextLen > 0) and IsWindow(FIEHandle);
  if TimerOne.Enabled then
    ActionExecuteScriptTimer.ImageIndex := 17
  else ActionExecuteScriptTimer.ImageIndex := 16;

  vPoint := MemoScriptEditor.CaretPos;
  StatusBarOne.Panels[0].Text :=
    Format('��: %d, ��: %d', [vPoint.Y + 1, vPoint.X + 1]);
  if FScriptChanging then
    StatusBarOne.Panels[1].Text := '�༭��...'
  else StatusBarOne.Panels[1].Text := '';
  if not FEditorShowing and not TreeViewScriptList.ReadOnly and
    (TreeView_GetEditControl(TreeViewScriptList.Handle) = 0) then
    TreeViewScriptList.ReadOnly := True; // ���ⵥ�������޸�

  ActionEditCopy.Enabled := Length(MemoScriptEditor.SelText) > 0;
  ActionEditDelete.Enabled := ActionEditCopy.Enabled and not MemoScriptEditor.ReadOnly;
  ActionEditCut.Enabled := ActionEditDelete.Enabled;

  ActionEditPaste.Enabled := Clipboard.HasFormat(CF_TEXT);
  //ActionEditRedo.Enabled := MemoScriptEditor.CanRedo;
  ActionEditUndo.Enabled := MemoScriptEditor.CanUndo;
  ActionEditSelectAll.Enabled := not MemoScriptEditor.ReadOnly;
  ActionSelectRoot.Enabled := Assigned(TreeViewScriptList.Selected);
end;

procedure TFormIEScript.ActionViewSourceExecute(Sender: TObject);
begin
  if not FHasDocument then Exit;
  SendMessage(FIEHandle, WM_COMMAND, IDM_VIEWSOURCE, Handle);
end;

procedure TFormIEScript.ActionPropertiesExecute(Sender: TObject);
begin
  if not FHasDocument then Exit;
  SendMessage(FIEHandle, WM_COMMAND, IDM_PROPERTIES, Handle);
end;

procedure TFormIEScript.ActionSavePageExecute(Sender: TObject);
var
  pvaIn, pvaOut: OleVariant;
begin
  if not FHasDocument then Exit;

  FDocument := nil;
  FHasDocument := DocumentFromHWND(FIEHandle, FDocument) = 0;
  if not FHasDocument then Exit;
  (FDocument.parentWindow as IServiceprovider).QueryService(
    IWebbrowserApp, IWebbrowser2, FWebBrowser2);
  if FWebBrowser2 = nil then Exit;
  FWebBrowser2.ExecWB(4, 1, pvaIn, pvaOut);
end;

procedure TFormIEScript.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  WindowState := wsNormal;
end;

procedure TFormIEScript.ActionSaveExecute(Sender: TObject);                     //2008-03-14 ZswangY37 No.4
var
  vFileName: string;
begin
  //if not Assigned(TreeViewScriptList.Selected) then Exit;
  vFileName := GetNodePath(TreeViewScriptList.Selected);
  if DirectoryExists(vFileName) then
    vFileName := IncludeTrailingPathDelimiter(vFileName) + 'default';
  ForceDirectories(ExtractFilePath(vFileName));
  SetFileTextAndEncoding(vFileName, FEncoding, MemoScriptEditor.Text);
  FScriptChanging := False;
  MemoScriptEditor.Modified := False;
end;

procedure TFormIEScript.MemoScriptEditorChange(Sender: TObject);
var
  vSpace: string;
begin
  TimerOne.Enabled := False;
  FScriptChanging := True;
  case FLastChar of
    #13:
      if (TMemo(Sender).CaretPos.X = 0) and (TMemo(Sender).CaretPos.Y > 0) then
      begin
        FLastChar := #0;
        vSpace := FirstSpace(TMemo(Sender).Lines[TMemo(Sender).CaretPos.Y - 1],
          [#9, #32]);
        if Length(vSpace) > 0 then TMemo(Sender).SelText := vSpace;
      end;
  end;
end;

procedure TFormIEScript.TreeViewScriptListChange(Sender: TObject;
  Node: TTreeNode);
var
  vFileName: string;
begin
  FSelectPath := GetNodePath(TTreeView(Sender).Selected);
  MemoScriptEditor.Clear;
  vFileName := FSelectPath;
  if DirectoryExists(vFileName) then
    vFileName := IncludeTrailingPathDelimiter(vFileName) + 'default';
  if FileExists(vFileName) then
  begin
    MemoScriptEditor.Lines.Text := GetFileTextAndEncoding(vFileName, FEncoding);
    FEncoding := GetFileEncoding(vFileName);
    if (FEncoding = 'ansi') or (FEncoding = '~ansi') then
      ActionCodingAnsi.Checked := True
    else if FEncoding = 'utf8' then
      ActionCodingUtf8.Checked := True
    else if FEncoding = '~utf8' then
      ActionCodingUtf8NoneBOM.Checked := True
    else if (FEncoding = 'unicode') or (FEncoding = '~unicode') then
      ActionCodingUnicode.Checked := True;
    FScriptChanging := False;
  end;
end;

function TFormIEScript.GetNodePath(ATreeNode: TTreeNode): string;
begin
  while Assigned(ATreeNode) do
  begin
    Result := '\' + ATreeNode.Text + Result;
    ATreeNode := ATreeNode.Parent;
  end;
  Result := ExcludeTrailingPathDelimiter(ExePath) + Result;
end;

procedure TFormIEScript.ActionCopyShortExecute(Sender: TObject);                     //2008-03-14 ZswangY37 No.3
begin
  if MemoScriptEditor.GetTextLen <= 0 then Exit; // �ű�������
  Clipboard.AsText := ShortScript;
end;

procedure TFormIEScript.ActionOpenDirExecute(Sender: TObject);
begin
  WinExec(PChar(Format('explorer "%s",/n,/select',
    [GetNodePath(TreeViewScriptList.Selected)])), SW_SHOW);
end;

procedure TFormIEScript.TreeViewScriptListKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F2: ActionRename.Execute;
  end;
end;

function TFormIEScript.ShortScript: string;
var
  S: string;
begin
  Result := '';
  if MemoScriptEditor.GetTextLen <= 0 then Exit; // �ű�������
  S := Trim(MemoScriptEditor.Text);
  if Length(S) <= 0 then Exit;
  S := StringReplace(S, #9, #32, [rfReplaceAll]);
  S := StringReplace(S, #13, #32, [rfReplaceAll]);
  S := StringReplace(S, #10, #32, [rfReplaceAll]);
  while Pos(#32#32, S) > 0 do
    S := StringReplace(S, #32#32, #32, [rfReplaceAll]);
  Result := 'javascript:' + S;
end;

procedure TFormIEScript.ActionAddFavoriteExecute(Sender: TObject);              //2008-03-19 ZswangY37 No.1
var
  vFileName: string;
  vShortScript: string;
begin
  vShortScript := ShortScript;
  if Length(vShortScript) <= 0 then Exit;
  with TAddFavoriteDialog.Create(nil) do try
    WindowHandle := Handle;
    DisplayName := ChangeFileExt(
      ExtractFileName(GetNodePath(TreeViewScriptList.Selected)), '');
    if not Execute then Exit;
    vFileName := Format('%s\%s', [Directory, DisplayName]);
    CreateUrlShortcut(vShortScript, vFileName);
  finally
    Free;
  end;
end;

procedure TFormIEScript.TreeViewScriptListEdited(Sender: TObject;
  Node: TTreeNode; var S: String);                                              //2008-03-21 ZswangY37 No.1
var
  vOldFile, vNewFile: string;
begin
  if SameText(Node.Text, S) then Exit; // δ�ı�
  if IsValidFileName(S) and (Assigned(Node.Data) or
    (SameText(ExtractFileExt(S), '.js') or
      SameText(ExtractFileExt(S), '.vbs'))) then
  begin
    vOldFile := GetNodePath(Node);
    vNewFile := ExtractFilePath(vOldFile) + S;
    if not FileExists(vOldFile) and not DirectoryExists(vOldFile) then
    begin
      S := Node.Text;
      MessageBox(Handle, 'ԭʼ�ļ��ѱ�ɾ����', '����', MB_ICONERROR or MB_OK);
    end else if FileExists(vNewFile) or DirectoryExists(vNewFile) then
    begin
      S := Node.Text;
      MessageBox(Handle, 'Ŀ���ļ��Ѿ����ڡ�', '����', MB_ICONERROR or MB_OK);
    end else if MoveFile(PChar(vOldFile), PChar(vNewFile)) then
    begin
      FSelectPath := vNewFile;
      Node.Text := S;
      Node.ImageIndex := PathIconIndex(vNewFile);
      Node.SelectedIndex := Node.ImageIndex;
    end else
    begin
      S := Node.Text;
      MessageBox(Handle, '�ļ�����ʧ�ܡ�', '����', MB_ICONERROR or MB_OK);
    end;
  end else
  begin
    S := Node.Text;
    MessageBox(Handle, '������������Ч���ļ���������չ������Ϊ"js"��"vbs"��', '��ʾ',
      MB_ICONINFORMATION or MB_OK);
  end;
end;

procedure TFormIEScript.ActionRenameExecute(Sender: TObject);
begin
  if not Assigned(TreeViewScriptList.Selected) then Exit;
  if TreeViewScriptList.CanFocus then TreeViewScriptList.SetFocus;
  TreeViewScriptList.ReadOnly := False;
  FEditorShowing := True;
  TreeViewScriptList.Selected.EditText();
  FEditorShowing := False;
end;

procedure TFormIEScript.TreeViewScriptListMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
(*
  I: Integer;
  vFileName: string;
*)
  vTreeNode: TTreeNode;
begin
  vTreeNode := TTreeView(Sender).GetNodeAt(X, Y);
  if not Assigned(vTreeNode) then Exit;
  if Button = mbRight then TTreeView(Sender).Selected := vTreeNode;
  if Button <> mbLeft then Exit;
  if TTreeView(Sender).SelectionCount <= 0 then Exit;
  (*if not DragDetectPlus(TWinControl(Sender).Handle, Point(X, Y)) then Exit;
  DropFileSourceOne.Files.Clear;                                                //2008-03-21 ZswangY37 No.2
  for I := 0 to TTreeView(Sender).SelectionCount - 1 do
  begin
    vFileName := GetNodePath(TTreeView(Sender).Selections[I]);
    if FileExists(vFileName) then
      DropFileSourceOne.Files.Add(vFileName);
  end;
  if DropFileSourceOne.Files.Count <= 0 then Exit;
  DropFileSourceOne.Execute;*)
end;

procedure TFormIEScript.ActionNewScriptExecute(Sender: TObject);                //2008-03-25 Zswangy37 No.1
var
  vExt: string;
  vPath: string;
  vFileName: string;
  I: Integer;
  vTreeNode: TTreeNode;
begin
  vTreeNode := nil;
  if (Sender = ActionNewJavaScript) or (Sender = ActionNewVBScript) then
  begin // �����ű�
    if not ((TreeViewScriptList.Items.Count <= 0) or
      Assigned(TreeViewScriptList.Selected)) then Exit;

    if (TreeViewScriptList.Items.Count <= 0) or
      Assigned(TreeViewScriptList.Selected.Data) then // ��Ŀ¼
      vTreeNode := TreeViewScriptList.Selected
    else vTreeNode := TreeViewScriptList.Selected.Parent;

    vPath := GetNodePath(vTreeNode);
    if Sender = ActionNewJavaScript then
      vExt := '.js'
    else vExt := '.vbs';
    I := 0;
    vFileName := Format('%s\�½�%.3d%s', [vPath, I, vExt]);
    while FileExists(vFileName) do
    begin
      Inc(I);
      vFileName := Format('%s\�½�%.3d%s', [vPath, I, vExt]);
    end;
    FileClose(FileCreate(vFileName));
  end else if (Sender = ActionNewSibling) or (Sender = ActionNewChildren) then // ����Ŀ¼
  begin
    if Sender = ActionNewSibling then // ����ͬ��Ŀ¼
    begin
      if not ((TreeViewScriptList.Items.Count <= 0) or
        Assigned(TreeViewScriptList.Selected)) then Exit;
      if Assigned(TreeViewScriptList.Selected) then // ��Ŀ¼
        vTreeNode := TreeViewScriptList.Selected.Parent;
    end else // ������Ŀ¼
    begin
      if not (Assigned(TreeViewScriptList.Selected) and
        Assigned(TreeViewScriptList.Selected.Data)) then Exit;
      vTreeNode := TreeViewScriptList.Selected
    end;
    vPath := GetNodePath(vTreeNode);
    I := 0;
    vFileName := Format('%s\�½�%.3d', [vPath, I]);
    while DirectoryExists(vFileName) do
    begin
      Inc(I);
      vFileName := Format('%s\�½�%.3d', [vPath, I]);
    end;
    ForceDirectories(vFileName);
  end;
  vTreeNode := TreeViewScriptList.Items.AddChild(
    vTreeNode, ExtractFileName(vFileName));
  vTreeNode.Data := Pointer(Ord(TAction(Sender).Tag < 0));
  vTreeNode.ImageIndex := PathIconIndex(vFileName);
  vTreeNode.SelectedIndex := vTreeNode.ImageIndex;
  TreeViewScriptList.Selected := vTreeNode;
  TreeViewScriptList.Selected.MakeVisible;
end;

procedure TFormIEScript.ActionCopyEvalExecute(Sender: TObject);
begin
  if MemoScriptEditor.GetTextLen <= 0 then Exit; // �ű�������
  Clipboard.AsText := EvalScript;
end;

function TFormIEScript.EvalScript: string;
var
  S: WideString;                                                                //2008-06-11 ZswangY37 No.1
  I: Integer;
begin
  Result := '';
  if MemoScriptEditor.GetTextLen <= 0 then Exit; // �ű�������
  S := Trim(MemoScriptEditor.Text);
  if Length(S) <= 0 then Exit;
  S := StringReplace(S, #9, #32, [rfReplaceAll]);
  S := StringReplace(S, #13, #32, [rfReplaceAll]);
  S := StringReplace(S, #10, #32, [rfReplaceAll]);
  while Pos(#32#32, S) > 0 do
    S := StringReplace(S, #32#32, #32, [rfReplaceAll]);
  Result := '';
  for I := 1 to Length(S) do
    if S[I] > #255 then
      Result := Result + Format('\u%.4x', [Ord(S[I])])
    else Result := Result + Format('\x%.2x', [Ord(S[I])]);
  Result := 'javascript:eval("' + Result + '");';
end;

procedure TFormIEScript.ActionDeleteExecute(Sender: TObject);                   //2008-03-24 Zswangy37 No.1
begin
  if not Assigned(TreeViewScriptList.Selected) then Exit;
  with TRecycleDialog.Create(nil) do try
    FileName := GetNodePath(TreeViewScriptList.Selected);
    WindowHandle := Handle;
    if Execute then TreeViewScriptList.Selected.Delete;
  finally
    Free;
  end;
end;

procedure TFormIEScript.MemoScriptEditorKeyPress(Sender: TObject;
  var Key: Char);
var
  vStart, vEnd: Integer; // ��ʼ�У�������
  vRemove: Boolean;
  vSelStart, vSelLength: Integer;
  vColIndex: Integer;
  I, J: Integer;
  S: string;
begin
  case Key of
    #9:
      begin
        if (TMemo(Sender).SelLength > 0) or
          (GetKeyState(VK_SHIFT) and $80 <> 0) then
        begin
          vRemove := GetKeyState(VK_SHIFT) and $80 <> 0;
          vStart := TMemo(Sender).Perform(EM_LINEFROMCHAR,
            TMemo(Sender).SelStart, 0);
          vEnd := TMemo(Sender).Perform(EM_LINEFROMCHAR,
            TMemo(Sender).SelStart + TMemo(Sender).SelLength, 0);
          vSelStart := TMemo(Sender).SelStart;
          vSelLength := TMemo(Sender).SelLength;
          vColIndex := vSelStart - TMemo(Sender).Perform(EM_LINEINDEX, vStart, 0);
          if vSelStart + vSelLength - TMemo(Sender).Perform(EM_LINEINDEX, vEnd, 0) = 0 then
            Dec(vEnd);
          TMemo(Sender).Lines.BeginUpdate;
          for I := vStart to vEnd do
          begin
            if not vRemove then
            begin
              TMemo(Sender).Lines[I] := #9 + TMemo(Sender).Lines[I];

              if (I = vStart) and (vColIndex > 0) then
                Inc(vSelStart)
              else Inc(vSelLength);
            end else
            begin
              S := TMemo(Sender).Lines[I];
              if Copy(S, 1, 1) = #9 then
              begin
                Delete(S, 1, 1);
                if (I = vStart) and (vColIndex > 0) then Dec(vSelStart);
                Dec(vSelLength);
              end else
              begin
                J := 0;
                while (Copy(S, 1, 1) = #32) and (J < 4) do
                begin
                  Delete(S, 1, 1);
                  if (I = vStart) and (vColIndex > 0) then Dec(vSelStart);
                  Dec(vSelLength);
                end;
              end;
              TMemo(Sender).Lines[I] := S;
            end;
          end;
          TMemo(Sender).Lines.EndUpdate;
          TMemo(Sender).SelStart := vSelStart;
          TMemo(Sender).SelLength := Max(vSelLength, 0);
          Key := #0;
        end;
      end;
  end;
  FLastChar := Key;
end;

procedure TFormIEScript.ActionFindTextExecute(Sender: TObject);
var
  S: string;
begin
  S := MemoScriptEditor.SelText;
  if S <> '' then FindDialogOne.FindText := S;
  FindDialogOne.Execute;
end;

procedure TFormIEScript.FindDialogOneFind(Sender: TObject);
begin
  FSearchEvent := FindDialogOneFind;
  FSearchDialog := TFindDialog(Sender);
  if not SearchMemo(MemoScriptEditor, TFindDialog(Sender).FindText,             //2008-05-22 ZswangY37 No.1
    TFindDialog(Sender).Options) then
    MessageDlg(Format('����"%s"û���ҵ���', [TFindDialog(Sender).FindText]),
      mtInformation, [mbOK], 0);
end;

procedure TFormIEScript.ActionReplaceExecute(Sender: TObject);
var
  S: string;
begin
  S := MemoScriptEditor.SelText;
  if S <> '' then ReplaceDialogOne.FindText := S;
  ReplaceDialogOne.Execute;
end;

procedure TFormIEScript.ReplaceDialogOneReplace(Sender: TObject);
var
  vFound: Boolean;
begin
  FSearchEvent := FindDialogOneFind;
  FSearchDialog := TFindDialog(Sender);
  if frReplaceAll in TReplaceDialog(Sender).Options then // ȫ���滻
  begin
    MemoScriptEditor.SelStart := 0;
    vFound := SearchMemo(MemoScriptEditor,
      TReplaceDialog(Sender).FindText, TReplaceDialog(Sender).Options);
    while vFound do
    begin
      MemoScriptEditor.SelText := TReplaceDialog(Sender).ReplaceText;
      vFound := SearchMemo(MemoScriptEditor,
        TReplaceDialog(Sender).FindText, TReplaceDialog(Sender).Options);
    end;
  end else
  begin
    if SameText(MemoScriptEditor.SelText, TReplaceDialog(Sender).FindText) then
      MemoScriptEditor.SelText := TReplaceDialog(Sender).ReplaceText;
    FindDialogOneFind(Sender);
  end;
end;

procedure TFormIEScript.MemoScriptEditorKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_A: if ssCtrl in Shift then TMemo(Sender).SelectAll;
  end;
end;

procedure TFormIEScript.ActionExecuteScriptTimerExecute(Sender: TObject);       //2008-11-15 ZswangY37 No.1
var
  S: string;
begin
  if TimerOne.Enabled then
    TimerOne.Enabled := False
  else
  begin
    S := '1000';
    if not InputQuery('����', '���������', S) then Exit;
    TimerOne.Interval := StrToIntDef(S, 0);
    TimerOne.Enabled := True;
  end;
end;

procedure TFormIEScript.TimerOneTimer(Sender: TObject);
begin
  ActionExecuteScript.Execute;
end;

procedure TFormIEScript.ActionNoForegroundExecute(Sender: TObject);
begin
  ;
end;

procedure TFormIEScript.ActionCodingAnsiExecute(Sender: TObject);
begin
  if TAction(Sender).Checked then Exit;
  TAction(Sender).Checked := True;
  FEncoding := 'ansi';
  FScriptChanging := True;
end;

procedure TFormIEScript.ActionCodingUtf8Execute(Sender: TObject);
begin
  if TAction(Sender).Checked then Exit;
  TAction(Sender).Checked := True;
  FEncoding := 'utf8';
  FScriptChanging := True;
end;

procedure TFormIEScript.ActionCodingUtf8NoneBOMExecute(Sender: TObject);
begin
  if TAction(Sender).Checked then Exit;
  TAction(Sender).Checked := True;
  FEncoding := '~utf8';
  FScriptChanging := True;
end;

procedure TFormIEScript.ActionCodingUnicodeExecute(Sender: TObject);
begin
  if TAction(Sender).Checked then Exit;
  TAction(Sender).Checked := True;
  FEncoding := 'unicode';
  FScriptChanging := True;
end;

procedure TFormIEScript.TreeViewScriptListChanging(Sender: TObject;
  Node: TTreeNode; var AllowChange: Boolean);                                   //2008-12-14 ZswangY37 No.1
begin
  if (Node = FLastChangingNode) then Exit;
  FLastChangingNode := Node;
  if FScriptChanging then
  begin
    case Application.MessageBox('�Ƿ񱣴��Ѿ��޸ĵĽű���',
        'ѯ��', MB_YESNOCANCEL) of
      IDOK, IDYES: ActionSave.Execute;
      IDCANCEL: AllowChange := False;
    else
    end;
  end;
end;

procedure TFormIEScript.ActionEditCutExecute(Sender: TObject);
begin
  MemoScriptEditor.CutToClipboard;
end;

procedure TFormIEScript.ActionEditCopyExecute(Sender: TObject);
begin
  MemoScriptEditor.CopyToClipboard;
end;

procedure TFormIEScript.ActionEditRedoExecute(Sender: TObject);
begin
  //MemoScriptEditor.Redo;
end;

procedure TFormIEScript.ActionEditSelectAllExecute(Sender: TObject);
begin
  MemoScriptEditor.SelectAll;
end;

procedure TFormIEScript.ActionEditDeleteExecute(Sender: TObject);
begin
  MemoScriptEditor.SelText := '';
end;

procedure TFormIEScript.ActionEditPasteExecute(Sender: TObject);
begin
  MemoScriptEditor.PasteFromClipboard;
end;

procedure TFormIEScript.ActionEditUndoExecute(Sender: TObject);
begin
  MemoScriptEditor.Undo;
end;

procedure TFormIEScript.ActionSearchAgainExecute(Sender: TObject);
begin
  if Assigned(FSearchEvent) and Assigned(FSearchDialog) then
    FSearchEvent(FSearchDialog);
end;

procedure TFormIEScript.ActionSelectRootExecute(Sender: TObject);
begin
  TreeViewScriptList.Selected := nil;
  TreeViewScriptListChange(TreeViewScriptList, nil);
end;

var
  vBuffer: array[0..MAX_PATH] of Char;

initialization

  GetModuleFileName(HInstance, vBuffer, SizeOf(vBuffer)); // ��ȡģ�������ļ�
  GetFileVersionInfomation(vBuffer, vModuleVersionInfomation); // ��ȡģ��汾��Ϣ
  OleInitialize(nil);

finalization
  OleUninitialize;

end.
