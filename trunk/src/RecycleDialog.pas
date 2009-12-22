(*//
标题：将文件放到回收站的对话框
说明：调用Shell函数，避免不能模态显示
设计：ZswangY37
日期：2008-03-24
支持：wjhu111#21cn.com
//*)

unit RecycleDialog;

interface

uses
  Windows, Messages, SysUtils, Classes, Dialogs;

type
  TRecycleDialog = class(TCommonDialog)
  private
    { Private declarations }
    FFileName: string;
    FWindowHandle: THandle;
  protected
    { Protected declarations }
  public
    { Public declarations }
    property WindowHandle: THandle read FWindowHandle write FWindowHandle;
    function Execute: Boolean; override;
  published
    { Published declarations }
    property FileName: string read FFileName write FFileName;
  end;

procedure Register;

implementation

uses ComObj, SHDocVw, ShellAPI;

procedure Register;
begin
  RegisterComponents('Zswang', [TRecycleDialog]);
end;

{ TRecycleDialog }

type
  TDialogData = packed record
    rWindowHandle: THandle;
    rFileName: string;
  end;
  PDialogData = ^TDialogData;

function DoExecute(var ADialogData: TDialogData): Bool; stdcall;
var
  vSHFileOpStruct: TSHFileOpStruct;
  vBuffer: array[0..MAX_PATH] of Char;
begin
  Result := False;
  if not FileExists(ADialogData.rFileName) and
    not DirectoryExists(ADialogData.rFileName) then Exit;
  FillChar(vBuffer, SizeOf(vBuffer), 0);
  Move(ADialogData.rFileName[1], vBuffer[0], Length(ADialogData.rFileName));
  FillChar(vSHFileOpStruct, SizeOf(vSHFileOpStruct), 0);
  vSHFileOpStruct.wFunc := FO_DELETE;
  vSHFileOpStruct.pFrom := vBuffer;
  vSHFileOpStruct.fFlags := FOF_ALLOWUNDO;
  Result := (SHFileOperation(vSHFileOpStruct) = 0) and
    not vSHFileOpStruct.fAnyOperationsAborted;
end;

function TRecycleDialog.Execute: Boolean;
var
  vDialogData: TDialogData;
begin
  FillChar(vDialogData, SizeOf(vDialogData), 0);
  vDialogData.rWindowHandle := FWindowHandle;
  vDialogData.rFileName := FileName;
  Result := TaskModalDialog(@DoExecute, vDialogData);
end;

end.
