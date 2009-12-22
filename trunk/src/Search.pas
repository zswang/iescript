unit Search;

interface

uses Windows, SysUtils, StdCtrls, Dialogs;

const
  cWordDelimiters: set of Char = [#0..#255] - ['a'..'z','A'..'Z','1'..'9','0'];

function SearchMemo( // 搜索编辑框中的字符串
  AMemo: TCustomEdit; // 编辑框
  ASearchString: string; // 搜索的字符串
  AOptions: TFindOptions // 选项
): Boolean; // 返回是否搜索到结果

function SearchBuf( // 搜索缓冲区的字符串
  ABuf: PChar; // 缓冲区
  ABufLen: Integer; // 缓冲区大小
  ASelStart, ASelLength: Integer; // 选择的位置和长度
  ASearchString: string; // 搜索的字符串
  AOptions: TFindOptions // 搜索选项
): PChar; // 返回搜索到的位置

implementation


function SearchMemo( // 搜索编辑框中的内容
  AMemo: TCustomEdit; // 编辑框
  ASearchString: string; // 搜索的字符串
  AOptions: TFindOptions // 选项
): Boolean; // 返回是否搜索到结果
var
  Buffer, P: PChar;
  Size: Word;
begin
  Result := False;
  if not Assigned(AMemo) then Exit;
  if Length(ASearchString) <= 0 then Exit;
  Size := AMemo.GetTextLen;
  if (Size = 0) then Exit;
  Buffer := StrAlloc(Size + 1);
  try
    AMemo.GetTextBuf(Buffer, Size + 1);
    P := SearchBuf(Buffer, Size, AMemo.SelStart, AMemo.SelLength,
      ASearchString, AOptions);
    if P <> nil then
    begin
      AMemo.SelStart := P - Buffer;
      AMemo.SelLength := Length(ASearchString);
      Result := True;
    end;
  finally
    StrDispose(Buffer);
  end;
end; { SearchMemo }

function SearchBuf( // 搜索缓冲区的字符串
  ABuf: PChar; // 缓冲区
  ABufLen: Integer; // 缓冲区大小
  ASelStart, ASelLength: Integer; // 选择的位置和长度
  ASearchString: string; // 搜索的字符串
  AOptions: TFindOptions // 搜索选项
): PChar; // 返回搜索到的位置
var
  SearchCount, I: Integer;
  C: Char;
  Direction: Shortint;
  CharMap: array [Char] of Char;

  function FindNextWordStart(var BufPtr: PChar): Boolean;
  begin
    while (SearchCount > 0) and
      ((Direction = 1) xor (BufPtr^ in cWordDelimiters)) do
    begin
      Inc(BufPtr, Direction);
      Dec(SearchCount);
    end;
    while (SearchCount > 0) and
      ((Direction = -1) xor (BufPtr^ in cWordDelimiters)) do
    begin
      Inc(BufPtr, Direction);
      Dec(SearchCount);
    end;
    Result := SearchCount >= 0;
    if (Direction = -1) and (BufPtr^ in cWordDelimiters) then
    begin
      Dec(BufPtr, Direction);
      Inc(SearchCount);
    end;
  end;

begin
  Result := nil;
  if ABufLen <= 0 then Exit;
  if Length(ASearchString) <= 0 then Exit;
  if frDown in AOptions then // 向后搜索
  begin
    Direction := 1;
    Inc(ASelStart, ASelLength);
    SearchCount := ABufLen - ASelStart - Length(ASearchString);
    if SearchCount < 0 then Exit;
    if Longint(ASelStart) + SearchCount > ABufLen then Exit;
  end else
  begin
    Direction := -1;
    Dec(ASelStart, Length(ASearchString));
    SearchCount := ASelStart;
  end;
  if (ASelStart < 0) or (ASelStart > ABufLen) then Exit;
  Result := @ABuf[ASelStart];

  for C := Low(CharMap) to High(CharMap) do
    CharMap[C] := C;

  if not (frMatchCase in AOptions) then
  begin
    AnsiUpperBuff(PChar(@CharMap), sizeof(CharMap));
    AnsiUpperBuff(@ASearchString[1], Length(ASearchString));
  end;

  while SearchCount >= 0 do
  begin
    if frWholeWord in AOptions then
      if not FindNextWordStart(Result) then Break;
    I := 0;
    while CharMap[Result[I]] = ASearchString[I + 1] do
    begin
      Inc(I);
      if I >= Length(ASearchString) then
      begin
        if (not (frWholeWord in AOptions)) or
          (SearchCount = 0) or
          (Result[I] in cWordDelimiters) then
          Exit;
        Break;
      end;
    end;
    Inc(Result, Direction);
    Dec(SearchCount);
  end;
  Result := nil;
end; { SearchBuf }

end.

