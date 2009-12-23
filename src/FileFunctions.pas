unit FileFunctions;

interface

// 获得文件的编码类型 ~utf8 表示为无BOM的文件
function GetFileEncoding(AFilename: string): string;

function GetFileTextAndEncoding(AFilename: string; var AEncoding: string): string;

function SetFileTextAndEncoding(AFilename: string; AEncoding: string; AText: string): Boolean;

// 判断缓冲区是否为utf8编码
function isTextUTF8(buffer: PChar; len: Integer): Boolean;

implementation

uses SysUtils, Classes, Windows;

function GetFileEncoding(AFileName: string): string;
var
  S: string;
  L: Integer;
begin
  Result := 'ansi';
  if not FileExists(AFileName) then Exit;
  with TMemoryStream.Create do try
    LoadFromFile(AFileName);
    if Size < 3 then Exit;
    SetLength(S, 3);
    Read(S[1], Length(S));
    if Copy(S, 1, 2) = #$FF#$FE then
    begin
      Result := 'unicode';
      Exit;
    end;
    if Copy(S, 1, 3) = #$EF#$BB#$BF then
    begin
      Result := 'utf8';
      Exit;
    end;
    Position := 0;
    SetLength(S, $1000);
    L := Read(S[1], Length(S));
    if isTextUTF8(@S[1], L) then
    begin
      Result := '~utf8';
      Exit;
    end;
  finally
    Free;
  end;
end;

function isTextUTF8(buffer: PChar; len: Integer): Boolean;
var
  octets: Integer;
  I: Integer;
  chr: Byte;
  ascii: Boolean;
begin
  octets := 0;
  ascii := True;
  Result := False;
  for I := 0 to len - 1 do
  begin
    chr := Byte(buffer[I]);
    if chr and $80 <> 0 then ascii := False;
    if octets = 0 then
    begin
      if chr >= $80 then
      begin
	chr := Byte(chr * 2);
	Inc(octets);
	while (chr and $80 <> 0) do
        begin
	  chr := Byte(chr * 2);
	  Inc(octets);
        end;
	Dec(octets); // count includes this character
	if octets = 0 then Exit; // must start with 11xxxxxx
      end;  
    end else
    begin
      if chr and $c0 <> $80 then Exit; // non-leading bytes must start as 1$xxxxx
      Dec(octets); // processed another octet in encoding
    end;
  end;
  Result := (octets <= 0) and not ascii;
end;

function GetFileTextAndEncoding(AFileName: string; var AEncoding: string): string;
var
  W: WideString;
  S: string;
  L: Integer;
begin
  Result := '';
  AEncoding := 'ansi';
  if not FileExists(AFileName) then Exit;
  AEncoding := GetFileEncoding(AFileName);
  if AEncoding = 'ansi' then
  begin
    with TMemoryStream.Create do try
      LoadFromFile(AFileName);
      if Size <= 0 then Exit;
      SetLength(Result, Size);
      Read(Result[1], Length(Result));
    finally
      Free;
    end;
    Exit;
  end;

  if AEncoding = 'unicode' then
  begin
    with TMemoryStream.Create do try
      LoadFromFile(AFileName);
      L := Size;
      Position := 2;
      L := L - 2;
      if L <= 0 then Exit;
      SetLength(W, L div SizeOf(WideChar));
      Read(W[1], Length(W) * SizeOf(WideChar));
      Result := W;
    finally
      Free;
    end;
    Exit;
  end;

  if (AEncoding = 'utf8') or (AEncoding = '~utf8') then
  begin
    with TMemoryStream.Create do try
      LoadFromFile(AFileName);
      L := Size;
      if AEncoding = 'utf8' then Position := 3;
      if AEncoding = 'utf8' then L := L - 3;
      if L <= 0 then Exit;
      SetLength(S, L);
      Read(S[1], Length(S));
      Result := Utf8ToAnsi(S);
    finally
      Free;
    end;
    Exit;
  end;
end;

function SetFileTextAndEncoding(AFileName: string;
  AEncoding: string; AText: string): Boolean;
var
  S: string;
  W: WideString;
begin
  if AEncoding = 'ansi' then
  begin
    with TMemoryStream.Create do try
      Write(AText[1], Length(AText));
      SaveToFile(AFileName);
    finally
      Free;
    end;
  end;
  if AEncoding = 'unicode' then
  begin
    with TMemoryStream.Create do try
      S := #$FF#$FE;
      Write(S[1], Length(S));
      W := AText;
      Write(W[1], Length(W) * SizeOf(WideChar));
      SaveToFile(AFileName);
    finally
      Free;
    end;
  end;
  if (AEncoding = 'utf8') or (AEncoding = '~utf8') then
  begin
    with TMemoryStream.Create do try
      if AEncoding = 'utf8' then
      begin
        S := #$EF#$BB#$BF;
        Write(S[1], Length(S));
      end;
      S := AnsiToUtf8(AText);;
      Write(S[1], Length(S));
      SaveToFile(AFileName);
    finally
      Free;
    end;
  end;
  Result := True;
end;

end.
