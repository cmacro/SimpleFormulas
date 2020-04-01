unit uFormulas;

///
///  表达式解析
///
///
///  Check(Calc(' 1') = 1, '基础公式错误');
///  Check(Calc('1.1') = 1.1, '基础公式错误');
///  Check(Calc(' 1 + 1 ') = 2, '基础计算错误');
///  Check(Calc('pi') = Unassigned, '变量大小写敏感'); //
///  Check(Calc('PI') = 3.14159265358979, '变量大小写敏感'); //
///  Check(Calc('sin(PI/2)') = 1, '一次函数计算错误'); //
///  Check(Calc('min(1, 2)') = 1, '两次函数计算错误');
///  Check(Calc('iif(1<>1, 1, 2)') = 2, '多次次函数计算错误'); //
///  Check(Calc('-1 + -1') = -2, '前导符号计算错误');
///  Check(Calc('100%') = 1, '百分比计算错误');
///  Check(Calc('100% + 1') = 2, '百分比计算错误');
///  Check(Calc('100% + 100%') = 2, '百分比计算错误');
/// -- moguf.com

interface

uses
  SysUtils, Classes, Variants, Generics.Collections;

type
  ExpValue = Variant;
  PExpValue = ^ExpValue;

const
  MAX_NAME_LEN = 255;

type
  EExpParserError = class(Exception);
  ECalcError = class(Exception);

  IExpVariable = interface
    ['{8AF79710-8FDB-4841-A4F2-D01896C871F2}']
    function GetVarValue(AName: PChar; var Value: Variant): boolean; stdcall;
  end;

  TFuncParamType = (fptNoParam, fptUnary, fptBinary, fptTernary, fptMulti);

  TUnaryFunc = function(const x: ExpValue): ExpValue;
  TBinaryFunc = function(const x: ExpValue; const y: ExpValue): ExpValue;
  TTernaryFunc = function (const x, y, z: ExpValue): ExpValue;
  TMultiFunc = function(const x: array of ExpValue): ExpValue;
  TVarFunEvent = procedure(Sender: TObject; const AName: string; var Value: ExpValue) of object;
  TBuildVarNodeEvent = procedure (Sender: TObject; const AName: string; var AEvent: TVarFunEvent) of object;

  // 计算节点
  TCalcNode = class
  public
    function GetValue: ExpValue; virtual; abstract;
    function IsUsed(Addr: Pointer): boolean; virtual; abstract;
    procedure Optimize; virtual;
    property Value: ExpValue read GetValue;
  end;
  TExpValueArray = array of ExpValue;

  // 计算公式
  TFormula = class
  private
    FOwner: TObject;
    FExpression: string;
    FDirty: boolean;
    FNode: TCalcNode;
    FExpVariable: IExpVariable;
    FVariables: TDictionary<string, ExpValue>;
    FOptimization: boolean;
    FOnVarValue: TVarFunEvent;
    FOnBuildVarNode: TBuildVarNodeEvent;
    procedure DoOnGetVarValue(Sender: TObject; const AName: string; var Value: ExpValue);
  protected
    function GetExpression: string;
    procedure SetExpression(const str: string);
    function FindLastOper(const s: string): integer;

    function BuildCalcNode(var s: string; var ANode: TCalcNode): boolean;
    function IsUnaryOperFunc(const s: string; var AParam: string; var funcAddr: Pointer; CurrChar: integer): boolean;
    function IsBinaryOperFunc(const s: string; var ALParam, ARParam: string; var funcAddr: Pointer; AOperIdx: integer): boolean;
    function IsVariableName(const s: string): Boolean;
    procedure Optimize;
    procedure WriteErr(ACode: integer; const AMsg: string);
  public
    ErrCode: Integer;
    ErrMsg: string;
    constructor Create(AOwner: TObject); virtual;
    destructor Destroy; override;
    procedure Parse; virtual;
    function  Evaluate: ExpValue; virtual;
    procedure CreateVar(const AName: string; const AValue: ExpValue); virtual;
    procedure AssignVar(const AName: string; const val: ExpValue); virtual;
    function  GetVar(const AName: string): ExpValue; virtual;
    procedure DeleteVar(const AName: string); virtual;
    procedure GetFunctions(AList: TStrings); virtual;
    function  GetVariables(AList: TStrings): Integer; virtual;
    procedure Randomize; virtual;
    procedure FreeParseTree; virtual;
    function  IsFunctionUsed(const funcName: string): boolean; virtual;

    function  Calc(const s: string): ExpValue;

    property Value: ExpValue read Evaluate;
    property Variable[const name: string]: ExpValue read GetVar write AssignVar;
    function IsParamFunc(const s: string; var AFunName: string; var AParams: TStringList): Boolean;

    property Expression: string read GetExpression write SetExpression;
    property Optimization: boolean read FOptimization write FOptimization default true;
    property OnVarValue: TVarFunEvent read FOnVarValue write FOnVarValue;
    property OnBuildVarNode: TBuildVarNodeEvent read FOnBuildVarNode write FOnBuildVarNode;
  end;

  // 对外注册的函数
  //   全局 函数、变量，常量，接口
  FormulaReg = class
  public
    class procedure AddFun(const AName: string; AFuncAddr: Pointer; AType: TFuncParamType; const ANote: string = '');
    class procedure AddConst(const AName: string; const AValue: ExpValue);
    class procedure AddVar(const AName: string; AFuncAddr: TVarFunEvent);
    class procedure AddIntf(const AIntfName: string; const AIntf: TVarFunEvent);
  end;

  // test
  function _StrContains(const x, y, z: ExpValue): ExpValue;

implementation

uses
  math, StrUtils, untCommFuns;

const
  EXPERR_Parser = 1;    // 解析错误
  EXPERR_InvalidOp = 2; // 除零错误
  EXPERR_Calc = 3;      // 计算错误： 变量错误
  EXPERR_Other = 100;   // 其他错误

type
  TFunItemInfo = class
    Name: string;
    FuncAddr: Pointer;
    ParamType: TFuncParamType;
    Note: string;
  end;

  TFunDict = class
  private
    FList: TObjectList<TFunItemInfo>;
    FNames: TDictionary<String,TFunItemInfo>;
    FConsts: TDictionary<String, Variant>; // 全局常量
    FVars: TDictionary<string,TVarFunEvent>;  // 全局变量
    FIntfs: TDictionary<string,TVarFunEvent>;  // 全局变量回调, 内部会有多个变量，由外部控制
    function GetItems(const AName: string): TFunItemInfo;
  public
    constructor Create;
    destructor Destroy; override;
    procedure GetVarValue(Sender: TObject; const AName: string; var Value: ExpValue);
    function Add(const AName: string; AFuncAddr: Pointer; AType: TFuncParamType; const ANote: string = ''): TFunItemInfo;
    function Find(const AName: string): TFunItemInfo;
    function Exists(const AName: string): boolean;
    function FuncAddr(const AName: string): Pointer;
    property Items[const AName: string]: TFunItemInfo read GetItems; default;

    class procedure RegDefaultFuns;
  end;

var
  FunDict: TFunDict = nil;

function IsValidNameChar(index: integer; c: char): boolean;
begin
  // 检查是否是有效的名称字符
  //   1、前缀不能是数字
  //   2、不能有表达式类型字符
  //
  c := UpCase(c);
  if (index = 1) then
  begin
    if (not ((c < 'A') or (c > 'Z'))) or (c = '_') then
      Result := true
    else
      Result := false;
  end else
  begin
    if (not ((c < '0') or (c > 'Z'))) or (c = '_') then
      Result := true
    else
      Result := false;
  end;
end;

function TrimCopy(const s: string; idx, l: integer): string;
begin
  Result := Trim(Copy(s, idx, l));
end;

function IsValidName(const name: string): boolean;
var i, len: integer;
begin
  len := Length(name);
  if len > 0 then
  begin
    for i := 1 to len do
    begin
      if not IsValidNameChar(i, name[i]) then
      begin
        Result := false;
        exit;
      end;
    end;
  end;
  Result := true;
end;

procedure DisposeList(var List: TList);
var i: Integer;
begin
  for i := List.Count - 1 downto 0 do
  begin
    Dispose(List.Items[i]);
    List.Delete(i);
  end;
  List.Free;
  List := nil;
end;

procedure OptimizeNode(var ANode: TCalcNode); forward;

type
  // 常量节点
  TConstNode = class(TCalcNode)
  private
    FValue: ExpValue;
  public
    constructor Create(const Val: ExpValue);
    destructor Destroy; override;
    function GetValue: ExpValue; override;
    function IsUsed(Addr: Pointer): boolean; override;
  end;

  // 单参数函数
  TUnaryParamNode = class(TCalcNode)
  private
    x: TCalcNode;
    Func: TUnaryFunc;
  public
    constructor Create(AFunc: TUnaryFunc; AXParam: TCalcNode);
    destructor Destroy; override;
    function  GetValue: ExpValue; override;
    function  IsUsed(Addr: Pointer): boolean; override;
    procedure Optimize; override;
  end;

  // 双参数函数 max(x, y)
  TBinaryParamNode = class(TCalcNode)
  private
    x, y : TCalcNode;
    Func: TBinaryFunc;
  public
    constructor Create(AFunc: TBinaryFunc; AXParam, AYParam: TCalcNode);
    destructor Destroy; override;

    function   GetValue: ExpValue; override;
    function   IsUsed(Addr: Pointer): boolean; override;
    procedure  Optimize; override;
  end;

  TTernaryParamNode = class(TCalcNode)
  private
    x, y, z : TCalcNode;
    Func: TTernaryFunc;
  public
    constructor Create(AFunc: TTernaryFunc; AXParam, AYParam, AZParam: TCalcNode);
    destructor Destroy; override;
    function  GetValue: ExpValue; override;
    function  IsUsed(Addr: Pointer): Boolean; override;
    procedure Optimize; override;
  end;

  // 多参数函数
  TMultiParamNode = class(TCalcNode)
  private
    Params: TList;
    Func: TMultiFunc;
  public
    constructor Create(AFunc: TMultiFunc; ANodes: TList);
    destructor Destroy; override;

    function  GetValue: ExpValue; override;
    function  IsUsed(Addr: Pointer): boolean; override;
    procedure Optimize; override;
  end;

  // 变量节点
  TVarNode = class(TCalcNode)
  private
    FName: string;
    FFunc: TVarFunEvent;
  public
    constructor Create(const AName: string; AFunc: TVarFunEvent);
    destructor  Destroy; override;

    function  GetValue: ExpValue; override;
    function  IsUsed(Addr: Pointer): boolean; override;
  end;

function _greater(const x, y: ExpValue): ExpValue;
begin
  if (x > y) then Result := 1
  else Result := 0;
end;

function _less(const x, y: ExpValue): ExpValue;
begin
  if (x < y) then Result := 1
  else Result := 0;
end;

function _equal(const x, y: ExpValue): ExpValue;
begin
  if (x = y) then Result := 1
  else Result := 0;
end;

function _ltequals(const x, y: ExpValue): ExpValue;
begin
  if (x <= y) then Result := 1
  else Result := 0;
end;

function _gtequals(const x, y: ExpValue): ExpValue;
begin
  if (x >= y) then Result := 1
  else Result := 0;
end;

function _notequals(const x, y: ExpValue): ExpValue;
begin
  if (x <> y) then Result := 1
  else Result := 0;
end;

function _and(const x, y: ExpValue): ExpValue;
begin
  if ((x <> 0) and (y <> 0)) then Result := 1
  else Result := 0;
end;

function _or(const x, y: ExpValue): ExpValue;
begin
  if ((x <> 0) or (y <> 0)) then Result := 1
  else Result := 0;
end;

function _not(const x: ExpValue): ExpValue;
begin
  if (x = 0) then Result := 1
  else Result := 0;
end;

function _unaryadd(const x: ExpValue): ExpValue;
begin
  Result := x;
end;

function _add(const x, y: ExpValue): ExpValue;
begin
  if VarIsNumeric(x) and VarIsNumeric(y) then
    Result := x + y
  else
    Result := VarToStrDef(x, '') + VarToStrDef(y, '');
end;

function _subtract(const x, y: ExpValue): ExpValue;
begin
  Result := x - y;
end;

function _multiply(const x, y: ExpValue): ExpValue;
begin
  Result := x * y;
end;

function _divide(const x, y: ExpValue): ExpValue;
begin
  Result := x / y;
end;

function _modulo(const x, y: ExpValue): ExpValue;
begin
  Result := trunc(x) mod trunc(y);
end;

function _intdiv(const x, y: ExpValue): ExpValue;
begin
  Result := trunc(x) div trunc(y);
end;

function _negate(const x: ExpValue): ExpValue;
begin
  Result := -x;
end;

function _percentage(const x: ExpValue): ExpValue;
begin
  Result := x / 100;
end;

function _intpower(const x, y: ExpValue): ExpValue;
begin
  Result := IntPower(x, trunc(y));
end;

function _square(const x: ExpValue): ExpValue;
begin
  Result := sqr(x);
end;

function _power(const x, y: ExpValue): ExpValue;
begin
  Result := Power(x, y);
end;

function _sin(const x: ExpValue): ExpValue;
begin
  Result := sin(x);
end;

function _cos(const x: ExpValue): ExpValue;
begin
  Result := cos(x);
end;

function _arctan(const x: ExpValue): ExpValue;
begin
  Result := arctan(x);
end;

function _sinh(const x: ExpValue): ExpValue;
begin
  Result := (exp(x) - exp(-x)) * 0.5;
end;

function _cosh(const x: ExpValue): ExpValue;
begin
  Result := (exp(x) + exp(-x)) * 0.5;
end;

function _cotan(const x: ExpValue): ExpValue;
begin
  Result := cotan(x);
end;

function _tan(const x: ExpValue): ExpValue;
begin
  Result := tan(x);
end;

function _exp(const x: ExpValue): ExpValue;
begin
  Result := exp(x);
end;

function _ln(const x: ExpValue): ExpValue;
begin
  Result := ln(x);
end;

function _log10(const x: ExpValue): ExpValue;
begin
  Result := log10(x);
end;

function _logN(const x, y: ExpValue): ExpValue;
begin
  Result := logN(x, y);
end;

function _sqrt(const x: ExpValue): ExpValue;
begin
  Result := sqrt(x);
end;

function _abs(const x: ExpValue): ExpValue;
begin
  Result := abs(x);
end;

function _min(const x, y: ExpValue): ExpValue;
begin
  if x < y then
    Result := x
  else
    Result := y;
end;

function _max(const x, y: ExpValue): ExpValue;
begin
  if x > y then
    Result := x
  else
    Result := y;
end;

function _sign(const x: ExpValue): ExpValue;
begin
  if x < 0 then
    Result := -1
  else if x > 0 then
    Result := 1.0
  else
    Result := 0.0;
end;

function _round(const x,y: ExpValue): ExpValue;
var
  v,f: Extended;
begin
  // 小数四舍五入
  //
  //  注： 0.1045  取 3位小数 SimpleRound中Trunc会变成 0.104问题
  v := x;
  f := IntPower(10, -y);
  //  Delphi Bug： 增加后缀振荡，防止出现错误
  if v < 0 then v := (v / f) - 0.5000000009
  else v := (v / f) + 0.5000000009;
  Result := Trunc(v) * f;
end;

function _trunc(const x: ExpValue): ExpValue;
begin
  Result := int(x);
end;

function _ceil(const x: ExpValue): ExpValue;
begin
  if frac(x) > 0 then
    Result := int(x + 1)
  else
    Result := int(x);
end;

function _floor(const x: ExpValue): ExpValue;
begin
  if frac(x) < 0 then
    Result := int(x - 1)
  else
    Result := int(x);
end;

function _rnd(const x: ExpValue): ExpValue;
begin
  Result := int(Random * int(x));
end;

function _random(const x: ExpValue): ExpValue;
begin
  Result := Random * x;
end;

function _iif(const x, y, z: ExpValue): ExpValue;
begin
  if x <> 0 then Result := y
  else Result := z;
end;

function _StrPos(const x, y: ExpValue): ExpValue;
var
  s,sub: string;
begin
  s := VarToStrDef(x, '');
  sub := VarToStrDef(y, '');
  Result := Pos(sub, s);
end;

function _StrCopy(const x, y, z: ExpValue): ExpValue;
var
  iIdx, iLen: Integer;
  s: string;
begin
  s := VarToStrDef(x, '');
  iIdx := 0;
  if VarIsNumeric(y) then
    iIdx := Trunc(y);
  iLen := 0;
  if VarIsNumeric(z) then
    iLen := Trunc(z);

  if (iIdx <=0) or (iLen <= 0) or (s = '') then
    Result := ''
  else
    Result := Copy(s, iIdx, iLen);
end;

function _StrContains(const x, y, z: ExpValue): ExpValue;
var
  bInStr: Boolean;
  I: integer;
  iSubLen, iLen: Integer;
  iWord: Integer;
  s, sub: string;
  sTmp: string;
  sp: Char;
begin
  ///
  ///  参数：
  ///    第一个： 子字符串
  ///    第二个： 列表字符串
  ///    第三个： 字符串分割字符， 缺省 ',' 逗号
  Result := 0;
  sub := VarToStrDef(x, '');
  s := VarToStrDef(y, '');
  if (s = '') or (sub = '') then Exit;

  sp := ',';
  sTmp := VarToStrDef(z, '');
  if  sTmp <> '' then sp := sTmp[1];

  iLen := Length(s);
  iSubLen := Length(sub);
  if iLen = iSubLen then
  begin
    if SameText(s, sub) then
      Result := 1;
    Exit;
  end;

  iWord := 1;
  bInStr := False;

  for I := 1 to iLen do
  begin
    if s[i] = #39 then
      bInStr := not bInStr
    else if s[i] = sp then
    begin
      if bInStr then Continue;
      // a,b =
      if (iSubLen = i - iWord) then
      begin
        if StrLIComp(PChar(@s[iWord]), PChar(sub), iSubLen) = 0 then
        begin
          Result := 1;
          break;
        end;
      end;
      iWord := i+1;
    end;
  end;

  if (Result = 0) and (iSubLen = iLen - iWord + 1) then
  begin
    if StrLIComp(PChar(@s[iWord]), PChar(sub), iSubLen) = 0 then
      Result := 1;
  end;
end;

function _StrLeft(const x, y: ExpValue): ExpValue;
var
  cnt: Integer;
  s: string;
begin
  s := VarToStrDef(x, '');
  cnt := 0;
  if VarIsNumeric(y) then cnt := Trunc(y);
  Result := LeftStr(s, cnt);
end;

function _StrRight(const x, y: ExpValue): ExpValue;
var
  cnt: Integer;
  s: string;
begin
  s := VarToStrDef(x, '');
  cnt := 0;
  if VarIsNumeric(y) then cnt := Trunc(y);
  Result := RightStr(s, cnt);
end;


{ TConstNode }

constructor TConstNode.Create(const Val: ExpValue);
begin
  inherited Create;
  FValue := Val;
end;

destructor TConstNode.Destroy;
begin
  inherited Destroy;
end;

function TConstNode.GetValue;
begin
  Result := FValue;
end;

function TConstNode.IsUsed(Addr: Pointer): boolean;
begin
  Result := false;
end;

{ TUnaryParamNode }

constructor TUnaryParamNode.Create(AFunc: TUnaryFunc; AXParam: TCalcNode);
begin
  inherited Create;
  Func := AFunc;
  x := AXParam;
end;

destructor TUnaryParamNode.Destroy;
begin
  if Assigned(x) then x.Free;
  inherited Destroy;
end;

function TUnaryParamNode.GetValue;
begin
  Result := Func(x.value);
end;

function TUnaryParamNode.IsUsed(Addr: Pointer): boolean;
begin
  Result := (Addr = @Func) or x.IsUsed(Addr);
end;

procedure TUnaryParamNode.Optimize;
begin
  OptimizeNode(x);
end;

{ TBinaryParamNode }

constructor TBinaryParamNode.Create(AFunc: TBinaryFunc; AXParam, AYParam:
    TCalcNode);
begin
  inherited Create;
  Func := AFunc;
  x := AXParam;
  y := AYParam;
end;

destructor TBinaryParamNode.Destroy;
begin
  if assigned(x) then x.Free;
  if assigned(y) then y.Free;
  inherited Destroy;
end;

function TBinaryParamNode.GetValue;
begin
  Result := Func(x.GetValue, y.GetValue);
end;

function TBinaryParamNode.IsUsed(Addr: Pointer): boolean;
begin
  Result := (Addr = @Func) or x.IsUsed(Addr) or y.IsUsed(Addr);
end;

procedure TBinaryParamNode.Optimize;
begin
  OptimizeNode(x);
  OptimizeNode(y);
end;

{ TMultiParamNode }

constructor TMultiParamNode.Create(AFunc: TMultiFunc; ANodes: TList);
begin
  inherited Create;
  Func := AFunc;
  Params := TList.Create;
  Params.Assign(ANodes);
end;

destructor TMultiParamNode.Destroy;
var i: Integer;
begin
  for i := Params.Count - 1 downto 0 do
    TCalcNode(Params.Items[i]).Free;
  Params.Free;
  inherited Destroy;
end;

function TMultiParamNode.GetValue;
var Values: TExpValueArray;
  i: Integer;
begin
  SetLength(Values, Params.Count);
  try
    for i := 0 to Params.Count - 1 do
    begin
      Values[i] := TCalcNode(Params.Items[i]).GetValue;
    end;
    Result := Func(Values);
  finally
    SetLength(Values, 0);
  end;
end;

function TMultiParamNode.IsUsed(Addr: Pointer): boolean;
var
  i: Integer;
begin
  Result := (Addr = @Func);
  for i := 0 to Params.Count - 1 do
  begin
    if (TCalcNode(Params.Items[i])).IsUsed(Addr) then
    begin
      Result := true;
      exit;
    end;
  end;
end;

procedure TMultiParamNode.Optimize;
var
  i: Integer;
  cNode: TCalcNode;
begin
  for i := 0 to Params.Count - 1 do
  begin
    cNode := TCalcNode(Params.Items[i]);
    OptimizeNode(cNode);
    Params.Items[i] := cNode;
  end;
end;

{ TVarNode }
constructor TVarNode.Create(const AName: string; AFunc: TVarFunEvent);
begin
  inherited Create;
  FName := AName;
  FFunc := AFunc;
end;

destructor TVarNode.Destroy;
begin
  inherited Destroy;
end;

function TVarNode.GetValue: ExpValue;
begin
  FFunc(Self, FName, Result);
end;

function TVarNode.IsUsed(Addr: Pointer): boolean;
begin
  Result := Assigned(FFunc);
end;


function TFormula.Calc(const s: string): ExpValue;
begin
  Expression := s;
  Result := Value;
end;

constructor TFormula.Create(AOwner: TObject);
begin
  FOwner := AOwner;
  // 创建时自动注册默认的函数
  TFunDict.RegDefaultFuns;

  FExpression := '';
  FNode := nil;
  FDirty := true;
  FOptimization := true;

  FVariables:= TDictionary<string, ExpValue>.Create;
  Supports(AOwner, IExpVariable, FExpVariable);
end;

destructor TFormula.Destroy;
begin
  FExpVariable  := nil;
  FreeAndNil(FNode);
  FVariables.Free;
  inherited Destroy;
end;

procedure TFormula.DoOnGetVarValue(Sender: TObject; const AName: string; var
    Value: ExpValue);
begin
  Value := GetVar(AName);
end;

procedure TFormula.Parse;
var
  sExp: string;
begin
  FreeAndNil(FNode);
  if not (Length(FExpression) > 0) then
    WriteErr(EXPERR_Parser, '没有计算公式.');

  sExp := Trim(FExpression);
  if not BuildCalcNode(sExp, FNode) then
    FreeAndNil(FNode);

  if assigned(FNode) and FOptimization then
    Optimize;
  FDirty := false;
end;

function TFormula.Evaluate: ExpValue;
begin
  if (FDirty) then
    Parse;

  // 有错误直接退出
  if ErrCode > 0 then
  begin
    Result := Unassigned;
    Exit;
  end;

  try
    ErrCode := 0;
    Result := FNode.GetValue;
  except
    on e: EInvalidOp do
    begin
      Result := Unassigned;
      WriteErr(EXPERR_InvalidOp, e.ClassName + ':' + e.Message);
    end;
    on e: ECalcError do
    begin
      Result := Unassigned;
      WriteErr(EXPERR_Calc, e.ClassName + ':' + e.Message);
    end;
    on E: Exception do
    begin
      Result := Unassigned;
      WriteErr(EXPERR_Other, e.ClassName + ':' + e.Message);
    end;
  end;
end;

function CheckBrackets(const formula: string): boolean;
var
  i, n: integer;
begin
  // 检查括号是否对齐
  Result := false;
  n := 0;
  for i := 1 to Length(formula) do
  begin
    case formula[i] of
      '(': inc(n);
      ')': dec(n);
    end;
    if n < 0 then exit;
  end;
  Result := (n = 0);
end;

procedure RemoveOuterBrackets(var formula: string);
var temp: string;
  Len: integer;
begin
  // 移除外围括号
  //  ((X+1)-(Y-1))  = (X+1)-(Y-1)
  //  (X+1)-(Y-1)  这种情况不能移除
  // 函数
  //   ShowMessage(copy('hello', 2, 0))  -- 不能移除
  //   Copy('hello', 2, 0)  -- 不能移除
  //
  temp := formula;

  Len := Length(temp);
  while (Len > 2) and (temp[1] = '(') and (temp[Len] = ')') do
  begin
    temp := TrimCopy(temp, 2, Length(temp) - 2);
    if CheckBrackets(temp) then
      formula := temp;
    Len := Length(temp);
  end;
end;

function IsValidConstValue(const s: string; var ExpValue: ExpValue): boolean;
var
  code: integer;
  i: Integer;
  iLen: Integer;
  v: Extended;
begin
  Result := False;
  iLen := Length(s);
  if Length(s) = 0 then
    Exit;

  if (s[1] = #39) and (s[iLen] = #39) then
  begin
    // 字符串类型常量
    code := 0;
    for i := 2 to iLen - 1 do
      if s[i] = #39 then
      begin
        inc(code);
        break;
      end;
    if code = 0 then
    begin
      ExpValue := AnsiDequotedStr(s, #39);
      Result := True;
    end;
  end
  else
  begin
    Val(s, v, code);
    if code = 0 then
    begin
      ExpValue := v;
      Result := True;
    end;
  end;
end;


function TFormula.BuildCalcNode(var s: string; var ANode: TCalcNode): boolean;
var
  val: ExpValue;
  iIdxOper: integer;
  sLP, sRP: string;
  cParams: TStringList;
  cNodes: TList;
  cSubNode, cRightNode: TCalcNode;
  pFuncAddr: Pointer;
  i: Integer;
  pEvent: TVarFunEvent;
  sFunName: string;
  sSubExp: string;

  function GetNodeOf(idx: Integer): TCalcNode;
  begin
    /// 子参数设置
    ///   为简化处理，所有的 1~3元参数 必须全
    ///   没有参数的使用 Null 常量处理
    if cNodes.Count > idx then Result := TCalcNode(cNodes[idx])
    else Result := TConstNode.Create(Null);
  end;

begin
  Result := false;
  RemoveOuterBrackets(s);
  if (not Length(s) > 0) or (not CheckBrackets(s)) then
    exit;

  ANode := nil;
  if IsValidConstValue(s, val) then
    ANode := TConstNode.Create(val)
  else if IsVariableName(s) then
    ANode := TVarNode.Create(s, DoOnGetVarValue)
  else
  begin
    // 函数，一元，二元解析
    iIdxOper := FindLastOper(s);
    if (iIdxOper = 0) and IsParamFunc(s, sFunName, cParams) then
    begin
      pFuncAddr := FunDict[sFunName].FuncAddr;
      if Assigned(pFuncAddr) then
      begin
        cNodes := TList.Create;
        try
          for i := 0 to cParams.Count - 1 do
          begin
            sSubExp := cParams[i];
            if not BuildCalcNode(sSubExp, cSubNode) then
              Break;
            cNodes.Add(cSubNode);
          end;

          /// 创建函数节点
          ///   函数允许有缺省参数，缺省传入为Null
          if cParams.Count = cNodes.Count then
          begin
            case FunDict[sFunName].ParamType of
              fptUnary    : ANode := TUnaryParamNode.Create(pFuncAddr, GetNodeOf(0));
              fptBinary   : ANode := TBinaryParamNode.Create(pFuncAddr, GetNodeOf(0), GetNodeOf(1));
              fptTernary  : ANode := TTernaryParamNode.Create(pFuncAddr, GetNodeOf(0), GetNodeOf(1), GetNodeOf(2));
              fptMulti    : ANode := TMultiParamNode.Create(pFuncAddr, cNodes);
              else ANode := TMultiParamNode.Create(pFuncAddr, cNodes);
            end;
          end
          else
          begin
            // 清除创建不成功的参数节点
            for i := 0 to cNodes.count - 1 do
              TCalcNode(cNodes[i]).Free;
          end;
        finally
          cNodes.free;
          cParams.free;
        end;
      end;
    end
    else if (iIdxOper > 0) and IsUnaryOperFunc(s, sLP, pFuncAddr, iIdxOper) then
    begin
      // 一元和二元操作符
      if BuildCalcNode(sLP, cSubNode) then
        ANode := TUnaryParamNode.Create(pFuncAddr, cSubNode);
    end
    else if (iIdxOper > 0) and IsBinaryOperFunc(s, sLP, sRP, pFuncAddr, iIdxOper) then
    begin
      if BuildCalcNode(sLP, cSubNode) and BuildCalcNode(sRP, cRightNode) then
        ANode := TBinaryParamNode.Create(pFuncAddr, cSubNode, cRightNode);
    end;
  end;

  if not Result and Assigned(OnBuildVarNode) then
  begin
    pEvent := nil;
    OnBuildVarNode(Self, s, pEvent);
    if assigned(pEvent) then
      ANode := TUnaryParamNode.Create(pFuncAddr, cSubNode);
  end;

  Result := ANode <> nil;
end;


function TFormula.FindLastOper(const s: string): integer;
var
  i, j, iBracketCount, iPrecedence, prevOperIndex: Integer;
  bInString: Boolean;
  ch: Char;
  iLastIdx: Integer;
  iCurrIdx: integer;

  function CheckOper(idx: integer; level: integer): boolean;
  begin
    Result := False;
    if (iLastIdx = 0) then
      Result := iPrecedence >= level
    else if s[iLastIdx] = '%' then // 后缀操作符号
      Result := True
  end;
  function CompSetOper(idx: integer; level:integer; saveOper: boolean = True): boolean;
  begin
    Result := False;
    if CheckOper(idx, level) then
    begin
      iPrecedence := level;
      iCurrIdx := i;
    end;
    if saveOper then
      iLastIdx := idx;
  end;

begin
  ///
  /// 获取最低优先级操作符号
  ///

  iPrecedence := 13; // 最高12级优先等级
  iBracketCount := 0;

  bInString := False;
  iLastIdx := 0;
  iCurrIdx := 0;

  for i := 1 to Length(s) do
  begin
    // 字符串解析
    if bInString then
    begin
      if (s[i] = #39) then bInString := False;
      Continue;
    end
    else if iBracketCount > 0 then
    begin
      // 括号优先
      if s[i] = ')' then dec(iBracketCount)
      else if s[i] = '(' then inc(iBracketCount);
      continue;
    end;

    case s[i] of
      ' ', #13,#10: ; // 忽略空格
      #39: begin bInString := True; iLastIdx := 0; end;
      ')': begin dec(iBracketCount); iLastIdx := 0; end;
      '(': begin inc(iBracketCount); iLastIdx := 0; end;

      '|': CompSetOper(i, 1);
      '&': CompSetOper(i, 2);
      '!': CompSetOper(i, 3);
      '=':
      begin
        CompSetOper(i, 4, False);
        if (iLastIdx > 0) then
        begin
          prevOperIndex := i - iLastIdx;
          if not ((s[prevOperIndex] = '<') or (s[prevOperIndex] = '>')) then
            iLastIdx := i;
        end
        else iLastIdx := i;
      end;

      '>':
        begin
          CompSetOper(i, 5, False);
          if (iLastIdx > 0) then
          begin
            if not (s[i - iLastIdx] = '<') then
              iLastIdx := i;
          end
          else iLastIdx := i;
        end;

      '<': CompSetOper(i, 5);
      '-': CompSetOper(i, 7);
      '+': CompSetOper(i, 7);
      '%': CompSetOper(i, 9);
      '/': CompSetOper(i, 9);
      '*': CompSetOper(i, 9);
      '^': CompSetOper(i, 12);
      'E':
      begin
        if i > 1 then
        begin
          ch := s[i - 1];
          if ((ch >= '0') and (ch <= '9')) then
          begin
            j := i;
            while (j > 1) do
            begin
              dec(j);
              ch := s[j];
              if ((ch = '.') or ((ch >= '0') and (ch <= '9'))) then
              begin
                continue;
              end;
              if ((ch = '_') or ((ch >= 'A') and (ch <= 'Z'))) then
              begin
                iLastIdx := 0;
                break;
              end;
              iLastIdx := i;
              break;
            end;
            if ((j = 1) and ((ch >= '0') and (ch <= '9'))) then
            begin
              iLastIdx := i;
            end
          end
        end
        else iLastIdx := 0;
      end
      else iLastIdx := 0;
    end;
  end;
  Result := iCurrIdx;
end;


function TFormula.IsBinaryOperFunc(const s: string; var ALParam, ARParam:
    string; var funcAddr: Pointer; AOperIdx: integer): boolean;
var
  iLen: integer;
  iOperLen: Integer;
begin
  iLen := Length(s);
  Result := false;
  if (iLen = 0) or (AOperIdx = 0) then exit;
  if (AOperIdx > iLen - 1) then exit;

  iOperLen := 1;
  case s[AOperIdx] of
    // 运算符
    '+': funcAddr := @_add;
    '-': funcAddr := @_subtract;
    '*': funcAddr := @_multiply;
    '/': funcAddr := @_divide;
    '^': funcAddr := @_power;
    '%': funcAddr := @_intdiv;

    // 逻辑符
    '<':
    begin
      iOperLen := 2;
      case s[AOperIdx + 1] of
        '>': funcAddr := @_notequals;
        '=': funcAddr := @_ltequals;
        else
        begin
          funcAddr := @_less;
          iOperLen := 1;
        end;
      end;
    end;
    '>':
    begin
      if s[AOperIdx + 1] = '=' then
      begin
        funcAddr := @_gtequals;
        iOperLen := 2;
      end
      else
        funcAddr := @_greater;
    end;
    '=': funcAddr := @_equal;
    '&': funcAddr := @_and;
    '|': funcAddr := @_or;
    else funcAddr := nil;
  end;

  if Assigned(funcAddr) then
  begin
    ALParam := TrimCopy(s, 1, AOperIdx - 1);
    ARParam := TrimCopy(s, AOperIdx + iOperLen, iLen - AOperIdx - iOperLen + 1);
  end;

  Result := Assigned(funcAddr) and (Length(ALParam) > 0) and (Length(ARParam) > 0);
end;


function TFormula.IsUnaryOperFunc(const s: string;
    var AParam: string; var funcAddr: Pointer; CurrChar: integer): boolean;
var
  iLen: integer;
begin
  Result := false;
  if s = '' then Exit;
  funcAddr := nil;

  iLen := Length(s);
  // 前导表达式
  if CurrChar = 1 then
  begin
    AParam := TrimCopy(s, 2, iLen - 1);
    case s[CurrChar] of
      '+': funcAddr := @_unaryadd;
      '-': funcAddr := @_negate;
      '!': funcAddr := @_not;
      else funcAddr := nil;
    end;
  end
  else if CurrChar = iLen then
  begin
    // 一元后导表达式
    //
    // 百分比支持
    if s[CurrChar] = '%' then
    begin
      AParam := TrimCopy(s, 1, iLen - 1);
      funcAddr := @_percentage;
    end
  end;

  Result := Assigned(funcAddr) and (Length(AParam) > 0);
end;

procedure TFormula.CreateVar(const AName: string; const AValue: ExpValue);
begin
  FVariables.AddOrSetValue(UpperCase(AName), AValue);
end;

procedure TFormula.AssignVar(const AName: string; const val: ExpValue);
begin
  CreateVar(AName, val);
end;

function TFormula.GetVar(const AName: string): ExpValue;
begin
  Result := Unassigned;

  /// 变量读取顺序（高到低）
  ///   1、本地事件回调
  ///   2、本地接口回调
  ///   3、本地注册变量
  ///   4、全局变量
  ///

  // 全局变量
  FunDict.GetVarValue(Self, AName, Result);
  // 局部变量优先于 全局
  if FVariables.ContainsKey(AName) then
    Result := FVariables.Items[AName];
  if assigned(FExpVariable) then
    FExpVariable.GetVarValue(PChar(AName), Result);
  if Assigned(FOnVarValue) then
    FOnVarValue(Self, AName, Result);

  if VarIsEmpty(Result) then
    raise ECalcError.CreateFmt('%s变量未定义，公式：%s', [AName, FExpression]);
end;

procedure TFormula.DeleteVar(const AName: string);
var
  sKey: string;
begin
  sKey := UpperCase(AName);
  if FVariables.ContainsKey(sKey) then
  begin
    FVariables.Remove(sKey);
    FDirty := true;
  end;
end;

function TFormula.GetExpression: string;
begin
  Result := FExpression;
end;

procedure TFormula.SetExpression(const str: string);
begin
  FExpression := str;
  FDirty := true;
  WriteErr(0,'');
end;

procedure TFormula.GetFunctions(AList: TStrings);
var
  cItem: TFunItemInfo;
begin
  for cItem in FunDict.FList do
    AList.Add(cItem.Name);
end;

function TFormula.GetVariables(AList: TStrings): Integer;
var
  sName: string;
begin
  for sName in FVariables.Keys do
    AList.Add(sName);
  Result := AList.Count;
end;

procedure TFormula.FreeParseTree;
begin
  FNode.Free;
  FNode := nil;
  FDirty := true; //so that next time we call Evaluate, it will call the Parse method.
end;

function TFormula.IsFunctionUsed(const funcName: string): boolean;
var
  pAddr: Pointer;
begin
  Parse;
  pAddr := FunDict.FuncAddr(funcName);
  if not (pAddr = nil) then
    Result := FNode.IsUsed(pAddr)
  else Result := false;
end;

procedure OptimizeNode(var ANode: TCalcNode);
var
  bCanOpt: Boolean;
  cNewNode: TConstNode;
  i: Integer;
begin
  ANode.Optimize;
  bCanOpt := False;
  if (ANode is TBinaryParamNode) then
  begin
    with ANode as TBinaryParamNode do
      bCanOpt := (x is TConstNode) and (y is TConstNode);
  end
  else if (ANode is TUnaryParamNode) then
  begin
    with ANode as TUnaryParamNode do
      bCanOpt := x is TConstNode;
  end
  else if ANode is TTernaryParamNode then
  begin
    with ANode as TTernaryParamNode do
      bCanOpt := (x is TConstNode) and (y is TConstNode) and (z is TConstNode);
  end
  else if (ANode is TMultiParamNode) then
  begin
    bCanOpt := True;
    with ANode as TMultiParamNode do
      for i := 0 to Params.Count - 1 do
        if not (TCalcNode(Params[i]) is TConstNode) then
        begin
          bCanOpt := False;
          Break;
        end;
  end;

  if bCanOpt then
  begin
    cNewNode := TConstNode.Create(ANode.Value);
    ANode.free;
    ANode := cNewNode;
  end;
end;

function TFormula.IsParamFunc(const s: string; var AFunName: string;
    var AParams: TStringList): Boolean;
var
  bInString: Boolean;
  I: Integer;
  iBracketCount: Integer;
  iIndex: Integer;
  iLen: Integer;
  sParam: string;
begin
  AFunName := '';
  Result := False;
  if (s = '') then Exit;
  iLen := Length(s);
  if s[iLen] <> ')' then Exit;

  iIndex := 0;
  for I := 1 to iLen - 1 do
  begin
    if (s[i] = '(') then
    begin
      iIndex := i;
      Break;
    end
    else if not IsValidNameChar(i, s[i]) then
      Break;
  end;

  AFunName := '';
  if (iIndex > 0) then
    AFunName := Copy(s, 1, iIndex-1);
  if (AFunName = '') or not FunDict.Exists(AFunName) then
    Exit;

  //assert(not assigned(AParams));
  AParams := TStringList.Create;

  inc(iIndex);
  iBracketCount := 0;
  bInString := False;
  for i := iIndex to iLen - 1 do
  begin
    case s[i] of
      #39: bInString := not bInString;
      '(': inc(iBracketCount);
      ')': dec(iBracketCount);
      ',':
      begin
        if bInString or (iBracketCount > 0) then
          Continue;

        sParam := TrimCopy(s, iIndex, i - iIndex);
        AParams.Add(sParam);
        iIndex := i + 1;
      end;
    end;
  end;

  sParam := TrimCopy(s, iIndex, iLen - iIndex);
  AParams.Add(sParam);
  Result := True;
end;

function TFormula.IsVariableName(const s: string): Boolean;
var
  i: Integer;
begin
  // 检查是否是有效的变量名称
  if s = '' then Exit(False);

  ///
  /// 变量类型允许使用  . 运算符号
  ///   点运算符由外部处理
  ///
  Result := True;
  for i := 1 to length(s) do
    if not (IsValidNameChar(i, s[i]) or (s[i] = '.')) then
    begin
      Result := False;
      Break;
    end;
end;

procedure TFormula.Optimize;
begin
  OptimizeNode(FNode);
end;

procedure TFormula.Randomize;
begin
  System.Randomize;
end;

procedure TFormula.WriteErr(ACode: integer; const AMsg: string);
begin
  ErrCode := ACode;
  ErrMsg := AMsg
end;

{ TFunDict }

function TFunDict.Add(const AName: string; AFuncAddr: Pointer;
    AType: TFuncParamType; const ANote: string = ''): TFunItemInfo;
var
  cItem: TFunItemInfo;
  sKey: string;
begin
  /// 注册函数
  ///
  /// 存在直接替换
  sKey := UpperCase(AName);
  if FNames.ContainsKey(sKey) then
  begin
    cItem := FNames[sKey];
    FNames.Remove(sKey);
    cItem.Free;
  end;

  cItem := TFunItemInfo.Create;
  cItem.Name := AName;
  cItem.FuncAddr := AFuncAddr;
  cItem.ParamType := AType;
  cItem.Note := ANote;
  FList.Add(cItem);
  FNames.Add(sKey, cItem);
  Result := cItem;
end;

constructor TFunDict.Create;
begin
  FList:= TObjectList<TFunItemInfo>.Create;
  FNames:= TDictionary<String,TFunItemInfo>.Create;
  FConsts:= TDictionary<String,Variant>.Create; // 全局常量
  FVars:= TDictionary<string,TVarFunEvent>.Create;  // 全局变量
  FIntfs:= TDictionary<string,TVarFunEvent>.Create;
end;

destructor TFunDict.Destroy;
begin
  FIntfs.free;
  FNames.free;
  FList.free;
  FVars.free;
  FConsts.free;
  inherited;
end;

function TFunDict.Exists(const AName: string): boolean;
begin
  Result := FNames.ContainsKey(UpperCase(AName));
end;

function TFunDict.Find(const AName: string): TFunItemInfo;
begin
  Result := nil;
  FNames.TryGetValue(UpperCase(AName), Result);
end;

function TFunDict.FuncAddr(const AName: string): Pointer;
var
  cItem: TFunItemInfo;
begin
  Result := nil;
  cItem := Find(AName);
  if assigned(cItem) then
    Result := cItem.FuncAddr;
end;

function TFunDict.GetItems(const AName: string): TFunItemInfo;
begin
  Result := Find(AName);
end;

procedure TFunDict.GetVarValue(Sender: TObject; const AName: string;
  var Value: ExpValue);
var
  pEvent: TVarFunEvent;
begin
  if FConsts.ContainsKey(AName) then
    Value := FConsts[AName];
  if FVars.TryGetValue(AName, pEvent) then
    pEvent(Sender, AName, Value);
  for pEvent in FIntfs.Values do
    pEvent(Sender, AName, Value);
end;

class procedure TFunDict.RegDefaultFuns;
begin
  // 注册默认函数
  if assigned(FunDict) then
    Exit;

  /// 函数类型
  ///  函数的变量数量

  FunDict := TFunDict.Create;
  FunDict.Add('SQR',    @_square  , fptUnary);
  FunDict.Add('SIN',    @_sin     , fptUnary);
  FunDict.Add('COS',    @_cos     , fptUnary);
  FunDict.Add('ATAN',   @_arctan  , fptUnary);
  FunDict.Add('SINH',   @_sinh    , fptUnary);
  FunDict.Add('COSH',   @_cosh    , fptUnary);
  FunDict.Add('COTAN',  @_cotan   , fptUnary);
  FunDict.Add('TAN',    @_tan     , fptUnary);
  FunDict.Add('EXP',    @_exp     , fptUnary);
  FunDict.Add('LN',     @_ln      , fptUnary);
  FunDict.Add('LOG',    @_log10   , fptUnary);
  FunDict.Add('SQRT',   @_sqrt    , fptUnary);
  FunDict.Add('ABS',    @_abs     , fptUnary);
  FunDict.Add('SIGN',   @_sign    , fptUnary);
  FunDict.Add('TRUNC',  @_trunc   , fptUnary);
  FunDict.Add('CEIL',   @_ceil    , fptUnary);
  FunDict.Add('FLOOR',  @_floor   , fptUnary);
  FunDict.Add('RND',    @_rnd     , fptUnary);
  FunDict.Add('RANDOM', @_random  , fptUnary);

  FunDict.Add('ROUND',  @_round   , fptBinary);
  FunDict.Add('INTPOW', @_intpower, fptBinary);
  FunDict.Add('POW',    @_power   , fptBinary);
  FunDict.Add('LOGN',   @_logn    , fptBinary);
  FunDict.Add('MIN',    @_min     , fptBinary);
  FunDict.Add('MAX',    @_max     , fptBinary);
  FunDict.Add('MOD',    @_modulo  , fptBinary);

  FunDict.Add('IIF',    @_iif     , fptTernary);

  // string funcionts
  FunDict.Add('POS',   @_StrPos   , fptBinary);
  FunDict.Add('COPY',  @_StrCopy  , fptTernary);
  FunDict.Add('Contains', @_StrContains, fptTernary); // 有默认值
  FunDict.Add('Left',  @_StrLeft  , fptBinary);
  FunDict.Add('Right', @_StrRight , fptBinary);

  // const values
  FunDict.FConsts.Add('PI', 3.14159265358979);
end;

procedure FreeFuns;
begin
  if assigned(FunDict) then
    FunDict.free;
end;

procedure TCalcNode.Optimize;
begin
  // do nothing.
end;

{ TTernaryParamNode }

constructor TTernaryParamNode.Create(AFunc: TTernaryFunc; AXParam, AYParam, AZParam: TCalcNode);
begin
  inherited Create;
  Func := AFunc;
  x := AXParam;
  y := AYParam;
  z := AZParam;
end;

destructor TTernaryParamNode.Destroy;
begin
  x.Free;
  y.Free;
  z.free;
  inherited;
end;

function TTernaryParamNode.GetValue: ExpValue;
begin
  Result := Func(x.Value, y.Value, z.Value);
end;

function TTernaryParamNode.IsUsed(Addr: Pointer): Boolean;
begin
  Result := (Addr = @Func) or x.IsUsed(Addr) or y.IsUsed(Addr) or z.IsUsed(Addr);
end;

procedure TTernaryParamNode.Optimize;
begin
  OptimizeNode(x);
  OptimizeNode(y);
  OptimizeNode(z);
end;

{ FormulaReg }

class procedure FormulaReg.AddConst(const AName: string;
  const AValue: ExpValue);
begin
  TFunDict.RegDefaultFuns;
  FunDict.FConsts.AddOrSetValue(AName, AValue);
end;

class procedure FormulaReg.AddFun(const AName: string; AFuncAddr: Pointer;
  AType: TFuncParamType; const ANote: string);
begin
  TFunDict.RegDefaultFuns;
  FunDict.Add(AName, AFuncAddr, AType, ANote);
end;

class procedure FormulaReg.AddVar(const AName: string; AFuncAddr: TVarFunEvent);
begin
  TFunDict.RegDefaultFuns;
  FunDict.FVars.AddOrSetValue(AName, AFuncAddr);
end;

class procedure FormulaReg.AddIntf(const AIntfName: string; const AIntf:
    TVarFunEvent);
begin
  TFunDict.RegDefaultFuns;
  FunDict.FIntfs.AddOrSetValue(AIntfName, AIntf);
end;

initialization

finalization
  FreeFuns;

end.

