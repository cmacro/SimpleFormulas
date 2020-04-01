unit TestuExpParser;

interface

uses
  TestFramework, uFormulas;

type
  TestExpParser = class(TTestCase)
  private
    procedure VarEvent(Sender: TObject; const AName: string; var Value: Variant);
    procedure IntfEvent(Sender: TObject; const AName: string; var Value: Variant);
  protected
    procedure DoOnVarValue(Sender: TObject; const AName: string; var Value: ExpValue);
  published
    procedure TestOpts;
    procedure TestCalcBase;
    procedure TestStrCalc;
    procedure TestFuns;
    procedure TestVarFuns;
    procedure TestErrs;
    procedure TestCalcFuns;
    procedure TestRegs;

  end;


implementation

uses
  StrUtils, SysUtils, Math, Variants;

type
  TacParser = Class(TFormula);



function Calc(const AExp: string): Variant;
var
  cParser: TFormula;
begin
  cParser := TFormula.Create(nil);
  try
    cParser.Optimization := False;
    Result := cParser.Calc(AExp);
  finally
    cParser.free;
  end;
end;

procedure TestExpParser.DoOnVarValue(Sender: TObject; const AName: string;
    var Value: ExpValue);
begin
  if SameText(AName, 'a') then
    Value := 100
  else if SameText(AName, 'b') then
    value := 200
  else
    Value := AName
end;

procedure TestExpParser.TestCalcBase;
begin
  Check(Sin(pi/2) = 1);
  Check(Calc(' 1') = 1, '基础公式错误');
  Check(Calc('1.1') = 1.1, '基础公式错误');
  Check(Calc(' 1 + 1 ') = 2, '基础计算错误');
  Check(Calc('pi') = Unassigned, '变量大小写敏感'); //
  Check(Calc('PI') = 3.14159265358979, '变量大小写敏感'); //
  Check(Calc('sin(PI/2)') = 1, '一次函数计算错误'); //
  Check(Calc('min(1, 2)') = 1, '两次函数计算错误');
  Check(Calc('iif(1<>1, 1, 2)') = 2, '多次次函数计算错误'); //
  Check(Calc('-1 + -1') = -2, '前导符号计算错误');
  Check(Calc('100%') = 1, '百分比计算错误');
  Check(Calc('100% + 1') = 2, '百分比计算错误');
  Check(Calc('100% + 100%') = 2, '百分比计算错误');
end;


procedure TestExpParser.TestCalcFuns;
  function Contains(const sub, s: string): Integer;
  var Values: TExpValueArray;
  begin
    SetLength(Values, 2 );
    Values[0] := sub;
    Values[1] := s;
    Result := _StrContains(sub, s, null);
  end;
begin
  Check(Contains('a', 'a')=1);
  Check(Contains('', 'a')=0);
  Check(Contains('a', 'ab')=0);
  Check(Contains('ab', ',ab')=1);
  Check(Contains('ab', 'ab,')=1);
  Check(Contains('ab', 'xxsss,ab,dd')=1);
  Check(Contains('ab', ',,ab')=1);
  Check(Contains('ab', 'abc,,')=0);
  Check(Contains('ab', 'xx,ab')=1);
  Check(Contains('ab', 'ss,' + QuotedStr('xx,ab'))=0);
end;

procedure TestExpParser.TestErrs;
begin
  Check(Calc('0/0') = Unassigned, '除数零错误处理错误');
  Check(Calc('10 + 0/0') = Unassigned, '除数零错误处理错误');
  Check(Calc('errfun(10,20)') = Unassigned, '无效函数处理错误');
  Check(Calc('errvar') = Unassigned, '无效变量处理处理错误');
end;

procedure TestExpParser.TestFuns;
begin
  Check(Calc('Round(2.345678,3)') = 2.346);
  Check(calc('Round(50*38*55/1000000,3)') = 0.105);
end;

procedure TestExpParser.TestOpts;
var
  cParser: TacParser;
begin
  cParser := TacParser.Create(nil);
  Check(cParser.FindLastOper('+1+-1') = 3);
  Check(cParser.FindLastOper(QuotedStr('1+1') + '+1') = 6);
  Check(cParser.FindLastOper('SIN(90)') = 0); //
  Check(cParser.FindLastOper('MIN(1,2)') = 0);
  Check(cParser.FindLastOper('IIF(1,1,-2)') = 0);
  Check(cParser.FindLastOper('POS(''abc+def'', '''+''')') = 0);
  check(cParser.FindLastOper('1%') = 2);
  check(cParser.FindLastOper('1%%') = 3);
  check(cParser.FindLastOper('-1%%') = 4);
  Check(cParser.FindLastOper('+1%% + -1%') = 6);
  cParser.Free;
end;

function TestStr1(const x: ExpValue): ExpValue;
begin
  Result :=  'Test1_' + VarToStr(x);
end;
function TestStr2(const x, y: ExpValue): ExpValue;
begin
  Result := 'Test2_' + VarToStr(x) + '_' + VarToStr(y);
end;
function TestStr3(const x, y, z: ExpValue): ExpValue;
begin
  Result := 'Test3_' + VarToStr(x) + '_' + VarToStr(y) + '_' + VarToStr(z);
end;
function TestStrN(const x: array of ExpValue): ExpValue;
var
  i: Integer;
begin
  Result := 'TestN_' + intToStr(Length(x));
  for i := low(x) to high(x) do
    Result := Result + '_' + VarToStr(x[i]);
end;

procedure TestExpParser.VarEvent(Sender: TObject; const AName: string; var Value: Variant);
begin
  Value := 'var_' + AName;
end;

procedure TestExpParser.IntfEvent(Sender: TObject; const AName: string; var Value: Variant);
begin
  if (AName <> '') then
  begin
    if Pos('Intf', AName) > 0 then
      Value := 'Intf_' + AName;
  end;
end;


procedure TestExpParser.TestRegs;
begin
  FormulaReg.AddFun('TestStr1', @TestStr1, fptUnary);
  FormulaReg.AddFun('TestStr2', @TestStr2, fptBinary);
  FormulaReg.AddFun('TestStr3', @TestStr3, fptTernary);
  FormulaReg.AddFun('TestStrN', @TestStrN, fptMulti);
  FormulaReg.AddConst('const', 'Const');
  FormulaReg.AddVar('varevet', VarEvent);
  FormulaReg.AddIntf('Intf', IntfEvent);

  Check(Calc('TestStr1(''abc'')') = 'Test1_abc');
  Check(Calc('TestStr2(''abc'',''def'')') = 'Test2_abc_def');
  Check(Calc('TestStr3(''a'',''b'',''c'')') = 'Test3_a_b_c');
  Check(Calc('TestStrN(''a'',''b'',''c'',''d'')') = 'TestN_4_a_b_c_d');
  Check(Calc('const') = 'Const');
  Check(Calc('varevet') = 'var_varevet');
  Check(Calc('Intfvarevet') = 'Intf_Intfvarevet');
  Check(Calc('Intvarevet') = Unassigned, '没有注册变量，不应有值');
end;

procedure TestExpParser.TestStrCalc;
begin
  Check(Calc(QuotedStr('abc')) = 'abc');
  Check(Calc('''abc'' + ''def''') = 'abcdef');
  check(calc('Pos(''abc+def'', ''+'')') = 4);
end;

procedure TestExpParser.TestVarFuns;
var
  cExp: TFormula;
begin
  cExp := TFormula.Create(nil);
  try
    cExp.Optimization := False;
    cExp.OnVarValue := DoOnVarValue;
    Check(cExp.Calc('a+b') = 300);
    Check(cExp.Calc('a+c') = '100c');
  finally
    cExp.free;
  end;
end;

initialization
  RegisterTest(TestExpParser.Suite);

end.
