简易计算公式
==============

主要功能

- 运算符 `+ - * / ^ %` 
- 逻辑符 `= > >= < <= <> & |`
- 字符串拼接    `'1'+'1'` = `11`
- 可扩充常量、变量和函数


## 测试

> **测试环境：**    
> Delphi 2010 （D7不兼容）  
> win10 VM 虚拟机  


```
function Calc(const AExp: string): Variant;
var cParser: TFormula;
begin
  cParser := TFormula.Create(nil);
  try
    Result := cParser.Calc(AExp);
  finally
    cParser.free;
  end;
end;

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
```
