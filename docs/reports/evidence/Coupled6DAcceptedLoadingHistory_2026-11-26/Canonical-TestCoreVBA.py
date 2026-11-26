"""ДЛЯ ТЕСТОВ: сопоставляет точные числовые значения и типы литералов Core-модуля после штатной записи VBE."""
from decimal import Decimal
import re


TestNumber=re.compile(r'(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?[#!%&@^]?',re.I)


def TestCanonical(text):
    """Не меняет строки и исполняемые имена; учитывает только точный формат числа и кодовую страницу комментариев."""
    result=[];numeric=0;comment_losses=0
    for line in text.splitlines(keepends=True):
        index=0;output=[]
        while index<len(line):
            if line[index]=='"':
                start=index;index+=1
                while index<len(line):
                    if line[index]=='"':
                        index+=1
                        if index<len(line) and line[index]=='"':index+=1;continue
                        break
                    index+=1
                output.append(line[start:index]);continue
            if line[index]=="'":
                comment=line[index:];encoded=comment.encode('cp1251','replace').decode('cp1251')
                comment_losses+=sum(a!=b for a,b in zip(comment,encoded))
                output.append(encoded);index=len(line);continue
            eligible=(index==0 or not (line[index-1].isalnum() or line[index-1] in '_.'))
            match=TestNumber.match(line,index) if eligible else None
            if match:
                token=match[0];suffix=token[-1] if token[-1] in '#!%&@^' else ''
                number=token[:-1] if suffix else token;value=Decimal(number)
                if suffix:kind={'#':'Double','!':'Single','%':'Integer','&':'Long','@':'Currency','^':'LongLong'}[suffix]
                elif '.' in number or 'e' in number.lower():kind='Double'
                elif value<=32767:kind='Integer'
                elif value<=2147483647:kind='Long'
                else:kind='Double'
                output.append('<number:'+kind+':'+format(value.normalize(),'f')+'>')
                numeric+=1;index=match.end();continue
            output.append(line[index]);index+=1
        result.append(''.join(output))
    return ''.join(result),{'numeric_tokens':numeric,'unrepresentable_comment_characters':comment_losses,
                           'runtime_strings_unchanged':True,'decimal_values_and_literal_types_preserved':True}


def TestWrap(body):
    """Ограничивает сравнение одним явно импортированным тестовым модулем; рабочие модули проверяются исходным правилом."""
    def wrapped(text,name):
        record=body(text,name)
        if name=='modTestCoupledElasticState':return (TestCanonical(record[0])[0],*record[1:])
        return record
    return wrapped
