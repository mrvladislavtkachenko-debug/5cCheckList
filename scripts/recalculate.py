"""Evaluate formulas for the initial preview and embed real formula/chart caches.

Excel/Calc still recalculate natively after edits. No engine is needed by the user.
This script is also the formula-level test harness (overridable input nodes).
"""
from __future__ import annotations
import argparse
import contextlib
import io
import math
from pathlib import Path
import re
import zipfile
from lxml import etree as ET
import formulas
import numpy as np
from openpyxl.utils.cell import range_boundaries, get_column_letter

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CNS='http://schemas.openxmlformats.org/drawingml/2006/chart'
RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PNS='http://schemas.openxmlformats.org/package/2006/relationships'


def split_key(key):
    if not isinstance(key,str) or '!' not in key or ']' not in key:
        return None
    left,cell=key.rsplit('!',1)
    if not re.fullmatch(r'\$?[A-Z]+\$?\d+',cell):
        return None
    sheet=left.split(']',1)[1].rstrip("'")
    return sheet.upper(),cell.replace('$','').upper()


def normalise(value):
    if isinstance(value,np.generic):value=value.item()
    if isinstance(value,(int,float)) and not isinstance(value,bool):
        if not math.isfinite(float(value)):raise ValueError(f'Non-finite formula result: {value}')
        return value
    if value is None:return ''
    if isinstance(value,str):return value
    if isinstance(value,bool):return value
    raise TypeError(f'Unsupported result {value!r}, {type(value)}')


class Evaluator:
    def __init__(self,path):
        self.path=Path(path)
        self.model=formulas.ExcelModel().loads(str(self.path)).finish()
        self.keys={split_key(k):k for k in self.model.cells if split_key(k)}
        self.solution=None
        self.values={}

    def calculate(self,overrides=None):
        inputs={}
        for (sheet,cell),value in (overrides or {}).items():
            key=self.keys[(sheet.upper(),cell.upper())]
            inputs[key]=[[value]]
        self.solution=self.model.calculate(inputs=inputs)
        self.values={}
        for key,result in self.solution.items():
            target=split_key(key)
            if target and hasattr(result,'value') and result.value.size==1:
                self.values[target]=normalise(result.value.reshape(-1)[0])
        return self

    def get(self,sheet,cell):
        return self.values[(sheet.upper(),cell.upper())]

    def errors(self):
        return {k:v for k,v in self.values.items() if isinstance(v,str) and v.startswith(('#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A','#SPILL!','#CALC!','#NULL!','#GETTING_DATA'))}

    def embed(self,path=None):
        if self.errors():raise ValueError(f'Formula errors: {self.errors()}')
        path=Path(path or self.path)
        with zipfile.ZipFile(path) as z: parts={n:z.read(n) for n in z.namelist()}
        workbook=ET.fromstring(parts['xl/workbook.xml'])
        relationships=ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
        targets={r.get('Id'):r.get('Target') for r in relationships}
        count=0
        for sheet in workbook.find(f'{{{NS}}}sheets'):
            name=sheet.get('name').upper()
            target=targets[sheet.get(f'{{{RNS}}}id')]
            key=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            xml=ET.fromstring(parts[key])
            for c in xml.iter(f'{{{NS}}}c'):
                if c.find(f'{{{NS}}}f') is None:continue
                cell=c.get('r').upper()
                if (name,cell) not in self.values:raise ValueError(f'No calculated cache: {name}!{cell}')
                value=self.values[name,cell]
                old=c.find(f'{{{NS}}}v')
                if old is not None:c.remove(old)
                v=ET.SubElement(c,f'{{{NS}}}v')
                if isinstance(value,bool):c.set('t','b');v.text='1' if value else '0'
                elif isinstance(value,(int,float)):
                    c.attrib.pop('t',None);v.text=format(value,'.15g')
                else:c.set('t','str');v.text=value
                count+=1
            parts[key]=ET.tostring(xml,xml_declaration=True,encoding='UTF-8',standalone=True)
        calc=workbook.find(f'{{{NS}}}calcPr')
        if calc is None:calc=ET.SubElement(workbook,f'{{{NS}}}calcPr')
        calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
        parts['xl/workbook.xml']=ET.tostring(workbook,xml_declaration=True,encoding='UTF-8',standalone=True)
        charts=0
        for key,data in list(parts.items()):
            if not re.fullmatch(r'xl/charts/chart\d+\.xml',key):continue
            xml=ET.fromstring(data)
            for kind in ['numRef','strRef']:
                for node in xml.iter(f'{{{CNS}}}{kind}'):
                    f=node.find(f'{{{CNS}}}f')
                    if f is None or not f.text or '!' not in f.text:continue
                    sheet,address=f.text.rsplit('!',1)
                    sheet=sheet.strip("'").upper()
                    lo_c,lo_r,hi_c,hi_r=range_boundaries(address)
                    values=[self.values.get((sheet,f'{get_column_letter(c)}{r}'),'')
                            for r in range(lo_r,hi_r+1) for c in range(lo_c,hi_c+1)]
                    cache_tag='numCache' if kind=='numRef' else 'strCache'
                    old=node.find(f'{{{CNS}}}{cache_tag}')
                    if old is not None:node.remove(old)
                    cache=ET.SubElement(node,f'{{{CNS}}}{cache_tag}')
                    if kind=='numRef':ET.SubElement(cache,f'{{{CNS}}}formatCode').text='General'
                    ET.SubElement(cache,f'{{{CNS}}}ptCount',val=str(len(values)))
                    for index,value in enumerate(values):
                        if kind=='numRef' and not isinstance(value,(int,float,bool)):continue
                        pt=ET.SubElement(cache,f'{{{CNS}}}pt',idx=str(index))
                        ET.SubElement(pt,f'{{{CNS}}}v').text=(format(value,'.15g') if isinstance(value,(int,float)) else str(value))
            parts[key]=ET.tostring(xml,xml_declaration=True,encoding='UTF-8',standalone=True);charts+=1
        tmp=path.with_suffix('.tmp.xlsx')
        with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
            for name,data in parts.items():z.writestr(name,data)
        tmp.replace(path)
        return count,charts


def main():
    parser=argparse.ArgumentParser();parser.add_argument('workbook',type=Path);args=parser.parse_args()
    e=Evaluator(args.workbook).calculate()
    errors=e.errors()
    if errors:
        for key,value in list(errors.items())[:60]:print(key,value)
        raise SystemExit(f'{len(errors)} formula errors')
    count,charts=e.embed()
    print(f'Cached {count} formulas and {charts} charts.')
    for cell in ['C10','C11','C12','C13','C14','C15','C16','C17','C18']:print('Начало!'+cell, e.get('Начало',cell))


if __name__=='__main__':main()
