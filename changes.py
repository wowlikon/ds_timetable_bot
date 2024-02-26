from openpyxl import Workbook, load_workbook, utils
from table2ascii import table2ascii as t2a, PresetStyle
from xls2xlsx import XLS2XLSX
from pathlib import Path
import re, ujson

class Change:
    def __init__(self, group: str, lesson: int, src: str, dst: str, teacher: str, room: str):
        self.group = group
        self.lesson = lesson
        self.room = room
        self.teacher = teacher
        self.replacing = (src, dst)
    
    def __str__(self) -> str:
        return f"{self.group}|{self.lesson}|[{self.replacing[0]}]=>[{self.replacing[1]}]|{self.room}|{self.teacher}"
    
    def as_list(self) -> list:
        return [self.group, self.lesson, self.replacing, self.teacher, self.room]

    def as_dict(self) -> dict:
        return {
            "room": self.room,
            "group": self.group,
            "lesson": self.lesson,
            "teacher": self.teacher,
            "replace": dict(zip(('from', 'to'), self.replacing))
            }

class Changes:
    def __init__(self):
        self.changes = list()
    
    def __str__(self) -> str:
        return "\n".join(map(str, self.changes))
    
    def append(self, elem: Change):
        self.changes.append(elem)
    
    def get_by_group(self, group: str) -> list:
        res = list()
        for i in self.changes:
            if i.group == group:
                res.append(i)
        return res

    def as_list(self) -> list:
        result = []
        for i in self.changes:
            result.append(i.as_list())
        return result
    
    def as_dict(self) -> dict:
        result = {}
        for i in self.changes:
            t = i.as_dict()
            tg = t['group']
            if tg not in result.keys():
                result[tg] = {"changes":[]}
            t.pop('group')
            result[tg]['changes'].append(t)
        return result

def parseFile(file: str, cache: str) -> str:
    filePL = Path(file) #ch_data
    x2x = XLS2XLSX(str(filePL.parent)+'/changes.xls')
    x2x.to_xlsx(str(filePL.parent)+'/changes.xlsx')

    wb = load_workbook(str(filePL.parent)+'/changes.xlsx', True)

    ws = wb.active
    row_count = ws.max_row
    column_count = ws.max_column

    print(f"{column_count}x{row_count}")

    changes = Changes()
    for row in range(row_count):
        line = list()
        for column in range(column_count):
            v = ws[utils.cell.get_column_letter(column+1)+str(row)].value
            if v: line.append(v)
        if len(line) == column_count-2:
            change = Change(*line)
            changes.append(change)

    return changes

def parseCache(filename: str, group: str, day: str) -> Changes:
    with open(filename+'ср_data.json', 'r', encoding='utf-8') as f:
        tt = ujson.load(f)['groups']
    print(tt)
    g = tt.get(group, {})
    print(group, ':', g)
    d = g.get(day, {})
    print(day, ':', d)
    return Changes.from_dict(d, day)

def makeTable(data: list) -> str:
    return t2a(
        header=["Group", "Lesson", "Replace", "Teacher", "Room"],
        body=data, first_col_heading=True
    )

if __name__ == '__main__':
    filename = input("Filename: ")
    print(filename)
    cf = input("Cache: ")
    print(cf)
    qGroup = input("Group: ")
    print(qGroup)
    qDay = input("Day of week: ")
    print(qDay)

    cf = parseFile(filename, cf)
    data = parseCache(cf, qGroup, qDay)
    
    print(data.as_list())
    print(data.as_dict())
    print(makeTable(data.as_list()))