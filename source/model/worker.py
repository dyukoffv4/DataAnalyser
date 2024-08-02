import re


class Worker:
    def __init__(self, sheet, from_column, to_column, res_columns: dict, size):
        self._in_column = from_column
        self._out_columns = [i for i in res_columns.values()]
        self._rules = []
        self._results = [[] for _ in range(len(res_columns))]

        keys = list(res_columns.keys())

        for i in range(2, size + 2):  # todo: Risk of 1'st string
            if sheet[f'{to_column}{i}'].value is not None:
                self._rules.append(sheet[f'{to_column}{i}'].value)
                for k in range(len(self._results)):
                    self._results[k].append(sheet[f'{keys[k]}{i}'].value)

    def process_sheet(self, sheet):
        for i in range(2, len(sheet[self._in_column])):  # todo: Risk of 1'st string
            if (value := sheet[f'{self._in_column}{i}'].value) is None:
                break

            for k in range(len(self._rules)):
                # todo: Made an external grader
                if re.search(self._rules[k], value) is not None:
                    for n in range(len(self._results)):
                        sheet[f'{self._out_columns[n]}{i}'].value = self._results[n][k]
