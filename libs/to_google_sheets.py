#!/usr/bin/python3
# -*- coding: utf-8 -*-


class SpreadsheetError(Exception):
    pass


class SpreadsheetNotSetError(SpreadsheetError):
    pass


class SheetNotSetError(SpreadsheetError):
    pass


class Spreadsheet:

    # Класс-оберта методов
    def __init__(self, spreadsheetId, sheetId, service, sheetTitle):
        self.requests = []
        self.valueRanges = []
        self.spreadsheetId = spreadsheetId
        self.sheetId = sheetId
        self.service = service
        self.sheetTitle = sheetTitle


    def prepare_setDimensionPixelSize(self, dimension, startIndex, endIndex, pixelSize):
        self.requests.append({"updateDimensionProperties": {
            "range": {"sheetId": self.sheetId,
                      "dimension": dimension,
                      "startIndex": startIndex,
                      "endIndex": endIndex},
            "properties": {"pixelSize": pixelSize},
            "fields": "pixelSize"}})

    def prepare_setColumnsWidth(self, startCol, endCol, width):
        self.prepare_setDimensionPixelSize("COLUMNS", startCol, endCol + 1, width)

    def prepare_setColumnWidth(self, col, width):
        self.prepare_setColumnsWidth(col, col, width)

    def prepare_setValues(self, cellsRange, values, majorDimension="ROWS"):
        self.valueRanges.append(
            {"range": self.sheetTitle + "!" + cellsRange, "majorDimension": majorDimension, "values": values})

    # spreadsheets.batchUpdate and spreadsheets.values.batchUpdate
    def runPrepared(self, valueInputOption="USER_ENTERED"):
        upd1Res = {'replies': []}
        upd2Res = {'responses': []}
        try:
            if len(self.requests) > 0:
                upd1Res = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId,
                                                                  body={"requests": self.requests}).execute()
            if len(self.valueRanges) > 0:
                upd2Res = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId,
                                                                           body={"valueInputOption": valueInputOption,
                                                                                 "data": self.valueRanges}).execute()
        finally:
            self.requests = []
            self.valueRanges = []
        return (upd1Res['replies'], upd2Res['responses'])

        # Converts string range to GridRange of current sheet; examples:
        #   "A3:B4" -> {sheetId: id of current sheet, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
        #   "A5:B"  -> {sheetId: id of current sheet, startRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
    def toGridRange(self, cellsRange):
        if self.sheetId is None:
            raise SheetNotSetError()
        if isinstance(cellsRange, str):
            startCell, endCell = cellsRange.split(":")[0:2]
            cellsRange = {}
            rangeAZ = range(ord('A'), ord('Z') + 1)
            if ord(startCell[0]) in rangeAZ:
                cellsRange["startColumnIndex"] = ord(startCell[0]) - ord('A')
                startCell = startCell[1:]
            if ord(endCell[0]) in rangeAZ:
                cellsRange["endColumnIndex"] = ord(endCell[0]) - ord('A') + 1
                endCell = endCell[1:]
            if len(startCell) > 0:
                cellsRange["startRowIndex"] = int(startCell) - 1
            if len(endCell) > 0:
                cellsRange["endRowIndex"] = int(endCell)
        cellsRange["sheetId"] = self.sheetId
        return cellsRange

    def prepare_mergeCells(self, cellsRange, mergeType = "MERGE_ALL"):
        self.requests.append({"mergeCells": {"range": self.toGridRange(cellsRange), "mergeType": mergeType}})

    # formatJSON should be dict with userEnteredFormat to be applied to each cell
    def prepare_setCellsFormat(self, cellsRange, formatJSON, fields="userEnteredFormat"):
        self.requests.append({"repeatCell": {"range": self.toGridRange(cellsRange),
                                             "cell": {"userEnteredFormat": formatJSON}, "fields": fields}})

    # formatsJSON should be list of lists of dicts with userEnteredFormat for each cell in each row
    def prepare_setCellsFormats(self, cellsRange, formatsJSON, fields="userEnteredFormat"):
        self.requests.append({"updateCells": {"range": self.toGridRange(cellsRange),
                                              "rows": [{"values": [{"userEnteredFormat": cellFormat} for cellFormat
                                                                   in rowFormats]} for rowFormats in formatsJSON],
                                              "fields": fields}})

    def prepare_setBorderFormats(self, cellsRange, fields = "userEnteredFormat"):
        self.requests.append({"updateBorders": {"range": self.toGridRange(cellsRange),
                                                "bottom": {'style': 'SOLID',
                                                           'width': 1,
                                                           'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                                "top": {'style': 'SOLID',
                                                           'width': 1,
                                                           'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                                "left": {'style': 'SOLID',
                                                           'width': 1,
                                                           'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                                "right": {'style': 'SOLID',
                                                           'width': 1,
                                                           'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                                }
                              })
