"""In-memory stand-in for the logging spreadsheet (worksheets by index and title), for tests and UI harnesses."""
import gspread


class FakeWorksheet:
    def __init__(self, title):
        self.title, self.cells, self.row_count = title, [], 200

    def get_all_values(self):
        return [list(row) for row in self.cells if any(row)] if self.title != "log" else [list(r) for r in self.cells]

    def update(self, a1, rows, value_input_option=None):
        assert a1 == "A1" and value_input_option == "RAW"
        for i, row in enumerate(rows):
            while len(self.cells) <= i:
                self.cells.append([])
            self.cells[i] = [str(x) for x in row]

    def batch_clear(self, ranges):
        start = int(ranges[0].split(":")[0][1:]) - 1
        for i in range(start, len(self.cells)):
            self.cells[i] = [""] * 4

    def add_rows(self, n):
        self.row_count += n

    def append_row(self, row, value_input_option=None):
        self.cells.append([str(x) for x in row])

    def delete_rows(self, number):
        del self.cells[number - 1]

    def get_all_records(self):
        header, *rows = self.cells or [["Timestamp", "User", "Event Type", "Details"]]
        return [dict(zip(header, row)) for row in rows]


class FakeSpreadsheet:
    def __init__(self):
        log = FakeWorksheet("log")
        log.cells.append(["Timestamp", "User", "Event Type", "Details"])
        self.tabs = {0: log}

    def worksheet(self, title):
        if title not in self.tabs:
            raise gspread.WorksheetNotFound(title)
        return self.tabs[title]

    def add_worksheet(self, title, rows, cols):
        self.tabs[title] = FakeWorksheet(title)
        return self.tabs[title]

    def get_worksheet(self, index):
        return self.tabs[index]
