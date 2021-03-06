#!/bin/env python

import os
import sys
from argparse import ArgumentParser
from linecache import getlines
from configparser import ConfigParser
from urllib.parse import urlparse, unquote

from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import QThread, pyqtSignal # punt on QProcess due to IPC complexity
from PyQt5.QtWidgets import QPushButton, QWidget, QApplication, QLabel, QGridLayout, QProgressBar, QMessageBox

import xlsxwriter as xlwr
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell

from emolog.emotool.ppxl_util import (
HALF_CYCLE_CELL_TO_FORMULA,
HALF_CYCLE_PREDEFINED_TITLES,
HALF_CYCLE_PREDEFINED_CELL_NAMES,
HALF_CYCLE_TITLE_TO_CELL_NAME,
HALF_CYCLE_CELL_TO_TITLE_NAME,
HALF_CYCLE_FORMULA_TITLES,
)

VERSION = 0.1

PARAMETERS_SHEET_NAME = 'Parameters'
HALF_CYCLES_SHEET_NAME = 'Half-Cycles'

DIRECTION_TEXT = 'Direction'
DOWN_AVERAGES_TEXT = 'DOWN Averages'
UP_AVERAGES_TEXT = 'UP Averages'
ALL_AVERAGES_TEXT = 'ALL Averages'
HALF_CYCLE_SUMMARY_TEXT = 'Half-Cycle Summary'

CONFIG_FILENAME = 'summary.ini'
OUTPUT_FILENAME = 'summary.xlsx'


def read_xlsx(d):
    entries = [entry for entry in os.scandir(d) if entry.is_file() and entry.path.endswith('xlsx')]
    filenames = list(sorted([entry.path for entry in entries]))
    return filenames


def get_readers(orig_filenames, progress=None):
    filename_to_reader = {}
    filenames = []
    for i, filename in enumerate(orig_filenames):
        reader = xlrd.open_workbook(filename=filename)
        sheet_names = reader.sheet_names()
        if progress:
            progress(i)
        if HALF_CYCLES_SHEET_NAME not in sheet_names:
            continue
        filename_to_reader[filename] = reader
        filenames.append(filename)
    filenames.sort()
    return [filename_to_reader[filename] for filename in filenames], filenames


def verify_cell_at(sheet, row, col, contents):
    value = sheet.cell(rowx=row, colx=col).value
    if value != contents:
        print(f"expected sheet {hc.name}[{row},{col}] to be {contents} but found {value}")
        raise SystemExit


def find_row(sheet, col, text, max_row=200):
    for i in range(max_row):
        if sheet.cell(rowx=i, colx=col).value == text:
            return i
    print(f"{sheet.name}: could not find a row containing {text} in column {col}")
    raise SystemExit


def colvals(sheet, col):
    return [x.value for x in sheet.col(col)]


def rowvals(sheet, col):
    return [x.value for x in sheet.row(col)]


def get_parameters(reader):
    parameters = reader.sheet_by_name(PARAMETERS_SHEET_NAME)
    keys = colvals(parameters, 0)
    values = colvals(parameters, 1)
    return dict(zip(keys, values))


def get_summary_data(reader):
    hc = reader.sheet_by_name(HALF_CYCLES_SHEET_NAME)
    half_cycle_summary_row_number = find_row(hc, col=0, text=HALF_CYCLE_SUMMARY_TEXT)
    titles_row_number = half_cycle_summary_row_number + 1
    down_averages_row_number = half_cycle_summary_row_number + 2
    up_averages_row_number = half_cycle_summary_row_number + 3
    all_averages_row_number = half_cycle_summary_row_number + 4
    for row, text in [
        (titles_row_number, DIRECTION_TEXT),
        (down_averages_row_number, DOWN_AVERAGES_TEXT),
        (up_averages_row_number, UP_AVERAGES_TEXT),
        (all_averages_row_number, ALL_AVERAGES_TEXT),
    ]:
        verify_cell_at(hc, row=row, col=1, contents=text)
    rowxs = [titles_row_number, down_averages_row_number, up_averages_row_number, all_averages_row_number]
    summary_titles, down, up, all = [rowvals(hc, rowx)[2:] for rowx in rowxs]
    return dict(titles=summary_titles, down=down, up=up, all=all)


def small_int_dict(arrays):
    """
    Allocate an integer starting with 0 for each new key found in the <arrays>
    going over them one by one. An example:

    small_int_dict([['a', 'b'], ['a', 'c']]) => {'a': 0, 'b': 1, 'c': 2}
    :param arrays: [[Object]]
    :return: dict(Object -> int)
    """
    ret = {} # val -> int
    for i, arr in enumerate(arrays):
        for val in arr:
            if val not in ret:
                ret[val] = len(ret)
    return ret


class Render():
    """
    Utilities for creating rows or columns from shorter descriptions
    """
    @staticmethod
    def points(points):
        """
        Take points = [(index, text)] and place them in a single row, i.e.:
        [(2, 'a'), (5, 'b')] => [None, None, 'a', None, None, 'b']
        :param points:
        :return:
        """
        max_i = max(i for i, v in points)
        ret = [None] * (max_i + 1)
        for i, v in points:
            ret[i] = v
        return ret

    @staticmethod
    def points_add(deltas):
        data = []
        ind = 0
        for d, v in deltas:
            ind += d
            data.append((ind, v))
        return Render.points(data)

    @staticmethod
    def subset(subset, d, default=None):
        return [d.get(param, default) for param in subset]


class IntAlloc():
    def __init__(self, init=0):
        self.val = init

    def inc(self, delta):
        self.val += delta
        return self.val


def summarize_dir(d, config):
    filenames = read_xlsx(d)
    output_filename = summarize_files(filenames=filenames, output_path=d, config=config)
    return output_filename


def enum_cum_len(vs, initial=0):
    if len(vs) == 0:
        return
    acc = initial
    for v in vs:
        yield acc, v
        acc += len(v)


class Output:
    def __init__(self, filename):
        self.filename = filename
        self.cell_formats = []
        self.data = []

    def add(self, row, col, value, cell_format):
        self.data.append((row, col, value, cell_format))

    def add_row(self, row, col, data, cell_format):
        for i, value in enumerate(data):
            self.add(row=row, col=col + i, value=value, cell_format=cell_format)

    def add_col(self, row, col, data, cell_format):
        for i, value in enumerate(data):
            self.add(row=row + i, col=col, value=value, cell_format=cell_format)

    def add_format(self, **kw):
        """
        Identical to add_format(properties=kw)
        :param kw:
        :return:
        """
        self.cell_formats.append(kw)
        return len(self.cell_formats) - 1

    def write(self):
        writer = xlwr.Workbook(self.filename)
        formats = {i: writer.add_format(properties=properties) for i, properties in enumerate(self.cell_formats)}
        summary_out = writer.add_worksheet('Summary')
        for row, col, value, format in self.data:
            summary_out.write(row, col, value, formats[format]) # cannot use named arguments due to (row, col, *args) def
        writer.close()


def dunion(d1, d2):
    d = dict(d1)
    d.update(d2)
    return d


def do_nothing(*args):
    pass


def required_cell_names_from_titles(titles):
    """ return the sum of cells required for given titles """

    def helper(cells):
        formulae = [HALF_CYCLE_CELL_TO_FORMULA[c] for c in cells if c in HALF_CYCLE_CELL_TO_FORMULA]

        def cell_names_for_formulae(ff):
            return [x for x in ff.__code__.co_varnames if x.endswith('_cell')]
        return set(sum([cell_names_for_formulae(f) for f in formulae], []))

    initial_cells = [HALF_CYCLE_TITLE_TO_CELL_NAME[k] for k in titles if k in HALF_CYCLE_FORMULA_TITLES]
    s = helper(initial_cells)
    N = 10
    additions = [s]
    for i in range(N):
        print(s)
        new_s = helper(s)
        if len(new_s) == 0:
            break
        additions.append(new_s)
        s = new_s
    assert len(new_s) == 0, f'Increase formulae resolution iterations from {N}; unresolved variables: {new_s!r}'
    s = additions[0]
    for s2 in additions[1:]:
        s |= s2
    return s


def summarize_files(filenames, output_path, config, progress=None):
    """
    read all .xls files in the directory that have a 'Half-Cycles' sheet, and
    create a new summary.xls file from them
    :param dir:
    :return: written xlsx filename full path if successful, else None
    """

    if progress is None:
        progress = do_nothing

    progress_count = [0]

    def update_progress():
        progress(progress_count[0])
        progress_count[0] += 1

    user_defined_fields = [x for x in config.user_defined_fields if x not in HALF_CYCLE_PREDEFINED_TITLES]
    half_cycle_directions = config.half_cycle_directions
    half_cycle_fields = config.half_cycle_fields
    parameter_names = HALF_CYCLE_PREDEFINED_TITLES + [x for x in config.parameters]

    # check we have all inputs required for the formula presented
    required_cell_names = required_cell_names_from_titles(half_cycle_fields)
    available_cell_names = {HALF_CYCLE_TITLE_TO_CELL_NAME[x] for x in set(HALF_CYCLE_TITLE_TO_CELL_NAME.keys()) & (set(parameter_names) | set(half_cycle_fields))}
    if required_cell_names - available_cell_names:
        missing_titles = [HALF_CYCLE_CELL_TO_TITLE_NAME[x] for x in sorted(list(required_cell_names - available_cell_names))]
        half_cycle_fields = half_cycle_fields + missing_titles
        print(f"added missing {missing_titles!r}")

    # compute titles - we have a left col for the 'Up/Down/All' caption
    summary_titles = half_cycle_fields
    N_par = len(parameter_names)
    N_sum = len(summary_titles)
    N_user = len(user_defined_fields)
    top_titles = [None] * N_par + sum([[d] * N_sum for d in half_cycle_directions], [])
    titles = parameter_names + (len(half_cycle_directions) * summary_titles)

    readers, filenames = get_readers(filenames, lambda *args: update_progress()) # the initial filenames contains xlsx that are not produced by the post processor

    N = len(readers)
    if N == 0:
        print("no files found")
        return

    output_filename = allocate_unused_file_in_directory(os.path.join(output_path, OUTPUT_FILENAME))

    print("reading parameters")
    all_parameters = [get_parameters(reader) for reader in readers]

    print("reading summaries")
    all_summaries = [get_summary_data(reader) for reader in readers]

    # aggregate all data to output: tuples of row, col, format, value
    output = Output(output_filename)

    # formats for titles and cells
    title_format = output.add_format(text_wrap=True, align='left', bold=True)
    col_format = output.add_format(text_wrap=True, align='left', num_format='0.000')
    user_format = output.add_format(align='left', num_format='0.000')

    row = IntAlloc()
    # create titles
    output.add_row(col=row.val + 1, row=0, data=[''] * N_user + top_titles, cell_format=title_format)
    output.add_row(col=row.val + 1, row=1, data=user_defined_fields + titles, cell_format=title_format)
    row.inc(2)

    # write column for each file
    def cells_from_d(keys, row, col):
        return {HALF_CYCLE_TITLE_TO_CELL_NAME[k]:
                                  xl_rowcol_to_cell(row=row, col=col + i) for i, k in enumerate(keys)
                if k in HALF_CYCLE_TITLE_TO_CELL_NAME}

    param_left_col = 1 + N_user + N_par # 1 - for file name; TODO: make this declarative (place cells on board with name, than use name)

    for filename, parameters, summary in zip(filenames, all_parameters, all_summaries):
        params_values = Render.subset(subset=parameter_names, d=parameters)
        sum_per_dir = [
            {k: HALF_CYCLE_CELL_TO_FORMULA.get(HALF_CYCLE_TITLE_TO_CELL_NAME.get(k, None), v) for k, v in zip(summary['titles'], summary[key.lower()])
             if k in summary_titles}
            for key in half_cycle_directions]
        cell_locations = {k: xl_rowcol_to_cell(row=row.val, col=1 + N_user + i) for i, k in enumerate(HALF_CYCLE_PREDEFINED_CELL_NAMES)}
        summary_rows_with_unfilled_formula = [
            (Render.subset(summary_titles, sum_row), dunion(cell_locations, cells_from_d(keys=summary_titles, row=row.val, col=col)))
                        for col, sum_row in enum_cum_len(sum_per_dir, initial=param_left_col)]
        for values, cells in summary_rows_with_unfilled_formula:
            assert len(cells.values()) == len(set(cells.values())), "error: allocated same cell to two variables"
        summary_rows = [[x(**cells) if callable(x) else x for x in values]
                        for values, cells in summary_rows_with_unfilled_formula]
        summary_values = sum(summary_rows, [])
        filename = os.path.split(filename)[-1]
        data = [filename] + [''] * N_user + params_values + summary_values
        output.add_row(col=0, row=row.val, data=data, cell_format=col_format)
        row.inc(1)
    #summary_out.set_row(firstrow=0, lastrow=2, width=8)
    #summary_out.set_row(firstrow=2, lastrow=N + 3, width=8)
    output.write()
    return output_filename


def allocate_unused_file_in_directory(initial):
    """look for a file at the dirname(initial) with basename(initial)
    file name. If one already exists, try adding _1, then _2 etc. right
    before the extention
    """
    i = 1
    d = os.path.dirname(initial)
    filename_with_ext = os.path.basename(initial)
    noext, ext = filename_with_ext.rsplit('.', 1)
    fname = initial
    while os.path.exists(fname) and i < 1000:
        fname = os.path.join(d, f'{noext}_{i}.{ext}')
        i += 1
    return fname


def button(parent, title, callback):
    class Button(QPushButton):
        def mousePressEvent(self, e):
            QPushButton.mousePressEvent(self, e)
            callback()
    return Button(title, parent)


def paths_from_file_urls(urls):
    ret = []
    for url in urls:
        if len(url) == 0:
            continue
        parsed = urlparse(url)
        if parsed.scheme != 'file':
            print(f'ignoring scheme = {parsed.scheme!r} ({url!r})')
            continue
        path = unquote(parsed.path if len(parsed.path) > 0 else parsed.netloc)
        if 'win' in sys.platform and path[:1] == '/':
            path = path[1:]
        if not os.path.exists(path):
            print(f"no such file: {path!r} ({url!r})")
            continue
        ret.append(path)
    return ret


def start(filename):
    if hasattr(os, 'startfile'):
        os.startfile(filename)
    else:
        os.system(f'xdg-open "{filename}"')


class Config:
    def __init__(self, d):
        ini_filename = os.path.join(d, CONFIG_FILENAME)
        if os.path.exists(ini_filename):
            print(f"reading config from {ini_filename}")
            self.config = ConfigParser()
            self.config.read(ini_filename)
        else:
            self.config = None
        # merge half_cycle and half_cycles sections, converting them to
        # half_cycle section
        half_cycles = self._get_sections(['half_cycle', 'half_cycles'])
        self.user_defined_fields = self._get_strings('user_defined', 'fields', ["Pump Head [m]", "Damper used?", "PSU or Solar Panels", "MPPT used?", "General Notes"])
        self.half_cycle_fields = half_cycles.get('fields', ['Average Velocity [m/s]', 'Flow Rate [LPM]'])
        self.half_cycle_directions = half_cycles.get('directions', ['down', 'up', 'all'])
        self.parameters = self._get_strings('global', 'parameters', [])

    def _get_sections(self, sections):
        ds = [self._get_section(s) for s in sections]
        dret = {}
        for d in ds:
            dret.update(d)
        return self._parse_strings(dret)

    def _parse_strings(self, d):
        def split(v):
            return [x.strip() for x in v.split(',')]
        return {k: split(v) for k, v in d.items()}

    def _get_section(self, section):
        if not self.config or not self.config.has_section(section):
            return {}
        return dict(self.config._unify_values(section, None))

    def _get(self, section, field, default):
        if self.config is not None and self.config.has_option(section, field):
            return self.config.get(section, field, raw=True) # avoid % interpolation, we want to have % values
        return default

    def _get_strings(self, section, field, default):
        if self.config is not None and self.config.has_option(section, field):
            return [x.strip() for x in [y for y in self.config.get(section, field, raw=True).split(',') if len(y) > 0]]
        return default


class SummarizeThread(QThread):
    sig = pyqtSignal(int)

    def __init__(self, files, output, parent):
        super().__init__(parent)
        self.files = files
        self.output = output

    def progress(self, val):
        self.sig.emit(val)

    def run(self):
        config = Config(self.output)
        output_file = summarize_files(list(self.files), self.output, config=config, progress=self.progress)
        self.output_file = output_file
        self.progress(-1) # TODO - type safe, nicer


parameters_help = """\
 Create a file named "summary.ini" in the file directory.
 Example contents:

[global]
parameters=Comm advance mode,Comm advance const delay up,Comm advance const delay down

[half_cycle]
;; Defaults to: Average Velocity [m/s],Flow Rate [LPM]
fields=Average Velocity [m/s],Flow Rate [LPM]
;; Defaults to: down,up,all
directions=down,up,all
"""


gui_help = f"""\
Drag files from a _Single Directory_ and click the resulting button. Opens
the result file (created in the same directory) when it is completely finished
(watch the pretty progress monitor).

Adding more parameters:
This is a manual process - no GUI:
{parameters_help}
"""


class GUI(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()
        self.files = set()
        self.output = None


    def show_help(self):
        QMessageBox.information(self, "Help", gui_help)


    def updateProgBar(self, *args, **kw):
        if args[0] == -1:
            self.onSummarizeDone(self.summarize_thread.output_file)
        self.progress.show()
        self.progress.setMaximum(len(self.files))
        self.progress.setValue(self.progress.value() + 1)
        #print(f"TODO: Progress: {args}, {kw}") # progress report is useless right now

    def summarize(self):
        summarize_thread = SummarizeThread(files=self.files, output=self.output, parent=self)
        # Connect signal to the desired function
        self.updateProgBar([0])
        summarize_thread.sig.connect(self.updateProgBar)
        summarize_thread.start()
        self.summarize_thread = summarize_thread

    def onSummarizeDone(self, output_file):
        start(output_file)
        raise SystemExit

    def initUI(self):
        self.setAcceptDrops(True)

        layout = self.layout = QGridLayout()
        layout.setSpacing(10)
        self.setLayout(layout)

        # TODO ugly hack to make grid give more space to label - better to use
        # spacing, once I learn how.
        l = ' ' * 10 + 'Drag files Here' + ' ' * 10
        self.drag_label = QLabel('/' + (' ' * (len(l) - 1 + 10)) + '\n' + l + '\n' + ' ' * (len(l) - 1 + 10) + '/')
        layout.addWidget(self.drag_label, 1, 0)
        layout.addWidget(button(title="Help", callback=self.show_help, parent=self), 2, 0)
        if os.path.exists(BUILD_NAME_FILENAME):
            with open(BUILD_NAME_FILENAME) as f:
                layout.addWidget(QLabel(f.read().strip()), 3, 0)

        self.summarize_button = button(title='Summarize', callback=self.summarize, parent=self)
        self.summarize_button.hide()
        layout.addWidget(self.summarize_button, 2, 0)
        self.setWindowTitle('Post Process xlsx summarizer')

        self.progress = QProgressBar()
        self.progress.hide()
        layout.addWidget(self.progress, 3, 0)

    def update_button_label(self, new_text):
        self.summarize_button.setText(new_text)

    def dragEnterEvent(self, e):
        e.accept()

    def dropEvent(self, e):
        # TODO: hide the drag label, show a button instead to do consolidation
        # get the relative position from the mime data
        mime = e.mimeData().text()
        files = paths_from_file_urls([x.strip() for x in mime.split('\n')])
        if len(files) == 0:
            print("no files dragged")
            return
        directory = os.path.dirname(files[0])
        self.output = directory
        if not os.path.exists(self.output):
            print(f"no such directory {self.output}")
            return
        for file in files:
            if not os.path.exists(file):
                print(f"no such file {file}")
            else:
                self.files.add(file)
        if len(self.files) > 0:
            self.update_button_label(f"{len(self.files)} to {self.output}")
            self.drag_label.hide()
            self.summarize_button.show()
        else:
            print("len(self.files) == 0")
        e.accept()


def start_gui():
    app = QApplication(sys.argv)
    ex = GUI()
    ex.show()
    app.exec_()


def main():
    parser = ArgumentParser()
    parser.add_argument('--dir')
    args = parser.parse_args()
    if args.dir is None:
        # gui mode
        start_gui()
        return

    # console mode
    output = summarize_dir(args.dir, Config(args.dir))
    if not output:
        return
    print(f"wrote {output}")
    start(output)


if __name__ == '__main__':
    main()
