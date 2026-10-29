"""Simple Windows/macOS desktop UI. Run with: python src/app.py."""
import calendar
import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from openpyxl.utils import get_column_letter

from mapping import build_report, read_rows, write_workbook
from audit_report import write_pdf

BLANK_POLICIES = {
    'Leave empty (default)': 'blank',
    'Keep existing values': 'preserve',
    'Set missing amounts to zero': 'zero',
    'Stop if amounts are missing': 'error',
}


def output_paths(folder, reporting_period, mode='workbook'):
    """Number the file being saved; other output types don't consume numbers."""
    month = calendar.month_name[int(reporting_period.split('-')[1])]
    number = 0
    while True:
        suffix = f'_{number}' if number else ''
        workbook = folder / f'TB_{month}{suffix}.xlsx'
        audit = folder / f'TB_{month}{suffix}_audit.{"pdf" if mode == "pdf" else "json"}'
        if not (workbook if mode == 'workbook' else audit).exists():
            return workbook, audit
        number += 1


def convert(options):
    """Run outside the UI thread so the window stays responsive."""
    source = Path(options['input'])
    template = Path(options['template'])
    mapping = Path(options['mapping'])
    folder = Path(options['folder']) if options['folder'] else source.parent
    for label, path in [('Trial balance', source), ('RE Only template', template), ('Mapping', mapping)]:
        if not path.is_file():
            raise ValueError(f'Cannot find the {label.lower()} file. Click Browse and select it again.\n{path}')
    if source.suffix.lower() != '.xlsx' or template.suffix.lower() != '.xlsx':
        raise ValueError('Trial balance and RE Only template must be .xlsx workbooks.')
    if not folder.is_dir():
        raise ValueError(f'Cannot find the output folder. Choose an existing folder or leave it empty to use the input folder.\n{folder}')

    report, amounts, column = build_report(
        read_rows(source, options['source_sheet']), read_rows(template, 'TB RE'),
        mapping, mapping_sheet=options['mapping_sheet'])
    mode = options.get('mode', 'workbook')
    if mode not in ('workbook', 'pdf', 'json'):
        raise ValueError('Choose Run, Audit PDF, or JSON (for developing).')
    audit_only = mode != 'workbook'
    output, audit = output_paths(folder, report['period'], mode)
    blockers = list(report['errors'])
    if report['unmapped_accounts'] and not options['allow_unmapped']:
        blockers.append('Some source accounts are unmapped. Update the mapping or enable partial output.')
    policy = BLANK_POLICIES[options['blank_policy']]
    if report['unresolved_destination_lines'] and policy == 'error':
        blockers.append('Some destination lines have no source amount. Choose how to handle missing amounts.')
    report.update(missing_lines_policy=policy, write_blockers=blockers,
                  output_written=False, audit_only=audit_only)
    if not audit_only and not blockers:
        # Reserve the filename exclusively, including when another app is running.
        with output.open('xb'):
            pass
        try:
            write_workbook(template, output, report, amounts, column, policy)
            report['output_written'] = True
        except Exception:
            output.unlink(missing_ok=True)
            raise
    if audit_only:
        if mode == 'pdf':
            write_pdf(audit, report)
        else:
            with audit.open('x', encoding='utf-8') as handle:
                json.dump(report, handle, indent=2, default=str)
    return output, audit if audit_only else None, report


class Application:
    """Selections live only in this window; nothing is saved as preferences."""
    def __init__(self, root):
        self.root = root
        root.title('Arcamed - Trial Balance to TB RE')
        root.minsize(700, 470)
        self.results = queue.Queue()
        self.running = False
        self.values = {name: tk.StringVar() for name in ('input', 'folder', 'template', 'mapping')}
        self.values.update(source_sheet=tk.StringVar(value='Report1 (6)'),
                           mapping_sheet=tk.StringVar(value='Mappings'),
                           blank_policy=tk.StringVar(value='Leave empty (default)'),
                           allow_unmapped=tk.BooleanVar(value=False))
        self.status = tk.StringVar(value='Choose your files, then click Run.')
        panel = ttk.Frame(root, padding=20)
        panel.pack(fill='both', expand=True)
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text='Trial Balance to TB RE', font=('', 18, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky='w', pady=(0, 18))
        self.controls = []
        for row, (name, label) in enumerate([
            ('input', 'Trial balance'), ('folder', 'Output folder'),
            ('template', 'RE Only template'), ('mapping', 'Mapping file')], 1):
            ttk.Label(panel, text=label).grid(row=row, column=0, sticky='w', padx=(0, 12), pady=6)
            entry = ttk.Entry(panel, textvariable=self.values[name])
            entry.grid(row=row, column=1, sticky='ew', pady=6)
            button = ttk.Button(panel, text='Browse...', command=lambda key=name: self.browse(key))
            button.grid(row=row, column=2, padx=(10, 0))
            self.controls.extend([entry, button])
        ttk.Label(panel, text='Leave output folder empty to use the trial balance folder.').grid(
            row=5, column=1, columnspan=2, sticky='w', pady=(0, 12))
        ttk.Label(panel, text='Missing amounts').grid(row=6, column=0, sticky='w')
        policy = ttk.Combobox(panel, textvariable=self.values['blank_policy'],
                              values=list(BLANK_POLICIES), state='readonly')
        policy.grid(row=6, column=1, sticky='ew')
        self.policy = policy
        ttk.Label(panel, text='Rows stay in place. Formulas and history are retained.').grid(
            row=7, column=1, columnspan=2, sticky='w', pady=(4, 10))
        partial = ttk.Checkbutton(panel, text='Allow partial output when source accounts are unmapped',
                                  variable=self.values['allow_unmapped'])
        partial.grid(row=8, column=0, columnspan=3, sticky='w')
        self.controls.append(partial)
        advanced = ttk.LabelFrame(panel, text='Sheet names (change only if needed)', padding=10)
        advanced.grid(row=10, column=0, columnspan=3, sticky='ew', pady=14)
        advanced.columnconfigure(1, weight=1)
        for row, (key, label) in enumerate([('source_sheet', 'Trial balance tab'), ('mapping_sheet', 'Excel mapping tab')]):
            ttk.Label(advanced, text=label).grid(row=row, column=0, sticky='w', padx=(0, 12))
            entry = ttk.Entry(advanced, textvariable=self.values[key])
            entry.grid(row=row, column=1, sticky='ew', pady=3)
            self.controls.append(entry)
        actions = ttk.Frame(panel)
        actions.grid(row=11, column=0, columnspan=3, sticky='w')
        self.run_button = ttk.Button(actions, text='Run', command=self.run)
        self.run_button.pack(side='left')
        for mode, label in [('pdf', 'Audit PDF'), ('json', 'JSON (for developing)')]:
            button = ttk.Button(actions, text=label, command=lambda selected=mode: self.run(selected))
            button.pack(side='left', padx=(10, 0))
            self.controls.append(button)
        ttk.Label(panel, textvariable=self.status, wraplength=520).grid(
            row=12, column=0, columnspan=3, sticky='w', pady=(10, 0))
        root.protocol('WM_DELETE_WINDOW', self.close)

    def browse(self, name):
        if name == 'folder':
            initial = self.values['folder'].get() or str(Path(self.values['input'].get()).parent)
            chosen = filedialog.askdirectory(parent=self.root, initialdir=initial)
        else:
            types = [('Excel workbook', '*.xlsx')]
            if name == 'mapping':
                types.append(('CSV mapping', '*.csv'))
            chosen = filedialog.askopenfilename(parent=self.root, filetypes=types)
        if chosen:
            self.values[name].set(chosen)

    def set_busy(self, busy):
        self.running = busy
        for control in self.controls + [self.run_button]:
            control.configure(state='disabled' if busy else 'normal')
        self.policy.configure(state='disabled' if busy else 'readonly')

    def run(self, mode='workbook'):
        options = {name: value.get().strip() if isinstance(value.get(), str) else value.get()
                   for name, value in self.values.items()}
        if not all(options[key] for key in ('input', 'template', 'mapping', 'source_sheet', 'mapping_sheet')):
            messagebox.showerror('Files needed', 'Choose the trial balance, RE Only template, mapping file and sheet names.')
            return
        self.set_busy(True)
        options['mode'] = mode
        self.status.set('Generating workbook...' if mode == 'workbook' else 'Checking files...')

        def work():
            try:
                self.results.put(('result', convert(options)))
            except Exception as exc:
                self.results.put(('error', friendly_error(exc)))

        threading.Thread(target=work, daemon=True).start()
        self.root.after(100, self.check_result)

    def check_result(self):
        # Tkinter widgets and dialogs must be updated on the main thread.
        try:
            kind, result = self.results.get_nowait()
        except queue.Empty:
            self.root.after(100, self.check_result)
            return
        self.set_busy(False)
        if kind == 'error':
            self.status.set('Could not generate output.')
            messagebox.showerror('Conversion failed', result)
            return
        output, audit, report = result
        if audit is not None:
            self.status.set(f'Audit saved: {audit.name}. No workbook created.')
            details = '\n'.join(report['write_blockers']) or 'All validation checks passed.'
            dialog = messagebox.showwarning if report['write_blockers'] else messagebox.showinfo
            dialog('Audit complete', f'Audit: {audit}\n\n{details}')
        elif report['write_blockers']:
            self.status.set('Needs review. No files saved.')
            messagebox.showwarning('Needs review', '\n'.join(report['write_blockers']) +
                                   '\n\nClick Audit PDF to save the full findings.')
        else:
            target = f"TB RE, {report['period']} (column {get_column_letter(report['month_column'])})"
            self.status.set(f'Updated {target}. Saved {output.name}')
            partial = '\nPartial output: some source accounts remain unmapped.' if report['unmapped_accounts'] else ''
            messagebox.showinfo('Success', f'Updated: {target}\nMapped reporting lines: {len(report["comparisons"])}'
                                f'\nSaved: {output}{partial}\n\nOpen this output file in Excel and save to refresh formulas.'
                                '\nThe original template has not been changed.')

    def close(self):
        if self.running:
            messagebox.showinfo('Conversion running', 'Wait for the conversion to finish before closing.')
        else:
            self.root.destroy()


def friendly_error(error):
    """Explain file/setup errors without showing a Python traceback."""
    if isinstance(error, PermissionError):
        return 'Cannot read or save this file. Close it in Excel and choose a folder you can write to.'
    if isinstance(error, FileExistsError):
        return 'Another file was saved with the same name. Click the button again to choose the next number.'
    if isinstance(error, ImportError):
        return 'A required package is missing. Run: python -m pip install -r requirements.txt'
    if isinstance(error, KeyError):
        return f'A required worksheet or field is missing: {error}. Check the selected files and sheet names.'
    if isinstance(error, ValueError):
        return str(error)
    return f'Could not read or save the selected files. Check that they are valid Excel workbooks and try again. Details: {error}'


if __name__ == '__main__':
    window = tk.Tk()
    Application(window)
    window.mainloop()
