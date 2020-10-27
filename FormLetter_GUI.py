#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 01:19:16 2020

@author: JÃƒÂ¼rgen Probst

"""

import sys, os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from threading import Thread, Event
import queue
import time
#import xlrd # just as a reminder that we need to install this package
import FormLetter


class PlaceholderEntry(tk.Entry):
    def __init__(self, master=None, placeholdertext="", color='grey', **kwargs):
        super().__init__(master, **kwargs)

        self.placeholder_text = placeholdertext
        self.placeholder_color = color
        self.default_fg_color = self['fg']

        self.bind("<FocusIn>", self.focus_in)
        self.bind("<FocusOut>", self.focus_out)

        self.write_placeholder()

    def write_placeholder(self):
        self.insert(0, self.placeholder_text)
        self['fg'] = self.placeholder_color

    def focus_in(self, *args):
        if self['fg'] == self.placeholder_color:
            self.delete('0', tk.END)
            self['fg'] = self.default_fg_color

    def focus_out(self, *args):
        if not self.get():
            self.write_placeholder()



class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(fill=tk.BOTH, expand=1, padx=5, pady=5)
        self.thread1 = None
        self.stop_thread = Event()
        self.queue = queue.Queue()
        self.num_to_convert = 0
        self.create_widgets()
        self.sheet_name = self.sheet_names = None
        self.xlbook = None
        self.data = None
        self.datafilename = None

    def create_widgets(self):
        self.ttkstyle = ttk.Style()
        self.ttkstyle.theme_use('clam')
        self.ttkstyle.map(
            'custom.TCombobox',
            fieldbackground=[('readonly','white'), ('disabled', 'gray80')],
            foreground=[('readonly','black')],
            # hide selection when out-of-focus:
            selectbackground=[('readonly', '!focus', 'white')],
            selectforeground=[('readonly', '!focus', 'black')]
            )
        self.ttkstyle.map(
            'TEntry',
            fieldbackground=[('disabled', 'gray80')])
        self.ttkstyle.configure("custom.TButton", padding=2)
        self.ttkstyle.configure("red.TButton", padding=2, foreground='red')

        # https://stackoverflow.com/questions/24768090/progressbar-in-tkinter-with-a-label-inside#40348163
        self.ttkstyle.layout("TProgressbar",
         [('TProgressbar.trough',
           {'children': [('LabeledProgressbar.pbar',
                          {'side': 'left', 'sticky': 'ns'}),
                         ("LabeledProgressbar.label",
                          {"sticky": ""})],
           'sticky': 'nswe'})])

        fileframe = tk.LabelFrame(self, text="Files", pady=5, padx=5)
        fileframe.pack(side=tk.TOP, fill=tk.X, expand=0)

        topframe = tk.Frame(fileframe, pady=0, padx=0)
        topframe.pack(side=tk.TOP, fill=tk.X, expand=0)

        tk.Label(topframe, text="Template file (*.html):   ").pack(side=tk.LEFT)
        self.templatefile_edt = ttk.Entry(topframe, width=1)
        self.templatefile_edt.pack(side=tk.LEFT, fill=tk.X, expand=1, ipady=2)
        self.templatefile_edt.bind('<Return>', self.templatefile_edt_return)
        ttk.Button(
            topframe, text="...", style="custom.TButton", width=4,
            command=self.open_template_dialog).pack(
                side=tk.LEFT, fill=tk.X)

        topframe = tk.Frame(fileframe, pady=5, padx=0)
        topframe.pack(side=tk.TOP, fill=tk.X, expand=0)

        tk.Label(topframe, text="Data file (*.xlsx / *.csv): ").pack(side=tk.LEFT)
        self.datafile_edt = ttk.Entry(topframe, width=1)
        self.datafile_edt.pack(side=tk.LEFT, fill=tk.X, expand=1, ipady=2)
        self.datafile_edt.bind('<Return>', self.datafile_edt_return)
        ttk.Button(
            topframe, text="...", style="custom.TButton", width=4,
            command=self.open_data_dialog).pack(
                side=tk.LEFT, fill=tk.X)
        tk.Label(topframe, text="  Sheet: ").pack(side=tk.LEFT)
        self.sheets_combo = ttk.Combobox(
                topframe, values=[], width=1,
                state='disabled', style="custom.TCombobox")
        self.sheets_combo.pack(side=tk.LEFT, fill=tk.X, expand=1, ipady=2)
        self.sheets_combo.bind('<<ComboboxSelected>>', self.update_sheet)


        tk.Label(self, text="").pack(side=tk.TOP)
        #ttk.Separator(self, orient="horizontal").pack(side=tk.TOP, pady=5)

        frame = tk.LabelFrame(self, text="Options", pady=5, padx=5)
        frame.pack(side=tk.TOP, fill=tk.X, expand=0)

        self.ttkstyle.configure("TCheckbutton", background=frame["background"])
        self.skip_var = tk.IntVar(0)
        ttk.Checkbutton(
            frame, text=" Skip dataset if ", command=self.on_skipcheck,
            var=self.skip_var).pack(side=tk.LEFT)
        self.skip_combo = ttk.Combobox(
                frame, values=[], width=1,
                state='disabled', style="custom.TCombobox")
        self.skip_combo.pack(
                side=tk.LEFT, fill=tk.X, expand=1, ipady=2, pady=5)
        tk.Label(frame, text=" equals ").pack(side=tk.LEFT)
        self.skip_edt = ttk.Entry(frame, width=1, state='disabled')
        self.skip_edt.pack(side=tk.LEFT, fill=tk.X, expand=1, ipady=2)



        tk.Label(self, text="").pack(side=tk.TOP)
        frame = tk.LabelFrame(self, text="Output", pady=5, padx=5)
        frame.pack(side=tk.TOP, fill=tk.X, expand=0)

        tk.Grid.columnconfigure(frame, 1, weight=1)
        tk.Grid.columnconfigure(frame, 3, weight=1)
        tk.Label(frame, text="Save pdf-files in: ").grid(
                row=0, column=0, sticky='w')
        self.dir_edt = ttk.Entry(frame, width=1)
        self.dir_edt.grid(row=0, column=1, sticky='nwse', columnspan=3)
        ttk.Button(
            frame, text="...", style="custom.TButton", width=4,
            command=self.choose_dest_folder).grid(row=0, column=4)

        tk.Label(frame, text="Output file name: ").grid(
                row=1, column=0, sticky='w')
        self.destfile_edt = ttk.Entry(frame, width=1)
        self.destfile_edt.grid(row=1, column=1, sticky='nwse', columnspan=3)
        self.destfile_edt.bind("<Key>", self.destfile_edt_focus_out)
        self.destfile_edt.bind("<Button>", self.destfile_edt_focus_out)
        self.destfile_edt.insert(0, "{i:04}_{RN}_{Person}.pdf")
        self.destfile_edt["state"] = "readonly" # TODO
        ttk.Button(
            frame, text="?", style="custom.TButton", width=4,
            command=self.show_destfile_help).grid(row=1, column=4)


        self.conversion_selection_var = tk.IntVar()

        tk.Label(frame, text="").grid(row=2)

        self.ttkstyle.configure("TRadiobutton", background=frame["background"])
        r1 = ttk.Radiobutton(
                frame, text='convert all',
                variable=self.conversion_selection_var, value=1)
        r1.grid(row=3, column=0, sticky='w')
        r2 = ttk.Radiobutton(
                frame, text='convert range:',
                variable=self.conversion_selection_var, value=2)
        r2.grid(row=4, column=0, sticky='w')
        r3 = ttk.Radiobutton(
                frame, text='convert selection: ',
                variable=self.conversion_selection_var, value=3)
        r3.grid(row=5, column=0, sticky='w')
        self.convert_from_spinbox = tk.Spinbox(
                frame, from_=1, to=1, bg='white', state='normal',
                command=self.select_r2)
        self.convert_from_spinbox.grid(row=4, column=1, sticky='we')
        tk.Label(frame, text='to').grid(row=4, column=2, sticky='w')
        self.convert_to_spinbox = tk.Spinbox(
                frame, from_=1, to=1, bg='white', state='normal',
                command=self.select_r2)
        self.convert_to_spinbox.grid(row=4, column=3, sticky='we')
        self.convert_from_spinbox.bind("<Key>", self.select_r2)
        self.convert_to_spinbox.bind("<Key>", self.select_r2)
        self.convert_selection_entry = PlaceholderEntry(
                frame, " for example: 1, 3-5, 7, 9", bg='white')
        self.convert_selection_entry.grid(
                row=5, column=1, sticky='we', columnspan=3)
        self.convert_selection_entry.bind("<Key>", self.select_r3)
        self.conversion_selection_var.set(1)



        tk.Label(self, text="").pack(side=tk.TOP)
        frame = tk.LabelFrame(self, text="Progress", pady=5, padx=5)
        frame.pack(side=tk.BOTTOM, fill=tk.X, expand=0)

        self.progressbar = ttk.Progressbar(frame, value=0, maximum=100,
                     mode="determinate",
                     #orient=tk.HORIZONTAL,
                     style="TProgressbar")
        self.progressbar.pack(side=tk.LEFT, fill=tk.X, expand=1)
        # change the text of the progressbar,
        # the trailing spaces are here to properly center the text
        self.ttkstyle.configure("TProgressbar", text="0 %      ")



        #tk.Label(self, text="").pack(side=tk.TOP)
        #tk.Label(self, text="").pack(side=tk.BOTTOM)
        frame = tk.Frame(self, pady=5, padx=5)
        frame.pack(side=tk.BOTTOM, fill=tk.Y, expand=1, anchor="center")

        self.go_button = ttk.Button(
            frame, text="Go!", style="custom.TButton",
            command=self.run_conversion)
        self.go_button.pack(
                side=tk.LEFT, fill=tk.X)
        tk.Label(frame, text=" ").pack(side=tk.LEFT, padx=5)
        ttk.Button(
            frame, text="Exit", style="custom.TButton",
            command=self.leave).pack(
                side=tk.LEFT, fill=tk.X)

    def templatefile_edt_return(self, event):
        if not self.templatefile_edt.get():
            # no filename entered yet:
            self.open_template_dialog()
        else:
            # do nothing, filename will be read later
            pass

    def datafile_edt_return(self, event):
        if not self.datafile_edt.get():
            # no filename entered yet:
            self.open_data_dialog()
        else:
            self.open_data_file(self.datafile_edt.get())

    def open_template_dialog(self):
        fname = filedialog.askopenfilename(
            filetypes=[('HTML', '*.html'), ('all files', '*.*')])
        if fname:
            self.templatefile_edt.delete(0, tk.END)
            self.templatefile_edt.insert(0, fname)

    def open_data_dialog(self):
        fname = filedialog.askopenfilename(
            filetypes=[('Excel files', '*.xlsx *.xls *.csv'), ('all files', '*.*')])
        if fname:
            self.datafile_edt.delete(0, tk.END)
            self.datafile_edt.insert(0, fname)
            self.open_data_file(fname)

    def open_data_file(self, fname, keep_selected_sheet=False):
        self.datafilename = fname
        ext = os.path.splitext(fname)[-1]
        if ext in ['.xlsx', '.xls']:
            # open excel file
            self.xlbook = pd.ExcelFile(fname)
            self.sheet_names = self.xlbook.sheet_names
            self.sheets_combo.config(state='readonly')
            self.sheets_combo["values"] = self.sheet_names
            if (keep_selected_sheet
                and self.sheet_name is not None
                and self.sheet_name in self.sheet_names):
                self.sheets_combo.set(self.sheet_name)
            else:
                # choose first sheet:
                self.sheets_combo.current(newindex=0)
            self.update_sheet()
        else:
            self.sheets_combo.set("")
            self.sheets_combo.config(state='disabled')
            self.sheets_combo["values"] = []
            self.xlbook = None
            self.sheet_names = self.sheet_name = None
            try:
                self.data = pd.read_csv(fname)
                self.clean_up_data()
                self.update_data_columns()
            except pd.errors.ParserError as pe:
                print("unknown data file format")
                raise pe


        self.convert_from_spinbox["to"] = self.data.shape[0]
        self.convert_from_spinbox.delete(0, tk.END)
        self.convert_from_spinbox.insert(0, 1)
        self.convert_to_spinbox["to"] = self.data.shape[0]
        self.convert_to_spinbox.delete(0, tk.END)
        self.convert_to_spinbox.insert(0, self.data.shape[0])

    def update_sheet(self, event=None):
        #if event is not None:
        #    event.widget.selection_clear()
        self.sheet_name = self.sheets_combo.get()
        self.data = self.xlbook.parse(self.sheet_name)
        self.clean_up_data()
        self.update_data_columns()

    def clean_up_data(self):
        # drop emtpy lines (where all values are nan):
        self.data = self.data.dropna(axis=0, how='all')

    def update_data_columns(self):
        self.data_columns = list(self.data.columns)
        self.skip_combo["values"] = self.data_columns

    def on_skipcheck(self):
        if self.skip_var.get():
            self.skip_combo["state"] = "readonly"
            self.skip_edt["state"] = "normal"
        else:
            self.skip_combo["state"] = "disabled"
            self.skip_edt["state"] = "disabled"

    def choose_dest_folder(self):
        result = filedialog.askdirectory(
            initialdir=os.path.curdir)
        if result:
            self.dir_edt.delete(0, tk.END)
            self.dir_edt.insert(0, result)

    def destfile_edt_focus_out(self, event=None):
        #TODO check if right format, add ending .pdf
        messagebox.showinfo("Under construction", "Sorry, the destination file name cannot be changed yet.")

    def show_destfile_help(self):
        messagebox.showinfo("Help", "under construction...") #TODO

    def select_r2(self, event=None):
        self.conversion_selection_var.set(2)

    def select_r3(self, event=None):
        self.conversion_selection_var.set(3)

    def get_indexes_to_convert(self):
        start = 1
        count = self.data.shape[0]
        end = count - start + 1
        if self.conversion_selection_var.get() == 1:
            # convert all
            return range(start - 1, end)
        elif self.conversion_selection_var.get() == 2:
            # convert from..to
            start = max(start, int(self.convert_from_spinbox.get()))
            end = min(end, int(self.convert_to_spinbox.get()))
            return range(start - 1, end)
        else:
            rangetxt = self.convert_selection_entry.get().strip()
            if rangetxt.startswith("for example") or not rangetxt:
                raise ValueError("Please specify pages to convert")
            try:
                entries = rangetxt.split(",")
                lst = []
                for entry in entries:
                    if not entry:
                        continue
                    if '-' in entry:
                        f, t = entry.split('-')
                        lst.extend(
                            range(
                                max(start, int(f.strip())) - 1,
                                min(end, int(t.strip()))))
                    else:
                        num = int(entry.strip())
                        if num <= end and num >= start:
                            lst.append(num - 1)
                return list(set(sorted(lst)))
            except ValueError:
                raise ValueError("Wrong format used in page selection text")

    # https://stackoverflow.com/questions/15323574/how-to-connect-a-progress-bar-to-a-function
    def periodic_call(self):
        # check for progress updates from queue:
        self.check_queue()
        if self.thread1 and self.thread1.is_alive():
            self.master.after(100, self.periodic_call)
        else:
            self.thread1 = None
            # finished! Return button to original state:
            self.go_button["style"] = 'custom.TButton'
            self.go_button["text"] = "Go!"
            self.go_button["command"] = self.run_conversion

    def check_queue(self):
        last = None
        while self.queue.qsize():
            try:
                last = self.queue.get_nowait()
            except queue.Empty:
                pass
        if last:
            self.progressbar["value"] = last
            self.ttkstyle.configure(
                "TProgressbar",
                text="%i/%i       " % (last, self.num_to_convert))

    def stop(self):
        self.stop_thread.set()
        self.thread1.join()

    def run_conversion(self):

        # check that everything is filled in nicely:

        tempf = self.templatefile_edt.get()
        if not tempf:
            messagebox.showerror("Error", "Please choose template file.")
            return
        # TODO do this check on <Focus_Out> of entry:
        if not os.path.isfile(tempf):
            messagebox.showerror(
                "Error", 'Template file "%s" does not exist.' % tempf)
            return

        dataf = self.datafile_edt.get()
        if not dataf:
            messagebox.showerror("Error", "Please choose data file.")
            return
        # TODO do this check on <Focus_Out> of entry:
        if self.datafilename != dataf:
            # user changed file name without hitting enter, reload:
            if not os.path.isfile(dataf):
                messagebox.showerror(
                    "Error", 'Data file "%s" does not exist.' % dataf)
                return
            self.open_data_file(dataf, keep_selected_sheet=True)

        skip_data = bool(self.skip_var.get())
        if skip_data:
            skip_column = self.skip_combo.get()
            if skip_column not in self.data_columns:
                messagebox.showerror(
                    "Error",
                    'Data column "%s" does not exist (skip data).' % skip_column)
                return
            skip_value = self.skip_edt.get()
            if not skip_value:
                messagebox.showerror(
                    "Error",
                    "Please fill in a proper value when to skip data.")
                return
        else:
            skip_value = skip_column = None

        destfile_format = self.destfile_edt.get()
        if not destfile_format:
            messagebox.showerror(
                "Error",
                "Please fill in a proper destination file name.")
            return

        destdir = self.dir_edt.get()
        if not destdir:
            messagebox.showerror(
                "Error",
                "Please fill in a proper destination folder.")
            return
        if os.path.isfile(destdir):
            messagebox.showerror(
                "Error",
                "Destination folder is an already existing file.")
            return
        if os.path.isdir(destdir):
            result = messagebox.askokcancel(
                "Directory exists",
                "Warning: Destination directory already exists. "
                "If you continue, contained files will be overwritten.")
            if not result:
                # user chose 'cancel'
                return
        else:
            # create directory:
            os.makedirs(destdir)

        indexes = self.get_indexes_to_convert()
        self.num_to_convert = len(indexes)
        self.progressbar["max"] = self.num_to_convert

        self.stop_thread.clear()
        self.thread1 = Thread(
            target=self.secondary_thread_loop,
            kwargs={
                "templatefile": tempf,
                "datafile": dataf,
                "sheet": self.sheet_name,
                "do_skip_data": skip_data,
                "skip_data_column": skip_column,
                "skip_data_value": skip_value,
                "destfile_format": destfile_format,
                "destdir": destdir,
                "indexes": indexes})
        self.thread1.start()
        self.go_button["style"] = 'red.TButton'
        self.go_button["text"] = "Stop"
        self.go_button["command"] = self.stop
        self.periodic_call()

    def secondary_thread_loop(
            self, templatefile, datafile, sheet,
            do_skip_data, skip_data_column, skip_data_value,
            destfile_format, destdir, indexes):
        print('using template file: %s' % templatefile)
        print('using data file: %s' % datafile)
        if sheet is not None:
            print('using sheet name: %s' % sheet)

        fl = FormLetter.FormLetter(templatefile, datafile, sheet)
        # workaround: fix skip_data_column name:
        if do_skip_data:
            skip_data_column = skip_data_column.replace(" ", "_")
            # another workaround: convert skip_data_value into proper type:
            skip_data_dtype = fl.data[skip_data_column].dtype
            skip_data_value = pd.np.array(
                    [skip_data_value,]).astype(skip_data_dtype)[0]
        total = len(indexes)

        for i, rownum in enumerate(indexes):
            if self.stop_thread.is_set():
                break
            row = fl.get_data_row(rownum)
            fname = os.path.join(
                destdir, destfile_format.format(i=rownum+1, **row))
            if do_skip_data and row[skip_data_column] == skip_data_value:
                print("skipping %i/%i: file %s" % (rownum+1, total, fname))
                time.sleep(0.001) # just in case we skip a lot
                # communicate progress:
                self.queue.put(i + 1)
                continue

            print("procesing %i/%i: file %s" % (rownum+1, total, fname))
            fl.write_to_pdf(i, fname)
            # communicate progress:
            self.queue.put(i + 1)

    def leave(self):
        if self.thread1:
            self.stop()
        self.master.destroy()


def main():
    root = tk.Tk()
    app = Application(master=root)
    # set window title
    root.title("FormLetter")
    root.geometry("640x500")

    def on_closing():
        #if messagebox.askokcancel("Quit", "Do you want to quit?"):
        app.leave()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    app.mainloop()

if __name__ == '__main__':
    main()
