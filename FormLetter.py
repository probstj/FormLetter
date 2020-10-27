#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct  1 08:38:16 2020

@author: Jürgen Probst
"""
import pandas as pd
import jinja2
import babel.numbers, babel.dates
import weasyprint
import os.path
import sys, os
import locale
#import xlrd # just as a reminder that we need to install this package


# Jinja2 documentation: https://jinja.palletsprojects.com/en/2.11.x/
# https://jinja.palletsprojects.com/en/2.11.x/api/
# https://jinja.palletsprojects.com/en/2.11.x/templates/

# http://babel.pocoo.org/en/latest/numbers.html#numbers
# http://babel.pocoo.org/en/latest/dates.html#date-and-time

# different ways to center text in html: https://www.computerhope.com/issues/ch001474.htm
# https://weasyprint.readthedocs.io/en/stable/tips-tricks.html
# Optional whitespace escaping: https://svn.python.org/projects/external/Jinja-1.1/docs/build/escaping.html


class FormLetter(object):

    def __init__(self, template, datafile, sheet_name=None):
        """Create a FormLetter object.

        :param template: filename of template file (.html file) which will
            be filled by the datafile table entries.

        :param datafile: filename of table file (e.g. .csv or .xlsx) which
            will be used to fill template.

        :param sheet_name: Sheet name of datafile to be used, if .xlsx file.
            If None (default), the first sheet will be used.

        """
        self.template_file = template

        self.datafile = datafile
        self.data_ext = os.path.splitext(datafile)[-1]
        if self.data_ext == '.xlsx':
            # open excel file
            xl = pd.ExcelFile(datafile)
            if sheet_name is None:
                self.sheet_name = xl.sheet_names[0]
            elif sheet_name in xl.sheet_names:
                self.sheet_name = sheet_name
            else:
                raise ValueError("sheet name '%s' not in file" % sheet_name)
            self.data = xl.parse(self.sheet_name)
        else:
            try:
                self.data = pd.read_csv(datafile)
            except pd.errors.ParserError as pe:
                print("unknown data file format")
                raise pe

        # drop emtpy lines (where all values are nan):
        self.data = self.data.dropna(axis=0, how='all')

        # find column names with spaces:
        columns = list(self.data.columns)
        for i, column in enumerate(columns):
            if " " in column:
                new_name = column.replace(" ", "_")
                print('WARNING: data column name "%s" contains spaces. '
                      'Will be renamed to "%s". '
                      'Please use new name in template' % (column, new_name))
                columns[i] = new_name
        self.data.columns = columns

        # prepare substitution dictionary, will be used for every row:
        self.subdict = {key: "" for key in self.data.columns}

        self.env = jinja2.Environment(
                loader=jinja2.FileSystemLoader(
                        os.path.split(self.template_file)[0]))

        # On a tested windows machine, babel wouldn't work because there
        # was no current locale set. Apparently, there is no `LC_NUMERIC`
        # environment variable, so the babel default locale of
        # `babel.default_locale('LC_NUMERIC')` returns `None`
        # Solution: Set the environment variable 'LC_ALL' since it will
        # also help with the non-existing 'LC_TIME' for example:
        self.locale = locale.getlocale()[0]
        if babel.default_locale('LC_NUMERIC') is None:
            if os.getenv('LC_ALL') is None:
                os.environ['LC_ALL'] = self.locale

        # Add some formatters to Jinja environment, so they can be
        # used in the template:
        self.env.filters['format_currency'] = babel.numbers.format_currency
        self.env.filters['format_percent'] = babel.numbers.format_percent
        self.env.filters['format_decimal'] = babel.numbers.format_decimal
        self.env.filters['format_amount'] = lambda x: babel.numbers.format_decimal(
            x, format=u'#,##0.00')
        self.env.filters['format_date'] = babel.dates.format_date
        self.env.filters['format_datetime'] = babel.dates.format_datetime
        # TODO self.env.filters['format_adapted_date'] = date_formatter

        self.template = self.env.get_template(
                os.path.split(self.template_file)[1])

        print(self.data.columns) # TODO debugging only
        print(self.data.head())
        #print(self.data.dtypes)
        print()
        print()


    def get_filled_html(self, row,
            #TODO custom_formatters=None,
            #TODO locale=None
            ):
        """Return the template, filled with the data of the specified row,
        as HTML-formatted string.

        :param row:
            the row number of data which will be used to fill the template.
            start counting at 0.
        :param custom_formatters: None or dict of one-parameter functions;
            If None (default), a set of default formatters for each
            column will be used, using babel.numbers and babel.dates.
            Individual formatting functions can be supplied with the
            column names as keys. The result of each function
            must be a unicode string.
        :param locale: None or locale identifier, e.g. 'de_DE' or 'en_US';
            The locale used for formatting numeric and date values with
            babel. If None (default), the locale will be taken from the
            `LC_NUMERIC` or `LC_TIME` environment variables on your
            system, for numeric or date values, respectively.
        :returns:
            HTML-formatted string

        """
        # prepare substitution dictionary for current row:
        self.subdict.update(self.data.iloc[row].items())
        html = self.template.render(self.subdict)

        return html


    def write_to_pdf(self, row, file_name):
        """Save the template, filled with the data of the specified row,
        as PDF file

        :param row:
            the row number of data which will be used to fill the template.
            start counting at 0.
        :param file_name: string;
            Destination file name.

        """
        html = self.get_filled_html(row)
        doc = weasyprint.HTML(string=html)
        doc.write_pdf(file_name)

    def get_data_row(self, row):
        return self.data.iloc[row]

    def get_number_of_rows(self):
        return self.data.shape[0]

def main(argv=sys.argv[1:]):
    if len(argv) < 2:
        print("usage: python Formletter.py templatefile.html datafile.xlsx "
              "[sheet_name]")
        return
    print('using template file: %s' % argv[0])
    print('using data file: %s' % argv[1])
    if len(argv) > 2:
        print('using sheet name: %s' % argv[2])
        fl = FormLetter(argv[0], argv[1], argv[2])
    else:
        fl = FormLetter(argv[0], argv[1])

    total = fl.get_number_of_rows()
    for i in range(total):
        row = fl.get_data_row(i)
        if row["1_wenn_RN_verschickt"]:
            print("skipping %i/%i: %s %s" % (i + 1, total, row["RN"], row["Person"]))
            continue
        fname = "pdf%0i_%s_%s.pdf" % (i, row["RN"], row["Person"])
        print("procesing %i/%i: file %s" % (i + 1, total, fname))
        fl.write_to_pdf(i, fname)



if __name__ == '__main__':
    main()
