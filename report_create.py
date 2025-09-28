# import os, time, re, sys
import traceback, sys # for pretty-printing any issues that happened during runtime; if we hit FileNotFound I don't appreciate when a log traceback is shown, the error should be simple and clear
from datetime import datetime
# from dateutil import tz
import argparse
from pathlib import Path
import re
import json

import pandas as pd
import xlsxwriter



if __name__ == '__main__':
    # run as a program
    import util_format_mddread_inputs
    import plugins
    # import format_sheet as excel_openpyxl_format_sheet
    # import format_sheet_overview as excel_openpyxl_format_sheet_overview
    import format_sheet_xlsxwriter as excel_xlsxwriter_format_sheet
    import format_sheet_xlsxwriter_overview as excel_xlsxwriter_format_sheet_overview
elif '.' in __name__:
    # package
    from . import util_format_mddread_inputs
    from . import plugins
    # from . import format_sheet_openpyxl as excel_openpyxl_format_sheet
    # from . import format_sheet_openpyxl_overview as excel_openpyxl_format_sheet_overview
    from . import format_sheet_xlsxwriter as excel_xlsxwriter_format_sheet
    from . import format_sheet_xlsxwriter_overview as excel_xlsxwriter_format_sheet_overview
else:
    # included with no parent package
    import util_format_mddread_inputs
    import plugins
    # import format_sheet as excel_openpyxl_format_sheet
    # import format_sheet_overview as excel_openpyxl_format_sheet_overview
    import format_sheet_xlsxwriter as excel_xlsxwriter_format_sheet
    import format_sheet_xlsxwriter_overview as excel_xlsxwriter_format_sheet_overview











class ReportDocument:
    class CellNotFound(Exception):
        """Cell not found"""
    def __init__(self,inp={},config={}):

        self.config = config

        # if "columns" not in self.config or not self.config["columns"]:
        #     self.config["columns"] = ['name']
        
        self.config['columns'] = [ c for c in (inp['report_scheme']['columns'] if ('columns' in inp['report_scheme']) else []) ] if 'report_scheme' in inp else []
        if 'flags' not in self.config:
            self.config['flags'] = []

        self.dataframes = util_format_mddread_inputs.prep_dataframes(inp,self.config)

        def call_plugin_on_df(df):
            for plugin in plugins.plugins:
                if plugin['active'] and plugin['should_run'](self.config):
                    if 'on_dataframe' in plugin:
                        df['df'] = plugin['on_dataframe'](df['df'],df['name'],self.config)
            return df
        self.dataframes = [
            call_plugin_on_df(df) for df in self.dataframes
        ]
    


    def write_to_file(self,out_filename):

        # # config = self.config

        # with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
        #     for o in self.dataframes:
        #         o['df'].to_excel(writer, sheet_name=o['name'])
        #         format_fn = excel_openpyxl_format_sheet.format_sheet if not(o['name']=='overview') else excel_openpyxl_format_sheet_overview.format_sheet
        #         format_fn(writer.sheets[o['name']])

        workbook = xlsxwriter.Workbook(out_filename)

        def write_cell(worksheet,r,c,cell):
            def sanitize_final(val):
                if isinstance(val, str):
                    # Replace carriage returns with linefeed
                    val = val.replace("\r", "\n")
                    # Remove other invalid control chars
                    val = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", val)
                    return val
                else:
                    # For numbers, bools, None → leave as is
                    return val
            if isinstance(cell, list):
                # Build args for write_rich_string: [format, text, format, text, ...]
                args = []
                text_warning_add = '(Error: text length exeeds Excel limits of 32767 characters) '
                allowed_text_limit = 32765 - len(text_warning_add)
                reached_text_limit = 0
                for text, color_id in cell:
                    text = sanitize_final(text)
                    color = None
                    if not color_id or (color_id==''):
                        color = None
                    elif color_id=='changed':
                        color = '#ffe49c'
                    elif color_id=='added':
                        color = '#6bc795'
                    elif color_id=='removed':
                        color = '#f59278'
                    else:
                        color = '#dddddd'
                    fmt = workbook.add_format({"color": color}) if color else workbook.add_format()
                    reached_text_limit = reached_text_limit + len(text)
                    if reached_text_limit>allowed_text_limit:
                        text = text[:len(text)+allowed_text_limit-reached_text_limit]
                        if len(text)>0:
                            args.extend([fmt, text])
                        args.insert(0,text_warning_add)
                        args.insert(0,workbook.add_format())
                        break
                    if len(text)>0:
                        args.extend([fmt, text])
                try:
                    if len(args)==0:
                        worksheet.write(r, c, '')
                    elif len(args)>2:
                        worksheet.write_rich_string(r, c, *args)
                    else:
                        worksheet.write(r, c, args[1], args[0])
                except TypeError as e:
                    if len(args)==0:
                        worksheet.write(r, c, '')
                    elif len(args)>2:
                        worksheet.write_rich_string(r, c,[a if i%2==0 else '{t}'.format(t=a) for i,a in enumerate(args)])
                    else:
                        worksheet.write(r, c, '{t}'.format(t=args[1]), args[0])
            else:
                cell = sanitize_final(cell)
                try:
                    worksheet.write(r, c, cell)
                except TypeError as e:
                    worksheet.write(r, c, '{t}'.format(t=cell))

        for o in self.dataframes:
            worksheet = workbook.add_worksheet(o['name'])

            df = o['df']

            write_cell(worksheet, 0, 0, df.index.name or "")
            for c, col_name in enumerate(df.columns):
                write_cell(worksheet,0, c + 1, col_name)
            for r, idx in enumerate(df.index):
                write_cell(worksheet,r + 1, 0, idx)  # +1 because row 0 is for headers

            for r, row in enumerate(df.itertuples(index=False)):
                for c, cell in enumerate(row):
                    write_cell(worksheet,r+1,c+1,cell)
                    
            # format_fn = excel_openpyxl_format_sheet.format_sheet if not(o['name']=='overview') else excel_openpyxl_format_sheet_overview.format_sheet
            # format_fn(worksheet)
            format_fn = excel_xlsxwriter_format_sheet.format_sheet if not(o['name']=='overview') else excel_xlsxwriter_format_sheet_overview.format_sheet
            format_fn(workbook, worksheet, nrows=len(df.index)+1, ncols=len(df.columns)+1, config=self.config)
            for plugin in plugins.plugins:
                if plugin['active'] and plugin['should_run'](self.config):
                    if 'on_format_sheet' in plugin:
                        plugin['on_format_sheet'](workbook, worksheet, nrows=len(df.index)+1, ncols=len(df.columns)+1, name=o['name'], config=self.config)

        workbook.close()



    # @staticmethod
    # def has_value_numeric(arg):
    #     if pd.isna(arg):
    #         return False
    #     if arg is None:
    #         return False
    #     if arg==0:
    #         return True
    #     if arg == False:
    #         return True # false evaluates to 0 which is numeric
    #     if arg=='':
    #         return False
    #     return not not arg

    # @staticmethod
    # def has_value_text(arg):
    #     if pd.isna(arg):
    #         return False
    #     if arg is None:
    #         return False
    #     if arg==0:
    #         return True
    #     if arg == False:
    #         return False
    #     if arg=='':
    #         return False
    #     return not not arg





def entry_point(config={}):
    try:
        time_start = datetime.now()
        script_name = 'mdmtoolsap excel report script'

        parser = argparse.ArgumentParser(
            description="Produce a summary of input file in excel (read from json)",
            prog='mdmtoolsap --program report_excel'
        )
        parser.add_argument(
            '--inpfile',
            help='JSON with File Data',
            type=str,
            required=True
        )
        parser.add_argument(
            '--output-format',
            help='Set output format: html or excel',
            type=str,
            required=False
        )
        parser.add_argument(
            '--flags',
            help='(Optional) set of flags (comma-delimited) to control program behavior)',
            type=str,
            required=False
        )
        args = None
        args_rest = None
        if( ('arglist_strict' in config) and (not config['arglist_strict']) ):
            args, args_rest = parser.parse_known_args()
        else:
            args = parser.parse_args()
        input_map_filename = None
        if args.inpfile:
            input_map_filename = Path(args.inpfile)
            # input_map_filename = '{input_map_filename}'.format(input_map_filename=input_map_filename.resolve())
        # input_map_filename_specs = open(input_map_filename_specs_name, encoding="utf8")
        config_output_format = 'excel'
        if args.output_format:
            config_output_format = args.output_format

        print('{script_name}: script started at {dt}'.format(dt=time_start,script_name=script_name))

        #print('{script_name}: reading {fname}'.format(fname=input_map_filename,script_name=script_name))
        if not(Path(input_map_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=input_map_filename))
        
        inpfile_map_in_json = None
        with open(input_map_filename, encoding="utf8") as input_map_file:
            try:
                inpfile_map_in_json = json.load(input_map_file)
            except json.JSONDecodeError as e:
                # just a more descriptive message to the end user
                # can happen if the tool is started two times in parallel and it is writing to the same json simultaneously
                raise TypeError('Diff: Can\'t read input file as JSON: {msg}'.format(msg=e))

        config = {}

        if args.flags:
            config['flags'] = '{s}'.format(s=args.flags).split(',')

        result = None
        if config_output_format=='excel':
            result = ReportDocument(inpfile_map_in_json,config)
        else:
            raise ValueError('report.py: unsupported output format: {fmt}'.format(fmt=config_output_format))
        
        result_fname = ( Path(input_map_filename).parents[0] / '{basename}{ext}'.format(basename=re.sub(r'\.json\s*?$','','{n}'.format(n=Path(input_map_filename).name),flags=re.I),ext='.xlsx') if Path(input_map_filename).is_file() else re.sub(r'^\s*?(.*?)(?:\.json)?\s*?$',lambda m: '{base}{added}'.format(base=m[1],added='.xlsx'),'{path}'.format(path=input_map_filename)) )
        print('{script_name}: saving as "{fname}"'.format(fname=result_fname,script_name=script_name))
        # with open(result_fname, "w") as outfile:
        #     outfile.write(result)
        if not not result:
            result.write_to_file(result_fname)
        else:
            raise Exception('Error: inp file was not opened and loaded, something was wrong')

        time_finish = datetime.now()
        print('{script_name}: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start,script_name=script_name))

    except Exception as e:
        # for pretty-printing any issues that happened during runtime; if we hit FileNotFound I don't appreciate when a log traceback is shown, the error should be simple and clear
        # the program is designed to be user-friendly
        # that's why we reformat error messages a little bit
        # stack trace is still printed (I even made it longer to 20 steps!)
        # but the error message itself is separated and printed as the last message again

        # for example, I don't write 'print('File Not Found!');exit(1);', I just write 'raise FileNotFoundErro()'
        print('',file=sys.stderr)
        print('Stack trace:',file=sys.stderr)
        print('',file=sys.stderr)
        traceback.print_exception(e,limit=20)
        print('',file=sys.stderr)
        print('',file=sys.stderr)
        print('',file=sys.stderr)
        print('Error:',file=sys.stderr)
        print('',file=sys.stderr)
        print('{e}'.format(e=e),file=sys.stderr)
        print('',file=sys.stderr)
        exit(1)


if __name__ == '__main__':
    entry_point({'arglist_strict':True})
