# import os, time, re, sys
from datetime import datetime
# from dateutil import tz
import argparse
from pathlib import Path
import re
import json

import pandas as pd




if __name__ == '__main__':
    # run as a program
    import util_dataframe_wrapper
    import format_sheet as excel_format_sheet
    import format_sheet_overview as excel_format_sheet_overview
elif '.' in __name__:
    # package
    from . import util_dataframe_wrapper
    from . import format_sheet as excel_format_sheet
    from . import format_sheet_overview as excel_format_sheet_overview
else:
    # included with no parent package
    import util_dataframe_wrapper
    import format_sheet as excel_format_sheet
    import format_sheet_overview as excel_format_sheet_overview






class Map:
    class CellNotFound(Exception):
        """Cell not found"""
    def __init__(self,inp={},config={}):

        self.config = config

        # if "columns" not in self.config or not self.config["columns"]:
        #     self.config["columns"] = ['name']
        
        self.config['columns'] = [ c for c in (inp['report_scheme']['columns'] if ('columns' in inp['report_scheme']) else []) ] if 'report_scheme' in inp else []


        self.update(inp)






    def update(self,inp):
        
        if not inp:
            inp = {}
        config = self.config

        report_data_sections = []

        report_data_sections.append(self.prep_index_section_obj(inp))

        config["section_ids_used"] = []
        for section_obj in ( inp['sections'] if 'sections' in inp else [] ):
            report_data_sections.append(self.prep_section_obj_from_inp(section_obj))

        self.dataframes = [ { 'name': o['name'], 'df': self.df_prep(o) } for o in report_data_sections ]
        
        return None
    
    def prep_index_section_obj(self,inp):
        
        def sanitize_text_extract_filename(s):
            return re.sub(r'^.*[/\\](.*?)\s*?$',lambda m: m[1],'{sstr}'.format(sstr=s))

        data_add = []

        result_ins_htmlmarkup_title = '???'
        result_ins_htmlmarkup_heading = '???'
        result_ins_htmlmarkup_reporttype = inp['report_type'] if 'report_type' in inp else '???'
        result_ins_htmlmarkup_headertext = '{reporttype} Report'.format(reporttype=result_ins_htmlmarkup_reporttype)
        result_ins_htmlmarkup_banner = ''
        if result_ins_htmlmarkup_reporttype=='MDD':
            result_ins_htmlmarkup_title = 'MDD: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']))
            result_ins_htmlmarkup_heading = 'MDD: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']))
            result_ins_htmlmarkup_headertext = '' # it's too obvious, we shouldn't print unnecessary line; it says "MDD" with a very big font size in h1
        elif result_ins_htmlmarkup_reporttype=='diff':
            result_ins_htmlmarkup_title = 'Diff: {MDD_A} vs {MDD_B}'.format(MDD_A=sanitize_text_extract_filename(inp['source_left']),MDD_B=sanitize_text_extract_filename(inp['source_right']))
            result_ins_htmlmarkup_heading = 'Diff'
        else:
            if( result_ins_htmlmarkup_reporttype and (len(result_ins_htmlmarkup_reporttype)>0) and not (result_ins_htmlmarkup_reporttype=='???') ):
                result_ins_htmlmarkup_title = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=result_ins_htmlmarkup_reporttype)
                result_ins_htmlmarkup_heading = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=result_ins_htmlmarkup_reporttype)
            elif len([flag for flag in ( (inp['report_scheme']['flags'] if 'flags' in inp['report_scheme'] else []) if 'report_scheme' in inp else []) if re.match(r'^\s*?data-type\s*?:',flag)])>0:
                flags_indicating_data_type = [flag for flag in ( (inp['report_scheme']['flags'] if 'flags' in inp['report_scheme'] else []) if 'report_scheme' in inp else []) if re.match(r'^\s*?data-type\s*?:',flag)]
                data_type_str = '/'.join([re.sub(r'^\s*?data-type\s*?:\s*?(.*?)\s*?$',lambda m: m[1],flag) for flag in flags_indicating_data_type])
                result_ins_htmlmarkup_title = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=data_type_str)
                result_ins_htmlmarkup_heading = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=data_type_str)
            else:
                result_ins_htmlmarkup_title = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc='File')
                result_ins_htmlmarkup_heading = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc='File')
        result_ins_htmlmarkup_banner = []+[{'name':'datetime','value':inp['report_datetime_utc']}]+inp['source_file_metadata']
        
        data_add.append(['',result_ins_htmlmarkup_heading])
        
        for o in result_ins_htmlmarkup_banner:
            data_add.append([o['name'],o['value']])


        section = {
            'columns': ['name','value'],
            'column_headers': {
                "name": "",
                "value": ""
            },
            'content': None,
            'name': 'overview',
            'title': 'overview',
            'id': 'overview',
            # 'statistics': section_obj['statistics'],
            'data': data_add
        }
        return section

    def prep_section_obj_from_inp(self,section_obj):

        config = self.config

        data_add = []
        result_column_headers = [ '{col}'.format(col=col) for col in section_obj['columns'] ] if 'columns' in section_obj else config["columns"]
        for row in ( section_obj['content'] if section_obj['content']else [] ):
            row_add = []
            for col in result_column_headers:
                row_add.append( row[col] if col in row else '' )
            data_add.append(row_add)
        
        section_title = section_obj['name']
        # if 'title' in section_obj:
        #     section_title = section_obj['title']
        
        section_id = section_obj['name']
        section_id = re.sub(r'([^a-zA-Z])',lambda m: '_x{d}_'.format(d=ord(m[1])),section_id,flags=re.I)
        # while True:
        #     if not (section_id in config["section_ids_used"]):
        #         config["section_ids_used"].append(section_id)
        #         break
        #     else:
        #         section_id = section_id+ '_' + str(config["section_ids_used"].index(section_id)+2)

        section = {
            **section_obj,
            'columns': section_obj['columns'] if 'columns' in section_obj else self.config['columns'],
            'content': None,
            'name': section_obj['name'],
            'title': section_title,
            'id': section_id,
            # 'statistics': section_obj['statistics'],
            'data': data_add
        }
        return section
    

    def df_prep(self,section_data):

        data = util_dataframe_wrapper.PandasDataframeWrapper(section_data['columns'])

        for row in section_data['data']:
            row_add = [item for item in row]
            data.append(*row_add)

        return data.to_df()


    def write_to_file(self,out_filename):

        config = self.config

        df_dataframes = self.dataframes

        
        with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
            for o in self.dataframes:
                o['df'].to_excel(writer, sheet_name=o['name'])
                format_fn = excel_format_sheet.format_sheet if not(o['name']=='overview') else excel_format_sheet_overview.format_sheet
                format_fn(writer.sheets[o['name']])


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

    result = None
    if config_output_format=='excel':
        result = Map(inpfile_map_in_json)
    else:
        raise ValueError('report.py: unsupported output format: {fmt}'.format(fmt=config_output_format))
    
    result_fname = ( Path(input_map_filename).parents[0] / '{basename}{ext}'.format(basename=Path(input_map_filename).name,ext='.xlsx') if Path(input_map_filename).is_file() else re.sub(r'^\s*?(.*?)\s*?$',lambda m: '{base}{added}'.format(base=m[1],added='.xlsx'),'{path}'.format(path=input_map_filename)) )
    print('{script_name}: saving as "{fname}"'.format(fname=result_fname,script_name=script_name))
    # with open(result_fname, "w") as outfile:
    #     outfile.write(result)
    if not not result:
        result.write_to_file(result_fname)
    else:
        raise Exception('Error: inp file was not opened and loaded, something was wrong')

    time_finish = datetime.now()
    print('{script_name}: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start,script_name=script_name))


if __name__ == '__main__':
    entry_point({'arglist_strict':True})
