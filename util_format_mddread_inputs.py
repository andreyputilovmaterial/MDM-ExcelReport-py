
from pathlib import Path # to format file name on overview sheet
import re



if __name__ == '__main__':
    # run as a program
    import util_dataframe_wrapper
elif '.' in __name__:
    # package
    from . import util_dataframe_wrapper
else:
    # included with no parent package
    import util_dataframe_wrapper




def isnonempty(o):
    if o==0:
        return True
    elif o==False:
        return True
    else:
        return not not o

def sanitize(o):
    if not isnonempty(o):
        yield ('','')
    elif isinstance(o,list):
        for part in o:
            yield from sanitize(part)
    elif isinstance(o,dict):
        def modify_as_per_flags(item,flags):
            result = item
            for flag in flags:
                if (flag=='role-time') or (flag=='role-date') or (flag=='role-datetime'):
                    result = result
                elif flag=='role-added':
                    if result[1]=='removed':
                        result = (result[0],'changed')
                    else:
                        result = (result[0],'added')
                elif flag=='role-removed':
                    if result[1]=='added':
                        result = (result[0],'changed')
                    else:
                        result = (result[0],'removed')
                elif flag=='role-sronly':
                    result = result
                elif flag=='role-label':
                    result = result
            return result
        flags = []
        if 'flags' in o:
            flags = o['flags']
        if 'role' in o:
            flags = [*flags,*['role-{f}'.format(f=f) for f in [o['role']]]]
        # yield from sanitize(o['parts'])
        if 'parts' in o:
            nested = sanitize(o['parts'])
        elif 'name' in o and 'value' in o:
            nested = ['Name: "',o['name'],'", Value: "',o['value'],'", ']
            nested = [p for p in sanitize(nested)]
        elif 'text' in o:
            nested = sanitize(o['text'])
        else:
            nested = sanitize('')
        for n in nested:
            yield modify_as_per_flags(n,flags)
    elif isinstance(o,str):
        o = '{o}'.format(o=o)
        # o = sanitize_text_with_markers(...)
        yield(o,'')
    else:
        yield (o,'')














def prep_dataframes(inp,config):
    def sanitize_wrapped(s):
        # return [(s,'')]
        result = sanitize(s)
        return [r for r in result] # we need an instance of list, not generator

    def df_prep(section_data):
        data = util_dataframe_wrapper.PandasDataframeWrapper(section_data['columns'])
        for row in section_data['data']:
            row_formatted = [sanitize_wrapped(r) for r in row]
            data.append(*row_formatted)
        return data.to_df()
    
    report_data_sections = []

    # add overview df
    report_data_sections.append(prep_overview_section(inp,config))

    # add remaining dataframes
    # config["section_ids_used"] = []
    for section_obj in ( inp['sections'] if 'sections' in inp else [] ):
        report_data_sections.append(prep_datasection_from_mddread_section(section_obj,config))

    return [ { 'name': o['name'], 'df': df_prep(o) } for o in report_data_sections ]





def prep_overview_section(inp,config):
    
    def sanitize_text_extract_filename(s):
        # return re.sub(r'^.*[/\\](.*?)\s*?$',lambda m: m[1],'{sstr}'.format(sstr=s))
        return Path(s).name

    data_add = []

    result_ins_title = '???'
    result_ins_heading = '???'
    result_ins_reporttype = inp['report_type'] if 'report_type' in inp else '???'
    result_ins_headertext = '{reporttype} Report'.format(reporttype=result_ins_reporttype)
    result_ins_banner = ''
    if result_ins_reporttype=='MDD':
        result_ins_title = 'MDD: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']))
        result_ins_heading = 'MDD: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']))
        result_ins_headertext = '' # it's too obvious, we shouldn't print unnecessary line; it says "MDD" with a very big font size in h1
    elif result_ins_reporttype=='diff':
        result_ins_title = 'Diff: {MDD_A} vs {MDD_B}'.format(MDD_A=sanitize_text_extract_filename(inp['source_left']),MDD_B=sanitize_text_extract_filename(inp['source_right']))
        result_ins_heading = 'Diff'
    else:
        if( result_ins_reporttype and (len(result_ins_reporttype)>0) and not (result_ins_reporttype=='???') ):
            result_ins_title = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=result_ins_reporttype)
            result_ins_heading = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=result_ins_reporttype)
        elif len([flag for flag in ( (inp['report_scheme']['flags'] if 'flags' in inp['report_scheme'] else []) if 'report_scheme' in inp else []) if re.match(r'^\s*?data-type\s*?:',flag)])>0:
            flags_indicating_data_type = [flag for flag in ( (inp['report_scheme']['flags'] if 'flags' in inp['report_scheme'] else []) if 'report_scheme' in inp else []) if re.match(r'^\s*?data-type\s*?:',flag)]
            data_type_str = '/'.join([re.sub(r'^\s*?data-type\s*?:\s*?(.*?)\s*?$',lambda m: m[1],flag) for flag in flags_indicating_data_type])
            result_ins_title = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=data_type_str)
            result_ins_heading = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc=data_type_str)
        else:
            result_ins_title = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc='File')
            result_ins_heading = '{report_desc}: {filepath}'.format(filepath=sanitize_text_extract_filename(inp['source_file']),report_desc='File')
    result_ins_banner = []+[{'name':'datetime','value':inp['report_datetime_utc']}]+inp['source_file_metadata']
    
    data_add.append(['',result_ins_heading])
    
    for o in result_ins_banner:
        data_add.append([o['name'],o['value']])


    section_obj = {
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
    return section_obj



def prep_datasection_from_mddread_section(section_obj_from_json,config):

    data_add = []
    result_column_headers = [ '{col}'.format(col=col) for col in section_obj_from_json['columns'] ] if 'columns' in section_obj_from_json else config["columns"]
    for row in ( section_obj_from_json['content'] if section_obj_from_json['content']else [] ):
        row_add = []
        for col in result_column_headers:
            row_add.append( row[col] if col in row else '' )
        data_add.append(row_add)
    
    section_title = section_obj_from_json['name']
    # if 'title' in section_obj_from_json:
    #     section_title = section_obj_from_json['title']
    
    section_id = section_obj_from_json['name']
    section_id = re.sub(r'([^a-zA-Z])',lambda m: '_x{d}_'.format(d=ord(m[1])),section_id,flags=re.I)
    # while True:
    #     if not (section_id in config["section_ids_used"]):
    #         config["section_ids_used"].append(section_id)
    #         break
    #     else:
    #         section_id = section_id+ '_' + str(config["section_ids_used"].index(section_id)+2)

    section_obj = {
        **section_obj_from_json,
        'columns': section_obj_from_json['columns'] if 'columns' in section_obj_from_json else config['columns'],
        'content': None,
        'name': section_obj_from_json['name'],
        'title': section_title,
        'id': section_id,
        # 'statistics': section_obj_from_json['statistics'],
        'data': data_add
    }
    return section_obj




