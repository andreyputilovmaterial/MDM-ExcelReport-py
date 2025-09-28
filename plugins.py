
import numpy as np, pandas as pd




def plugin__prep_excel_for_translation_util__on_dataframe(df,name,config):
    if name=='overview':
        return df
    # 1. add a column
    # df.set_index('name', inplace=True)
    if df.index.name=='name':
        df.insert(0, 'update', '')
        df.insert(len(df.columns), 'comment', '')
    # 2. if there is an empty "root" row - remove it
    def is_row_blank(row):
        return row.isna().all() or (row.replace('', np.nan).isna().all())
    # Check if the first row is completely blank
    # if df.iloc[0].isna().all():
    # if is_row_blank(df.iloc[0]):
    #     df = df.iloc[1:].reset_index(drop=True)
    def is_row_blank(row):
        for cell in row:
            # Each cell is a list of tuples
            if not cell:  # empty list → blank
                continue
            # If any tuple has non-empty value, row is not blank
            if any(value != '' and not pd.isna(value) for value, _ in cell):
                return False
        return True
    if is_row_blank(df.iloc[0]):
        df = df.iloc[1:]
    return df

def plugin__prep_excel_for_translation_util__on_format_sheet(workbook, worksheet, nrows, ncols, name, config):
    if name=='overview':
        return
    worksheet.set_column(1, 0, 4)
    worksheet.set_column(0, 0, 35)






plugins = [
    {
        'name': 'prep_excel_for_translation_util',
        'active': True,
        'should_run': lambda config: 'mdd_translationoverlays_excel' in (config['flags'] if 'flags' in config else []),
        'on_dataframe': plugin__prep_excel_for_translation_util__on_dataframe,
        'on_format_sheet': plugin__prep_excel_for_translation_util__on_format_sheet,
    },
]



