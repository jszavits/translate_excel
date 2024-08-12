import sys
import pandas as pd
import numpy as np
import translators as ts


def translate_df(df_old, lang_src, lang_dest, engine='google', delay=1):
    """
    This function takes in a dataframe in language determined by
    <lang_src> parameter and translates it into a new dataframe 
    in language determined by <lang_dest> parameter using engine
    <engine> (default is 'google'). For a full list of engines visit 
    https://pypi.org/project/translators
    
    Translation uses Google translate service which may return a 
    HTTP error 429 (too many requests). This can be avoided using 
    parameter <delay> which will set up a delay of <delay> seconds 
    between each HTTP request (default is 1 second).
    """

    nrows, ncols = df_old.shape
    df_new = df_old.copy()

    for i in np.arange(nrows):
        for j in np.arange(ncols):
            old_text = str(df_old.iloc[i,j]).strip()
            if old_text.lower() != 'nan':
                new_text = ts.translate_text(
                    query_text=old_text, 
                    translator=engine,
                    from_language=lang_src,
                    to_language=lang_dest,
                    sleep_seconds=delay
                )
                df_new.iloc[i,j] = new_text

    return df_new

# input file
input1 = sys.argv[1].strip()

# source language
lang_src = sys.argv[2]

# destination language
lang_dest = sys.argv[3]

if input1[-5:] == '.xlsx':

    fn_old = input1[:-5]

    # translate to new filename and remove all leading and trailing whitespace 
    fn_new = ts.translate_text(
        query_text=fn_old, 
        translator='google', 
        from_language=lang_src, 
        to_language=lang_dest,
    ).strip()

    print(f'[INFO] Translated filename {fn_old} to {fn_new}.')

    # get all sheet names in a list
    xl = pd.ExcelFile(f"{fn_old}.xlsx")
    sheets = xl.sheet_names
    n=len(sheets)

    print(f'[INFO] Found {n} sheet(s) in the original excel file.')

    # export translated sheets to a new excel file
    with pd.ExcelWriter(f"{fn_new}.xlsx") as writer:
   
        # loop over all sheets in the excel file
        i=0
        for sheet in sheets:
            i+=1
            print(f"[INFO] Processing sheet {sheet} [{i}/{n}]")
            df_old = pd.read_excel(f"{fn_old}.xlsx", sheet, index_col=None, header=None)
            df_new = translate_df(df_old, lang_src, lang_dest)
            df_new.to_excel(writer, sheet_name=sheet, index=False, header=False)

else:
    raise('Input file must have .xlsx extension.')
