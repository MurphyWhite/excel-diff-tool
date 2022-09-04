import argparse  # Module to integrate Python code with command-line interfaces
import json

import pandas as pd
import numpy as np


def report_diff(x):
    """Helper function to use with groupby.apply to highlight changes in cell values"""
    return x[0] if x[0] == x[1] or pd.isna(x).all() else f'{x[0]} ---> {x[1]}'


def strip(x):
    """Helper function to use with applymap to strip all whitespaces from the dataframe"""
    return x.strip() if isinstance(x, str) else x


def diff_pd(old_df, new_df, idx_col):
    """Identify differences between two pandas DataFrames using a key column containing unique row identifers
    Args:
        old_df (pd.DataFrame): first dataframe
        new_df (pd.DataFrame): second dataframe
        idx_col (str): column name of the index, needs to be present in both DataFrames
    Note: Key column is assumed to have a unique row identifier, i.e. no duplicates
    """
    # setting the column name as index for fast operations
    old_df = old_df.set_index(idx_col)
    new_df = new_df.set_index(idx_col)
    # get the added and removed rows
    old_keys = old_df.index
    new_keys = new_df.index
    removed_keys = np.setdiff1d(old_keys, new_keys)
    added_keys = np.setdiff1d(new_keys, old_keys)
    out_data = {
        'removed': old_df.loc[removed_keys],
        'added': new_df.loc[added_keys]
    }
    # focusing on common data of both dataframes
    common_keys = np.intersect1d(old_keys, new_keys, assume_unique=True)
    common_columns = np.intersect1d(old_df.columns, new_df.columns, assume_unique=True)
    new_common = new_df.loc[common_keys, common_columns].applymap(strip)
    old_common = old_df.loc[common_keys, common_columns].applymap(strip)
    # get the changed rows keys by dropping identical rows
    # (indexes are ignored, so we'll reset them)
    common_data = pd.concat([old_common.reset_index(), new_common.reset_index()])
    changed_keys = common_data.drop_duplicates(keep=False)[idx_col].unique()
    # combining the changed rows via multi level columns
    df_all_changes = pd.concat([old_common.loc[changed_keys], new_common.loc[changed_keys]],
                               axis='columns', keys=['old', 'new'])
    df_all_changes = df_all_changes.swaplevel(axis='columns')[new_common.columns]
    # using report_diff to merge the changes in a single cell with "-->"
    df_changed = df_all_changes.groupby(level=0, axis=1).apply(
        lambda frame: frame.apply(report_diff, axis=1))
    out_data['changed'] = df_changed

    return out_data


def compare_excel(path1, path2, out_path, sheet_name, index_col_name=None, header_map=None, **kwargs):
    old_df = pd.read_excel(path1, sheet_name=sheet_name, **kwargs)
    new_df = pd.read_excel(path2, sheet_name=sheet_name, **kwargs)
    if header_map:
        headers = get_header_map(header_map)
        replace_excel_header(old_df, headers)
    diff = diff_pd(old_df, new_df, index_col_name)
    with pd.ExcelWriter(out_path) as writer:
        for sname, data in diff.items():
            data.to_excel(writer, sheet_name=sname)
    print(f"Differences saved in {out_path}")


def get_header_map(headers_path):
    with open(headers_path, 'r', encoding="utf-8") as f:
        load_dict = json.load(f)
        return load_dict


def replace_excel_header(dataframe, headers_map):
    """
    replace the src excels headers if they have different header name
    :param dataframe:
    :param headers_map:
    :return:
    """
    dataframe.rename(columns=headers_map, inplace=True)


# if __name__ == '__main__':
#     test()

if __name__ == '__main__':
    cfg = argparse.ArgumentParser(
        description="Compares two Excel sheets and outputs the differences to a separate Excel file. "
                    "A column name can be specified as the unique row identifier."
    )

    cfg.add_argument("path1", help="Fist Excel file")
    cfg.add_argument("path2", help="Second Excel file")
    cfg.add_argument("sheetname", help="Name of the sheet to compare.")
    cfg.add_argument("--header-map", help="Path of header map.")
    cfg.add_argument("-c", "--index-column", help="Name of the column with unique row identifier",
                     required=True)
    cfg.add_argument("-o", "--output-path", help="Path of the comparison results",
                     default="compared.xlsx")
    cfg.add_argument("--skiprows", help='Excel row containing the table headers', type=int, action='append',
                     default=None)
    opt = cfg.parse_args()

    compare_excel(opt.path1, opt.path2, opt.output_path, opt.sheetname, opt.index_column,
                  header_map=opt.header_map, skiprows=opt.skiprows)
