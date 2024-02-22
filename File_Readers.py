import pandas as pd
from typing import Union, List
from pathlib import Path


def csv_reader(file_path: Union[str, Path], columns: List[str]) -> pd.DataFrame:
    """Imports a .csv file into a pandas dF. Enter the list of columns you wish to import.

    :param file_path: (var) the path of the file you wish to read in. variable or a raw string. note; no backslashes.
    :param columns: (str) the columns you wish to use, separated by commas
    :return: a df with the columns you specified in the input

    Note:
        When selecting your columns you wish to read in, enclose them in a list in the parameter. See below example.

    Example:
        my_path = "C:/Users/Dan/Downloads/test.csv"

        df = csv_reader(my_path, ['col1','col2']
    """
    return pd.read_csv(file_path, encoding='latin1', usecols=columns)


def txt_reader(file_path: Union[str, Path], columns: List[str]) -> pd.DataFrame:
    """Imports a .txt file into a pandas dF. Enter the list of columns you wish to import.

    :param file_path: (var) the path of the file. can be a variable or a raw string. no backslashes. only fwd slashes
    :param columns: (str) the columns you wish to use, separated by commas, do not enclose in a list
    :return: a df with the columns you specified in the input

    Note:
        When selecting your columns you wish to read in, enclose them in a list in the parameter. See below example.

    Example:
        my_path = "C:/Users/Dan/Downloads/test.txt"

        df = txt_reader(my_path, ['col1','col2']
    """
    return pd.read_csv(file_path, sep='\t', encoding='latin1', usecols=columns)


def concat_df(x: List[pd.DataFrame]) -> pd.DataFrame:
    """Converts a list of dF's - that you imported and appended to a var - into one dF.

    :param x: The list of dataframes you wish to concat into one.
    :return: the concat'd dataframe, with refreshed index

    Example:
        df = df1.append(df2)  # creates a list of two dF's
        df = concat_df(df) # merges the two dF's into one dF
    """
    return pd.concat(x, ignore_index=True)

