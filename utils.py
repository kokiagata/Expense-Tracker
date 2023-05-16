import pandas as pd


def add_new_rows(dataframe, name, value):
    dataframe.loc[len(dataframe.index)] = [name, value]
