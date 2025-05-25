import pandas as pd

def cal_ema(df, column="Close", span = 20):

    return df[column].ewm(span=span, adjust=False).mean()