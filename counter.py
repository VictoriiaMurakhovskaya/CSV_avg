import pandas as pd
from datetime import time
import xlsxwriter


def convert_time(string_time):
    values = string_time.split(':')
    hour = int(values[0]) if len(values) == 3 else 0
    mins = int(values[-2])
    sec = int(values[-1])
    return time(hour, mins, sec)


def change_time(cur_time, delta):
    sec = cur_time.hour*24*60 + cur_time.minute*60 + cur_time.second
    sec = sec + delta
    return time(sec // 3600, (sec % 3600) // 60, sec % 60)


def calculate(intervals, csvfile, outputfile):
    df = pd.read_csv(csvfile)
    for index, row in df.iterrows():
        df.at[index, 'time_formatted'] = convert_time(row['time'])
    df = df.drop(columns=['time'])

    res = {}
    for i in range(0, len(intervals) - 1):
        y = df.loc[(df['time_formatted'] >= intervals[i]) & (df['time_formatted'] < intervals[i+1])]
        to_drop = [item for item in y if item not in ['neutral', 'happy', 'sad', 'surprise', 'angry', 'fear', 'disgust']]
        y = y.drop(columns=to_drop)
        res.update({(intervals[i], change_time(intervals[i+1], -1)): y.mean(axis=0)})
    # pd.DataFrame(res).to_excel(outputfile)
    return pd.DataFrame(res)