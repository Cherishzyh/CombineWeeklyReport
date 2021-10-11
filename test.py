import pandas as pd
import datetime


def state(statedate):
    # 周日期对应
    # 往前取365天
    df = pd.DataFrame([str(d)[:10] for d in pd.date_range(end=statedate, periods=365)],columns=['date'])
    # 生成周
    df['week'] = df['date'].apply(lambda x: int(datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%W")))
    df['week'] = df['week'].apply(lambda x: '%s%s' % ('week', str(x)))
    # 取年月
    df['new_date'] = df['date'].apply(lambda x: x[-5:])
    # 排序
    df = df.sort_values(by=['new_date'], ascending=[True])
    # 周开始日期
    df_first = df.drop_duplicates(subset=['week'], keep='first')
    df_first = df_first.rename(columns={'new_date': 'start_date'})
    # 周结束日期
    df_last = df.drop_duplicates(subset=['week'], keep='last')
    df_last = df_last.rename(columns={'new_date': 'end_date'})
    df_date = pd.merge(df_first[['start_date', 'week']], df_last[['end_date', 'week']], on='week')
    tt = datetime.datetime.strptime(statedate, '%Y-%m-%d')
    this_year = tt.year
    df_date['start_date'] = df_date['start_date'].apply(lambda x: '%s-%s' % (str(this_year), str(x)))
    df_date['end_date'] = df_date['end_date'].apply(lambda x: '%s-%s' % (str(this_year), str(x)))
    return df_date


def MyState(statedate):
    df = pd.DataFrame([str(d)[:10] for d in pd.date_range(end=statedate, periods=365)],columns=['date'])

    df['week'] = df['date'].apply(lambda x: int(datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%W")))
    df['week'] = df['week'].apply(lambda x: '%s%s' % ('WK', str(x)))

    df_first = df.drop_duplicates(subset=['week'], keep='first')
    df_first = df_first.rename(columns={'date': 'start_date'})

    df_last = df.drop_duplicates(subset=['week'], keep='last')
    df_last = df_last.rename(columns={'date': 'end_date'})

    df_date = pd.merge(df_first[['start_date', 'week']], df_last[['end_date', 'week']], on='week')
    return df_date


wc_df = state('2020-12-31')
my_df = MyState('2020-12-31')
time_df = pd.Timestamp('2012-05-01')

