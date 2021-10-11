import os
import numpy as np
import pandas as pd
import datetime
from openpyxl import load_workbook

'''
1. load weekly report 
    √(the best way: get sheet name; load df according to sheet name (decrease=True) until the columns include '日期')**

2. load final report
    input the week number
    if there is week number:     
        search the rows need to add
        combine these rows to a new sheet
    else:
        search the rows need to add
        combine these rows to a new excel in a new sheet

'''


class CombineWeeklyReport:
    def __init__(self, year=int(2021)):
        super(CombineWeeklyReport, self).__init__()
        self.year = year
        self._WeeklyCalendar()

    def LoadWeeklyReport(self, weekly_path_list, week):
        df_list = []
        self.week_num_list = []
        for weekly_report in weekly_path_list:
            try:
                excel = pd.ExcelFile(weekly_report)
            except Exception:
                print('Can not Load {}, Please Check file'.format(weekly_report))
                break
            sheet_list = excel.sheet_names

            for sheet in reversed(sheet_list):
                weekly_report_df = pd.read_excel(weekly_report, sheet_name=sheet)
                columns_list = weekly_report_df.columns.tolist()
                is_retain = [columns for columns in columns_list if '日期' in str(columns)]
                if len(is_retain) == 1:
                    try:
                        weekly_report_df.rename(columns={is_retain[0]: '日期'}, inplace=True)
                        weekly_report_df.rename(columns={columns_list[0]: 'SA #'}, inplace=True)
                        weekly_report_df['日期'] = weekly_report_df['日期'].apply(
                            lambda x: np.NaN if str(x).isspace() else x)
                        weekly_report_df['日期'] = weekly_report_df['日期'].apply(
                            lambda x: np.NaN if isinstance(x, str) else x)
                        weekly_report_df = weekly_report_df[weekly_report_df['日期'].notnull()]

                        year_df = weekly_report_df['日期'].apply(lambda x: pd.to_datetime(x).year)
                        weekly_report_df = weekly_report_df.loc[
                            [index for index in year_df.index if year_df[index] == self.year]]

                        weekly_report_df['日期'] = weekly_report_df['日期'].apply(
                            lambda x: datetime.datetime.strptime(str(x)[:10], "%Y-%m-%d").strftime('%d-%b-%y'))
                        wk_num = weekly_report_df['日期'].apply(
                            lambda x: int(datetime.datetime.strptime(str(x)[:9], "%d-%b-%y").strftime("%W")))
                        retain_index = [index for index in wk_num.index if wk_num[index] == week]
                        if week == None:  # 如果没有输入指定的周次，则load所有的含有日期的数据
                            self.week_num_list.extend(wk_num.tolist())
                            df_list.append(weekly_report_df.fillna(value=''))
                        elif len(retain_index) == 0 and len(sheet_list) == 1:  # 该sheet无相应周次数据但是只有一个sheet，load该sheet
                            df_list.append(weekly_report_df.fillna(value=''))
                            continue  # 该sheet无相应周次数据，则继续load其余sheet，直到有数据
                        elif len(retain_index) == 0 and len(sheet_list) > 1:
                            continue
                        else:
                            df_list.append(weekly_report_df.fillna(value='').loc[retain_index])
                            break  # 如果有输入指定的周次，且该sheet有相应数据，则停止加载其余sheet
                    except Exception as e:
                        print(weekly_report)
                        print(e)
                        continue
        return df_list

    def _WeeklyCalendar(self):
        if self.year % 400 == 0 or self.year % 4 == 0:
            periods = 366
        else:
            periods = 365
        year_end = '{}-12-31'.format(self.year)
        df = pd.DataFrame([str(d)[:10] for d in pd.date_range(end=year_end, periods=periods)], columns=['date'])

        df['week'] = df['date'].apply(lambda x: int(datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%W")))

        df_first = df.drop_duplicates(subset=['week'], keep='first')
        df_first = df_first.rename(columns={'date': 'start_date'})

        df_last = df.drop_duplicates(subset=['week'], keep='last')
        df_last = df_last.rename(columns={'date': 'end_date'})

        self.weekly_calendar = pd.merge(df_first[['start_date', 'week']], df_last[['end_date', 'week']], on='week')
        return self.weekly_calendar

    def _SheetName(self, wk):
        wk = int(wk)
        month_abbr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        start_d = self.weekly_calendar.loc[wk]['start_date'][-2:]
        start_m = int(self.weekly_calendar.loc[wk]['start_date'][-5:-3])
        end_d = self.weekly_calendar.loc[wk]['end_date'][-2:]
        end_m = int(self.weekly_calendar.loc[wk]['end_date'][-5:-3])
        if start_m == end_m:
            sheet_name = 'WK{} {}-{} {}'.format(wk, start_d, end_d, month_abbr[start_m - 1])
        else:
            sheet_name = 'WK{} {} {}-{} {}'.format(wk, start_d, month_abbr[start_m - 1], end_d, month_abbr[end_m - 1])
        return sheet_name

    def _ExcelWriter(self, total_df, target_path):
        if not os.path.exists(target_path):
            df = pd.DataFrame()
            df.to_excel(target_path, index=False)

        book = load_workbook(target_path)
        writer = pd.ExcelWriter(target_path, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        if isinstance(total_df, dict):
            for key in total_df.keys():
                total_df[key].to_excel(writer, sheet_name=key, index=False)
            writer.close()
        else:
            total_df.to_excel(writer, sheet_name=self.sheet_name, index=False)
            writer.close()

    def Run(self, weekly_path_folder, target_path, week=None):
        weekly_path_list = [os.path.join(weekly_path_folder, path) for path in os.listdir(weekly_path_folder)]
        weekly_path_list = [path for path in weekly_path_list if path.endswith('.xlsx')]

        df_list = self.LoadWeeklyReport(weekly_path_list, week)
        if week:
            self.sheet_name = self._SheetName(week)
            new_df = pd.DataFrame()
            for df in df_list:
                new_df = new_df.append(df, ignore_index=True)
            new_df.fillna(value='')
            self._ExcelWriter(new_df, target_path)
        else:
            self.week_num_list = list(set(self.week_num_list))
            self.sheet_name = [self._SheetName(value) for value in self.week_num_list]
            df_dict = {}
            df_dict = df_dict.fromkeys(self.sheet_name, pd.DataFrame())
            for df in df_list:
                df.drop_duplicates()
                for index in df.index:
                    wk_df = int(datetime.datetime.strptime(str(df.loc[index, '日期'])[:9], '%d-%b-%y').strftime("%W"))
                    wk_index = self.week_num_list.index(wk_df)
                    df_dict[self.sheet_name[wk_index]] = df_dict[self.sheet_name[wk_index]].append(df.loc[index])
            self._ExcelWriter(df_dict, target_path)
        return df_list


if __name__ == "__main__":
    weekly_path_folder = r'C:\Users\82375\Desktop\weekly_report'
    # target_path = r'C:\Users\82375\Desktop\target.xlsx'
    target_path = r'C:\Users\82375\Desktop\weekly_report\target.xlsx'
    # test = pd.read_excel(target_path, sheet_name=None)

    WR = CombineWeeklyReport(year=2021)
    WR.Run(weekly_path_folder, target_path=target_path, week=33)

