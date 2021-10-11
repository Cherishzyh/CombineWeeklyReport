import numpy as np
import pandas as pd
import datetime


class CombineWeeklyReport():
    def __init__(self):
        super(CombineWeeklyReport, self).__init__()

    def LoadWeeklyReport(self, weekly_path_list):
        df_list = []
        for weekly_report in weekly_path_list:
            weekly_report_df = pd.read_excel(weekly_report, sheet_name=None)
            if isinstance(weekly_report_df, dict):
                df_list.extend([weekly_report_df[key] for key in weekly_report_df.keys()])
            else:
                df_list.extend(weekly_report_df)
        return df_list

    def _Delete(self, df_list):
        date_df = []
        for index, weekly_report_df in enumerate(df_list):
            column_name = [d for d in weekly_report_df.columns.tolist() if '日期' in d]  # 排除如果列名“日期”中含有空格或者别的特殊符号找不到该列的情况
            if len(column_name) == 1:
                weekly_report_df.rename(columns={column_name[0]: '日期'}, inplace=True)
                # 删除没有写日期的行
                weekly_report_df['日期'] = weekly_report_df['日期'].apply(lambda x: np.NaN if str(x).isspace() else x)
                weekly_report_df = weekly_report_df[weekly_report_df['日期'].notnull()]
                weekly_report_df['日期'] = weekly_report_df['日期'].apply(lambda x: datetime.datetime.strptime(str(x)[:10], "%Y-%m-%d").strftime('%d-%b-%Y'))
            date_df.append(weekly_report_df.fillna(value=' '))
        return date_df

    def Date2Week(self, df_list):
        self.week_num_list = []
        date_df = self._Delete(df_list)
        for weekly_report_df in date_df:
            try:
                week_df = weekly_report_df['日期'].apply(lambda x: int(datetime.datetime.strptime(str(x)[:11], '%d-%b-%Y').strftime("%W")))
                self.week_num_list.extend(week_df.tolist())
            except Exception as e:
                pass
        self.week_num_list = list(set(self.week_num_list))
        return date_df

    def _WeeklyCalendar(self, year):
        df = pd.DataFrame([str(d)[:10] for d in pd.date_range(end=year, periods=365)], columns=['date'])

        df['week'] = df['date'].apply(lambda x: int(datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%W")))
        # df['week'] = df['week'].apply(lambda x: '%s%s' % ('WK', str(x)))

        df_first = df.drop_duplicates(subset=['week'], keep='first')
        df_first = df_first.rename(columns={'date': 'start_date'})

        df_last = df.drop_duplicates(subset=['week'], keep='last')
        df_last = df_last.rename(columns={'date': 'end_date'})

        self.weekly_calendar = pd.merge(df_first[['start_date', 'week']], df_last[['end_date', 'week']], on='week')
        return self.weekly_calendar

    def _SheetName(self):
        sheet_name = []
        month_abbr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        for wk in self.week_num_list:
            start_date = self.weekly_calendar.loc[wk]['start_date'][-2:]
            start_month = int(self.weekly_calendar.loc[wk]['start_date'][-5:-3])
            end_date = self.weekly_calendar.loc[wk]['end_date'][-2:]
            end_month = int(self.weekly_calendar.loc[wk]['end_date'][-5:-3])
            if start_month == end_month:
                sheet_name.append('WK{} {}-{} {}'.format(wk, start_date, end_date, month_abbr[start_month-1]))
            else:
                sheet_name.append('WK{} {} {}-{} {}'.format(wk, start_date, month_abbr[start_month-1], end_date, month_abbr[end_month-1]))
        return sheet_name

    def _ExcelWriter(self, df_dict, target_path):
        writer = pd.ExcelWriter(target_path)
        for key in df_dict.keys():
            df_dict[key].to_excel(writer, sheet_name=key, index=False)
        writer.close()

    def Run(self, weekly_path_list, target_path):
        self._WeeklyCalendar('2021-12-31')
        df_list = self.LoadWeeklyReport(weekly_path_list)
        date_df = self.Date2Week(df_list)

        df_dict = {}
        df_dict = df_dict.fromkeys(self._SheetName(), pd.DataFrame(columns=date_df[0].columns.tolist()))
        other_sheet = 0
        for df in date_df:
            if '日期' in df.columns.tolist():
                if len(df.columns.tolist()) == len(date_df[0].columns.tolist()):
                    df.columns = date_df[0].columns.tolist()
                else:
                    df.columns = date_df[0].columns.tolist()[:-1]
                for index in df.index:
                    wk = int(datetime.datetime.strptime(str(df.loc[index, '日期'])[:11], '%d-%b-%Y').strftime("%W"))
                    wk_index = self.week_num_list.index(wk)
                    df_dict[self._SheetName()[wk_index]] = df_dict[self._SheetName()[wk_index]].append(df.loc[index])
            else:
                df_dict['sheet{}'.format(str(other_sheet))] = df
                other_sheet += 1
        self._ExcelWriter(df_dict, target_path)
        return df_dict


if __name__ == "__main__":
    com1_path = r'C:\Users\82375\Desktop\combination1.xlsx'
    com2_path = r'C:\Users\82375\Desktop\combination2.xlsx'
    com3_path = r'C:\Users\82375\Desktop\combination3.xlsx'
    weekly_path_list = [com1_path, com2_path, com3_path]
    target_path = r'C:\Users\82375\Desktop\my_target.xlsx'

    # sheet_name = pd.ExcelFile(com1_path).sheet_names
    # target = pd.read_excel(com1_path, sheet_name=sheet_name[0])
    # column_list = target.columns.tolist()
    # column_name = [d for d in column_list if '日期' in d]
    # assert len(column_name) == 1
    # date_list = target[column_name[0]]
    # date_list = [time.strptime(str(date_list[index]), '%Y-%m-%d %H:%M:%S') for index in range(len(date_list))]
    # # a = time.strptime(str(date_list.values.tolist()), '%Y-%m-%d %H:%M:%S')
    # weeknum = [datetime.datetime(int(d.tm_year), int(d.tm_mon), int(d.tm_mday)).isocalendar()[1] for d in date_list]
    #
    # print()
    WR = CombineWeeklyReport()
    WR.Run(weekly_path_list, target_path)
