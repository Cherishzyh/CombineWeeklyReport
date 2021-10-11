import os
import numpy as np
import pandas as pd
import difflib
from pypinyin import lazy_pinyin


class CombineWeeklyReportByName:
    def __init__(self,):
        super(CombineWeeklyReportByName, self).__init__()

    def LoadWeeklyReport(self, weekly_report):
        try:
            excel = pd.ExcelFile(weekly_report)
            sheet_name = excel.sheet_names
            sheet_list = [sheet for sheet in sheet_name if 'Sheet' not in sheet]
            total_df = pd.DataFrame()
            new_columns = ['SA #', '科学家', '日期', '医院', '科室主任/医生', '沟通方式\n现场客户；远程客户；内部工作）',
                           '售前/售后/其他', 'Professional Service\n(Yes/No）', '工作内容']
            for sheet in sorted(sheet_list):
                weekly_report_df = pd.read_excel(weekly_report, sheet_name=sheet)
                columns_list = weekly_report_df.columns.tolist()
                is_retain = [columns for columns in columns_list if '科学家' in str(columns)]
                if len(is_retain) != 1:
                    print('科学家 is not in sheet {}'.format(sheet))
                    continue
                else:
                    weekly_report_df.rename(columns=dict(zip(columns_list[:9], new_columns)), inplace=True)
                    weekly_report_df = weekly_report_df[weekly_report_df['科学家'].notnull()]
                    total_df = total_df.append(weekly_report_df)
            return total_df
        except Exception as e:
            print('Can not Load {}, Please Check file, \n The error info: {}'.format(weekly_report, e))
            return None

    def _DropRepeatName(self, scientists_name):

        def __GetEqualRate(str1, str2):
            return difflib.SequenceMatcher(None, str1, str2).quick_ratio()

        def __GetShortName(name_list):
            len_array = np.array([len(name) for name in name_list])
            index = np.argmin(len_array)
            return name_list[index]

        new_name_list = []
        new_all_name_list = []
        scientists_name_copy = [''.join(lazy_pinyin(name)).lower() for name in scientists_name]
        for name in scientists_name:
            name = ''.join(lazy_pinyin(name)).lower()
            similarity_1 = [__GetEqualRate(name, name_copy) for name_copy in scientists_name_copy]
            similarity_2 = [__GetEqualRate(name, name_copy) for name_copy in scientists_name]
            index_list = [index for index in range(len(similarity_1)) if similarity_1[index] >= 0.9 or similarity_2[index] >= 0.9]
            name_list = [scientists_name[index] for index in index_list]
            if name_list not in new_all_name_list:
                new_all_name_list.append(name_list)
                new_name_list.append(__GetShortName(name_list))
        return new_name_list, new_all_name_list

    def Run(self, weekly_path_folder, new_excel_path):
        total_df = self.LoadWeeklyReport(weekly_path_folder)
        rename_columns = lambda column: 'Supplement' if 'Unnamed' in column else column
        columns_list = [rename_columns(column) for column in total_df.columns.tolist()]
        total_df.rename(columns=dict(zip(total_df.columns.tolist(), columns_list)), inplace=True)
        scientists_name = list(set(total_df['科学家'].tolist()))
        sheet_name, sheet_all_name = self._DropRepeatName(scientists_name)
        writer = pd.ExcelWriter(new_excel_path, engine='openpyxl')
        total_index = 0
        for index, name in enumerate(sheet_all_name):
            sheet_df = pd.concat([total_df.loc[total_df.loc[:, '科学家'] == n, :] for n in name], axis=0)
            print(name[0], len(sheet_df.index.tolist()))
            total_index += len(sheet_df.index.tolist())
            sheet_df['日期'] = sheet_df['日期'].apply(lambda x: str(x)[:10])
            sheet_df.replace('nan', '', inplace=True)
            sheet_df.to_excel(writer, sheet_name=sheet_name[index], index=False)
        print('Load row: {}\t Save row: {}'.format(len(total_df.index.tolist()), total_index))
        writer.close()


if __name__ == "__main__":
    weekly_path_folder = r'C:\Users\82375\Desktop\Weekly+report_Scientific Solution Team 2021.xlsx'
    new_excel_path = r'C:\Users\82375\Desktop\WeeklyReportByName.xlsx'

    WR = CombineWeeklyReportByName()
    WR.Run(weekly_path_folder, new_excel_path)


