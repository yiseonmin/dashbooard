import pandas as pd
import xlwings as xw
import openpyxl
import re
# from excel_config import *

class ExcelManager:
    def __init__(self, filename, project):
        self.filename = filename    # 엑셀 파일명
        self.project = project
        self.book = xw.Book(filename)

        # 1) Result sheet DataFrame
        self.sheet_result = self.book.sheets(1)
        print('\n\n>> 1.sheet result: ', self.sheet_result)
        # excel_range = 'A2:E10'  # % (3 + len(PROJECT_INFO[project]))
        self.df_summary = self.sheet_result.used_range.options(pd.DataFrame, index=False).value
        # print('\tEXCEL RANGE:', excel_range)
        print(self.df_summary)

        # 2) project detail data sheet DataFrame
        self.sheet_detail = self.book.sheets(2)
        print('\n>> 2.sheet detail: ', self.sheet_detail)

        # column: A=1, B=2, C=3 ..., R=18
        if self.project in manual_condition:
            last_col = 8 + len(PROJECT_INFO[project]) * 3 + 2   # result/logs/times -> 3 / +2: total_time, manual condition
        else:
            last_col = 8 + len(PROJECT_INFO[project]) * 3 + 1   # result/logs/times -> 3 / +1: total time
        last_col_letter = openpyxl.utils.get_column_letter(last_col)
        las_row = self.sheet_detail.range('B2').end('down').row
        excel_range = 'B2:%s%d' % (last_col_letter, las_row)

        self.df_detail = self.sheet_detail.range(excel_range).options(pd.DataFrame, index=False).value

        self.first_result_colnum = 7
        self.first_log_colnum = self.first_result_colnum + len(PROJECT_INFO[self.project])
        self.end_log_colnum = self.first_log_colnum + len(PROJECT_INFO[self.project])
        self.first_time_colnum = self.first_result_colnum + len(PROJECT_INFO[self.project]) * 2
        self.end_time_colnum = self.first_time_colnum + len(PROJECT_INFO[self.project])
        self.domain_cols = self.df_detail.iloc[:, self.first_result_colnum:self.first_log_colnum].columns

    def delete_previous_data(self):
        prev_na = {}
        # delete result
        for col in range(self.first_result_colnum, self.first_log_colnum):
            domain = self.domain_cols[col - self.first_result_colnum]
            prev_dom_na = []
            for row in self.df_detail.iloc[:, col].index:
                if 'N/A' == self.df_detail.iloc[row, col]:
                    prev_dom_na.append('%03d' % int(self.df_detail.iloc[row, 0]))
                else:
                    self.df_detail.iloc[row, col] = ''
            prev_na[domain] = prev_dom_na

        # delete logs
        for col in range(self.first_log_colnum, self.end_log_colnum):
            for row in self.df_detail.iloc[:, col].index:
                self.df_detail.iloc[row, col] = ''

        # delete test time
        for col in range(self.first_time_colnum, self.end_time_colnum):
            print('delete col: ', col)
            for row in self.df_detail.iloc[:, col].index:
                self.df_detail.iloc[row, col] = ''

        return prev_na

    # def set_summary_test_time(self, test_time, project):
    #     domain_elapsed = {'vm': 0, 'bm': 0}
    #     for i, domain in enumerate(PROJECT_INFO[project]):
    #         print(f"start: {test_time[domain]['start']} / end: {test_time[domain]['end']}")
    #         self.df_summary.iloc[1+i, 2] = test_time[domain]['start']
    #         self.df_summary.iloc[1+i, 3] = test_time[domain]['end']
    #         self.df_summary.iloc[1+i, 4] = round(int(test_time[domain]['elapsed'])/60) # minute
    #
    #         # bm/vm 구분해서 시간 저장
    #         if 'BM' in domain:
    #             domain_elapsed['bm'] += round(int(test_time[domain]['elapsed'])/60)
    #             print(f"cur dom: {domain}, time: {round(int(test_time[domain]['elapsed'])/60)}")
    #         else:
    #             domain_elapsed['vm'] += round(int(test_time[domain]['elapsed'])/60)
    #             print(f"cur dom: {domain}, time: {round(int(test_time[domain]['elapsed']) / 60)}")
    #     print('domain test time: ', domain_elapsed)
    #     return domain_elapsed

    def set_summary_test_time(self, test_time, project):
        domain_elapsed = {'vm': 0, 'bm': 0}
        for i, domain in enumerate(PROJECT_INFO[project]):
            print(f"start: {test_time[domain]['start']} / end: {test_time[domain]['end']}")
            self.df_summary.iloc[1+i, 2] = test_time[domain]['start']
            self.df_summary.iloc[1+i, 3] = test_time[domain]['end']
            self.df_summary.iloc[1+i, 4] = round(int(test_time[domain]['elapsed'])/60) # minute

            # bm/vm 구분해서 시간 저장
            if re.search('BM|BA|BL', domain):
                domain_elapsed['bm'] += round(int(test_time[domain]['elapsed'])/60)
                print(f"cur dom: {domain}, time: {round(int(test_time[domain]['elapsed'])/60)}")
            else:
                domain_elapsed['vm'] += round(int(test_time[domain]['elapsed'])/60)
                print(f"cur dom: {domain}, time: {round(int(test_time[domain]['elapsed']) / 60)}")
        print('domain test time: ', domain_elapsed)
        return domain_elapsed

    def set_summary_test_result(self, project, domain_elapsed):
        for i, domain in enumerate(PROJECT_INFO[project]):
            if re.search('BM|BA|BL', domain):
                # Test Time (minute)
                self.df_summary.iloc[9, 1] = domain_elapsed['bm']    # spent
                self.df_summary.iloc[9, 2] = manual_review           # review
                total_bm = domain_elapsed['bm'] + manual_review
                self.df_summary.iloc[9, 3] = total_bm                # total
                # Test Result (hour)
                self.df_summary.iloc[4, 16] = '%.2f H' % (auto_test_review_bm/60)   # Auto Test Review
                self.df_summary.iloc[4, 17] = '%.2f H' % (total_bm / 60)            # Manual Test
                self.df_summary.iloc[4, 18] = '%.2f H' % (test_setup/60)            # Test Setup
                self.df_summary.iloc[4, 19] = '%.2f H' % (create_report_bm/60)      # Create Report
            else:
                # Test Time (minute)
                self.df_summary.iloc[8, 1] = domain_elapsed['vm']    # spent
                self.df_summary.iloc[8, 2] = manual_review           # review
                total_vm = domain_elapsed['vm'] + manual_review
                self.df_summary.iloc[8, 3] = total_vm                # total
                # Test Result (hour)
                self.df_summary.iloc[3, 16] = '%.2f H' % (auto_test_review_vm/60)   # Auto Test Review
                self.df_summary.iloc[3, 17] = '%.2f H' % (total_vm / 60)            # Manual Test
                self.df_summary.iloc[3, 18] = '%.2f H' % (test_setup/60)            # Test Setup
                self.df_summary.iloc[3, 19] = '%.2f H' % (create_report_vm/60)      # Create Report

    # def set_summary_test_result(self, project, domain_elapsed):
    #     for i, domain in enumerate(PROJECT_INFO[project]):
    #         if 'BM' in domain:
    #             # Test Time (minute)
    #             self.df_summary.iloc[9, 1] = domain_elapsed['bm']    # spent
    #             self.df_summary.iloc[9, 2] = manual_review           # review
    #             total_bm = domain_elapsed['bm'] + manual_review
    #             self.df_summary.iloc[9, 3] = total_bm                # total
    #             # Test Result (hour)
    #             self.df_summary.iloc[4, 16] = '%.2f H' % (auto_test_review_bm/60)   # Auto Test Review
    #             self.df_summary.iloc[4, 17] = '%.2f H' % (total_bm / 60)            # Manual Test
    #             self.df_summary.iloc[4, 18] = '%.2f H' % (test_setup/60)            # Test Setup
    #             self.df_summary.iloc[4, 19] = '%.2f H' % (create_report_bm/60)      # Create Report
    #         else:
    #             # Test Time (minute)
    #             self.df_summary.iloc[8, 1] = domain_elapsed['vm']    # spent
    #             self.df_summary.iloc[8, 2] = manual_review           # review
    #             total_vm = domain_elapsed['vm'] + manual_review
    #             self.df_summary.iloc[8, 3] = total_vm                # total
    #             # Test Result (hour)
    #             self.df_summary.iloc[3, 16] = '%.2f H' % (auto_test_review_vm/60)   # Auto Test Review
    #             self.df_summary.iloc[3, 17] = '%.2f H' % (total_vm / 60)            # Manual Test
    #             self.df_summary.iloc[3, 18] = '%.2f H' % (test_setup/60)            # Test Setup
    #             self.df_summary.iloc[3, 19] = '%.2f H' % (create_report_vm/60)      # Create Report


    def set_detail_test_result(self, test_info):
        for i, dom in enumerate(self.domain_cols):
            cur_result_idx = self.first_result_colnum + i
            cur_logs_idx = self.first_log_colnum + i
            cur_time_idx = self.first_time_colnum + i
            print(f'DOMAIN: {dom}, res_idx: {cur_result_idx}, log_idx: {cur_logs_idx}, time_idx: {cur_time_idx}')
            for id, detail in test_info[dom].items():
                if float(id) in self.df_detail['SITL ID'].tolist():
                    row_idx = self.df_detail[self.df_detail['SITL ID'] == float(id)].index
                    self.df_detail.iloc[row_idx, cur_result_idx] = detail['result']
                    self.df_detail.iloc[row_idx, cur_logs_idx] = detail['logs']
                    if detail['elapsed'] != 0:
                        self.df_detail.iloc[row_idx, cur_time_idx] = detail['elapsed']

    def set_detail_total_test_time(self, elapsed_total):
        col_idx = self.end_time_colnum
        for id, elapsed in elapsed_total.items():
            if float(id) in self.df_detail['SITL ID'].tolist():
                row_idx = self.df_detail[self.df_detail['SITL ID'] == float(id)].index
                self.df_detail.iloc[row_idx, col_idx] = elapsed
