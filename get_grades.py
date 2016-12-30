import pandas as pd
import json
from decimal import Decimal


def get_sorted_and_ranked_grades(path_to_grades_excel, id_column, grade_column):
    # example: path_to_grades_excel = 'grades.xlsx', id_column = '学号',
    # grade_column = '成绩'
    grades = pd.read_excel(path_to_grades_excel, sheetname=0)

    sorted_grades = grades.sort_values(by=[grade_column], ascending=False)
    sorted_grades['rank'] = sorted_grades[
        grade_column].rank(method='min', ascending=False)
    sorted_grades['percentage'] = sorted_grades.apply(
        lambda row: row['rank']/len(sorted_grades.index), axis=1)

    raw_records = [{'id': row[id_column], 'grade': row[grade_column],
                    'rank': row['rank'], 'percentage': row['percentage']}
                   for row in [sorted_grades.iloc[index]
                               for index in range(len(sorted_grades.index))]]

    temp_records = json.loads(
        pd.DataFrame(raw_records).reset_index().to_json(orient='records'))

    final_records = []

    for row in temp_records:
        final_row = {}
        final_row['id'] = str(row['id']).split('.')[0]
        final_row['grade'] = '{0:.2f}'.format(Decimal(row['grade']))
        final_row['rank'] = str(int(str(row['rank']).split('.')[0]))
        final_row['percentage'] = '{0:.2f}%'.format(
            Decimal(row['percentage']) * 100)

        final_records.append(final_row)

    return final_records
