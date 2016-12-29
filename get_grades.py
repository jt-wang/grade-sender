import pandas as pd
import json


def get_sorted_and_ranked_grades(path_to_grades_excel, id_column, grade_column):
    # example: path_to_grades_excel = 'grades.xlsx', id_column = '学号',
    # grade_column = '成绩'
    grades = pd.read_excel(path_to_grades_excel, sheetname=0)

    sorted_grades = grades.sort_values(by=[grade_column], ascending=False)
    sorted_grades['rank'] = sorted_grades[
        grade_column].rank(method='min', ascending=False)
    sorted_grades['percentage'] = sorted_grades.apply(
        lambda row: row['rank']/len(sorted_grades.index), axis=1)

    records = [{'id': row[id_column], 'grade': row[grade_column],
                'rank': row['rank'], 'percentage': row['percentage']}
               for row in [sorted_grades.iloc[index]
                           for index in range(len(sorted_grades.index))]]

    return json.loads(pd.DataFrame(records).reset_index().to_json(orient='records'))
