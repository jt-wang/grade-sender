# README

**This repo is written for Teacher Assitants of Fudan University to send grades to each student respectively with his/her own student-id, grade, rank, and percentage.**

All you need is:

- **a `fudan.edu.cn` email account**
- **an excel file**, in which their are at least columns for **student-id** and **grade**.

You don't need to manually use Excel to generate ranks or percentages since it has been integrated in this repo leveraging `pandas`. The entire procedure is done in memory therefore there is no change to your excel file or extra file on your disk.

requirements:

- **python3**
- **xlrd** (pip install xlrd)
- **pandas** (pip install pandas)

if you already have python3 installed, you can use the `requirements.txt` as below to install `xlrd` and `pandas` at the same time:

```shell
pip install -r requirements.txt
```

run `send_grades.py` and follow the instructions.

enjoy!
