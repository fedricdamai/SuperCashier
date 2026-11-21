"""Row map shared by the workbook builder and the docx / slides exporters (v6 layout)."""
# Miscellaneous
MI = dict(S=10, A=11, V=12, VF=13, ST=14, D=15, REACH=16, RT=17, SEG0=22, PROD0=27, DL0=39, GROW0=45, SRC0=53)
# Task 1
T1 = dict(ANS0=10, A0=18, CHK=26, RK0=31, N0=42, TOT=50, SRC=52)
# Task 2: answers (rows 10-16), inputs (21-39), scenarios (43-54), lowest OKX share (56), top 3 (60-63)
T2 = dict(ANS0=10, NANS=7, VOL=21, FEE=22, ASK=23, BR1=24, BR2=25, BR3=26, S1=27, S2=28, S3=29, S4=30,
          MS1=31, MS2=32, MS3=33, MS4=34, PY1=35, PY2=36, PY3=37, PY4=38, ORDER=39, SC0=43, MINSH=56, TOP0=60)
CASES = ("50% of forecast", "Forecast", "$10M a month", "$20M a month")
# Task 3
T3 = dict(ITEMS0=10, PLAN1=20, PLAN0=22, OFFER0=31, EARN1=45, EARNT=46, KEEPT=47, SUP0=51)
# Task 4 (our 8 partners; pilot C is its own track)
T4 = dict(ANS0=10, FOCUS=19, W0=20, LIVE=24, ACT=25, FTD=26, TPD=27, SW8=28, REV=29, REVN=30, RUN=31,
          PIL0=35, R0=43, RTOT=51, P0=52, DEL0=58, MON0=69, CAD0=81)
# Task 4 Gantt tab
GT = dict(G0=10, KPI0=27)
