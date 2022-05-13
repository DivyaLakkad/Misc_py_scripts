import datetime

today = datetime.date.today()

idx = (today.weekday() + 1) % 7 # MON = 0, SUN = 6 -> SUN = 0 .. SAT = 6

for i in range(0, 7):
    x = today - datetime.timedelta(7+idx - i)
    final_x = x.strftime(f"%a \n%b %d")
    print(final_x)