from datetime import date
import datetime

today = str(datetime.datetime.now().isoformat(sep="-",timespec="seconds"))

s = today
s = s.replace("/", "_").replace(":", "-")
print(s)