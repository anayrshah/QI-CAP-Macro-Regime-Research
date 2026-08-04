import calendar
from datetime import date
import csv
from pathlib import Path

dates = []
for year in range(2010, 2027):
    for month in range(1, 13):
        if year == 2026 and month > 5:
            break
        last_day = calendar.monthrange(year, month)[1]
        dates.append(date(year, month, last_day))

output_path = Path("Monthly_Calendar_2010_2026.csv")

with open(output_path, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Last Day of Month"])
    for d in dates:
        writer.writerow([d.isoformat()])

print(f"Saved to: {output_path.resolve()}")
print(f"Rows: {len(dates)}")
