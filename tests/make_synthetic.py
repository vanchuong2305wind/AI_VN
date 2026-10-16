import sys, random, datetime as dt, openpyxl
out = sys.argv[1]
wb = openpyxl.Workbook()
ws = wb.active; ws.title = "Sales"
ws.append(["Sales report 2024"]); ws.append([])
ws.append(["Order ID", "Date", "Region", "Product", "Units", "Revenue ($)"])
random.seed(0)
for i in range(5000):
    ws.append([f"O{i:05d}", dt.datetime(2024,1,1)+dt.timedelta(days=random.randint(0,365)),
               random.choice(["North","South","East","West"]), random.choice(["Pen","Book","Bag"]),
               random.randint(1,50), round(random.uniform(5,500),2)])
ws2 = wb.create_sheet("Summary")
ws2.append(["Table 1. Revenue by region"]); ws2.append(["Region", "Q1", "Q2", "Q3", "Q4"])
for r in ["North","South","East","West"]: ws2.append([r]+[random.randint(100,999) for _ in range(4)])
ws2.append(["Total", 1,2,3,4]); ws2.append([]); ws2.append([])
ws2.append(["Table 2. Headcount by department"]); ws2.append(["Department", "Male", "Female"])
ws2.merge_cells(start_row=ws2.max_row+1, start_column=1, end_row=ws2.max_row+1, end_column=3)
ws2.append(["ENGINEERING"])
for d in ["Backend","Frontend"]: ws2.append([d, random.randint(5,50), random.randint(5,50)])
ws2.append(["SALES"])
for d in ["Inside","Field"]: ws2.append([d, random.randint(5,50), random.randint(5,50)])
ws2.append(["Source: HR system"])
wb.save(out)
