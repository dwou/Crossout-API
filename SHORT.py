import openpyxl as xl, requests as rq, itertools as it, re

alpha = "abcdefghijklmnopqrstuvwxyz"
products = [''.join(p) for r in range(1, 4) for p in it.product(alpha, repeat=r)]
a_to_i_map = {i: sum([(int(str(j), 36) - 9) * 26 ** (len(i) - j - 1) for j in range(len(i))]) for i in products}

a_to_i = lambda x: a_to_i_map[x.lower()]
i_to_a = lambda x: i_to_a((x-1)//26) + alpha[x%26-1] if x > 0 else ""
xy_to_c = lambda x,y: i_to_a(x)+str(y)
#broken
lrange = lambda x: [(col, row) for col in range(ord(x.split(':')[0][:1]), ord(x.split(':')[1][:1]) + 1) for row in range(int(x.split(':')[0][1:]), int(x.split(':')[1][1:]) + 1)] if re.match(r'[A-Z]+\d+:[A-Z]+\d+', x) else None
oid_replace = lambda x: x.replace('AB', 'AK')
update_items = lambda url="https://crossoutdb.com/api/v1/items": open("Items.txt", "w+").write(str(rq.get(url).json()))
get_data = lambda: eval(open("Items.txt", "r").read())

file = "Crossout.xlsx"
wb = xl.load_workbook(filename=file, data_only=False)
sheet = wb["Crafting+Market"]

update_items()
big_arr = get_data()
id_indx_map = {j['id']: i for i, j in enumerate(big_arr)}
name_id_map = {i['name']: i['id'] for i in big_arr}
avail_id_map = {i['availableName']: i['id'] for i in big_arr}

oid_range = lrange("T34:T194")
if oid_range is not None:
    for x, y in oid_range:
        print(x,y,xy_to_c(x, y))
        cell_value = sheet[xy_to_c(x, y)].value
        if cell_value and cell_value.startswith('='):
            sheet[f"U{y}"] = oid_replace(cell_value)


ranges = [["L8:L194", "AA{y}", "AB{y}", "AI{y}"], ["C7:C22", "D{y}", "E{y}", "A1"]]
for r in ranges:
    for x, y in lrange(r[0]):
        print(lrange(r[0]))
        name = sheet.cell(row=y, column=x).value
        if name is not None and name != "Name":
            print(f"Processing cell ({x}, {y}), name: {name}")
            Map = name_id_map.get(name) or avail_id_map.get(name)
            print(f"Map: {Map}")
            if Map is not None:
                item = big_arr[id_indx_map[Map]]
                if item["buyOrders"] >= 1:
                    sheet[eval(f"f'{r[1]}'")] = item["sellPrice"] / 100
                    sheet[eval(f"f'{r[2]}'")] = item["buyPrice"] / 100
                    sheet[eval(f"f'{r[3]}'")] = item["buyOrders"]
                else:
                    print(f"Skipping cell ({x}, {y}) - buyOrders < 1")
            else:
                print("Doesn't work:", (x, y), name)
        else:
            print(f"Skipping cell ({x}, {y}) - Invalid name")

wb.save(filename="outfile.xlsx")
