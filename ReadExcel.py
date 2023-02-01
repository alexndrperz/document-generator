import pandas

def reader(file):
    df = pandas.read_excel(file, header=None)

    y = df.fillna("miss")
    column = len(y.columns)
    # print(column)
    arr = []

    for x in range(column):
        for index, row in y.iterrows():
            if (row[x] != "miss" and row[x] != "Nombre"):
                arr.append(row[x])
        x +=1

    return arr

