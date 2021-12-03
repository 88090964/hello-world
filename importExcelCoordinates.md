import os
import openpyxl
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

path = r"C:\Users\Administrator\Desktop"
os.chdir(path)

workbook = openpyxl.load_workbook("新瑞机电（CGCS2000）(1).xlsx")
print(workbook.sheetnames)

sheet = workbook["Sheet1"]
print(sheet)


def sheetCell(name):
    listCell = []
    list = []
    for i in range(1, 72):
        listCell.append(name+str(i))
        print(listCell)

    for j in listCell:
        cell = sheet[j]
        print(cell.value)
        list.append(cell.value)
    return list

Xi = "A"
Yi = "B"
Zi = "C"
Xi = sheetCell(Xi)
Yi = sheetCell(Yi)
Zi = sheetCell(Zi)

fig = plt.figure()
ax = Axes3D(fig)
ax.scatter(Xi, Yi, Zi)

ax.set_zlabel('Z', fontdict={'size': 15, 'color': 'red'})
ax.set_ylabel('Y', fontdict={'size': 15, 'color': 'red'})
ax.set_xlabel('X', fontdict={'size': 15, 'color': 'red'})
plt.show()
