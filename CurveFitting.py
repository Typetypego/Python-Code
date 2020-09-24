import numpy as np
import xlrd
import matplotlib.pyplot as plt


def exceltolist(path, index):
    table = xlrd.open_workbook(path).sheet_by_index(index)

    d = []
    # 将表中数据按行逐步添加到列表中，最后转换为list结构
    for r in range(table.nrows):
        data1 = []
        for c in range(table.ncols):
            data1.append(table.cell_value(r, c))
        d.append(list(data1))
    return d


# Read file
file = 'chinaOpen.xlsx'
# file = r'ZJUNlictData-master\chinaOpen.xlsx'
FlatData = exceltolist(file, 0)
ChipData = exceltolist(file, 1)

aFlat = np.zeros(12)
bFlat = np.zeros(12)
cFlat = np.zeros(12)
minFlat = np.zeros(12)
maxFlat = np.zeros(12)
aChip = np.zeros(12)
bChip = np.zeros(12)
cChip = np.zeros(12)
minChip = np.zeros(12)
maxChip = np.zeros(12)


# Flat Calculation
for RobotID in range(12):
    if RobotID < 6:
        rowNum_low = 1
        rowNum_high = 11
        columnNum = 3 * RobotID
    else:
        rowNum_low = 14
        rowNum_high = 24
        columnNum = 3 * (RobotID - 6)

    for i in range(rowNum_low, rowNum_high + 1):
        if not FlatData[i][columnNum + 1]:
            rowNum_high = i - 1
            break


    FlatPower = [FlatData[i][columnNum] for i in range(rowNum_low, rowNum_high + 1)]
    FlatVel = [(FlatData[i][columnNum + 1] + FlatData[i][columnNum + 2]) / 2 for i in
               range(rowNum_low, rowNum_high + 1)]

    p = np.polyfit(FlatVel, FlatPower, 2)

    aFlat[RobotID] = p[0] * 100000
    bFlat[RobotID] = p[1]
    cFlat[RobotID] = p[2]

    minFlat[RobotID] = 30
    maxFlat[RobotID] = p[0] * 620 * 620 + p[1] * 620 + p[2]

    if maxFlat[RobotID] > 127:
        maxFlat[RobotID] = 127


# Chip Calculation
for RobotID in range(12):
    if RobotID < 6:
        rowNum_low = 3
        rowNum_high = 13
        columnNum = 3 * RobotID
    else:
        rowNum_low = 18
        rowNum_high = 28
        columnNum = 3 * (RobotID - 6)

    for i in range(rowNum_low, rowNum_high + 1):
        if not ChipData[i][columnNum + 1]:
            rowNum_high = i - 1
            break

    ChipPower = [ChipData[i][columnNum] for i in range(rowNum_low, rowNum_high + 1)]
    ChipDistance = [(ChipData[i][columnNum + 1] + ChipData[i][columnNum + 2]) / 2 for i in
                    range(rowNum_low, rowNum_high + 1)]

    p = np.polyfit(ChipDistance, ChipPower, 2)

    aChip[RobotID] = p[0]
    bChip[RobotID] = p[1]
    cChip[RobotID] = p[2]

    minChip[RobotID] = 40
    maxChip[RobotID] = p[0] * 400 * 400 + p[1] * 400 + p[2]

    if maxChip[RobotID] > 127:
        maxChip[RobotID] = 127


# Output
allData = np.zeros((10, 12))
allData[0, :] = aFlat
allData[1, :] = bFlat
allData[2, :] = cFlat
allData[3, :] = minFlat
allData[4, :] = maxFlat
allData[5, :] = aChip
allData[6, :] = bChip
allData[7, :] = cChip
allData[8, :] = minChip
allData[9, :] = maxChip

# Write txt
with open("Test_Right.txt", "w") as f:
    for RobotID in range(12):
        f.write('[Robot%d]\n' % RobotID)
        f.write('FLAT_A=%f\n' % aFlat[RobotID])
        f.write('FLAT_B=%f\n' % bFlat[RobotID])
        f.write('FLAT_C=%f\n' % cFlat[RobotID])
        f.write('FLAT_MIN=%f\n' % minFlat[RobotID])
        f.write('FLAT_MAX=%f\n' % maxFlat[RobotID])
        f.write('CHIP_A=%f\n' % aChip[RobotID])
        f.write('CHIP_B=%f\n' % bChip[RobotID])
        f.write('CHIP_C=%f\n' % cChip[RobotID])
        f.write('CHIP_MIN=%f\n' % minChip[RobotID])
        f.write('CHIP_MAX=%f\n' % maxChip[RobotID])

np.savetxt("Test.txt", allData, fmt='%f', delimiter=',')

# Plot
plt.plot([FlatData[i][1] for i in range(1, 8)], [FlatData[i][0] for i in range(1, 8)], color="blue")
plt.plot([FlatData[i][4] for i in range(1, 8)], [FlatData[i][3] for i in range(1, 8)], color="red")
plt.ylim([30, 90])
plt.xlabel("Ball Speed (cm/s)")
plt.ylabel("Flat Kick Power")
plt.legend(["Test Data 1", "Test Data 2"])
plt.show()
