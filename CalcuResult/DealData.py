# coding=utf-8
import math
import sys
import xlrd
import xlwt
import numpy as np
import random

np.set_printoptions(threshold=np.inf)


class DealData:

    @staticmethod
    def analyzeData(filepath):
        # f = open("outputmatrix.txt", 'w+')
        # 1、解析xlsx
        readData = xlrd.open_workbook(filepath)
        sheet = readData.sheet_by_name('Sheet3')
        nrows = sheet.nrows
        # testValue = sheet.cell(3, 0).value
        # print(nrows)

        # 2、将readData中的数据组成矩阵来计算Xestimate
        dataMatrix = np.zeros((32, 32))

        # 取表格中某一列的值
        dataValue = sheet.col_values(0)

        matrixRow = 0
        matrixCol = 0
        for index, data in enumerate(dataValue):
            if (index + 1) % 32 == 0:
                dataMatrix[matrixRow][matrixCol] = data
                matrixRow = matrixRow + 1
                matrixCol = 0
                continue
            dataMatrix[matrixRow][matrixCol] = int(data)
            matrixCol = matrixCol + 1
        # print(dataMatrix, file=f)
        # print(dataMatrix)

        gaussDis1 = []
        gaussDis2 = []
        gaussDis3 = []
        gaussDis4 = []

        meanDis1 = []
        meanDis2 = []
        meanDis3 = []
        meanDis4 = []

        randDis1 = []
        randDis2 = []
        randDis3 = []
        randDis4 = []

        nearDis1 = []
        nearDis2 = []
        nearDis3 = []
        nearDis4 = []

        # 将所有点的不同距离的误差值分别取出来放在对应数组里，供后面分别求平均误差值
        for row in range(7, 26): #边界：row--7, 26  col-5, 27
            for col in range(5, 23):  # （26-7）* （23-5）= 342 个点
                deviationOfdisArry = gaussCalcuX(dataMatrix, row, col, 4)
                gaussDis1.append(deviationOfdisArry[0])
                gaussDis2.append(deviationOfdisArry[1])
                gaussDis3.append(deviationOfdisArry[2])
                gaussDis4.append(deviationOfdisArry[3])

                meanDeviationOfdisArry = meanCalcuX(dataMatrix, row, col, 4)
                meanDis1.append(meanDeviationOfdisArry[0])
                meanDis2.append(meanDeviationOfdisArry[1])
                meanDis3.append(meanDeviationOfdisArry[2])
                meanDis4.append(meanDeviationOfdisArry[3])

                randDeviationOfdisArry = randCalcuX(dataMatrix, row, col, 4)
                randDis1.append(randDeviationOfdisArry[0])
                randDis2.append(randDeviationOfdisArry[1])
                randDis3.append(randDeviationOfdisArry[2])
                randDis4.append(randDeviationOfdisArry[3])

                nearDeviationOfdisArry = nearestCalcuX(dataMatrix, row, col, 4)
                nearDis1.append(nearDeviationOfdisArry[0])
                nearDis2.append(nearDeviationOfdisArry[1])
                nearDis3.append(nearDeviationOfdisArry[2])
                nearDis4.append(nearDeviationOfdisArry[3])


        writebook = xlwt.Workbook()  # 打开一个excel
        sheet = writebook.add_sheet('ResultSheet')  # 在打开的excel中添加一个sheet

        sheet.write(0, 0, np.mean(gaussDis1))  # 写入excel
        sheet.write(1, 0, np.mean(gaussDis2))
        sheet.write(2, 0, np.mean(gaussDis3))
        sheet.write(3, 0, np.mean(gaussDis4))

        sheet.write(0, 1, np.mean(meanDis1))  # 写入excel
        sheet.write(1, 1, np.mean(meanDis2))
        sheet.write(2, 1, np.mean(meanDis3))
        sheet.write(3, 1, np.mean(meanDis4))

        sheet.write(0, 2, np.mean(randDis1))  # 写入excel
        sheet.write(1, 2, np.mean(randDis2))
        sheet.write(2, 2, np.mean(randDis3))
        sheet.write(3, 2, np.mean(randDis4))

        sheet.write(0, 3, np.mean(nearDis1))  # 写入excel
        sheet.write(1, 3, np.mean(nearDis2))
        sheet.write(2, 3, np.mean(nearDis3))
        sheet.write(3, 3, np.mean(nearDis4))

        writebook.save('SoftwareCalibrationData.xls')  # 一定要记得保存

    # DealData.calcuX(dataMatrix, 10, 15, 3)


"""
   功能：高斯法求解估计值与真实值误差
   参数：源点坐标（row，col），groupNum--组数（由近到远取groupNum组值）
"""


# @staticmethod
def gaussCalcuX(datamatrix, row, col, groupNum):
    # # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14],
    # #          datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
    # # sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15],
    # #          datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]

    result = []
    # groupNum = 5
    i = 1
    while i <= groupNum:

        s0Row = row - i
        s0Col = col + i
        s0 = datamatrix[s0Row][s0Col]
        print("s0--%r" % s0)

        numL = 1
        sDataL = []
        sDataL.append(s0)
        while numL <= 4:
            sDataL.append(datamatrix[s0Row + numL][s0Col + numL])
            numL = numL + 1

        print("sDataL--%r" % sDataL)

        numR = 1
        sDataR = []
        while numR <= 4:
            sDataR.append(datamatrix[s0Row - numR][s0Col - numR])
            numR = numR + 1

        print("sDataR--%r" % sDataR)

        # 算Gi
        con = 2
        gArray = []
        for index, data in enumerate(sDataL):
            gi = (1 / (2 * math.pi * np.sqrt(con))) * math.exp(-(np.sqrt(index) / (2 * np.sqrt(con))))
            gArray.append(gi)

        for index, data in enumerate(sDataR):
            gi = (1 / (2 * math.pi * np.sqrt(con))) * math.exp(-(np.sqrt(index + 1) / (2 * np.sqrt(con))))
            gArray.append(gi)

        print("gArray--%r" % gArray)

        # 算sum(Gi * si)
        sData = sDataL + sDataR
        print("sData--%r" % sData)
        sumArray = np.multiply(np.array(sData), np.array(gArray))
        print("sumArray--%r" % sumArray)
        sumGS = sum(sumArray)
        print("sumGS--%r" % sumGS)

        # 算sum(Gi)
        sumG = sum(gArray)
        print("sumGS--%r" % sumG)

        # 算Xestimate
        xEstimate = sumGS / sumG
        xTruth = datamatrix[row][col]
        print("xEstimate--%r" % xEstimate)
        print("xTruth--%r" % xTruth)

        # 算误差
        deviation = abs(xTruth - xEstimate)
        print("deviation--%r" % deviation)

        result.append(deviation)

        i = i + 1

    print("result--%r" % result)
    return result

"""
   功能：均值法求解估计值与真实值误差
   参数：源点坐标（row，col），groupNum--组数（由近到远取groupNum组值）
"""
def meanCalcuX(datamatrix, row, col, groupNum):
    # # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14],
    # #          datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
    # # sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15],
    # #          datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]

    result = []
    # groupNum = 5
    i = 1
    while i <= groupNum:

        s0Row = row - i
        s0Col = col + i
        s0 = datamatrix[s0Row][s0Col]
        print("s0--%r" % s0)

        numL = 1
        sDataL = []
        sDataL.append(s0)
        while numL <= 4:
            sDataL.append(datamatrix[s0Row + numL][s0Col + numL])
            numL = numL + 1

        print("sDataL--%r" % sDataL)

        numR = 1
        sDataR = []
        while numR <= 4:
            sDataR.append(datamatrix[s0Row - numR][s0Col - numR])
            numR = numR + 1

        print("sDataR--%r" % sDataR)

        sData = sDataL + sDataR
        print("sDataR--%r" % sData)

        # 算Xestimate
        xEstimate = np.mean(sData)
        xTruth = datamatrix[row][col]
        print("xEstimate--%r" % xEstimate)
        print("xTruth--%r" % xTruth)

        # 算误差
        deviation = abs(xTruth - xEstimate)
        print("deviation--%r" % deviation)

        result.append(deviation)

        i = i + 1

    print("result--%r" % result)
    return result

"""
   功能：随机数法求解估计值与真实值误差
   参数：源点坐标（row，col），groupNum--组数（由近到远取groupNum组值）
"""
def randCalcuX(datamatrix, row, col, groupNum):
    # # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14],
    # #          datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
    # # sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15],
    # #          datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]

    result = []
    # groupNum = 5
    i = 1
    while i <= groupNum:

        s0Row = row - i
        s0Col = col + i
        s0 = datamatrix[s0Row][s0Col]
        print("s0--%r" % s0)

        numL = 1
        sDataL = []
        sDataL.append(s0)
        while numL <= 4:
            sDataL.append(datamatrix[s0Row + numL][s0Col + numL])
            numL = numL + 1

        print("sDataL--%r" % sDataL)

        numR = 1
        sDataR = []
        while numR <= 4:
            sDataR.append(datamatrix[s0Row - numR][s0Col - numR])
            numR = numR + 1

        print("sDataR--%r" % sDataR)

        sData = sDataL + sDataR
        print("sDataR--%r" % sData)

        # 算Xestimate
        xEstimate = random.sample(sData,1)
        xTruth = datamatrix[row][col]
        print("xEstimate--%r" % xEstimate)
        print("xTruth--%r" % xTruth)

        # 算误差
        deviation = abs(xTruth - xEstimate)
        print("deviation--%r" % deviation)

        result.append(deviation)

        i = i + 1

    print("result--%r" % result)
    return result

"""
   功能：离源点最近值作为估计值，与真实值误差
   参数：源点坐标（row，col），groupNum--组数（由近到远取groupNum组值）
"""
def nearestCalcuX(datamatrix, row, col, groupNum):
    # # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14],
    # #          datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
    # # sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15],
    # #          datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]

    result = []
    # groupNum = 5
    i = 1
    while i <= groupNum:

        s0Row = row - i
        s0Col = col + i
        s0 = datamatrix[s0Row][s0Col]
        print("s0--%r" % s0)
        #
        # numL = 1
        # sDataL = []
        # sDataL.append(s0)
        # while numL <= 4:
        #     sDataL.append(datamatrix[s0Row + numL][s0Col + numL])
        #     numL = numL + 1
        #
        # print("sDataL--%r" % sDataL)
        #
        # numR = 1
        # sDataR = []
        # while numR <= 4:
        #     sDataR.append(datamatrix[s0Row - numR][s0Col - numR])
        #     numR = numR + 1
        #
        # print("sDataR--%r" % sDataR)
        #
        # sData = sDataL + sDataR
        # print("sDataR--%r" % sData)

        # 算Xestimate
        xEstimate = s0
        xTruth = datamatrix[row][col]
        print("xEstimate--%r" % xEstimate)
        print("xTruth--%r" % xTruth)

        # 算误差
        deviation = abs(xTruth - xEstimate)
        print("deviation--%r" % deviation)

        result.append(deviation)

        i = i + 1

    print("result--%r" % result)
    return result


# """
# 功能：高斯法求解估计值与真实值误差
# 参数：源点坐标（row，col），groupNum--组数（由近到远取groupNum组值）
# """
# @staticmethod
# def calcuX(datamatrix, row, col, groupNum):
#     # # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14],
#     # #          datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
#     # # sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15],
#     # #          datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]
#
#     result = []
#     # groupNum = 5
#     i = 1
#     while i <= groupNum:
#
#
#         s0Row = row - i
#         s0Col = col + i
#         s0 = datamatrix[s0Row][s0Col]
#         print("s0--%r" % s0)
#
#         numL = 1
#         sDataL = []
#         sDataL.append(s0)
#         while numL <= 4:
#             sDataL.append(datamatrix[s0Row+numL][s0Col+numL])
#             numL = numL + 1
#
#         print("sDataL--%r" % sDataL)
#
#         numR = 1
#         sDataR = []
#         while numR <= 4:
#             sDataR.append(datamatrix[s0Row - numR][s0Col - numR])
#             numR = numR + 1
#
#         print("sDataR--%r" % sDataR)
#
#         # 算Gi
#         con = 2
#         gArray = []
#         for index, data in enumerate(sDataL):
#             gi = (1 / (2 * math.pi * np.sqrt(con))) * math.exp(-(np.sqrt(index) / (2 * np.sqrt(con))))
#             gArray.append(gi)
#
#         for index, data in enumerate(sDataR):
#             gi = (1 / (2 * math.pi * np.sqrt(con))) * math.exp(-(np.sqrt(index + 1) / (2 * np.sqrt(con))))
#             gArray.append(gi)
#
#         print("gArray--%r" % gArray)
#
#         # 算sum(Gi * si)
#         sData = sDataL + sDataR
#         print("sData--%r" % sData)
#         sumArray = np.multiply(np.array(sData), np.array(gArray))
#         print("sumArray--%r" % sumArray)
#         sumGS = sum(sumArray)
#         print("sumGS--%r" % sumGS)
#
#         # 算sum(Gi)
#         sumG = sum(gArray)
#         print("sumGS--%r" % sumG)
#
#         # 算Xestimate
#         xEstimate = sumGS / sumG
#         xTruth = datamatrix[row][col]
#         print("xEstimate--%r" % xEstimate)
#         print("xTruth--%r" % xTruth)
#
#         # 算误差
#         deviation = abs(xTruth - xEstimate)
#         print("deviation--%r" % deviation)
#
#         result.append(deviation)
#
#         i = i + 1
#
#     writebook = xlwt.Workbook()  # 打开一个excel
#     sheet = writebook.add_sheet('ResultSheet')  # 在打开的excel中添加一个sheet
#
#     for index,data in enumerate(result):
#         sheet.write(index,0, result[index])  # 写入excel
#
#     writebook.save('SoftwareCalibrationData.xls')  # 一定要记得保存
