# coding=utf-8
import math
import sys
import xlrd
import xlwt
import numpy as np

np.set_printoptions(threshold=np.inf)


class DealData:

    @staticmethod
    def analyzeData(filepath):
        # f = open("outputmatrix.txt", 'w+')
        # 1、解析xlsx
        readData = xlrd.open_workbook(filepath)
        sheet = readData.sheet_by_name('Sheet1')
        nrows = sheet.nrows
        # testValue = sheet.cell(3, 0).value
        # print(nrows)

        # 2、将readData中的数据组成矩阵来计算Xestimate
        dataMatrix = np.zeros((32, 32))
        # print(dataMatrix)
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
        DealData.calcuX(dataMatrix, 9, 9)

    @staticmethod
    def calcuX(datamatrix, row, col):
        # # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14],
        # #          datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
        # # sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15],
        # #          datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]

        result = []
        groupNum = 5
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
                sDataL.append(datamatrix[s0Row+numL][s0Col+numL])
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

        writebook = xlwt.Workbook()  # 打开一个excel
        sheet = writebook.add_sheet('ResultSheet')  # 在打开的excel中添加一个sheet

        for index,data in enumerate(result):
            sheet.write(index,0, result[index])  # 写入excel

        writebook.save('SoftwareCalibrationData.xls')  # 一定要记得保存

    # @staticmethod
    # def calcuX(datamatrix):
    #
    #     # sData = [46,46,41,42,42,44,43,44,45]
    #     xTruth = datamatrix[9][9]
    #
    #     # sData = [datamatrix[8][10], datamatrix[9][11], datamatrix[10][12], datamatrix[11][13], datamatrix[12][14], datamatrix[7][9], datamatrix[6][8], datamatrix[5][7], datamatrix[4][6]]
    #     sData = [datamatrix[7][11], datamatrix[8][12], datamatrix[9][13], datamatrix[10][14], datamatrix[11][15], datamatrix[6][10], datamatrix[5][9], datamatrix[4][8], datamatrix[3][7]]
    #     print("sData--%r" % sData)
    #
    #     sDataL = sData[0:5]
    #     print("sDataL--%r" % sDataL)
    #
    #     sDataR = sData[5:]
    #     print("sDataR--%r" % sDataR)
    #
    #     # 算Gi
    #     con = 2
    #     gArray = []
    #     for index,data in enumerate(sDataL):
    #         gi = (1/(2*math.pi*np.sqrt(con)))*math.exp(-(np.sqrt(index)/(2*np.sqrt(con))))
    #         gArray.append(gi)
    #
    #     for index,data in enumerate(sDataR):
    #         gi = (1/(2*math.pi*np.sqrt(con)))*math.exp(-(np.sqrt(index+1)/(2*np.sqrt(con))))
    #         gArray.append(gi)
    #
    #     print("gArray--%r" % gArray)
    #
    #     # 算sum(Gi * si)
    #     sumArray = np.multiply(np.array(sData), np.array(gArray))
    #     print("sumArray--%r" % sumArray)
    #     sumGS = sum(sumArray)
    #     print("sumGS--%r" % sumGS)
    #
    #     # 算sum(Gi)
    #     sumG = sum(gArray)
    #     print("sumGS--%r" % sumG)
    #
    #     # 算Xestimate
    #     xEstimate = sumGS/sumG
    #     print("xEstimate--%r" % xEstimate)
    #     print("xTruth--%r" % xTruth)
    #
    #     # 算误差
    #     deviation = abs(xTruth - xEstimate)
    #     print("deviation--%r" % deviation)
