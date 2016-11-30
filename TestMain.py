import os
import re
class Converter:
    def __init__(self):
        self.jarPath = os.getcwd() + "\\out\\artifacts\\xls2csv2_jar\\xls2csv2.jar"
    def convert(self, inputFileName, outputFileName, minColumns = -1, sheetName = 'Sheet1'):
        os.system("java -jar " + self.jarPath + " " + inputFileName + " " + outputFileName + " " + str(minColumns) + " " + sheetName)
    def convertAll(self, inputDir, outputDir):
        for inputFile in os.listdir(inputDir):
            if inputFile.endswith("xls") or inputFile.endswith("xlsx"):
                self.convert(inputDir + inputFile, outputDir)
                # print ""
            elif inputFile.endswith("csv"):
                fileName = os.path.splitext(inputFile)[0]
                sheetName = "Sheet1"
                match = re.match("(.+)_(.+).csv", inputFile, re.M|re.I)
                if match:
                    fileName = match.group(1)
                    sheetName = match.group(2)
                self.convert(inputDir + inputFile, outputDir + fileName, -1, sheetName)
            else:
                print "Invalid File"


if __name__ == '__main__':
    inputDir = os.getcwd() + r"/input/"
    outputDir = os.getcwd() + r"/output/"
    Converter().convertAll(inputDir, outputDir)