import sys
from .ABCAnalysis import *
from .XYZAnalysis import *
from .ToExcel import *
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def main():
    if len(sys.argv) < 2:
        return('Please insert file!')
    
    dataABC = pd.read_excel(sys.argv[1], sheet_name=0) # assume
    dataXYZ = pd.read_excel(sys.argv[1], sheet_name=1)

    dataResultsABC = ABCVEN(dataABC)
    dataResultsXYZ = AnalysisXYZ(dataXYZ)

    wb = Workbook()
    ws = wb.active
    ws.title = "ABC_VEN_Data"
    for r in dataframe_to_rows(pd.concat(dataResultsABC), index=False, header=True):
        ws.append(r)
    ABCAnalysisDisplay(wb, pd.concat(dataResultsABC))

    ws = wb.create_sheet(title="XYZ_Data")
    for r in dataframe_to_rows(dataResultsXYZ, index=False, header=True):
        ws.append(r)
    XYZAnalysisDisplay(wb, dataResultsXYZ)

    wb.save('Test.xlsx')
    
if __name__ == '__main__':
    main()
