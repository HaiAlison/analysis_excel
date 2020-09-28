# ABC-VEN analysis

import pandas as pd
import numpy as np
import sklearn.cluster

def SplitIndex(data, lower, upper, trials=10):
    result = np.zeros(trials)
    for i in np.arange(trials):
        kmeans = sklearn.cluster.KMeans(n_clusters=2).fit(data[(data["CumSumPercentage"] > lower) & (data["CumSumPercentage"] < upper)].loc[:, ["CumSumPercentage", "THÀNH TIỀN"]])
        result[i] = len(kmeans.labels_) - sum(kmeans.labels_)
    return max(result)

def AnalysisABC(data):
    # Total value
    TotalValue = data["THÀNH TIỀN"].sum()
    # Sort in descending order 
    result = data.sort_values(by=["THÀNH TIỀN"], ascending=False).reset_index(drop=True)
    result.loc[:, "Percentage"] = result["THÀNH TIỀN"] / TotalValue
    # Cumulative sum
    result["CumSumPercentage"] = 0
    result.loc[0, "CumSumPercentage"] = result.loc[0, "Percentage"].copy()
    # Delta of percentage
    result["Delta"] = 0
    result.loc[0, "Delta"] = result.loc[0, "Percentage"].copy()
    for i in range(1, result.shape[0]):
        result.loc[i, "CumSumPercentage"] =  result.loc[i-1, "CumSumPercentage"] + result.loc[i, "Percentage"]
        result.loc[i, "Delta"] =  result.loc[i, "Percentage"] / result.loc[i-1, "Percentage"]
    
    #ABC classify
    result["ABC"] = None
    ## A class
    result.loc[result["CumSumPercentage"] <= 0.59, "ABC"] = "A"
    AClassIndex = result[result["CumSumPercentage"] < 0.59].shape[0]
    result.loc[AClassIndex : AClassIndex + SplitIndex(result, 0.59, 0.82) - 1, "ABC"] = "A"
    ## B class
    BClassIndex = result[result["ABC"].notna()].shape[0] # Get first B index by remove number of A    
    result.loc[BClassIndex : result[result["CumSumPercentage"] <= 0.89].shape[0] + SplitIndex(result, 0.89, 0.96) - 1, "ABC"] = "B"
    ## C class 
    result.loc[result["ABC"].isna(), "ABC"] = "C"
    
    return result.drop(columns=['Delta'])

def ABCVEN(data):
    YearArr = data['NĂM'].unique()
    ResultByYear = [None for i in range(len(YearArr))]
    DataCols = ['MÃ HÀNG', 'MÃ MẶT HÀNG', 'TÊN THUỐC', 'HOẠT CHẤT', 'HÀM LƯỢNG',
       'ĐVT', 'TCKT', 'VEN', 'HÃNG SX', 'NƯỚC SX', 'NHÀ CUNG ỨNG', 'ĐƠN GIÁ', 'NHÓM THUỐC', 'HỆ ĐIỀU TRỊ', 'NĂM']
    DataColsOrder = ['MÃ HÀNG', 'MÃ MẶT HÀNG', 'TÊN THUỐC', 'HOẠT CHẤT', 'HÀM LƯỢNG',
       'ĐVT', 'TCKT', 'HÃNG SX', 'NƯỚC SX', 'NHÀ CUNG ỨNG', 'ĐƠN GIÁ', 'NHÓM THUỐC', 'HỆ ĐIỀU TRỊ', 'NĂM', 'THÀNH TIỀN', 
                     'Percentage', 'CumSumPercentage', 'ABC', 'VEN', 'ABC-VEN']
    # Remove non-use drug
    data.drop(data[data["THÀNH TIỀN"] == 0].index, inplace=True)
    
    for ind, y in enumerate(YearArr):
        # Get match timeline 
        DataOfYear = data[data['NĂM'] == y].copy()
        # Group by drug code
        DataOfYearGroup = DataOfYear.groupby("MÃ HÀNG").sum()["THÀNH TIỀN"].reset_index()
        # ABC analysis 
        DataOfYearABC = AnalysisABC(DataOfYearGroup)
#         print(DataOfYearABC.head())
        # VEN and ABC/VEN Analysis
        ## Merge VEN
        dataResult = pd.merge(DataOfYearABC, DataOfYear[DataCols].drop_duplicates(), 
                              left_on="MÃ HÀNG", 
                              right_on="MÃ HÀNG")
        dataResult["ABC-VEN"] = dataResult["ABC"] + dataResult["VEN"] 
        
        ResultByYear[ind] = dataResult[DataColsOrder]
        
    return ResultByYear
