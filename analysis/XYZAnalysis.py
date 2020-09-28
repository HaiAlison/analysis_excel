import pandas as pd
import numpy as np

def AnalysisXYZ(data):
    Months = ['T1', 'T2', 'T3', 'T4', 'T5', 'T6', 
              'T7', 'T8', 'T9', 'T10', 'T11', 'T12']
    FillData = data.copy()
    FillData[Months] = data[Months].fillna(0)
    # Remove non-use drug
    FillData['TotalDemand'] = FillData[Months].apply(np.sum, axis=1)
    FillData = FillData[(FillData[Months].T != 0).any()]
    FillData.reset_index(drop=True)
    
    FillData['CV'] =  FillData[Months].apply(np.std, axis=1) / (FillData[Months].apply(np.sum, axis=1) / 12) 
    
    FillData["XYZ"] = None

    FillData.loc[FillData["CV"] > 0.56, "XYZ"] = "Z"
    FillData.loc[FillData["CV"] <= 0.33, "XYZ"] = "X"
    FillData.loc[FillData["XYZ"].isna(), "XYZ"] = "Y"
    
    return FillData