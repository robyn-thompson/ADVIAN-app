import numpy as np
import math
from statistics import mean
import csv
import pandas as pd
def import_file(filename):
    with open(filename, 'r') as readData:
        data = list(csv.reader(readData, quoting=csv.QUOTE_NONNUMERIC))
    return data
def write_to_excel(results, filename):
    sheet = 'advian_results'
    df = pd.DataFrame([results["RDAS"], results["RDPS"], results["RIAS"], results["RIPS"], results["CRI"], results["INT"], results["STA"], results["PRE"], results["DRI"], results["DRE"]],
                      index=["RDAS", "RDPS", "RIAS", "RIPS", "CRI", "INT", "STA", "PRE", "DRI", "DRE" ], columns=None)
    df.columns = pd.RangeIndex(1, len(df.columns)+1)
    df1 = pd.DataFrame([['', 'FACTORS', '']],
                      index=None, columns = ['', '', 'RESULTS OF ADVIAN ANALYSIS'])
    df2 = pd.DataFrame([['SYSTEM',''],['STABILITY', results["FINAL_STAB"]]],
                       index=None, columns = None)
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df1.to_excel(writer, sheet_name=sheet, index=False, )
        df.to_excel(writer, sheet_name=sheet, startrow = 2)
        df2.to_excel(writer, sheet_name=sheet, startrow=14, index=False, header = False)

def advian(input):
    n = len(input)
    print(n, " rows processing")
    #declare variables - a list for each
    param ={"DAS":[0]*n, "DPS":[0]*n, "IAS":[0]*n, "IPS":[0]*n, "RDAS":[0]*n, "RDPS":[0]*n, "RIAS":[0]*n, "RIPS":[0]*n, "CRI":[0]*n, "INT":[0]*n, "STA":[0]*n, "PRE":[0]*n, "DRI":[0]*n, "DRE":[0]*n, "FINAL_STAB":0}
    #normalize the input matrix
    mx = np.amax(input)
    input = np.divide(input,mx)
    # aux variables - check if i need them
    omega = np.copy(input) #order 1 matrix
    # calculate direct sums, initial indirect sums & maxima
    for i in range(n):
        for p in range(n):
            param["DAS"][i] = param["DAS"][i] + omega[i][p]
            param["DPS"][i] = param["DPS"][i] + omega[p][i]
            param["IAS"][i] = param["IAS"][i] + omega[i][p]
            param["IPS"][i] = param["IPS"][i] + omega[p][i]
        #calc max values
    MDS = max(np.amax(param["DAS"]), np.amax(param["DPS"]))
    MIS = MDS
    # calculate indirect sums
    for k in range (1,n-1):
        omega = np.matmul(input,omega)
        for i in range(n):
            for p in range(n):
                param["IAS"][i] = param["IAS"][i] + omega[i][p]
                param["IPS"][i] = param["IPS"][i] + omega[p][i]
            MIS = max(MIS, np.amax(param["IAS"]), np.amax(param["IPS"]))
    #calculate classification measures
    for i in range (n):
        param["RDAS"][i] = 100 * (param["DAS"][i] / MDS)
        param["RDPS"][i] = 100 * (param["DPS"][i] / MDS)
        param["RIAS"][i] = 100 * (param["IAS"][i] / MIS)
        param["RIPS"][i] = 100 * (param["IPS"][i] / MIS)
        param["CRI"][i] = math.sqrt(param["RIAS"][i] * param["RIPS"][i])
        param["INT"][i] = 0.5 * (param["RIAS"][i] + param["RIPS"][i])
        v1 = param["RIAS"][i]
        v2 = param["RIPS"][i]
        if v1==0 and v2==0:
            param["STA"][i] = 100
        elif v1 == 0:
            param["STA"][i] = abs(100 - (2 / (0 + (1 / param["RIPS"][i]))))
        elif v2 == 0:
            param["STA"][i] = abs(100 - (2 / ((1 / param["RIAS"][i]) + 0)))
        else:
            param["STA"][i] = abs(100 - (2/((1/param["RIAS"][i]) + (1/param["RIPS"][i]))))
        param["PRE"][i] = math.sqrt(param["CRI"][i] * param["RIAS"][i])
        param["DRI"][i] = math.sqrt((100-param["CRI"][i]) * param["RIAS"][i])
        param["DRE"][i] = math.sqrt((100 - param["CRI"][i]) * param["RIPS"][i])
    param["FINAL_STAB"] = mean(param["STA"])
    return param
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    input_filename = 'AI_data.csv'
    output_filename = 'AI_results.xlsx'
    arr = import_file(input_filename)
    results = advian(arr)
    write_to_excel(results, output_filename)