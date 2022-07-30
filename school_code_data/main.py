import requests
import json
import xmltodict
import pandas as pd
from openpyxl import load_workbook

def main():
    url1 = "https://open.neis.go.kr/hub/schoolInfo?ATPT_OFCDC_SC_CODE=B10&KEY=6d3acd88db854d2d87ffe7dfb817845f&pSize=1000&pIndex=1"
    url2 = "https://open.neis.go.kr/hub/schoolInfo?ATPT_OFCDC_SC_CODE=B10&KEY=6d3acd88db854d2d87ffe7dfb817845f&pSize=1000&pIndex=2"


    content1 = requests.get(url2).content
    dict1 = xmltodict.parse(content1)
    data_json1 = json.dumps(dict1['schoolInfo'])
    obj_json1 = json.loads(data_json1)

    file1 = open("./dataEx.json", "w+")
    file1.write(json.dumps(obj_json1['row']))

    data_frame = pd.DataFrame(obj_json1['row'])
    print(data_frame.count)

    df = pd.read_json("./dataEx.json")

    book = load_workbook("seoul_school_data.xlsx")
    writer = pd.ExcelWriter("seoul_school_data.xlsx")
    writer.book = book

    df.to_excel(writer, "Sheet2")
    writer.save()



if __name__ == "__main__":
    main()