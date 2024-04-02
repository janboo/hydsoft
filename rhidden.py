#-- coding: utf-8 --

import xlwings as xw
import datetime as dt

try:
    print(f"当前时间: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    app = xw.App(visible=False)
    wbR = xw.Book('pos-service-24.xls')
    #shtname = []
    #for s in wbR.sheets:
    #    shtname.append(s.name)
    #print(shtname)

#'员工（新增 终止 资料 变更)申请表', '家属（新增 终止 资料 变更)申请表', '省', '省市', '市县', '市县2', 
#'Options', 'hidden02', 'hidden011', 'hidden012', 'hidden12', 'hidden114', 'hidden115', 'hidden116'
    dic_Sheethidden = {
        "E_Branchs": 7,
        "E_JobTypes": 8,
        "E_Plans": 9,
        "F_Branchs": 10,
        "F_JobTypes": 11,
        "F_Plans": 12,
        "F_Relations": 13
    }

    shtR = wbR.sheets[dic_Sheethidden['E_Plans']]
    column = 'A'
    nums = shtR.range(f'{column}1').expand('down').count
    #raise Exception('Debug Exit!')
    if nums:
        print(f"{column}1:{column}{nums}")
        print(shtR.range(f"{column}1:{column}{nums}").value)

    wbR.close()
    app.quit()

except Exception as e:
    wbR.close()
    app.quit()
    print(f"Error: {str(e)}\n")