import xlwings as xw
import json
import datetime as dt
import sys, getopt, re

#保全操作类型字典
dic_Operator = {
    "新增": "①新增",
    "变更": "③变更",
    "终止": "②终止"
}
#家属关系字典
dic_Relations = {
    "计划四": "子女",
    "计划十三": "子女",
    "计划九": "配偶"
}
#证件类型字典
dic_cerType = {
    "中国": "居民身份证",
    "中国香港": "港澳通行证",
    "中国台湾": "台湾通行证"
}
#保障计划字典
dic_InsurancePlans = {
    "计划一": "0000273086-老员工（社保）",
    "计划二": "0000273087-新员工（社保）",
    "计划三": "0000273088-出差新员工（社保）",
    "计划四": "0000273089-子女",
    "计划五": "0000273090-实习生计划无社保",
    "计划六": "0000273091-出差老员工计划",
    "计划七": "0000273092-出差实习生计划（无社保）",
    "计划八": "0000273093-新员工计划（无社保）",
    "计划九": "0000273094-配偶计划",
    "计划十一": "0000273096-管培生计划（无社保）",
    "计划十二": "0000273097-无社保外籍老员工计划",
    "计划十三": "0000273098-老员工子女",
    "计划十五": "0000273100-高管计划-2（社保）",
    "计划十四": "0000273099-员工计划-2",
    "计划十": "0000273095-无社保高管计划"
}
#分支机构字典
dic_Branches = {
    "慧博": "H00065086-慧博云通科技股份有限公司",
    "元年": "0043216999-北京慧博元年科技有限公司",
    "慧海青岛分": "0043218340-杭州慧海软件技术有限公司青岛分公司"
}
#职业类型字典
dic_JobTypes = {
    "员工": "一般职业-机关团体公司-机关内勤(不从事危险工作)-1",
    "配偶": "一般职业-机关团体公司-机关内勤(不从事危险工作)-1",
    "子女": "文教机构-教育机构-学生(PB等级以投保人职业代码的WP等级为准)-2",
    "儿童": "服务业-家政管理-儿童(PB等级以投保人职业代码的WP等级为准)-2"
}

#保全员工表数据列
dic_EmployeeCols = {
    "保全类型": "A",
    "变更项目": "B",
    "分支机构": "C",
    "员工姓名": "D",
    "证件类型": "E",
    "证件号码": "F",
    "职业类型": "L",
    "保障计划": "M",
    "生效日期": "N",
    "是否医保": "O",
    "国籍": "S" 
}

#保全家属表数据列
dic_FamilyCols = {
    "保全类型": "A",
    "变更项目": "B",
    "分支机构": "C",
    "员工姓名": "D",
    "员工证件类型": "E",
    "员工证件号码": "F",
    "家属姓名": "G",
    "家属证件类型": "H",
    "家属证件号码": "I",
    "职业类型": "O",
    "保障计划": "P",
    "家属关系": "Q",
    "生效日期": "R",
    "是否医保": "S",
    "国籍": "W" 
}

#保全员工表需要转义的列
dic_tranEmployeeCols = {
    'A': dic_Operator,     #保全类型
    'B': '计划',           #变更项目
    'C': dic_Branches,     #分支机构
    'E': dic_cerType,      #证件类型
    'L': dic_JobTypes,     #职业类型
    'M': dic_InsurancePlans,   #保障计划
    'O': '是',                 #是否医保
    'S': '中国'                #国籍 
}

#保全家属表需要转义的列
dic_tranFamilyCols = {
    'A': dic_Operator,    #保全类型
    'B': '计划',           #变更项目
    'C': dic_Branches,    #分支机构
    'E': dic_cerType,     #员工证件类型
    'H': dic_cerType,     #家属证件类型
    'O': dic_JobTypes,    #家属职业类型
    'P': dic_InsurancePlans,  #保障计划
    'Q': dic_Relations,       #家属关系
    'S': '否',                #是否医保
    'W': '中国'               #国籍  
}

def logsFunc(string, isExcep=0):
    if isExcep: 
        print(f"!!!Exception: {string}")
        return
    with open('log.txt', 'a') as f:
        f.write(f"{string}\n")

def cp_excel_column( shtR, colR, shtW, colW, row=0):
    nums = shtR.range(f'{colR}2').expand('down').count
    shtR.range(f'{colR}2:{colR}{nums+1}').copy()
    shtW.range(f'{colW}{row+2}:{colW}{row+nums+1}').paste(paste='values_and_number_formats',skip_blanks=True)

def read_excel_colum(sheet, column, nums=0):
    try:
        if nums: return sheet.range(f'{column}2:{column}{nums+1}').value
        return sheet.range(f'{column}2').expand('down').value
    except Exception as e:
        logsFunc(f"Error reading Column {sheet.name}/{sheet['{column}1'].value}: {str(e)}", 1)

def write_excel_colum(sheet, data, column, row=0):
    try:
        sheet.range(f'{column}{row+2}').options(transpose=True).value = data
    except Exception as e:
        logsFunc(f"Error writing Column {sheet.name}/{sheet['{column}1'].value}: {str(e)}", 1)

def GetPlanByKeyName(value):
    keys = dic_InsurancePlans.keys()
    for key in keys:
        if re.search(r'^'+key, value):
            return dic_InsurancePlans[key]
    return value

def GetRelationByKeyName(value):
    keys = dic_Relations.keys()
    for key in keys:
        if re.search(r'^'+key, value):
            return dic_Relations[key]
    return value

#表头关键字搜索安全函数
def GetColByKeyName(head, keyname):
    h = []
    if keyname == '姓名':
        h = [h for h in head if re.search(r'^((?!子女).)*姓名', h)]
    elif keyname == "子女姓名":
        h = [h for h in head if re.search(r'[子女].*姓名', h)]
    elif keyname == '身份证号':
        h = [h for h in head if re.search(r'^((?!子女).)*身份', h)]
    elif keyname == '子女身份证号':
        h = [h for h in head if re.search(r'[子女].*身份', h)]
    elif keyname == '分支机构':
        h = [h for h in head if re.search(r'分支', h)]
    elif keyname == '生效日期':
        h = [h for h in head if re.search(r'日期', h)]
    elif keyname == '保障计划':
        h = [h for h in head if re.search(r'^((?!新).)*计划', h)]
    elif keyname == '新保障计划':
        h = [h for h in head if re.search(r'[新].*计划', h)]
    else:
        if keyname in head:
            return chr(ord('A')+head.index(keyname))
    if len(h):
        return chr(ord('A')+head.index(h[0]))
    
    logsFunc(f"-------- 原数据表中未找到 [{keyname}] 列!")
    return None

#员工信息转义函数
def apply_transforEmployee(data, key, operator=0):
    try:
        col = dic_EmployeeCols[key]
        if isinstance(data, int):   # 如果传入的是数字，则该列都是同一内容
            if isinstance(dic_tranEmployeeCols[col],dict):
                return [dic_tranEmployeeCols[col].get(operator,operator)] * data
            else:
                return [dic_tranEmployeeCols.get(col, col)] * data
        else:    #如果传入的是数组，则每个数组项都要转义
            assert data, f"{key} 列读取数据为None!"
            if col == 'E':     #证件类型
                return [dic_tranEmployeeCols["E"].get(value, '外国护照') for value in data]
            elif col == 'M':   # 保障计划
                return [GetPlanByKeyName(value) for value in data]
            elif col == 'C':   #分支机构
                d = []
                for b in data:
                    if b:
                        d.append(dic_Branches.get(b, b))
                    else:
                        d.append(dic_Branches.get('慧博'))
                return d
            else: return [dic_tranEmployeeCols[col].get(value, value) for value in data]
    except Exception as e:
        logsFunc(f"Error in apply_transforEmployee: {str(e)}", 1)
        return None

#家属信息转义函数
def apply_transforFamily(data, key, operator=0):
    try:
        col = dic_FamilyCols[key]
        if isinstance(data, int):
            if isinstance(dic_tranFamilyCols[col],dict):
                return [dic_tranFamilyCols[col].get(operator, operator)] * data
            else:
                return [dic_tranFamilyCols.get(col, col)] * data
        else:
            assert data, f"{key} 列读取数据为None!"
            if col in ('E', 'H'):
                return [dic_tranFamilyCols[col].get(value, '外国护照') for value in data]
            elif col == 'P':   # 保障计划
                return [GetPlanByKeyName(value) for value in data]
            elif col == 'Q':   # 家属关系
                return [GetRelationByKeyName(value) for value in data]
            elif col == 'C':   #分支机构
                d = []
                for b in data:
                    if b:
                        d.append(dic_Branches.get(b, b))
                    else:
                        d.append(dic_Branches.get('慧博'))
                return d
            else: return [dic_tranFamilyCols[col].get(value, value) for value in data]
    except Exception as e:
        logsFunc(f"Error in apply_transforFamily: {str(e)}", 1)
        return None

#复制新增员工数据项
def cp_NewEmployeeData(shtR, shtW, headR, rowW=0):
    try:
        #namedata = read_excel_colum(shtR, dic_shtcolrec.get(f'{headR.index("姓名")}'))
        namedata = read_excel_colum(shtR, GetColByKeyName(headR, "姓名"))
        nums = len(namedata)
        assert nums > 0, "本期没有新增员工信息！"

        write_excel_colum(shtW, namedata, dic_EmployeeCols["员工姓名"], rowW)
        write_excel_colum(shtW, apply_transforEmployee(nums, "保全类型", "新增"), dic_EmployeeCols["保全类型"], rowW)
        branchcol = GetColByKeyName(headR, "分支机构")
        if branchcol:
            branchdata = read_excel_colum(shtR, branchcol, nums)
            write_excel_colum(shtW, apply_transforEmployee(branchdata, "分支机构"), dic_EmployeeCols["分支机构"], rowW)
        else:
            write_excel_colum(shtW, apply_transforEmployee(nums, "分支机构", "慧博"), dic_EmployeeCols["分支机构"], rowW)
        regiondata = read_excel_colum(shtR, GetColByKeyName(headR, "国籍"))
        write_excel_colum(shtW, regiondata, dic_EmployeeCols["国籍"], rowW)
        write_excel_colum(shtW, apply_transforEmployee(regiondata, "证件类型"), dic_EmployeeCols["证件类型"], rowW)
        iddata = read_excel_colum(shtR, GetColByKeyName(headR, "身份证号"))
        write_excel_colum(shtW, iddata, dic_EmployeeCols.get("证件号码"), rowW)
        write_excel_colum(shtW, apply_transforEmployee(nums, "职业类型", "员工"), dic_EmployeeCols["职业类型"], rowW)
        plandata = read_excel_colum(shtR, GetColByKeyName(headR, "保障计划"))
        write_excel_colum(shtW, apply_transforEmployee(plandata, "保障计划"), dic_EmployeeCols["保障计划"], rowW)
        ondate = read_excel_colum(shtR, GetColByKeyName(headR, "生效日期"))
        write_excel_colum(shtW, ondate, dic_EmployeeCols.get("生效日期"), rowW)
        write_excel_colum(shtW, apply_transforEmployee(nums, "是否医保"), dic_EmployeeCols["是否医保"], rowW)
            
        logsFunc(f"..... [{shtR.book.name}]->[{shtR.name}] Copy To [{shtW.book.name}]->[{shtW.name}] Done! Nums: {nums}")
        return nums

    except Exception as e:
        logsFunc(f"Error in cp_NewEmployeeData: {str(e)}", 1)
        return 0

#复制新增家属数据项
def cp_NewFamilyData(shtR, shtW, headR, rowW=0):
    try:
        namedata = read_excel_colum(shtR, GetColByKeyName(headR, "姓名"))
        nums = len(namedata)
        assert nums > 0, "本期没有新增员工家属信息！"

        iddata = read_excel_colum(shtR, GetColByKeyName(headR, "身份证号"))
        fnamedata = read_excel_colum(shtR, GetColByKeyName(headR, "子女姓名"))
        fiddata = read_excel_colum(shtR, GetColByKeyName(headR, "子女身份证号"))
        plandata = read_excel_colum(shtR, GetColByKeyName(headR, "保障计划"))
        ondate = read_excel_colum(shtR, GetColByKeyName(headR, "生效日期"))
        regioncol = GetColByKeyName(headR, "国籍")
        if regioncol: 
            regiondata = read_excel_colum(shtR, regioncol)
            write_excel_colum(shtW, apply_transforFamily(regiondata, "员工证件类型"), dic_FamilyCols["员工证件类型"], rowW)
            write_excel_colum(shtW, apply_transforFamily(regiondata, "家属证件类型"), dic_FamilyCols["家属证件类型"], rowW)
            write_excel_colum(shtW, regiondata, dic_FamilyCols.get("国籍"), rowW)
        else:
            write_excel_colum(shtW, apply_transforFamily(nums, "员工证件类型", "中国"), dic_FamilyCols["员工证件类型"], rowW)
            write_excel_colum(shtW, apply_transforFamily(nums, "家属证件类型", "中国"), dic_FamilyCols["家属证件类型"], rowW)
            write_excel_colum(shtW, apply_transforFamily(nums, "国籍"), dic_FamilyCols["国籍"], rowW)
        
        write_excel_colum(shtW, apply_transforFamily(nums, "保全类型", "新增"), dic_FamilyCols["保全类型"], rowW)
        branchcol = GetColByKeyName(headR, "分支机构")
        if branchcol:
            branchdata = read_excel_colum(shtR, branchcol, nums)
            write_excel_colum(shtW, apply_transforFamily(branchdata, "分支机构"), dic_FamilyCols["分支机构"], rowW)
        else:
            write_excel_colum(shtW, apply_transforFamily(nums, "分支机构", "慧博"), dic_FamilyCols["分支机构"], rowW)
        
        write_excel_colum(shtW, namedata, dic_FamilyCols["员工姓名"], rowW)
        write_excel_colum(shtW, iddata, dic_FamilyCols["员工证件号码"], rowW)
        write_excel_colum(shtW, fnamedata, dic_FamilyCols["家属姓名"], rowW) 
        write_excel_colum(shtW, fiddata, dic_FamilyCols["家属证件号码"], rowW)
        write_excel_colum(shtW, apply_transforFamily(nums, "职业类型", "子女"), dic_FamilyCols["职业类型"], rowW)
        write_excel_colum(shtW, apply_transforFamily(plandata, "保障计划"), dic_FamilyCols["保障计划"], rowW)
        write_excel_colum(shtW, apply_transforFamily(plandata, "家属关系"), dic_FamilyCols["家属关系"], rowW)
        write_excel_colum(shtW, ondate, dic_FamilyCols["生效日期"], rowW)
        write_excel_colum(shtW, apply_transforFamily(nums, "是否医保"), dic_FamilyCols["是否医保"], rowW)
            
        logsFunc(f"..... [{shtR.book.name}]->[{shtR.name}] Copy To [{shtW.book.name}]->[{shtW.name}] Done! Nums: {nums}")
        return nums
        
    except Exception as e:
        logsFunc(f"Error in cp_FamilyData: {str(e)}", 1)
        return 0

#复制减员员工数据项
def cp_DelEmployeeData(shtR, shtW, headR, rowW=0):
    try:
        namedata = read_excel_colum(shtR, GetColByKeyName(headR, "姓名"))
        nums = len(namedata)
        assert nums > 0, "本期没有终止员工信息"

        iddata = read_excel_colum(shtR, GetColByKeyName(headR, "身份证号"))
        offdate = read_excel_colum(shtR, GetColByKeyName(headR, "生效日期"))
        regioncol = GetColByKeyName(headR, "国籍")
        if regioncol: 
            regiondata = read_excel_colum(shtR, regioncol)
            write_excel_colum(shtW, apply_transforEmployee(regiondata, "证件类型"), dic_EmployeeCols["证件类型"], rowW)
        else:
            write_excel_colum(shtW, apply_transforEmployee(nums, "证件类型", "中国"), dic_EmployeeCols["证件类型"], rowW)
        
        write_excel_colum(shtW, apply_transforEmployee(nums, "保全类型", "终止"), dic_EmployeeCols["保全类型"], rowW)
        write_excel_colum(shtW, namedata, dic_EmployeeCols["员工姓名"], rowW)
        write_excel_colum(shtW, iddata, dic_EmployeeCols["证件号码"], rowW)
        write_excel_colum(shtW, offdate, dic_EmployeeCols["生效日期"], rowW)

        logsFunc(f"..... [{shtR.book.name}]->[{shtR.name}] Copy To [{shtW.book.name}]->[{shtW.name}] Done! Nums: {nums}")
        return nums
        
    except Exception as e:
        logsFunc(f"Error in cp_DelEmployeeData: {str(e)}", 1)
        return 0

#复制变更员工数据项
def cp_ChgEmployeeData(shtR, shtW, headR, rowW=0):
    try:
        namedata = read_excel_colum(shtR, GetColByKeyName(headR, "姓名"))
        nums = len(namedata)
        assert nums > 0, "本期没有变更员工信息"

        iddata = read_excel_colum(shtR, GetColByKeyName(headR, "身份证号"))
        plandata = read_excel_colum(shtR, GetColByKeyName(headR, "新保障计划"))
        ondate = read_excel_colum(shtR, GetColByKeyName(headR, "生效日期"))
        regioncol = GetColByKeyName(headR, "国籍")
        if regioncol: 
            regiondata = read_excel_colum(shtR, regioncol)
            write_excel_colum(shtW, apply_transforEmployee(regiondata, "证件类型"), dic_EmployeeCols["证件类型"], rowW)
        else:
            write_excel_colum(shtW, apply_transforEmployee(nums, "证件类型", "中国"), dic_EmployeeCols["证件类型"], rowW)
        
        write_excel_colum(shtW, apply_transforEmployee(nums, "保全类型", "变更"), dic_EmployeeCols["保全类型"], rowW)
        write_excel_colum(shtW, apply_transforEmployee(nums, "变更项目", "计划"), dic_EmployeeCols["变更项目"], rowW)
        #write_excel_colum(shtW, apply_transforEmployee(nums, "分支机构", "慧博"), dic_EmployeeCols.get("分支机构"), rowW)
        write_excel_colum(shtW, namedata, dic_EmployeeCols["员工姓名"], rowW)
        write_excel_colum(shtW, iddata, dic_EmployeeCols["证件号码"], rowW)
        write_excel_colum(shtW, apply_transforEmployee(plandata, "保障计划"), dic_EmployeeCols["保障计划"], rowW)
        write_excel_colum(shtW, ondate, dic_EmployeeCols["生效日期"], rowW)
        #write_excel_colum(shtW, apply_transforEmployee(nums, "是否医保", "是"), dic_EmployeeCols["是否医保"], rowW)
        logsFunc(f"..... [{shtR.book.name}]->[{shtR.name}] Copy To [{shtW.book.name}]->[{shtW.name}] Done! Nums: {nums}")
        return nums
    except Exception as e:
        logsFunc(f"Error in cp_ChgEmployeeData: {str(e)}", 1)
        return 0

def main(filepathR='', filepathW=''):
    try:
        app = xw.App(visible=False)
        if filepathR.endswith('.xlsx'):
            wbRead = xw.Book(filepathR)
        assert "wbRead" in vars() , "readfile args failed!"
        if filepathW.endswith('.xls'):
            wbWrite = xw.Book(filepathW)
        assert "wbWrite" in vars() , 'writefile args failed!'
        
        nowdt = dt.datetime.now()
        starttime = nowdt.timestamp()
        logsFunc("--------------------------------------------------------")
        logsFunc(f"本次启动转换开始时间: {nowdt.strftime('%Y-%m-%d %H:%M:%S')}")
        logsFunc(f">> 载入原始文件: {wbRead.name}, 载入目标文件: {wbWrite.name}")
        
        shtW_Empolyee = wbWrite.sheets[0]
        shtW_Family = wbWrite.sheets[1]
        rowWEmployee = 0
        rowWFamily = 0
        for shtR in wbRead.sheets:
            rcode = 0
            headR = shtR.range('A1').expand('right').value
            logsFunc(f">>> 开始处理原始表->[{shtR.name}]:\n      表头: {headR}")
            if shtR.name in ["新增", "增员"]:
                rcode = cp_NewEmployeeData(shtR, shtW_Empolyee, headR, rowWEmployee)
                rowWEmployee += rcode
            elif shtR.name in ["减少", "减员", "离职"]:
                rcode = cp_DelEmployeeData(shtR, shtW_Empolyee, headR, rowWEmployee)
                rowWEmployee += rcode
            elif shtR.name in ["变更"]:
                rcode = cp_ChgEmployeeData(shtR, shtW_Empolyee, headR, rowWEmployee)
                rowWEmployee += rcode
            elif shtR.name in ["子女"]:
                rcode = cp_NewFamilyData(shtR, shtW_Family, headR, rowWFamily)
                rowWFamily += rcode
            elif shtR.name in ["配偶"]:
                rcode = cp_NewFamilyData(shtR, shtW_Family, headR, rowWFamily)
                rowWFamily += rcode
            else:
                logsFunc(f"-------- !!!No Rules for Sheet: {shtR.name}")
            
            if rcode: wbWrite.save()
            logsFunc(f">>> 处理原始表->[{shtR.name}] Done!")

        wbRead.close()
        wbWrite.close()
        app.quit()
        endtime = dt.datetime.now().timestamp()
        logsFunc(f"本次转换结束，耗时：{endtime - starttime}s")
        
        print("Work done!")
        print(f"本次转换结束，耗时：{endtime - starttime}s")
    
    except Exception as e:
        if "wbRead" in vars(): wbRead.close()
        if "wbWrite" in vars(): wbWrite.close()
        app.quit()
        logsFunc(f"Error in main function: {str(e)}", 1)
    

if __name__ == "__main__":
    filepathR = ''
    filepathW = ''
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hr:w:", ["rfile=","wfile="])
        assert len(opts), "Need args!\nUseage: ai.py -r <readfile.xlsx> -w <wirtefile.xls>'"
        for opt, arg in opts:
            print(opt, arg)
            if opt == '-h':
                print ('Useage: ai.py -r <readfile.xlsx> -w <wirtefile.xls>')
                sys.exit()
            elif opt in ('-r', '--rfile'):
                filepathR = arg
            elif opt in ('-w', '--wfile'):
                filepathW = arg
        main(filepathR, filepathW)
    except getopt.GetoptError:
        print ('Wrong args!!!\nUseage: ai.py -r <readfile.xlsx> -w <wirtefile.xls>')
    except Exception as e:
        print (f"{str(e)}")