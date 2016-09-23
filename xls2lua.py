# -*- coding: utf-8 -*
# author: zfengzhen
# branched by: legendxcheng

import xlrd
import os.path
import time

import sys
reload(sys) # Python2.5 初始化后会删除 sys.setdefaultencoding 这个方法，我们需要重新载入
sys.setdefaultencoding('utf-8')

SCRIPT_HEAD = "-- this file is auto-generated!\n\
-- don't change it manaully.\n\
-- source file: %s\n\
-- created at: %s\n\
\n\
\n\
"
SCRIPT_END = "\n\
for i=1, #(%s.all_type) do\n\
    local item = %s.all_type[i]\n\
    for j=1, #item do\n\
        item[j].__index = item[j]\n\
        if j < #item then\n\
            setmetatable(item[j+1], item[j])\n\
        end\n\
    end\n\
end\n\
\n\
\n\
"

elog = None

# 检查单元格内容类型是否正确 todo
def checkValueType(target_type, value):
    type_wrong = False
    if target_type == "string":
        pass
    elif target_type == "int":
        pass
    elif target_type == "float":
        pass
    elif target_type == "boolean":
        pass
    elif target_type == "table":
        pass
    elif target_type.endswith("_list"):
        pass

    type_wrong = True
    return type_wrong

def errorHalt():
    elog.close();


def get_cell_value(config_type, cell_type, value):
    v = None
    if config_type == "int" and cell_type == 2:
        v = int(value)
    elif config_type == "string" and cell_type == 1:
        v = value
    elif config_type == "boolean" and cell_type == 4:
        if str(value).lower() == "true":
            v = "true"
        else:
            v = "false"
    elif config_type == "float" and cell_type == 2:
        v = value
    elif config_type.endswith("_list") and cell_type == 1:
        v = "{" + value + "}"
    elif config_type == "table":
        v = value

    return v

def make_table(filename):

    logPrint("Now exporting " + filename)

    if not os.path.isfile(filename):
        raise NameError, "%s 不是一个合法的文件名" % filename
    book_xlrd = xlrd.open_workbook(filename,formatting_info=True)

    excel = {}
    excel["filename"] = filename
    excel["data"] = {}
    excel["meta"] = {}

    #  读入配置信息
    data_sheet_names = {}
    sheet = book_xlrd.sheet_by_name("config")
    for i in range(1, sheet.nrows):
        tag1 = sheet.cell_value(i, 0)
        tag2 = sheet.cell_value(i, 1)

        if not data_sheet_names.has_key(tag1):
            data_sheet_names[tag1] = tag2
        else :
            logPrint("要求导出两个或以上同样的Table名称:" + tag1)
            errorHalt()
            return

    for sheet_name in data_sheet_names:
        table_name = data_sheet_names[sheet_name]
        sheet = book_xlrd.sheet_by_name(sheet_name)
        excel["data"][sheet_name] = {}
        excel["meta"][sheet_name] = {}
        excel["meta"][sheet_name]["table"] = table_name

        # 必须大于2行2列
        if sheet.ncols <= 2 or sheet.nrows <= 2:
            return {}, -1, "sheet[" + sheet_name + "]" + "行与列数应该至少 > 2"

        # 解析标题
        title = {}
        col_idx = 0
        for col_idx in xrange(sheet.ncols):
            value = sheet.cell_value(0, col_idx)
            vtype = sheet.cell_type(0, col_idx)
            if vtype != 1:
                return {}, -1, "title columns[" + str(col_idx) + "] 应该是 string"
            title[col_idx] = str(value).replace(" ", "_")

        if title[0] != "id":
            return {}, -1, "sheet[" + sheet_name + "]" + " 第一列的标题应该是 [id]"
        elif title[1] != "name":
            return {}, -1, "sheet[" + sheet_name + "]" + " 第二列的标题应该是 [name]"

        excel["meta"][sheet_name]["title"] = title

        # 类型解析
        type_dict = {}
        col_idx = 0
        row_idx = 1
        for col_idx in xrange(sheet.ncols):
            value = sheet.cell_value(1, col_idx)
            vtype = sheet.cell_type(1, col_idx)
            type_dict[col_idx] = str(value)
            if (not type_dict[col_idx].lower() in ["string", "int", "boolean", "float", "string_list", "int_list", "boolean_list", "float_list", "table"]):
                return {}, -1, "sheet[" + sheet_name + "]" + \
                    " row[" + str(row_idx) + "] column[" + str(col_idx) + \
                    "] type  只能为 string, int, boolean, float, string_list, int_list, boolean_list, float_list, table"

        if type_dict[0].lower() != "int":
            return {}, -1,"sheet[" + sheet_name + "]" + " 第一列的类型必须为 [int]"
        elif type_dict[1].lower() != "string":
            return {}, -1, "sheet[" + sheet_name + "]" + " 第二列的类型必须为 [string]"

        excel["meta"][sheet_name]["type"] = type_dict
        # 第3行为解释数据，不导表

        # 数据解析
        row_idx = 3
        for row_idx in xrange(3, sheet.nrows):
            item_name = sheet.cell_value(row_idx, 1)
            if (item_name != None and item_name != ""):
                tid = int(sheet.cell_value(row_idx, 0))
                excel["data"][sheet_name][tid] = {}
                excel["meta"][sheet_name][item_name] = tid

        row_idx = 3
        # 数据从第2行开始
        for row_idx in xrange(3, sheet.nrows):
            row = {}

            # 名字,如果没有采用上一行数据
            item_name = sheet.cell_value(row_idx, 1)
            item_name_type = sheet.cell_type(row_idx, 1)
            if (item_name == None or item_name == "" or item_name_type == 0):
                return {}, -1,"sheet[" + sheet_name + "] 第" + str(row_idx + 1) + "行的name不可为空"

            # id, 如果没有采用上一行数据
            item_id = sheet.cell_value(row_idx, 0)
            item_id_type = sheet.cell_type(row_idx, 0)
            col_idx = 0
            for col_idx in xrange(sheet.ncols):
                value = sheet.cell_value(row_idx, col_idx)
                vtype = sheet.cell_type(row_idx, col_idx)
                # 本行有数据
                v = None
                v = get_cell_value(type_dict[col_idx].lower(), vtype, value)


                # 如果值为空,采用上一行数据
                if v is not None and value != "":
                    row[col_idx] = v
                #elif item_id != pre_item_id:
                #    pre_item_id = item_id
                #elif (item_id == pre_item_id):
                #    row[col_idx] = pre_row[col_idx]

            if row:
                item_idx = excel["meta"][sheet_name][item_name]
                excel["data"][sheet_name][item_idx] = row
    return excel, 0 , "ok"

def format_str(v):
    s = ("%s"%(v)).encode("gbk")
    if s[-1] == "]":
        s = "%s "%(s)
    return s




def write_to_lua_script(excel, output_root):

    def write_value(value_type, col_idx, title, vv):
        ret = ""
        if value_type == "int":
            ret = "\t" + str(title[col_idx] + " = " + str(vv) + ",\n")
        elif value_type == "boolean":
            ret = "\t" + str(title[col_idx] + " = \"" + str(vv) + "\",\n")
        elif value_type == "string":
            ret = "\t" + str(title[col_idx] + " = \"" + str(vv) + "\",\n")
        elif value_type == "float":
            ret = "\t" + str(title[col_idx] + " = " + str(vv) + ",\n")
        elif value_type.endswith("_list"):
            ret = "\t" + str(title[col_idx] + " = " + str(vv) + ",\n")
        elif value_type == "table":
            ret = "\t" + str(title[col_idx] + " = " + str(vv) + ",\n")
        return ret
    #getattr(self, "a_%s" % str(type(b)))
    #def a_int(self)
    for (sheet_name, sheet) in excel["data"].items():
        table_name = excel["meta"][sheet_name]["table"]
        outfp = open(output_root + "/" + excel["meta"][sheet_name]["table"] + ".lua", 'w')
        create_time = time.strftime("%a %b %d %H:%M:%S %Y", time.gmtime(time.time()))
        outfp.write(SCRIPT_HEAD % (excel["filename"], create_time))
        outfp.write("local " + table_name + " = {}\n")
        outfp.write("\n\n")

        title = excel["meta"][sheet_name]["title"]
        type_dict= excel["meta"][sheet_name]["type"]

        # 写入所有的类型
        '''outfp.write(sheet_name + ".type_map = {}\n")
        outfp.write("local type_map = " + sheet_name + ".type_map\n")
        for (item_name, item) in sheet.items():
            # 第一个数据会存在
            outfp.write("type_map[" + format_str(item[0][0]) + "] = \"" \
                + format_str(item[0][1]).replace(" ", "_")  + "s\"\n")
            outfp.write("type_map[\"" + format_str(item[0][1]).replace(" ", "_") + "s\"] = " \
                + format_str(item[0][0]) + "\n")

        outfp.write("\n\n")
        '''
        for (item_id, item) in sheet.items():
            # 写入数据
            #real_item_name = format_str(item_name).replace(" ", "_")
            for (col_idx, vv) in item.items():
                if col_idx == 0:
                    # outfp.write(sheet_name + "[" + str(item_id) + "] = {}\n")
                    #outfp.write("local " + real_item_name + " = " + sheet_name + "." + real_item_name + "\n\n")
                    outfp.write(table_name + "[" + str(item_id) + "] = {\n")
                outfp.write(write_value(type_dict[col_idx], col_idx, title, vv))

            outfp.write("}\n\n")


        outfp.write(table_name + ".id_by_name= {}\n")
        outfp.write("local id_by_name = " + table_name + ".id_by_name\n")

        for (item_name, item) in sheet.items():
            outfp.write("id_by_name[\"" + str(item[1]) + "\"] = "\
                + format_str(item_name).replace(" ", "_")  + "\n")

        outfp.write("\n\n")
        outfp.write("return " + table_name)
        outfp.write("\n\n")
        outfp.close()

def logPrint(ss):
    global elog
    elog.write(ss + "\n")
    print ss

def main():
    import sys
    global elog
    if len(sys.argv) < 3:
        sys.exit('''usage: xls2lua.py excel_name output_path)''')
    xls_dir = sys.argv[1]
    export_dir = sys.argv[2]

    elog = open("log.txt", "w")

    for root, dirs, files in os.walk(xls_dir):
        for filename in files:
            if not filename.endswith(".xls") :
                continue
            t, ret, errstr = make_table(root + "/" + filename)
            if ret != 0:
                logPrint("error: " + errstr)
            else:
                #logPrint("res:")
                logPrint(errstr)
                logPrint("success!!!")
                write_to_lua_script(t, export_dir)
    elog.close()
if __name__=="__main__":
    main()
