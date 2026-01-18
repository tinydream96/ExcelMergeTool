import openpyxl
from openpyxl import Workbook

def create_demo_files():
    # 1. 创建目标表 (主表 - 需要填入数据的表)
    wb_target = Workbook()
    ws_target = wb_target.active
    ws_target.title = "2026年员工工资表"
    
    # 表头
    headers_target = ["工号", "姓名", "所属部门", "职位", "本月实发工资", "年终绩效评级"]
    ws_target.append(headers_target)
    
    # 数据 (工资和评级是空的，等待合并)
    data_target = [
        ["EMP001", "张三", "技术部", "后端工程师", None, None],
        ["EMP002", "李四", "产品部", "产品经理", None, None],
        ["EMP003", "王五", "市场部", "销售总监", None, None],
        ["EMP004", "赵六", "人事部", "HRBP", None, None],
        ["EMP005", "孙七", "技术部", "前端工程师", None, None],
    ]
    
    for row in data_target:
        ws_target.append(row)
        
    wb_target.save("demo_main_target.xlsx")
    print("已创建中文示例主表: demo_main_target.xlsx")

    # 2. 创建来源表 (数据源 - 包含工资信息的表)
    wb_source = Workbook()
    ws_source = wb_source.active
    ws_source.title = "财务部导出数据"
    
    # 表头 (故意使用不同的列名，测试模糊匹配)
    headers_source = ["员工编号", "银行卡号", "核算薪资", "考评结果", "备注"]
    ws_source.append(headers_source)
    
    # 数据
    data_source = [
        ["EMP001", "622202******1234", 18500, "A", "晋升"],
        ["EMP002", "622202******5678", 21000, "A+", "优秀员工"],
        ["EMP003", "622202******9012", 15000, "B", "无"],
        ["EMP004", "622202******3456", 12000, "A", "全勤"],
        ["EMP006", "622202******7890", 9000, "C", "新入职"], # 这是一个多余的数据，主表中没有
    ]
    
    for row in data_source:
        ws_source.append(row)
        
    wb_source.save("demo_data_source.xlsx")
    print("已创建中文示例来源表: demo_data_source.xlsx")

if __name__ == "__main__":
    create_demo_files()
