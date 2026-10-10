import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font,PatternFill

# 第一步：读第8周抓的数据
df = pd.read_excel("书店书单.xlsx")
df["价格"] = df["价格"].str.replace("£","").astype(float)

# 第二步：算统计
summary = pd.DataFrame({
    "指标":["总书数","均价","最贵","最便宜","中位数"],
    "数值":[
        len(df),
        round(df["价格"].mean(),2),
        df["价格"].max(),
        df["价格"].min(),
        round(df["价格"].median(),2)
    ]
})

bins = [0,20,30,40,50,60]
labels = ["0-20元","20-30元","30-40元","40-50元","50元以上"]
df["价格区间"] = pd.cut(df["价格"],bins=bins,labels=labels)
dist = df.groupby("价格区间")["书名"].count().reset_index()
dist.columns = ["价格区间","书的数量"]
print(dist)

# 第三步：导出到一个Excel，三个sheet
with pd.ExcelWriter("书店分析报告.xlsx") as writer:
    df.to_excel(writer,sheet_name="清洗后数据",index=False)
    summary.to_excel(writer,sheet_name="统计摘要",index=False)
    dist.to_excel(writer,sheet_name="价格分布",index=False)

path = "书店分析报告.xlsx"
wb = load_workbook(path)

header_font = Font(bold=True,color="FFFFFF")
header_fill = PatternFill("solid",fgColor="A3D5E8")

for ws in wb.worksheets:
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
    # 简单调列宽
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3,50)

wb.save(path)

print("报告已生成！")
print(summary)
input("按回车退出...")