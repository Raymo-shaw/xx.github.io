from flask import Flask, request, render_template
import pandas as pd
import io

app = Flask(__name__)

# 处理Excel文件
def process_excel(file):
    # 读取Excel文件
    df = pd.read_excel(file)
    
    # 删除指定的列
    columns_to_delete = [5, 8, 11, 12, 13, 14, 15, 16, 18]  # 注意索引从0开始，所以实际删除的是第6, 9, 12, 13, 14, 15, 16, 17, 19列
    df.drop(df.columns[columns_to_delete], axis=1, inplace=True)

    # 插入自定义列名
    new_columns = ['工序', '工序', '工单编号', '物料编码', '物料描述', '生产数量', '单位', '开始日期', '需求日期', '创建日期']
    df.columns = new_columns + df.columns[len(new_columns):].tolist()
    
    # 统计第二列中不同类型的出现次数
    type_counts = df.iloc[:, 1].value_counts()
    
    # 按照指定顺序排序
    type_order = ['注塑', '压铸', '静电', '喷油', '移印', '装配']
    type_counts = type_counts.reindex(type_order, fill_value=0)
    
    # 处理数据：添加空行和小计
    final_data = []
    total_count = 0
    for t in type_order:
        type_data = df[df.iloc[:, 1] == t]
        type_count = type_counts[t]
        
        # 添加类型分隔行
        final_data.append([''] * len(df.columns))
        final_data.append([f'{t} 小计 = {type_count}'] + [''] * (len(df.columns) - 1))
        
        # 添加类型数据
        final_data.extend(type_data.values.tolist())
        total_count += type_count
    
    # 添加总计
    final_data.append([''] * len(df.columns))
    final_data.append([f'工单合计 = {total_count}'] + [''] * (len(df.columns) - 1))

    return df.columns.tolist(), final_data, total_count

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    
    # 处理Excel文件
    columns, data, total_count = process_excel(file)
    
    # 将处理后的数据转换为HTML表格
    return render_template('result.html', columns=columns, data=data, total_count=total_count)

if __name__ == "__main__":
     app.run(debug=True)
