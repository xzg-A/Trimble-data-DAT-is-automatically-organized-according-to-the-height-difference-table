import pandas as pd
import re,tqdm,os

# 获取高差表起点终点元组，构建列表
def get_high_difference(file):
    df = pd.read_excel(file)
    datalist = list(zip(df['起点'], df['终点']))
    return list(datalist)

# 获取111AAA里的原始数据列表，里面放字典，键为起点终点字符串，值为该起点终点的每行数据
def get_data_list(file):
    # 创建空字典
    data_dict = {}
    # 读取1111AAA.DAT文件
    with open(file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # print(lines)
    # 设置索引计数器和匹配模式
    count = 0
    pattern1 = re.compile(r'(.*)\|(.*)\|(.*)\|(.*)\|(.*)\|(.*)\|')
    pattern2 = re.compile(r'KD\d\s*(\S+)\s')


    for line in lines:
        if 'Start-Line' in line:
            # 匹配下一行起点字符串
            start1 = re.match(pattern1,lines[count + 1]).group(3)
            start2 = re.search(pattern2,start1).group(1).strip()
            start = start2
            start_index = count

        elif 'End-Line' in line:
            # 匹配上一行终点字符串
            end1 = re.match(pattern1,lines[count - 1]).group(3)
            end2 = re.search(pattern2,end1).group(1).strip()
            end = end2
            end_index = count + 1
            # 将起点终点字符串作为键，行数据作为值，存入字典
            data_dict[start+'--'+end] = lines[start_index:end_index]
        count += 1
        
    
    return data_dict

# 以往测为文件名，先存入往测数据，再存入返测数据
def write_to_file(qidian, zhongdian, file, data_dict):
    with open(file, 'w', encoding='utf-8') as f:
        # 分开添加数据
        for key, value in data_dict.items():
            # 先存入往测数据
            if key in file:
                for line in value:
                    f.write(line)
        for key, value in data_dict.items():
            # 再存入返测数据
            if (zhongdian + '--' + qidian) == key:
                for line in value:
                    f.write(line)
            

# 调试代码
# get_high_difference('高差表.xls')
# get_data_list('111AAA.DAT')

# 主函数
if __name__ == "__main__":
    # 获取高差表起点终点元组，构建列表
    high_difference = get_high_difference('高差表.xls')
    # 获取111AAA里的原始数据
    data_dict = get_data_list('111AAA.DAT')
    # 创建结果文件夹
    if os.path.exists('./结果') == False:
        os.mkdir('./结果')

    # 遍历高差表，逐个写入文件
    for i in tqdm.tqdm(range(1,len(high_difference)+1)):
        qidian = high_difference[i-1][0]
        zhongdian = high_difference[i-1][1]
        file = './结果/' + str(i).zfill(3) + '--' + str(qidian) + '--' + str(zhongdian) + '.DAT'
        write_to_file(qidian, zhongdian, file, data_dict)