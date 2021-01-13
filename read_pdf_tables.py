import camelot.io as camlot
import pandas as pd
from PyPDF2 import PdfFileReader as pdf_read
import os

#提取当前页中所有表格每一行的信息,空值进行删除
def camlot_method(pdf_path,page=7):
    lines=[]
    tables=camlot.read_pdf(pdf_path,pages=str(page),flavor='stream')
    for i in range(len(tables)):
        temp_tab=tables[i]
        datas=temp_tab.data
        for j in range(len(datas)):
            temp_list = [x.strip() for x in datas[j] if x.strip() != '']
            lines.append(temp_list)
    return lines

def read_table(file_path,page,keywords):
    '''
    :param file_path: 当前pdf路径
    :param page:根据书签获取到pdf的页码
    :param keywords 是将表格中的数据提取后根据关键词及进行提取，这边默认的是总收入和税后收入(根据不同的上市公司的进行修改对应的内容)
    :return:
    '''
    general=0
    general_percent=''
    pure=0
    pures_percent=''
    lines=camlot_method(file_path,page)
    for line in lines:
        try:
            if general ==0 or pure ==0:
                for k in range(len(line)):
                    if line[k] == keywords[0] and general == 0:
                        general = int(line[k+1].replace(',', ''))
                        if len(line) == 6:
                            general_percent = line[5]
                        elif len(line) == 4:
                            general_percent = line[3]
                        else:
                            general_percent = line[-1]
                    if line[k] == keywords[1] and pure == 0:
                        pure = int(line[k+1].replace(',', ''))
                        if len(line) == 4:
                            pures_percent = line[3]
                        else:
                            pures_percent = lines[-1]
        except:
            continue

    #返回结果
    return general,general_percent,pure,pures_percent


#书签信息的读取
def _setup_page_id_to_num(pdf,pages=None,_result=None,_num_pages=None):
    if _result is None:
        _result = {}
    if pages is None:
        _num_pages = []
        pages = pdf.trailer["/Root"].getObject()["/Pages"].getObject()
    t = pages["/Type"]
    if t == "/Pages":
        for page in pages["/Kids"]:
            _result[page.idnum] = len(_num_pages)
            _setup_page_id_to_num(pdf,page.getObject(),_result,_num_pages)
    elif t == "/Page":
        _num_pages.append(1)
    return _result

#将数据写入表格中，这边使用的是pandas的方法，也可以使用其他的pyhton包 xlws等
def write_excel(years,generals,general_percents,pures,pures_percents,excel_path):
    if(len(years)==len(generals) and len(pures)==len(generals)):
        df={'年份':years,'总收益(港幣百萬元)':generals,'总收益同比下降':general_percents,'净利润(港幣百萬元)':pures,'净利润同比下降':pures_percents}
        #将数据转换成pandas的DataFrame的格式
        df=pd.core.frame.DataFrame(df)
        #写入excel
        df.to_excel(excel_path)
    else:
        print('输出的信息长度不一致')

def get_company_data(file_folder,excel_path):
    '''
    file_folder pdf路径
    excel_path 数据提取后保存的excel路径
    '''
    file_lists = os.listdir(file_folder)
    # 创建不同的数组用于保存提取的数据
    years = []  # 年份
    generals = []  # 总收入
    general_percents = []  # 总收入同比下跌
    pures = []  # 净利润
    pures_percents = []  # 净利润同比下跌
    for file_name in file_lists:
        if os.path.splitext(file_name)[1] == '.pdf':  # 目录下包含.pdf的文件
            # 合并路径
            temp_path = os.path.join(file_folder, file_name)
            # pdf_read
            # 加入异常处理出现有问题的文件则进行下一个文件
            try:
                with open(temp_path, 'rb') as f:
                    # 根据下载到的文件名获取年报的年份，由于是第二年初的报告因此实际的统计年份需要减去一年
                    year = temp_path[temp_path.find('2'):temp_path.find('2') + 4]
                    year = int(year) - 1  # 获取到年报实际的年份
                    # 读取pdf文件
                    pdf = pdf_read(f)
                    # 检索文档中存在的书签信息,返回的对象是一个嵌套的列表
                    pg_id_num_map = _setup_page_id_to_num(pdf)
                    text_outline_list = pdf.getOutlines()
                    # 从书签列表中获取相关信息
                    # 首先根据看到的一些下载到的文件可以看到'按核心業務劃分的分析' 长江每年的年报信息主要在此书签之中
                    # 首先按照pdf文件中的信息查找此书签位置
                    for i in range(len(text_outline_list)):
                        # 这边我按照自己的理解认为是按核心業務劃分的分析这个书签，也可以是其他的，如果是其他的后续的代码也要做相应的修改
                        if text_outline_list[i].title == '按核心業務劃分的分析':
                            # 获取书签的所在的页码 由于页码是从第一页开始不是从0 开始加上1
                            pg_num = pg_id_num_map[text_outline_list[i].page.idnum] + 1
                            # 这边使用文字识别方法对pdf的表格进行处理，尝试了集中pdf文件提取表格数据的方法，对于不同的文件效果不一致
                            # 如果有其他的方法也可以替换
                            keywords=['收益總額','除稅後溢利']
                            general, general_percent, pure, pures_percent = read_table(temp_path, pg_num,keywords)
                            # #添加提取的信息
                            years.append(year)
                            generals.append(general)
                            general_percents.append(general_percent)
                            pures.append(pure)
                            pures_percents.append(pures_percent)
                            break
            except:
                #将有问题的文件文件路径进行打印
                print('文件获取信息时存在问题，需要人工检核',temp_path)
                continue
    # 在读取完当前文档的信息后将几个list表格的信息写入excel表格
    print(years, generals, general_percents, pures, pures_percents)
    write_excel(years, generals, general_percents, pures, pures_percents,excel_path)


if __name__ == '__main__':
    file_folder = './00001'  # 年报所在文件
    excel_path='result.xlsx' # excel文件名称
    get_company_data(file_folder,excel_path)

