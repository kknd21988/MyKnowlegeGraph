########################## SCIRPT 1 ###############################
# 功能：从xls中指定的pdf清单下载文件，保存到指定文件夹
# 参数如下：
# 1.产品清单，为xls文件
path = ''
# 2.指定下载的文件清单在哪个sheet中
sheetIndex = 0
# 3.指定从第几行开始下载，默认为0
start = 1
end = 117
# 4.指定下载的文件存放在哪个目录下
folder = ''


# 脚本运行
productCompanyIndex = 1 # “产品公司名”所在的列
productNameIndex = 0    # “产品名”所在的列
productDownloadIndex = 2    # “下载链接”所在的列
workbook = xlrd.open_workbook(path)
sheetTraining = workbook.sheets()[sheetIndex]
rowNum = sheetTraining.nrows
colNum = sheetTraining.ncols
for i in np.arange(start, end):
    if i != 0:
        # 获取文件参数
        productName = '#' + str(i) + sheetTraining.cell_value(i, productNameIndex) + '@' + sheetTraining.cell_value(i, productCompanyIndex)
        productLink = sheetTraining.cell_value(i, productDownloadIndex)
        # 下载文件
        download_file(productLink, productName, folder)
        print("File#%d OK! %f%% completed!" % (i, (i*100.0)/rowNum) )
