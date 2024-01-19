import pandas as pd
import numpy as np
import os
import sys
import shutil


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILES = os.path.join(SCRIPT_DIR, '../ExcelFiles/')
JSON_FILES = os.path.join(SCRIPT_DIR, '../JsonFiles/')


def BackupExcel():
    """
    excel源数据备份，拷贝到独立目录在做处理。
    
    返回：
    None
    """
    
    # excel源数据备份，拷贝到独立目录在做处理
    # 源目录
    sourceDir = os.path.join(SCRIPT_DIR, '../')
    # 目标目录
    targetDir = EXCEL_FILES

    # 确保目标目录存在，如果不存在则创建
    if not os.path.exists(targetDir):
        os.makedirs(targetDir)

    # 遍历源目录下的所有文件和子目录
    for filename in os.listdir(sourceDir):
        # 只处理.xlsx后缀的文件
        if filename.endswith('.xlsx'):
            sourceFile = os.path.join(sourceDir, filename)
            # 将源文件名作为目标文件名（也可以自定义新的文件名）
            # targetFile = os.path.join(targetDir, filename)
            
            # 复制文件
            shutil.copy(sourceFile, targetDir)
            
            
    # print('excel文件备份成功')


def CompleteCardLogicExcel():
    """
    如果选项ID匹配到卡牌选项表，补全卡牌逻辑表中的选项名称、下一级选项名称字段值。

    返回：
    None
    """

    # 读取Excel文件
    cardOptions = pd.read_excel(os.path.join(SCRIPT_DIR, '../卡牌选项表.xlsx'))
    cardLogic = pd.read_excel(os.path.join(SCRIPT_DIR, '../卡牌逻辑表.xlsx'))
    
    # 确保它们的值类型
    cardOptions['ID'] = cardOptions['ID'].astype(int)
    cardLogic['选项ID'] = cardLogic['选项ID'].astype(int)
    cardLogic['下一级选项ID_tem'] = cardLogic['下一级选项ID']
    cardLogic['下一级选项ID_tem'] = cardLogic['下一级选项ID_tem'].astype(str)
    
    
    for k, row in cardLogic.iterrows():
        val = row['下一级选项ID_tem']
        if val != 'nan':
            # 通过|分割获取下一级选项ID列表
            idsList = [int(n) for n in val.split('|')]
            # 通过ID列表匹配获取选项名称
            matchedRows = cardOptions[cardOptions['ID'].isin(idsList)]
            # 将匹配到的值回填回对应行号的 下一级选项名称 字段中
            if not matchedRows.empty:
                joinStr = '|'.join(matchedRows['选项名称'].astype('str'))
                cardLogic.loc[k, '下一级选项名称'] = joinStr

    
    # 使用merge操作结合两个数据表
    optionMerged = pd.merge(cardLogic, cardOptions[['ID', '选项名称']], how='left', left_on='选项ID', right_on='ID', suffixes=('_logic', '_option'))
    
    optionMerged['选项名称'] = optionMerged['选项名称_option']
    optionMerged['ID'] = optionMerged['ID_logic']

    # 获取原来选项名称字段所在的位置索引
    optionNameIdx = optionMerged.columns.get_loc('选项名称_logic')

    # 删除临时用的*_option、*_tmp字段，重复ID字段，旧的选择名称字段
    optionMerged.drop(columns=['选项名称_option', 'ID_logic', 'ID_option', '选项名称_logic', '下一级选项ID_tem'], inplace=True)
    
    # 插入到指定排序位置
    optionMerged.insert(0, 'ID', optionMerged.pop('ID')) # 复原逻辑表ID列到最前的位置
    optionMerged.insert(optionNameIdx, '选项名称', optionMerged.pop('选项名称')) # 复原逻辑表选项名称列到原来的位置
    
    # 最后保存结果到新的Excel文件或覆盖原文件
    optionMerged.to_excel(os.path.join(SCRIPT_DIR, '../卡牌逻辑表.xlsx'), index=False)
    
    # print("卡牌逻辑表选项名称补全成功")


def ExcelToJson(excelFiles, outputDir = JSON_FILES):
    """
    将Excel文件列表转换为对应的JSON文件。

    参数：
    - excel_files (List[str]): Excel文件名列表（包含完整路径或相对路径）。
    - output_dir (str): 输出JSON文件的目录，默认当前目录。

    返回：
    None
    """
    
    print(f"---------- Step 3: 开始生成Json文件 ----------")
    
    if not os.path.exists(outputDir):
        os.makedirs(outputDir)
        
        
    for file in excelFiles:
        fileName = EXCEL_FILES + file
        # 检查Excel文件是否存在
        if not os.path.isfile(fileName):
            print(f"警告：{fileName} 文件不存在或不是文件。")
            continue
        
        # 读取Excel文件
        df = pd.read_excel(fileName)

        # 将NaN替换为None或其他合适的值以便JSON可以正确处理
        df.replace({np.nan: ''}, inplace=True)
        
        # 从文件名中提取基础名称以创建JSON文件名
        base_name = os.path.splitext(os.path.basename(fileName))[0] + '.json'
        json_file = os.path.join(outputDir, base_name)

        # 转换并写入JSON文件
        df.to_json(json_file, orient='records', force_ascii=False, indent=4)
        
        print(f"[{file}]转换Json文件成功")


def Start():
    # 补全匹配项
    CompleteCardLogicExcel()
    print("---------- Step 1: 卡牌逻辑表补全匹配项完成 ----------")
    # 备份源数据表
    BackupExcel()
    print("---------- Step 2: Excel文件备份成功 ----------")
    # 生成json文件
    ExcelToJson(['卡牌资源表.xlsx', '卡牌信息表.xlsx', '卡牌选项表.xlsx', '卡牌逻辑表.xlsx'])


def dd():
    sys.exit()
    


# 启动执行
Start()
