# 磐石电气运单信息统计分析脚本
供本公司对装货流水信息进行分组统计

# 功能列表
## 从装货流水表格导出装托清单数据
    
## 从装货流水表格导出发货清单数据

# 使用方法

## 安装依赖
运行以下命令安装依赖程序
```shell 
    pip3 install -r requirements.txt
```

## 放入装货流水数据
将装货原始流水数据放入项目 `source` 目录下

## 运行程序
可直接运行  `main.py` 开始统计
```shell script
python main.py
```

## 查看结果
运行完上述命令后将在项目`target`目录下看到生成的文件
例如有源文件名称为`test.xlsx`，则可在`target`目录下看到如下两个文件：
- test.xlsx_part.txt 发货清单数据
- test.xlsx_pallet.txt 装托清单数据
