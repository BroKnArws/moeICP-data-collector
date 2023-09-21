# moeICP-data-collector
基于备案号的萌国 ICP 备案信息数据抓取

## 使用说明

1. 安装
```pip install requests lxml argparse openpyxl ```
2. 启动
``` python main.py --start 20220010 --end 20220015 --append```

启动后，程序会从 start 开始，依次扫描到 end 为止的备案号对应的网站信息，这些数据整合后，将输出格式为 .xlsx 的表格文件，并存入到指定的目录下（默认是程序运行所在的目录）。
参数说明：
- -- start 开始的萌号（备案号）
- -- end 结束的萌号（备案号）
- -- append 是否追加写（如果是，那么本次抓取时，会继续往表格后面追加数据）
- -- output 输出 xlsx 表格文件对应的路径（含文件名）（默认值：```./萌备案数据.xlsx```）
