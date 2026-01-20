# CAD-Annotation-Serial-Number-Write-back
CAD Annotation Serial Number Write-backCAD 批量标注处理工具专为 CAD设计，可批量提取 DWG 文件中的标注信息并生成结构化 Excel 报表，自动将序号回写至图纸对应位置。支持带圈 / 括号序号灵活切换，解决字体兼容性问题，搭载美观易用的 GUI 界面，操作流程可视化，有效提升工程图纸标注处理的效率与准确性。
CAD 批量标注处理工具专为 CAD设计，可批量提取 DWG 文件中的标注信息并生成结构化 Excel 报表，自动将序号回写至图纸对应位置。支持带圈 / 括号序号灵活切换，解决字体兼容性问题，搭载美观易用的 GUI 界面，操作流程可视化，有效提升工程图纸标注处理的效率与准确性。

The CAD Batch Annotation Tool is designed for CAD 2023. It batch extracts annotation information from DWG files, generates structured Excel reports, and automatically writes serial numbers back to corresponding positions in drawings. Supporting flexible switch between circled/bracketed numbers to solve font compatibility issues, it features an intuitive and user-friendly GUI with visualized operation flow, improving the efficiency and accuracy of engineering drawing annotation processing.

 ZwCAD Batch Annotation Tool
ZwCAD批量标注提取与序号回写工具

[![Python Version](https://img.shields.io/badge/Python-3.7%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

项目介绍 / Project Introduction
 中文
一款专为ZwCAD 2023设计的批量标注处理工具，通过可视化GUI界面实现DWG文件标注信息的批量提取、结构化Excel报表生成，以及序号自动回写至图纸对应位置。解决带圈序号字体兼容性问题，支持灵活切换序号样式，大幅提升工程图纸标注处理效率。

 English
A batch annotation processing tool designed for ZwCAD 2023. It implements batch extraction of annotation information from DWG files, generation of structured Excel reports, and automatic writing of serial numbers back to corresponding positions in drawings via a visual GUI interface. It solves font compatibility issues of circled serial numbers, supports flexible switching of serial number styles, and greatly improves the efficiency of engineering drawing annotation processing.

功能特性 / Key Features
 中文
1. 批量提取DWG文件中的尺寸标注、文本标注等信息，自动记录坐标位置
2. 生成包含序号、标注内容、坐标的结构化Excel报表，格式规整易查阅
3. 自动将序号回写至图纸标注位置，支持带圈/括号序号样式灵活切换
4. 美观易用的GUI界面，可视化操作流程，进度实时展示，日志清晰可查
5. 兼容ZwCAD COM接口，自动启动/连接ZwCAD程序，异常处理机制完善

 English
1. Batch extract dimension annotations, text annotations and other information from DWG files, and automatically record coordinate positions
2. Generate structured Excel reports containing serial numbers, annotation content and coordinates with regular format for easy reference
3. Automatically write serial numbers back to annotation positions in drawings, supporting flexible switching between circled/bracketed serial number styles
4. Elegant and easy-to-use GUI interface with visualized operation flow, real-time progress display and clear log records
5. Compatible with ZwCAD COM interface, automatically start/connect to ZwCAD program with improved exception handling mechanism

环境要求 / Environment Requirements
 中文
1. 操作系统：Windows 10/11（64位）
2. Python版本：3.7及以上
3. 依赖库：pywin32、openpyxl
4. 软件环境：ZwCAD 2023（其他版本需适配路径）

 English
1. Operating System: Windows 10/11 (64-bit)
2. Python Version: 3.7 or higher
3. Dependencies: pywin32, openpyxl
4. Software Environment: ZwCAD 2023 (other versions require path adaptation)

安装步骤 / Installation Steps
 中文
1. 克隆本仓库或下载源码包
bash
   git clone https://github.com/your-username/zwcad-batch-annotation.git
   cd zwcad-batch-annotation

2. 安装依赖库
bash
   pip install pywin32 openpyxl


 English
1. Clone this repository or download the source code package
bash
   git clone https://github.com/your-username/zwcad-batch-annotation.git
   cd zwcad-batch-annotation

2. Install dependencies
bash
   pip install pywin32 openpyxl


配置说明 / Configuration Instructions
 中文
修改脚本头部「用户可改区域」的参数：
| 参数名 | 说明 | 默认值 |
|--------|------|--------|
| ZWCAD_EXE | ZwCAD程序路径 | C:\Program Files\ZWSOFT\ZWCAD 2023\ZWCAD.exe |
| WORK_DIR | 输出文件保存目录 | D:\CAD标识\标识后 |
| EXCEL_NAME | 生成的Excel文件名 | 数值表.xlsx |
| TEXT_HEIGHT | 回写序号的文字高度 | 2.5 |
| TEXT_OFFSET_Y | 序号Y轴偏移量（避免遮挡原标注） | 3.0 |
| SUPPORT_FONT | 支持特殊字符的CAD字体 | gbcbig.shx |
| USE_BRACKET_NUMBERS | 是否使用括号序号（True/False） | True |

 English
Modify parameters in the "User Configurable Area" at the top of the script:
| Parameter Name | Description | Default Value |
|----------------|-------------|---------------|
| ZWCAD_EXE | ZwCAD program path | C:\Program Files\ZWSOFT\ZWCAD 2023\ZWCAD.exe |
| WORK_DIR | Output file save directory | D:\CAD标识\标识后 |
| EXCEL_NAME | Generated Excel file name | 数值表.xlsx |
| TEXT_HEIGHT | Text height of written-back serial numbers | 2.5 |
| TEXT_OFFSET_Y | Y-axis offset of serial numbers (avoid covering original annotations) | 3.0 |
| SUPPORT_FONT | CAD font supporting special characters | gbcbig.shx |
| USE_BRACKET_NUMBERS | Whether to use bracketed serial numbers (True/False) | True |

使用方法 / Usage
 中文
1. 运行脚本
bash
   python zwcad_batch_annotation.py

2. 在GUI界面中点击「选择DWG文件」，选中一个或多个需要处理的DWG文件
3. 点击「开始批量处理」，程序会自动启动ZwCAD并执行以下操作：
   - 提取DWG文件中的标注信息
   - 生成Excel报表至指定目录
   - 将序号回写至图纸对应位置并保存
4. 处理完成后可点击「打开输出文件夹」查看结果，日志窗口可查看详细执行过程

 English
1. Run the script
bash
   python zwcad_batch_annotation.py

2. Click "Select DWG Files" in the GUI interface and select one or more DWG files to process
3. Click "Start Batch Processing", the program will automatically start ZwCAD and perform the following operations:
   - Extract annotation information from DWG files
   - Generate Excel report to the specified directory
   - Write serial numbers back to corresponding positions in drawings and save
4. After processing, click "Open Output Folder" to view results, and check the log window for detailed execution process

注意事项 / Notes
 中文
1. 确保ZwCAD 2023已正确安装，且路径与配置中的ZWCAD_EXE一致
2. 运行脚本时需以管理员权限执行，避免权限不足导致ZwCAD启动失败
3. DWG文件路径和输出目录路径请勿包含中文、空格或特殊字符
4. 处理大文件时请耐心等待，ZwCAD启动和文件加载需要一定时间
5. 若出现COM接口连接失败，关闭ZwCAD后重新运行脚本

 English
1. Ensure ZwCAD 2023 is correctly installed and the path is consistent with ZWCAD_EXE in the configuration
2. Run the script with administrator privileges to avoid ZwCAD startup failure due to insufficient permissions
3. DWG file path and output directory path should not contain Chinese characters, spaces or special characters
4. Please be patient when processing large files, as ZwCAD startup and file loading take time
5. If COM interface connection fails, close ZwCAD and run the script again

许可证 / License
本项目采用MIT许可证开源 - 详见 [LICENSE](LICENSE) 文件  
This project is open source under the MIT License - see the [LICENSE](LICENSE) file for details.

免责声明 / Disclaimer
本工具仅用于学习和工作效率提升，请勿用于商业用途。使用本工具产生的任何问题，作者不承担相关责任。  
This tool is only for learning and work efficiency improvement, and shall not be used for commercial purposes. The author is not responsible for any problems arising from the use of this tool.
