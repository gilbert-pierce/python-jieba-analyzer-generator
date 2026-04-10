【Jieba 分词统计工具 - 离线安装指南】

==================================================
该程序用于在无网 Windows 环境下对 Excel 数据进行中文分词和词频统计。
本程序专为轻量化设计，移除庞大的 Pandas 库，改用 openpyxl，确保打包体积最小。
==================================================

【核心功能】
1. 全局词频统计：统计整列数据的高频词。
2. 行级分词导出：自动对每一行进行分词处理（去除停用词），并写入新的一列。
3. TF-IDF 关键词：使用 TF-IDF 算法提取整份文档的高权重关键词并导出结果。
4. 自定义停用词：界面提供文本框，可随时调整过滤词。

一、环境准备 (Windows 10/11)
------------------------------------------------
1. 安装 Python (推荐 3.9+)
   - 访问 python.org 下载并安装。
   - 【重要】勾选 "Add Python to PATH"。

2. 安装必要的依赖库
   打开 CMD 输入：
   
   pip install openpyxl jieba pyinstaller

   * openpyxl : Excel 处理库
   * jieba : 中文分词库
   * pyinstaller : 打包工具


二、运行源码
------------------------------------------------
1. 在网页界面右侧操作栏最底部，点击蓝色按钮【下载 Python 独立工具源码】，保存为 'analyzer.py'。
2. 将文件放入任意文件夹。
3. 在该文件夹打开 CMD 或 PowerShell。
4. 运行：
   
   python analyzer.py


三、打包为 EXE (体积优化版)
------------------------------------------------
由于移除了 Pandas，直接打包体积通常已经在 20MB 左右。

1. 在 'analyzer.py' 所在目录打开命令行。
2. 执行打包命令：

   pyinstaller -F -w -i NONE analyzer.py

   * -F : 单文件
   * -w : 无黑框窗口
   
3. 打包完成后，在 'dist' 文件夹中查找 exe 文件。


四、常见问题 FAQ
------------------------------------------------
Q: 如果我的电脑上已经安装了 Pandas，PyInstaller 打包时会把它打进去吗？
A: 通常不会。PyInstaller 会自动分析代码中的 'import' 语句。因为本工具代码中没有引用 pandas，所以即使环境中存在，正常情况下也不会被打包。

Q: 如何确保打出的 EXE 文件最小？
A: 最佳实践是创建一个干净的虚拟环境进行打包：
   1. python -m venv myenv
   2. myenv\Scripts\activate
   3. pip install openpyxl jieba pyinstaller (只安装这三个)
   4. pyinstaller -F -w analyzer.py
   这样可以保证 100% 纯净，体积通常 < 20MB。
