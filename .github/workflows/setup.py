"""
setup.py 用于配置 py2app 打包参数
"""

from setuptools import setup

APP = ['main.py'] # 你的主程序文件名
DATA_FILES = []   # 如果有额外的图片或配置文件需要打包，放在这里

OPTIONS = {
    'py2app': {
        'plist': {
            'CFBundleName': "销售毛利分析系统", # 应用名称
            'CFBundleDisplayName': "销售毛利分析系统", # 显示名称
            'CFBundleGetInfoString': "专业销售毛利智能分析工具 v4.4",
            'CFBundleIdentifier': "com.yourcompany.marginanalyzer", # 唯一标识符
            'CFBundleVersion': "4.4.0",
        },
        # 关键配置：包含必要的包
        'packages': [
            'pandas',
            'openpyxl',
            'matplotlib',
            'numpy',
            'tkinter'
        ],
        # 关键配置：显式包含 matplotlib 的 Tkinter 后端，否则图表可能无法显示
        'includes': [
            'matplotlib.backends.backend_tkagg',
            'matplotlib.figure',
            'matplotlib.backends.backend_agg'
        ],
        'iconfile': None, # 如果有 .icns 图标文件，在这里填路径，例如 'app.icns'
        'site_packages': True, # 包含 site-packages
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
    name="销售毛利分析系统",
    version="4.4"
)