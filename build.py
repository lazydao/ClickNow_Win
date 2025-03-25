import PyInstaller.__main__
import os
import shutil
import sys
import site

def build():
    # 清理之前的构建文件
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')

    # 获取虚拟环境路径
    venv_path = os.path.dirname(sys.executable)
    site_packages = site.getsitepackages()[0]

    # PyInstaller参数
    params = [
        'src/app.py',  # 主程序文件
        '--name=ClickNow',  # 生成的exe名称
        '--windowed',  # 使用GUI模式
        '--noconsole',  # 不显示控制台
        '--add-data=src/icons;icons',  # 添加图标文件夹
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不确认覆盖
        f'--paths={site_packages}',  # 添加site-packages路径
        '--hidden-import=win32com',
        '--hidden-import=win32com.client',
        '--hidden-import=win32gui',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
        '--hidden-import=win32api',
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.sip',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=PyQt5.QtWidgets',
        '--collect-all=win32com',
        '--collect-all=win32gui',
        '--collect-all=pythoncom',
        '--collect-all=pywintypes',
        '--collect-all=PyQt5',
    ]

    # 执行打包
    PyInstaller.__main__.run(params)

    # 创建发布文件夹
    release_dir = 'release'
    if not os.path.exists(release_dir):
        os.makedirs(release_dir)

    # 复制必要文件到发布文件夹
    if os.path.exists(f'{release_dir}/ClickNow'):
        shutil.rmtree(f'{release_dir}/ClickNow')
    shutil.copytree('dist/ClickNow', f'{release_dir}/ClickNow')

    print('打包完成！')
    print(f'发布文件位于: {os.path.abspath(release_dir)}')

if __name__ == '__main__':
    build()
