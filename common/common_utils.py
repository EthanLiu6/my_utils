import os
from pathlib import Path


def find_project_root():
    # 当前脚本路径
    current_file = Path(__file__).resolve()

    # 项目根目录
    project_root = current_file.parents[1]  # 逐级向上：common -> utils -> project_root
    return project_root


# 默认环境变量获取函数
def get_env(key: str, default=None) -> str:
    return os.getenv(key, default)


def get_dir_files_name(dir_path) -> list:
    """获取指定文件夹下的所有文件名，返回所有文件名的列表格式"""

    dir_path = Path(dir_path)
    if not os.path.isdir(dir_path):
        return []

    files = []
    for entry in os.listdir(dir_path):
        full_path = os.path.join(dir_path, entry)
        if os.path.isfile(full_path):
            files.append(entry)

    return files


def get_path_file_name(file_path) -> str:
    file_path = Path(file_path)
    if os.path.isfile(file_path):
        return file_path.name
    else:
        raise '输入文件路径有问题！'


if __name__ == '__main__':
    project_path = find_project_root()
    laws_files = get_dir_files_name(project_path.joinpath('processed_data/laws'))
    print(laws_files)
    law_file = get_path_file_name(project_path.joinpath('processed_data/laws/中华人民共和国招标投标法实施条例.csv'))
    print(law_file)