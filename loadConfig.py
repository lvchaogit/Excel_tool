
import os
from configparser import ConfigParser


def get_config():
    """ 获取配置文件  """
    print("#获得当前工作目录".format(os.getcwd()))
    print("#获得当前工作目录--{}".format(os.path.abspath('.')))
    print("#获得当前工作目录的父目录--{}".format(os.path.abspath('..')))
    print("#获得当前工作目录--{}".format(os.path.abspath(os.curdir)))
    cfg = ConfigParser()
    cfg.read('config.ini','UTF-8')
    count_column_config = cfg.get('installation', 'count_column')
    group_column_config = cfg.get('installation', 'group_column')

    if count_column_config is not None and group_column_config is not None:
        count_columns = count_column_config.split("、")
        group_columns = group_column_config.split("、")
        return ConfigObj(count_columns, group_columns)
    return None


class ConfigObj:
    """ 配置类OBJ """
    def __init__(self, count_columns, group_columns) -> None:
        super().__init__()
        self.count_columns = count_columns
        self.group_columns = group_columns
        self.all_columns = group_columns + count_columns

    def __str__(self):
        return str(self.group_columns) + str(self.count_columns)


if __name__ == '__main__':
    # open_excel("C:\\Users\\l2503\\Desktop\\线上不开票客户4月份销售明细.xlsx")
    print(get_config())
    pass
