import pandas as pd


def save(origin_path, new_path):
    df = pd.read_excel(origin_path)

    not_cal_avg_columns = {
        0: "姓名",
        2: "班级名称",
        3: "试卷类型"
    }
    data = [col for i, col in enumerate(list(df.columns)) if i not in not_cal_avg_columns.keys()]
    print(data)

    _map = {'姓名': '平均分'}

    means = df.mean(skipna=True)
    print(means)

    for n, mean in enumerate(means):
        if n > 0:
            # n = 0 学号，不显示平均值
            _map[data[n]] = mean
    print(_map)

    df = df.append(_map, ignore_index=True)
    print(df)

    df.to_excel(new_path, index=False)


if __name__ == '__main__':
    origin_path = '/Users/dong/work/fltrp/workspace/tools/潍坊区县-年级成绩单/统考汇总/坊子区/坊子区七年级/坊子区弘信学校.xlsx'
    new_path = '坊子区弘信学校_new.xlsx'
    save(origin_path, new_path)