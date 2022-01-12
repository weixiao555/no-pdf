# --coding:utf-8--
from wujiexi import data as dt
from wujiexi import sort as st


def get_data():
    # _user = 'huangxiao'
    # _pass = '88888'
    # project_number = 'G069338'

    _user = input('账号: ')
    _pass = input('密码: ')
    project_number = input('工程名称: ')

    cookies = dt.MyCookie()  # 通过selenium生成cookie
    cookie = cookies.selenium_login(_user, _pass)

    param = dt.MyParam()  # 生成请求页面的参数
    few = param.get_few_param(cookie, project_number)
    few_headers = few['headers']
    few_form_data = few['form_data']
    more = param.get_more_param(cookie, project_number)
    more_headers = more['headers']
    more_form_data = more['form_data']

    get = dt.MyGet()  # 获取数据"所有"和"查看"的数据
    few_data = get.get_few_data(few_headers, few_form_data)
    project_name = few_data['project_name']
    few_data_list = few_data['few_data_list']
    more_data_list = get.get_more_data(more_headers, more_form_data)

    parse = dt.MyParse()
    all_data = parse.sort_service(few_data_list, more_data_list)  # 合并数据,"所有"和"查看"

    mySave = dt.MySave()
    save_path = r'resource/excel/《{}》[{}]汇总表.xlsx'.format(project_name, project_number)
    mySave.my_save(all_data, project_name, save_path)  # 保存为xls
    return save_path


def sort_data(save_path):
    file_path = save_path
    get = st.GetDataframe()  # excel转为dataframe
    all_sheet_data = get.get_excel(file_path)
    data = get.merge_basic_main(all_sheet_data)

    sort = st.MySort()  # 排序,分楼层,分宿舍车间
    final_data = sort.merge(data)

    mySave = st.MySave()  # 保存为xlsx
    mySave.my_save(file_path, final_data)


def main():
    save_path = get_data()
    sort_data(save_path)


if __name__ == '__main__':
    main()
