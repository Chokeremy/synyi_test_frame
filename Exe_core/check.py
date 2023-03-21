# -*- coding:utf-8 -*-
# __author__ = 'szj'
import os
import re

from common.operateDB import DBInterface
from common.util import *
from common.log import Logger

sy_log = Logger(__name__)


def check_sort(params, key_value, db_config, inter_value=None, case_path=None):
    """
    针对依赖参数分类执行替换
    :param inter_value:
    :param db_config:
    :param case_path:
    :param params: 要处理的参数
    :param key_value: 中间传递值
    :return: 返回参数
    """
    # 判断是否为None
    if params:
        # 判断是否要从其他sheet页获取数据
        if re.findall('\$var{|\$sheet\d+.var{', params, re.I):
            params = check_var(params, key_value, inter_value)
            if "synyi_test_frame_track" in params:
                return params

        # 判断是否存在$sql
        if re.findall('\$sql{', params, re.I):
            config = db_config['pg']
            if re.findall('\$sql{select', params, re.I):
                params = check_select(params, 'pg', config)
            else:
                condition = re.findall('(?<=\$sql{)[\s\S]+?(?=})', params, re.I)
                check_db_other(condition[0], 'pg', config)

            if "synyi_test_frame_track" in params:
                return params

        # 判断是否存在$emrsql
        if re.findall('\$orasql{', params, re.I):
            config = db_config['oracle']
            if re.findall('\$orasql{select', params, re.I):
                params = check_select(params, 'oracle', config)
            else:
                condition = re.findall('(?<=\$orasql{)[\s\S]+?(?=})', params, re.I)
                check_db_other(condition[0], 'oracle', config)

            if "synyi_test_frame_track" in params:
                return params

        # 判断是否存在$hissql
        if re.findall('\$mssql{', params, re.I):
            config = db_config['sqlserver']
            if re.findall('\$mssql{select', params, re.I):
                params = check_select(params, 'sqlserver', config)
            else:
                condition = re.findall('(?<=\$mssql{)[\s\S]+?(?=})', params, re.I)
                check_db_other(condition[0], 'sqlserver', config)

            if "synyi_test_frame_track" in params:
                return params

        # 判断是否存在$function
        if re.findall('\$func{', params, re.I):
            params = check_func(params)

        if re.findall('\$upfile{', params, re.I):
            params = check_upFiles(params, case_path)

        if re.findall('\$refile{', params, re.I):
            params = check_reFiles(params, case_path)

        if re.findall('\$sleep{', params, re.I):
            timespan = re.findall('(?<={)\d+?(?=})', params)
            time.sleep(int(timespan[0]))

    return params


def check_var(params, key_value, inter_value=None):
    """
    替换中间变量
    :param inter_value:
    :param params:
    :param key_value:
    :return:
    """
    if re.findall('\$sheet\d+.var{', params, re.I):
        sheet_number = re.findall('(?<=\$sheet)\d+(?=.var{)', params, re.I)
        index = sheet_number[0] - 2
        try:
            ref_value = inter_value[index][inter_value]
        except:
            return {"synyi_test_frame_track": str(inter_value[index][inter_value]) + ' 参数依赖的接口出错，此条用例失败'}
    else:
        ref_value = key_value

    param_list = re.findall('\$var{[\s\S]+?}', params, re.I)
    for p in param_list:
        key = re.findall('(?<=\$var{)[\s\S]+?(?=})', p, re.I)
        referenceData = ref_value.get(key[0])
        if referenceData is not None:
            referenceData = str(referenceData)
            if referenceData == 'synyi_test_frame的错误标识':
                return {"synyi_test_frame_track": referenceData + ' 参数依赖的接口出错，此条用例失败'}
            else:
                params = params.replace(p, referenceData)
        else:
            return {"synyi_test_frame_track": key[0] + ' 没有在之前的接口用例中定义'}
    return params


def check_select(params, db_type, config):
    """
    执行select操作
    :param params:
    :param db_type:
    :param config:
    :return:
    """
    if "\n" in params:
        params = params.replace("\n", " ")
    if db_type == 'oracle':
        temp_param = re.findall('\$orasql{[\s\S]+?}', params, re.I)
        # 匹配所有sql，形成list
        temp_sql = re.findall('(?<=\$orasql{)[\s\S]+?(?=})', params, re.I)
    elif db_type == 'sqlserver':
        temp_param = re.findall('\$mssql{[\s\S]+?}', params, re.I)
        # 匹配所有sql，形成list
        temp_sql = re.findall('(?<=\$mssql{)[\s\S]+?(?=})', params, re.I)
    else:
        # 匹配所有需要被替换的完整结构，形成list
        temp_param = re.findall('\$sql{[\s\S]+?}', params, re.I)
        # 匹配所有sql，形成list
        temp_sql = re.findall('(?<=\$sql{)[\s\S]+?(?=})', params, re.I)

    new_db = DBInterface(db_type, config)

    # 执行sql,获取结果
    results = new_db.select_sql(temp_sql[0])
    if "synyi_test_frame_track" in results:
        return results
    else:
        value = results['data']
        # 判断value是否为str类型，若非则变换类型为str
        if isinstance(value[0][0], str):
            # 替换参数的值
            params = params.replace(temp_param[0], value[0][0])
        else:
            # 替换参数的值
            params = params.replace(temp_param[0], str(value[0][0]))

        return params


def check_db_other(params, db_type, config):
    """
    执行其他更新插入删除操作
    :param params:
    :param db_type:
    :param config:
    :return:
    """
    if "\n" in params:
        params = params.replace("\n", " ")
    new_db = DBInterface(db_type, config)
    result = new_db.op_sql(params)
    if result:
        return result


def check_func(params):
    """
    替换引用对应的函数值
    :param params:
    :return:
    """
    if re.findall('rd_str\([\w\d\-_.]+\)', params, re.I):
        temp_params = re.findall('\$func{rd_str\(\d+\)?}', params, re.I)
        for temp_param in temp_params:
            x = re.findall('(?<=rd_str\().*(?=\))', temp_param, re.I)
            value = gen_random_string(int(x[0]))
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('rd_int\(\d+,\d+\)', params, re.I):
        temp_params = re.findall('\$func{rd_int\(\d+,\d+\)?}', params, re.I)
        for temp_param in temp_params:
            min_num1 = re.findall('(?<=\()\d+(?=,)', temp_param, re.I)
            max_num2 = re.findall('(?<=,)\d+(?=\))', temp_param, re.I)
            value = gen_random_int(min_num1[0], max_num2[0])
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('timestamp\(\)', params, re.I):
        temp_params = re.findall('\$func{timestamp\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = get_timestamp()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('utc_time\(\)', params, re.I):
        temp_params = re.findall('\$func{utc_time\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = get_time()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('\.str_time\(\)', params, re.I):
        temp_params = re.findall('\$func{str_time\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = get_strftime()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('td_timestamp\(\)', params, re.I):
        temp_params = re.findall('\$func{td_timestamp\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = get_today_timestamp()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('tm_timestamp\(\)', params, re.I):
        temp_params = re.findall('\$func{tm_timestamp\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = get_tomorrow_timestamp()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('date\(\)', params, re.I):
        temp_params = re.findall('\$func{date\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = get_date()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('int\(\)', params, re.I):
        temp_params = re.findall('\$func{int\(\)?}', params, re.I)
        for temp_param in temp_params:
            x = re.findall('(?<=\()\d+.\d+(?=\))', temp_param)
            value = change_int(str(x[0]))
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    elif re.findall('gen_uuid\(\)', params, re.I):
        temp_params = re.findall('\$func{gen_uuid\(\)?}', params, re.I)
        for temp_param in temp_params:
            value = gen_uuid()
            if isinstance(value, str):
                params = params.replace(temp_param, value)
            else:
                params = params.replace(temp_param, str(value))

    return params


def check_upFiles(params, case_path):
    """
    替换body为上传的文件
    :param params:
    :param case_path:
    :return:
    """
    temp_params = re.findall('\$upfile{[\s\S]+?}', params, re.I)
    temp_path = re.findall('(?<=\$upfile{)[\s\S]+?(?=})', params, re.I)
    if re.findall('/', case_path):
        ios_win = '/'
    else:
        ios_win = '\\'
    if os.path.isdir(case_path):
        path = os.path.join(case_path, 'uploadFile' + ios_win + temp_path[0])
    else:
        parent_path = os.path.dirname(case_path)
        path = os.path.join(parent_path, 'uploadFile' + ios_win + temp_path[0])

    params = params.replace(temp_params[0], path)
    return params


def check_reFiles(params, case_path):
    """
    替换文件中的内容为请求值
    :param params:
    :param case_path:
    :return:
    """
    temp_params = re.findall('\$refile{[\s\S]+?}', params, re.I)
    temp_path = re.findall('(?<=\$refile{)[\s\S]+?(?=})', params, re.I)
    if re.findall('/', case_path):
        ios_win = '/'
    else:
        ios_win = '\\'
    if os.path.isdir(case_path):

        path = os.path.join(case_path, 'replaceFile' + ios_win + temp_path[0])
    else:
        parent_path = os.path.dirname(case_path)

        path = os.path.join(parent_path, 'replaceFile' + ios_win + temp_path[0])
    with open(path, 'r', encoding='utf-8') as f:
        bodyData = f.read()
    params = params.replace(temp_params[0], bodyData)
    return params


def set_key_value(params, response, key_value):
    """
    保存中间变量
    :param response:
    :param params:
    :param key_value:
    :return:
    """
    params_list = params.split(';\n')
    for param in params_list:
        param = param.split('=')
        set_key = param[0]
        set_value = param[1]
        if re.findall('[\'\"“].*[\'\"”]', set_value):
            key_value[set_key] = re.findall('(?<=[\'\"“]).*(?=[\'\"”])', set_value)[0]
        else:
            path_key = set_value.split('.')
            res = response
            if isinstance(res, str):
                res = eval(res)
            for key in path_key:
                if re.findall('^\d+$', key):
                    key = int(key)
                try:
                    res = res[key]
                except:
                    key_value[set_key] = "synyi_test_frame的错误标识"
                    return key_value
            key_value[set_key] = res
    return key_value


if __name__ == "__main__":
    set_key_value("code='1';\ndata=data", "{\"data\":1}", {"code": "'1'", "data": 'data'})