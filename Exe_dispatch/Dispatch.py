# -*- coding:utf-8 -*-
# __author__ = 'szj'
import os
import time
import re
import sys
from .Report import Report
from common.log import Logger
from .ExecuteCases import Execute
from .QYWX import QYWX
from .Report import report_data, make_zip
from .ReadCases import ReadCaseInfo
from .QYWX import make_mes

sy_log = Logger(__name__)


def dispatch(path, *case_suites):
    """
    调度执行文件
    :param path:
    :param case_suites: 可选参数，接受参数为列表对象，内容为sheet页名称。不传入，默认执行所有Case，传入则执行列表中指定的sheet页
    :return:
    """
    now_time = time.strftime("%Y-%m-%d %H_%M", time.localtime())  # 获取当前时间，并转化为指定的格式
    abs_path = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
    if re.findall('/', path):
        os_win = '/'
    else:
        os_win = '\\'

    if not os.path.isdir(path) and path.endswith('.xlsx'):
        excel_name = path.split(os_win)[-1].split('.')[0]
        report_name = excel_name + '_' + now_time + '_result.xlsx'
        report_dir = os.path.join(abs_path, 'Result')
        if not os.path.exists(report_dir):
            os.makedirs(report_dir)
        report_path = os.path.join(report_dir, report_name)

        sy_log.logger.info("------------------- 开始执行指定文件：%s -------------------" % excel_name)

        if len(case_suites) > 0:  # 判断是否指定sheet页执行
            all_cases, inter_value, db_config, timeout = ReadCaseInfo(path).pretreated_case(case_suites[0])
            exe_result = Execute().execute_case(all_cases, inter_value, db_config, timeout, path)
        else:
            all_cases, inter_value, db_config, timeout = ReadCaseInfo(path).pretreated_case()
            exe_result = Execute().execute_case(all_cases, inter_value, db_config, timeout, path)

        sy_log.logger.info("------------------- 执行结束，开始生成报告 -------------------")

        # 按照固定的路径生成报告
        report_value = report_data(exe_result)
        project_info = ReadCaseInfo(path).project_info()
        new_reportstyle = Report(report_path, report_value, project_info)
        new_reportstyle.totali()
        new_reportstyle.detail()

        sy_log.logger.info("------------------- 生成报告成功，请查看企业微信消息 -------------------")
        send_info = project_info.get('send_info')
        zip_name = str(report_path.split(".xlsx")[0]) + '.zip'
        make_zip(zip_name, source_file=report_path)
        if send_info['on_off'] == 'on':
            users = send_info['user']
            qywx = QYWX()
            qywx.get_token()
            media_id = qywx.upload_file(report_name, zip_name)
            qywx.push_file(media_id, users)
    elif os.path.isdir(path):
        dir_name = path.split(os_win)[-1]
        report_path = os.path.join(abs_path, 'Result' + os_win + dir_name)
        if not os.path.exists(report_path):
            os.makedirs(report_path)

        sy_log.logger.info("------------------- 开始执行文件夹  %s -------------------" % dir_name)
        total_data_sum = {"sum_case": 0, "success_case_num": 0, "fail_case": []}
        project_info = ''
        for file in os.listdir(path):
            case_path = path + os_win + file
            if not os.path.isdir(case_path):
                if case_path.endswith('.xlsx'):
                    sy_log.logger.info("------------------- 开始执行  %s ------------------- " % file)
                    all_cases, inter_value, db_config, timeout = ReadCaseInfo(case_path).pretreated_case()
                    result_value = Execute().execute_case(all_cases, inter_value, db_config, timeout, case_path)
                    report_file_path = os.path.join(report_path, file + '_' + now_time + '_result.xlsx')
                    report_value = report_data(result_value)
                    project_info = ReadCaseInfo(case_path).project_info()

                    new_report_style = Report(report_file_path, report_value, project_info)
                    new_report_style.totali()
                    new_report_style.detail()

                    total_data = report_data(result_value)
                    total_data_sum['sum_case'] += total_data['sum_case']
                    total_data_sum['success_case_num'] += total_data['success_case_num']
                    if total_data['sum_case'] != total_data['success_case_num']:
                        total_data_sum["fail_case"].append(file)
                else:
                    sy_log.logger.info("文件不是EXCEL，无法执行\n")

        make_zip(report_path + '.zip', report_path)
        sy_log.logger.info("------------------- 生成报告成功，请注意查收企业微信 -------------------")
        content = make_mes(total_data_sum)
        send_info = project_info.get('send_info')
        if send_info['on_off'] == 'on':
            users = send_info['user']
            qywx = QYWX()
            qywx.get_token()
            media_id = qywx.upload_file('Result_' + dir_name, report_path + '.zip')
            if 'synyi_test_frame_track' in media_id:
                sy_log.logger.info("------------------- 发送失败 -------------------")
            else:
                qywx.push_mes(content, users)
                qywx.push_file(media_id, users)
                sy_log.logger.info("------------------- 发送成功 -------------------")
    else:
        sy_log.logger.info("未能识别文件，无法执行")
        sys.exit(0)
