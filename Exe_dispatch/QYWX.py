import requests
import json
from common.log import Logger

sy_log = Logger(__name__)


def make_mes(total_data_sum):
    sumCase = total_data_sum['sum_case']
    success_case_num = total_data_sum['success_case_num']
    fail_num = sumCase - success_case_num
    pass_result = "%.2f%%" % (success_case_num / sumCase * 100)
    fail_case = str(total_data_sum['fail_case'])
    content = f'此次接口测试一共运行用例为：{sumCase}，通过个数为：{success_case_num}，失败个数为：{fail_num}，通过率为：{pass_result}, 有执行失败的Excel为：{fail_case}，报告详情请查看附件！'
    return content


class QYWX(object):
    def __init__(self):
        # 企业ID
        self.__qy_id = 'ww58d668fc292f65d4'
        # 森亿test应用的密钥
        self.__secret = 'Op9sApq3g7YBr_sEXBwTX-G7YQO0c5SUnzIlGPHnF_I'
        self.__agent_id = 1000053
        self.__token = ''

    def get_token(self):
        url = 'https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=' + self.__qy_id + '&corpsecret=' + self.__secret
        token = json.loads(requests.get(url).text)
        self.__token = token.get('access_token')

    def push_mes(self, content, user, toparty=""):
        url = 'https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=' + self.__token
        body = {
            "touser": user,
            "toparty": toparty,
            "totag": "",
            "msgtype": "text",
            "agentid": self.__agent_id,
            "text": {
                "content": content
            },
            "safe": 0,
            "enable_id_trans": 0,
            "enable_duplicate_check": 0,
            "duplicate_check_interval": 1800
        }
        res = requests.post(url, json=body)
        res = json.loads(res.text)
        if res.get('errcode') == 0:
            sy_log.logger.info('消息发送成功')

    def upload_file(self, filename, filepath):
        url = 'https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token=' + self.__token + '&type=file'
        payload = {'Content-Type': 'multipart/form-data;', 'filename': filename, 'filelength': '220'}
        files = [
            ('file', open(filepath, 'rb'))
        ]
        res = requests.post(url, headers={}, data=payload, files=files)
        res = json.loads(res.text)
        if res.get('errcode') == 0:
            media_id = res.get('media_id')
            return media_id
        else:
            return {'synyi_test_frame_track': '上传失败'}

    def push_file(self, media_id, user):
        url = 'https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=' + self.__token
        body = {
            "touser": user,
            "toparty": "",
            "totag": "",
            "msgtype": "file",
            "agentid": self.__agent_id,
            "file": {
                "media_id": media_id
            },
            "safe": 0,
            "enable_duplicate_check": 0,
            "duplicate_check_interval": 1800
        }
        res = requests.post(url=url, json=body)
        res = json.loads(res.text)
        if res.get('errcode') == 0:
            sy_log.logger.info('文件发送成功')


# if __name__ == '__main__':
#