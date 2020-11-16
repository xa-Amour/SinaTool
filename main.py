# -*- coding: utf-8 -*-

"""Get all Sina posts of somebody from id.
表中的内容是用Url方式表示的字段，是还没有下载图片的字段

Classes:

Functions:


"""

import re
import os
import math
import requests
import shutil
import time

import xlwings as xw

import user_ids

blog_num_max = 500
blog_report_num_max = 5
blog_comment_num_max = 500


def main():
    for user_id in user_ids.ids:
        print u'请输入微博用户ID：', user_id
        usr_info(user_id)


def convert_dict_gbk(input_dict):
    ret_dic = {}
    for dict_key, dict_value in input_dict:
        ret_dic[dict_key.decode('utf-8').encode('gbk')] = dict_value
    return ret_dic


def format_content(content):
    """
    对需要保存的内容进行过滤，去掉<a>这种标签
    :param content:
    :return: 返回过滤后的字符串
    """
    ret = re.sub(ur'(<span .*?>)', '', content)
    ret = re.sub(ur'(<img .*?>)', '', ret)
    ret = re.sub(ur'(</span>)', '', ret)
    ret = re.sub(ur'(<a .*?>)', '', ret)
    ret = re.sub(ur'(</a>)', '', ret)
    return ret


def format_excel_str(text):
    """
    Excel字符串，在文本前加'，防止公式错误卡死的情况
    :param text:
    :return:
    """
    if type(text) == unicode or type(text) == str:
        return '\'' + text
    return text


def download_image(pic_url, dir):
    return
    file_path = dir + get_file_name_from_url(pic_url)
    if not os.path.exists(file_path):
        with open(file_path, 'wb') as handle:
            try:
                response = requests.get(pic_url, stream=True, headers={'Connection': 'close'})
                if not response.ok:
                    print response
                for block in response.iter_content(1024):
                    if not block:
                        break
                    handle.write(block)
            except:
                return


def get_file_name_from_url(pic_url):
    return pic_url.split('/')[-1]


def save_comments(user_info, wb_id):
    """
    Get all comments from one post wb_id.
    :param user_info:
    :param wb_id: 某一个帖子的ID
    :return:
    """
    user_dir_path = './data_exported/' + user_info['user_name'] + '/'
    if os.path.exists(user_dir_path + 'comment-' + wb_id + '.xlsx'):
        return

    url = 'https://m.weibo.cn/api/comments/show?id={id}'.format(id=wb_id)
    page_url = 'https://m.weibo.cn/api/comments/show?id={id}&page={page}'
    response_raw = requests.get(url, headers={'Connection': 'close'})
    if not response_raw.text:
        print u'Comment not get: %d' % wb_id
        print u'被封了！！！！！'
        raise BaseException
        return
    response = response_raw.json()
    if not response.get('data', None):
        return
    global blog_comment_num_max
    page_num = min(response['data']['max'], int(math.ceil(blog_comment_num_max / 10.0)))
    if page_num <= 0:
        page_num = 1
    # 打开表格文档，格式化好需要的表格的列名和
    xl_book = xw.Book()
    xl_sheet = xl_book.sheets['Sheet1']
    try:
        xl_sheet.range('A1').value = [
            u'评论人ID',
            u'评论人名字',
            u'评论人头像',
            u'评论人主页',
            u'评论ID',
            u'评论时间',
            u'点赞数',
            u'评论内容',
            u'回复的父评论ID',
            u'回复的父评论内容'
        ]
        line_index = 2
        for i in range(1, page_num, 1):
            p_url = page_url.format(id=wb_id, page=i)
            try:
                resp = requests.get(p_url, headers={'Connection': 'close'})
            except:
                continue
            resp_data = resp.json()
            data = resp_data.get('data')
            if data:
                hot_data = data.get('hot_data')
                if hot_data:
                    for comment in hot_data:
                        xl_sheet.range('A' + str(line_index)).value = [
                            format_excel_str('id:' + str(comment['user']['id'])),
                            format_excel_str(comment['user']['screen_name']),
                            format_excel_str(comment['user']['profile_image_url']),
                            format_excel_str(comment['user']['profile_url']),
                            format_excel_str('id:' + str(comment['id'])),
                            format_excel_str(comment['created_at']),
                            format_excel_str(comment['like_counts']),
                            format_excel_str(format_content(comment['text'])),
                            format_excel_str('id:' + str(comment.get('reply_id', 'No'))),
                            format_excel_str(format_content(comment.get('reply_text', '')))
                        ]
                        # 下载头像
                        download_image(comment['user']['profile_image_url'], u'./profile_images/small/')
                        # 头像添加超链接
                        image_file = get_file_name_from_url(comment['user']['profile_image_url'])
                        url = u'../../profile_images/small/' + image_file
                        xl_sheet.range('C' + str(line_index)).add_hyperlink(url, image_file)
                        line_index += 1
                for comment in data['data']:
                    xl_sheet.range('A' + str(line_index)).value = [
                        format_excel_str('id:' + str(comment['user']['id'])),
                        format_excel_str(comment['user']['screen_name']),
                        format_excel_str(comment['user']['profile_image_url']),
                        format_excel_str(comment['user']['profile_url']),
                        format_excel_str('id:' + str(comment['id'])),
                        format_excel_str(comment['created_at']),
                        format_excel_str(comment['like_counts']),
                        format_excel_str(format_content(comment['text'])),
                        format_excel_str('id:' + str(comment.get('reply_id', 'No'))),
                        format_excel_str(format_content(comment.get('reply_text', '')))
                    ]
                    # 下载头像
                    download_image(comment['user']['profile_image_url'], u'./profile_images/small/')
                    # 头像添加超链接
                    image_file = get_file_name_from_url(comment['user']['profile_image_url'])
                    url = u'../../profile_images/small/' + image_file
                    xl_sheet.range('C' + str(line_index)).add_hyperlink(url, image_file)
                    line_index += 1
        # time.sleep(1)
    finally:
        user_dir_path = './data_exported/' + user_info['user_name'] + '/'
        xl_book.save(user_dir_path + 'comment-' + wb_id)
        xl_book.close()


def save_hot_comments(user_info, wb_id):
    """
    Get all hot comments from one post wb_id.
    热评
    :param user_info:
    :param wb_id: 某一个帖子的ID
    :return:
    """
    user_dir_path = './data_exported/' + user_info['user_name'] + '/'
    if os.path.exists(user_dir_path + 'comment-hot-' + wb_id + '.xlsx'):
        return

    test_url = 'https://m.weibo.cn/single/rcList?id={id}&type=comment&hot=1&tab=1'.format(id=wb_id)
    test_response = requests.get(test_url, headers={'Connection': 'close'})
    if not test_response.text:
        return
    url = 'https://m.weibo.cn/single/rcList?format=cards&id={id}&type=comment&hot=1&page=1'.format(id=wb_id)
    page_url = 'https://m.weibo.cn/single/rcList?format=cards&id={id}&type=comment&hot=1&page={page}'
    response = requests.get(url, headers={'Connection': 'close'})
    if not response.text:
        return
    global blog_comment_num_max
    page_num = min(response.json()[1]['maxPage'], int(math.ceil(blog_comment_num_max / 10.0)))
    if page_num <= 0:
        page_num = 1
    # 打开表格文档，格式化好需要的表格的列名和
    xl_book = xw.Book()
    xl_sheet = xl_book.sheets['Sheet1']
    try:
        xl_sheet.range('A1').value = [
            u'评论人ID',
            u'评论人名字',
            u'评论人头像',
            u'评论人主页',
            u'评论ID',
            u'评论时间',
            u'点赞数',
            u'评论内容',
            u'回复的父评论ID',
            u'回复的父评论内容'
        ]
        line_index = 2
        for i in range(1, page_num, 1):
            p_url = page_url.format(id=wb_id, page=i)
            resp = requests.get(p_url, headers={'Connection': 'close'})
            if not resp.text:
                break
            resp_data = resp.json()
            data = None
            if i == 1:
                data = resp_data[1].get('card_group')
            else:
                data = resp_data[0].get('card_group')
            if data:
                for comment in data:
                    xl_sheet.range('A' + str(line_index)).value = [
                        format_excel_str('id:' + str(comment['user']['id'])),
                        format_excel_str(comment['user']['screen_name']),
                        format_excel_str(comment['user']['profile_image_url']),
                        format_excel_str(comment['user']['profile_url']),
                        format_excel_str('id:' + str(comment['id'])),
                        format_excel_str(comment['created_at']),
                        format_excel_str(comment['like_counts']),
                        format_excel_str(format_content(comment['text'])),
                        format_excel_str('id:' + str(comment.get('reply_id', 'No'))),
                        format_excel_str(format_content(comment.get('reply_text', '')))
                    ]
                    # 下载头像
                    download_image(comment['user']['profile_image_url'], u'./profile_images/small/')
                    # 头像添加超链接
                    image_file = get_file_name_from_url(comment['user']['profile_image_url'])
                    url = u'../../profile_images/small/' + image_file
                    xl_sheet.range('C' + str(line_index)).add_hyperlink(url, image_file)
                    line_index += 1
            time.sleep(5)
    finally:
        user_dir_path = './data_exported/' + user_info['user_name'] + '/'
        xl_book.save(user_dir_path + 'comment-hot-' + wb_id)
        xl_book.close()


def save_reports(user_info, wb_id):
    """
    Get all comments from one post wb_id.
    :param user_info:
    :param wb_id: 某一个帖子的ID
    :return:
    """
    user_dir_path = './data_exported/' + user_info['user_name'] + '/'
    if os.path.exists(user_dir_path + 'report-' + wb_id + '.xlsx'):
        return

    url = 'https://m.weibo.cn/api/statuses/repostTimeline?id={id}'.format(id=wb_id)
    page_url = 'https://m.weibo.cn/api/statuses/repostTimeline?id={id}&page={page}'
    response = requests.get(url, headers={'Connection': 'close'}).json()
    if not response.get('data', None):
        return
    global blog_report_num_max
    page_num = min(response['data']['max'], int(math.ceil(blog_report_num_max / 10.0)))
    if page_num <= 0:
        page_num = 1
    # 打开表格文档，格式化好需要的表格的列名和
    xl_book = xw.Book()
    xl_sheet = xl_book.sheets['Sheet1']
    try:
        xl_sheet.range('A1').value = [
            u'评论人ID',
            u'评论人名字',
            u'评论人头像',
            u'评论人主页',
            u'评论ID',
            u'评论时间',
            u'点赞数',
            u'评论内容',
        ]
        line_index = 2
        for i in range(1, page_num, 1):
            p_url = page_url.format(id=wb_id, page=i)
            resp = requests.get(p_url, headers={'Connection': 'close'})
            resp_data = resp.json()
            data = resp_data.get('data')
            if data:
                for comment in data['data']:
                    xl_sheet.range('A' + str(line_index)).value = [
                        format_excel_str('id:' + str(comment['user']['id'])),
                        format_excel_str(comment['user']['screen_name']),
                        format_excel_str(comment['user']['profile_image_url']),
                        format_excel_str(comment['user']['profile_url']),
                        format_excel_str('id:' + str(comment['id'])),
                        format_excel_str(comment['created_at']),
                        format_excel_str(comment['like_counts']),
                        format_excel_str(format_content(comment['text'])),
                    ]
                    # 下载头像
                    download_image(comment['user']['profile_image_url'], u'./profile_images/small/')
                    # 头像添加超链接
                    image_file = get_file_name_from_url(comment['user']['profile_image_url'])
                    url = u'../../profile_images/small/' + image_file
                    xl_sheet.range('C' + str(line_index)).add_hyperlink(url, image_file)
                    line_index += 1
        # time.sleep(1)
    finally:
        user_dir_path = './data_exported/' + user_info['user_name'] + '/'
        xl_book.save(user_dir_path + 'report-' + wb_id)
        xl_book.close()


def save_blog_list(user_info):
    base_url = 'https://m.weibo.cn/api/container/getIndex?containerid={oid}&type=uid&value={uid}'
    page_url = 'https://m.weibo.cn/api/container/getIndex?containerid={oid}&type=uid&value={uid}&page={page}'
    url = base_url.format(oid=user_info['container_id'], uid=user_info['user_id'])
    resp = requests.get(url, headers={'Connection': 'close'})
    response = resp.json()
    global blog_num_max
    total = min(response['data']['cardlistInfo']['total'], blog_num_max)
    page_num = int(int(total) / 10) + 1
    # 打开表格文档，格式化好需要的表格的列名和
    xl_book = xw.Book()
    xl_sheet = xl_book.sheets['Sheet1']
    try:
        xl_sheet.range('A1').value = [
            u'发布ID',
            u'发布时间',
            u'点赞数',
            u'转发数',
            u'评论数',
            u'发布内容',
            u'转发表（链接）',
            u'评论表（链接）',
        ]
        line_index = 2
        for i in range(1, page_num, 1):
            p_url = page_url.format(oid=user_info['container_id'], uid=user_info['user_id'], page=i)
            page_resp = requests.get(p_url, headers={'Connection': 'close'})
            page_data = page_resp.json()
            cards = page_data['data']['cards']
            for card in cards:
                if card['card_type'] == 9:
                    mblog = card['mblog']
                    xl_sheet.range('A' + str(line_index)).value = [
                        format_excel_str('id:' + str(mblog['id'])),
                        format_excel_str(mblog['created_at']),
                        format_excel_str(mblog['attitudes_count']),
                        format_excel_str(mblog['reposts_count']),
                        format_excel_str(mblog['comments_count']),
                        format_excel_str(format_content(mblog['text'])),
                    ]
                    hyper_link_report = 'G' + str(line_index)
                    hyper_link_comment = 'H' + str(line_index)
                    line_index += 1
                    save_comments(user_info, mblog['id'])
                    xl_sheet.range(hyper_link_comment).add_hyperlink(u'comment-' + mblog['id'] + '.xlsx')
                    save_reports(user_info, mblog['id'])
                    xl_sheet.range(hyper_link_report).add_hyperlink(u'report-' + mblog['id'] + '.xlsx')
        # time.sleep(1)
    finally:
        user_dir_path = './data_exported/' + user_info['user_name'] + '/'
        xl_book.save(user_dir_path + 'posts')
        # xl_book.close()
        app = xw.apps.active
        app.quit()


def usr_info(usr_id):
    url = 'https://m.weibo.cn/api/container/getIndex?type=uid&value={usr_id}'.format(usr_id=usr_id)
    response = requests.get(url, headers={'Connection': 'close'})
    json_data = response.json()['data']
    # 每一个用户创建一个新的文件夹，如果已经存在删除重新建
    # 这里response导出的Json的字符串都是Unicode
    user_name = json_data.get('userInfo').get('screen_name')
    user_dir_path = './data_exported/' + user_name

    if not os.path.exists(user_dir_path):
        os.mkdir(user_dir_path + '/')

    # 把用户的数据写到一个文件里，以Key-Value的方式，写到excel文件，文件名就是这个微博账号的名字
    xl_book = xw.Book()
    xl_sheet = xl_book.sheets['Sheet1']
    try:
        # xl_sheet.range('A1').value = json_data.get('userInfo').get('screen_name')
        xl_sheet.range('A1').value = {
            u'用户ID': 'id:' + str(json_data.get('userInfo').get('id')),
        }
        xl_sheet.range('A2').value = {
            u'名字': format_excel_str(user_name),
        }
        xl_sheet.range('A3').value = {
            u'小头像': format_excel_str(json_data.get('userInfo').get('profile_image_url')),
        }
        download_image(json_data.get('userInfo').get('profile_image_url'), './profile_images/small/')
        image_name = get_file_name_from_url(json_data.get('userInfo').get('profile_image_url'))
        xl_sheet.range('B3').add_hyperlink(u'../../profile_images/small/' + image_name, image_name)

        xl_sheet.range('A4').value = {
            u'大头像': format_excel_str(json_data.get('userInfo').get('avatar_hd')),
        }
        xl_sheet.range('A5').value = {
            u'微博主页': format_excel_str(json_data.get('userInfo').get('profile_url')),
        }
        xl_sheet.range('A6').value = {
            u'主页背景': format_excel_str(json_data.get('userInfo').get('cover_image_phone')),
        }
        xl_sheet.range('A7').value = {
            u'发布状态数': format_excel_str(json_data.get('userInfo').get('statuses_count')),
        }
        xl_sheet.range('A8').value = {
            u'个人描述': format_excel_str(json_data.get('userInfo').get('description')),
        }
        xl_sheet.range('A9').value = {
            u'性别': format_excel_str(json_data.get('userInfo').get('gender')),
        }
        xl_sheet.range('A10').value = {
            u'粉丝': format_excel_str(json_data.get('userInfo').get('followers_count')),
        }
        xl_sheet.range('A11').value = {
            u'关注': format_excel_str(json_data.get('userInfo').get('follow_count')),
        }
        if isinstance(json_data.get('tabsInfo').get('tabs'), list):
            xl_sheet.range('A12').value = {
                u'贴子索引ID': format_excel_str('id:' + str(json_data.get('tabsInfo').get('tabs')[1].get('containerid'))),
            }
        elif isinstance(json_data.get('tabsInfo').get('tabs'), dict):
            xl_sheet.range('A12').value = {
                u'贴子索引ID': format_excel_str('id:' + str(json_data.get('tabsInfo').get('tabs')['1'].get('containerid'))),
            }
        else:
            raise TypeError()
    finally:
        xl_book.save(user_dir_path + '/' + json_data.get('userInfo').get('screen_name'))
        xl_book.close()

    user_id = json_data.get('userInfo').get('id')
    if isinstance(json_data.get('tabsInfo').get('tabs'), list):
        container_id = json_data.get('tabsInfo').get('tabs')[1].get('containerid')
    elif isinstance(json_data.get('tabsInfo').get('tabs'), dict):
        container_id = json_data.get('tabsInfo').get('tabs')['1'].get('containerid')
    else:
        raise TypeError()
    info = {
        'user_id': user_id,
        'container_id': container_id,
        'user_name': user_name,
    }
    save_blog_list(info)
    return info


if __name__ == '__main__':
    main()
