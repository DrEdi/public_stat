import requests, time, xlwt

# How to get token see README.md
token = ''
public_id = ''


class VkInfo:

    def __init__(self, token):
        self.__token = token

    def get_public_stat(self, owner_id):
        followers = self.__get_user_list(owner_id)
        wall_count = self.__get_wall_count(owner_id)
        wall_offset = 0

        while wall_offset <= wall_count:
            print(wall_offset, 'posts got')
            wall = self.__get_wall(owner_id, 1000, wall_offset)
            for post in wall:
                print(post['id'])
                comment_offset = 0
                while comment_offset <= post['comments']['count']:
                    try:
                        comments = self.__get_post_comments(owner_id, post['id'], comment_offset)
                    except Exception:
                        # if we got Exception that means we trid to get so much responses from server
                        # just wait for 1 second and server give us ordinagy response                        
                        time.sleep(1)
                        comments = self.__get_post_comments(owner_id, post['id'], comment_offset)
                    comment_offset += 100
                    for comment in comments:
                        if comment['from_id'] in followers:
                            followers[comment['from_id']]['comments'] += 1
                            followers[comment['from_id']]['charCount'] += len(comment['text'])
                            followers[comment['from_id']]['likedByPeople'] += comment['likes']['count']
                        else:
                            followers[comment['from_id']] = {
                                'comments': 1,
                                'charCount': len(comment['text']),
                                'is_member': False,
                                'published': 0,
                                'likedByPeople': comment['likes']['count']
                            }
                    # We must use sleep because of you have weak app server can give you only 3 responses
                    # by 1 second. Python works much more fater, so we have to wait a bit
                    time.sleep(0.2)
                try:
                    # This point is a bit curious
                    # Only if post were posted wih sign, it'll have a dict 'signer_id'
                    # if not, we could get an exeption
                    print(post['signer_id'])
                    if post['signer_id'] in followers:
                        followers[post['signer_id']]['published'] += 1
                        followers[post['signer_id']]['charCount'] += len(post['text'])
                        followers[post['signer_id']]['likedByPeople'] += post['likes']['count']
                    else:
                        followers[post['signer_id']] = {
                            'comments': 0,
                            'postCommented': 0,
                            'charCount': len(post['text']),
                            'is_member': False,
                            'published': 1,
                            'likedByPeople': post['likes']['count']
                        }
                except Exception:
                    print('This is post has no sign')

            wall_offset += 100
            time.sleep(0.1)

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Test')
        index = 1
        columns = ['Full name',
                   'Comments count',
                   'Middle count of chars',
                   'Is member',
                   'Count of published posts',
                   'Count of likes user\'s got']
        for i in range(len(columns)):
            ws.write(0, i, columns[i])

        ids = [user_id for user_id in followers]
        user_offset = 0
        while user_offset < len(ids):
            print(ids[user_offset:user_offset+1000])
            users_info = self.__get_user_info(ids[user_offset:user_offset+1000])
            print(users_info)
            for user in users_info:
                ws.write(index, 0, str(user['first_name'] + ' ' + user['last_name']))
                ws.write(index, 1, followers[user['id']]['comments'])
                try:
                    ws.write(index, 2, followers[user['id']]['charCount']/(followers[user['id']]['comments'] +
                                                                           followers[user['id']]['published']))
                except ZeroDivisionError:
                    ws.write(index, 2, 0)
                ws.write(index, 3, followers[user['id']]['is_member'])
                ws.write(index, 4, followers[user['id']]['published'])
                ws.write(index, 5, followers[user['id']]['likedByPeople'])
                index += 1
            user_offset += 1000
            time.sleep(1)
        wb.save('xl_rec.xls')

    def __get_group_count(self, owner_id):
        info = requests.get('https://api.vk.com/method/{method}'.format(method='groups.getById'), params={
            'group_id': -owner_id,
            'token': self.__token,
            'fields': 'members_count',
            'v': '5.38'})
        return info.json()['response'][0]['members_count']

    def __get_wall_count(self, owner_id):
        info = requests.get('https://api.vk.com/method/{method}'.format(method='wall.get'), params={
            'owner_id': owner_id,
            'token': self.__token,
            'v': '5.38'})
        return info.json()['response']['count']

    def __get_user_list(self, owner_id):
        follower_list = {}
        user_offset = 0
        followers_count = self.__get_group_count(owner_id)
        while user_offset <= followers_count:
            print(user_offset, 'users got')
            group_members = requests.get('https://api.vk.com/method/{method}'.format(method='groups.getMembers'),
                                         params={
                                            'group_id': -owner_id,
                                            'count': 1000,
                                            'offset': user_offset,
                                            'access_token': token,
                                            'v': '5.38'})
            for member in group_members.json()['response']['items']:
                follower_list[member] = {'comments': 0,
                                         'postCommented': 0,
                                         'charCount': 0,
                                         'is_member': True,
                                         'published': 0,
                                         'likedByPeople': 0}
            time.sleep(0.3)
            user_offset += 1000
        return follower_list

    def __get_wall(self,owner_id, count=10, offset=0):
        wall_get = requests.get('https://api.vk.com/method/{method}'.format(method='wall.get'), params={
            'owner_id': owner_id,
            'count': 100,
            'offset': offset,
            'access_token': token,
            'v': '5.38'}).json()['response']['items']
        return wall_get

    def __get_post_comments(self,owner_id, post_id, offset=0):
        post_get = requests.get('https://api.vk.com/method/{method}'.format(method='wall.getComments'), params={
            'owner_id': owner_id,
            'post_id': post_id,
            'access_token': token,
            'offset': offset,
            'preview_length': 0,
            'count': 100,
            'need_likes': 1,
            'v': '5.38'}).json()['response']['items']
        return post_get

    def __get_user_info(self, user_ids):
        user_info = requests.get('https://api.vk.com/method/{method}'.format(method='users.get'), params={
            'user_ids': str(user_ids).replace('[', '').replace(']', ''),
            'v': '5.38'}).json()['response']
        return user_info

info = VkInfo(token)
print(info.get_public_stat(public_id))
