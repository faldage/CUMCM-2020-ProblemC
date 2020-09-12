info = dict()
# 添加
info[0] = {'农', '林', '牧', '渔'}
info[1] = {'矿'}
info[2] = {'食品', '服装', '药', '制造'}
info[3] = {'电力', '热力', '燃气', '水'}
info[4] = {'建筑', '土木', '建设'}
info[5] = {'批发', '零售'}
info[6] = {'交通', '运输', '邮政'}
info[7] = {'住宿', '餐饮'}
info[8] = {'信息', '软件', '科技'}
info[9] = {'金融', '保险'}
info[10] = {'房地产'}
info[11] = {'租赁', '服务'}
info[12] = {'研究', '技术', '科技'}
info[13] = {'水利', '生态', '环境', '公共设施', '土地'}
info[14] = {'居民服务', '修理'}
info[15] = {'教育'}
info[16] = {'卫生', '社会工作'}
info[17] = {'新闻', '出版', '广播', '电视', '电影', '录音', '文化', '艺术', '体育', '娱乐'}
info[18] = {'机关', '机构', '组织'}
info[19] = {'国际组织'}

import datetime


begin = datetime.date(2017, 1, 1)
end = datetime.date(2020, 12, 31)
for i in range((end - begin).days+1):
    day = begin + datetime.timedelta(days=i)
    print(str(day))

