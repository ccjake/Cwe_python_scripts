import random
import pandas as pd
# letto_dict = {"a":"三等奖","b":"二等奖","c":"一等奖"}
cw_de = [
        "李明轩", "王子琪", "张博文", "刘佳慧", "陈浩宇",
        "杨乐瑶", "吴宇航", "徐静雯", "孙嘉诚", "郭悦然",
        "刘思雨", "陈伟杰", "杨雨欣", "吴俊熙", "徐梦瑶",
        "孙浩然", "郭子墨", "高诗涵", "林子豪", "郑晓琳",
        "张子涵", "赵心怡", "黄家豪", "周雅雯", "孙文涛",
        "郭晓婷", "高远航", "林雪儿", "郑志强", "谢佳琪"
    ]

cw_cn = [
        "赵瑞阳", "黄欣月", "周翰文", "孙婉清", "魏子杰",
        "蒋雨薇", "韩明哲", "谭诗雨", "卢浩然", "彭雅萱",
        "邓欣怡", "许志远", "付梓萌", "潘思琪", "丁俊霖",
        "邹佳妮", "薛宇辰", "杜悦溪", "钟文博", "姜晓薇",
        "冯明哲", "曹梦佳", "曾宇轩", "梁若曦", "宋昊天",
        "唐语嫣", "许梓豪", "邵思颖", "金奕博", "阎雨泽"
    ]

letto_dict = {}

cw_sum = cw_de + cw_cn
good_prize = random.sample(cw_sum, 10)
cw_sum = [cw for cw in cw_sum if cw not in good_prize]
cw_de = [cw for cw in cw_de if cw not in good_prize]
cw_cn = [cw for cw in cw_cn if cw not in good_prize]
for n in good_prize:
    letto_dict[n] = "参与奖"


third_prize = random.sample(cw_cn, 5)
cw_cn = [cw for cw in cw_cn if cw not in third_prize]
for n in third_prize:
    letto_dict[n] = "三等奖"


second_prize = random.sample(cw_de, 3)
cw_de = [cw for cw in cw_de if cw not in second_prize]
for n in second_prize:
    letto_dict[n] = "二等奖"



first_prize = random.sample(cw_de, 2)
cw_de = [cw for cw in cw_de if cw not in first_prize]
for n in first_prize:
    letto_dict[n] = "一等奖"

    
top_prize = random.sample(cw_de, 1)
cw_de = [cw for cw in cw_de if cw not in top_prize]
for n in top_prize:
    letto_dict[n] = "特等奖"

# print(letto_dict.keys())
letto_order_list = list(letto_dict.keys())
letto_dict.update({"a":"三等奖","b":"二等奖","c":"一等奖"})
# letto_order_list.reverse()
letto_order_list.insert(14, "a")
letto_order_list.insert(18, "b")
letto_order_list.insert(21, "c")

print(letto_dict)
ordered_letto_dict = {}
for name in letto_order_list:
    ordered_letto_dict[name] = letto_dict[name]

pd.DataFrame.from_dict(ordered_letto_dict, orient="index")
