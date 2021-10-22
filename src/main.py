import random
import datetime
import xlwt


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    week = ['一', '二', '三', '四', '五']
    breakfast = ['豆浆', '粥', '馍馍', '牛乳（或酸奶）', '煮荷包蛋', '大白菜猪肉包子', '玉米窝窝头']
    breakfast_must = ['鸡蛋', '馒头']
    meat_dishes = ['蒜苗回锅肉', '西兰花油豆腐', '芋头炖排骨', '青椒炒鸭肉', '香菇菜心', '虾皮冬瓜', '肉末茄子', '葱花土豆泥']
    vegetable_dish = ['青菜', '胡萝卜']

    current_time = datetime.datetime.now().microsecond
    random.seed(current_time)
    newb = xlwt.Workbook(encoding='utf-8')
    nws = newb.add_sheet("一周餐谱")

    nws.write(1, 0, '早饭')
    nws.write(2, 0, '午饭')
    nws.write(3, 0, '晚饭')

    for day in range(5):
        breakfast_option = random.choice(breakfast)
        noon_meat_dishes_options = random.choices(meat_dishes, k=2)
        noon_vegetable_dish_options = random.choice(vegetable_dish)
        night_meat_dishes_options = random.choices(meat_dishes, k=2)
        night_vegetable_dish_options = random.choice(vegetable_dish)

        col = day + 1
        nws.write(0, col, f'星期{week[day]}')
        nws.write(1, col, f'{breakfast_option} + {breakfast_must}')
        nws.write(2, col, f'{noon_meat_dishes_options} + {noon_vegetable_dish_options}')
        nws.write(3, col, f'{night_meat_dishes_options} + {night_vegetable_dish_options}')

    nws.col(day).width = 256 * 3
    for day in range(5):
        col = day + 1
        nws.col(col).width = 256 * 35
    for row in range(4):
        nws.row(row).height_mismatch = True
        nws.row(row).height = 20 * 40

    alignment = xlwt.Alignment()
    alignment.horz = 0x01
    alignment.vert = 0x01
    alignment.wrap = 1

    newb.save(f'一周餐谱_{current_time}.xls')
