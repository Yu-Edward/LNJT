import selenium,time,sqlite3,sys,PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.common.action_chains import   ActionChains



def select_db(n):
    '''
    获取数据库中运营商对应的Xpath，在SQLite3中
    :param n: SQL
    :return:
    '''
    pythondb = sqlite3.connect("D://python//Ve_Regu_Details.db")
    cursor = pythondb.cursor()
    sql = n
    cursor.execute(sql)
    values = cursor.fetchall()
    cursor.close()
    pythondb.close()
    return values

def checkPosition(Operator_name,Operator_xpath):
    '''
    运营商平台监控页面执行查岗操作
    :param Operator_name:运营商名称
    :param Operator_xpath:运营商Xpath
    :return:
    '''
    # 运营商平台
    try:
        link = browser.find_element_by_xpath(Operator_xpath)
        time.sleep(0.1)
        ActionChains(browser).click(link).perform()
    except:
        return
    time.sleep(0.5)
    img_src = browser.find_element_by_xpath('{}/img'.format(Operator_xpath))
    if img_src.get_attribute("src") == 'http://218.60.150.130:18080/lngps/res/img/index/pt_ioc2.gif':
        offline.append('{}{:>30}'.format(link.text,'状态：离线'))
        print('{}{:>30}'.format(link.text,'状态：离线'))
        return
    print(link.text)
    online.append(link.text)

    # 查岗
    link = browser.find_element_by_xpath('//*[@id="userListShow"]/div[1]/div/a[3]')
    time.sleep(0.1)
    ActionChains(browser).click(link).perform()
    time.sleep(0.5)

    # 下发
    browser.find_element_by_xpath('//*[@id="popInfoPlatForm"]/div/div/div/a[1]').click()
    time.sleep(0.5)

    # 告警提示弹窗
    alert = browser.switch_to.alert
    #获取警告提示信息
    # alert_text = alert.text
    # print(alert_text)
    time.sleep(0.5)
    #接取警告框，并关闭警告框
    alert.accept()
    return link.text

username = 'lnyg_yg'
password = 'yg_123'

# 加载火狐浏览器驱动
# browser = webdriver.Firefox()
# 加载谷歌浏览器
browser = webdriver.Chrome()

# 全屏打开
browser.maximize_window()
# 等待
browser.implicitly_wait(20)

# 打开网页
# browser.get('http://218.60.150.130:18080/lngps/')
browser.get('http://218.60.150.130:18080/lngps/')
# 打印网页title
print(browser.title)


#登录
browser.find_element_by_id("loginName").send_keys(username)
browser.find_element_by_id("loginPwd").send_keys(password)
browser.find_element_by_id("loginBtn").click()
time.sleep(1)
try:
    browser.find_element_by_xpath('//*[@id="loginForceBtn"]').click()
except:
    pass

try:
    user_name = browser.find_element_by_xpath('//*[@id="localUserInfo"]')
    print(user_name,user_name.text)
    print("login success") if user_name.text == username else print("login fail")
    operatorXpath = select_db('SELECT * FROM Operator_List_Detail')
except:
    browser.close()
    print('login fail! System out.')
    sys.exit()

time.sleep(5)

flag = True
counter = 0

online = []
offline = []
while flag:
    try:
        if counter == 10:
            sys.exit(0)
        # 平台监管
        link = browser.find_element_by_xpath('//*[@id="frame_tabs_top"]/li[3]/div[1]/label/em')
        time.sleep(0.1)
        ActionChains(browser).move_to_element(link).perform()
        time.sleep(0.5)

        # 运营商平台监管
        link = browser.find_element_by_xpath('//*[@id="menuVehMonitor3"]/ul[1]/li[1]/a/label')
        time.sleep(0.1)
        ActionChains(browser).click(link).perform()
        time.sleep(2)


        flag = False
    except:
        if counter == 10:
            sys.exit(0)
        counter += 1
        time.sleep(2)
        print(flag,counter)
        continue


for i in range(len(operatorXpath)):
    checkPosition(operatorXpath[i][0], operatorXpath[i][1])



print('Over')


time.sleep(1)
# browser.close()

[offline.append(i) for i in online]



sg.popup('\n'.join(offline) + "\n共有{}个运营商".format(len(offline)),title='服务商',keep_on_top = True)