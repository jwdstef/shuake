import re
import time
from time import ctime, sleep
from seleniumwire import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from PIL import Image
from log import GetLog
from mySqlHelper import MySqLHelper

import json
import requests
import pymysql
import traceback
import threading
import queue
import hashlib


# 优化的点
# 答题的时候，把没找到答案的题目记录下来存库。
# 可以设置不交卷，就搂题目
# 增加日志打印到文件
# 部分等待加载页面改用waitPageLoad方法，看情况改
#

# 安装selenium
# pip install selenium

# 安装获取请求的工具
# pip install selenium-wire

# 安装截图
# pip install pillow


# -----------链接参数-----------
# 首页  https://menhu.pt.ouchn.cn/
homePageUrl = 'https://menhu.pt.ouchn.cn/'
# 查题链接
queryAnswerUrl = 'http://121.37.181.45:8083/jeecg-boot/weTiku/postChaTi'
# 课程完成库
coursewaresUrl = "https://lms.ouchn.cn/api/course/"
docourseUrl = "https://lms.ouchn.cn/api/course/activities-read/"


# -----------刷课参数-----------
# 刷课开关
enableLesson = False
# 形考开关
enableExam = True
# 形考答题有答案的提交比例
examRate = 0.85
# 形考分数线 大于等于xx分不再答题
examLineScore = 75

# -----------数据库参数-----------
db_host = "47.95.194.30"
db_port = 3306
db_user = "root"
db_password = "Yuyan~9527"
db_database = "cargps"
userTable = "user2023_06"
tikuTable = "we_tiku"
db = MySqLHelper()


# -----------多线程-----------
# 队列
queue = queue.Queue()
# 定义线程数
threadCount = 1
# 线程启动之间的等待时间
threadSleepTime = 10
# 线程池子
threadPool = []


# -----------测试相关-----------
# 测试开关
testSwitch = True
# 测试的课程
testClassNames = ['形势与政策','电饭锅']
# 测试的考试ID
testExamId = []
# 是否隐藏浏览器
hiddenExplore = True


# 题型映射
subjectTypeMap = {
    'single_selection': '单选题',
    'multiple_selection': '多选题',
    'true_or_false': '判断题',
    'fill_in_blank': '填空题',
    'short_answer': '简答题',
    'text': '文本',
    'analysis': '综合题',
    'matching': '匹配题',
    'random': '随机题',
    'cloze': '完形填空题'}

# 初始化日志
logger = GetLog().get_log('')


# 初始化浏览器
def initExplore():
    if (hiddenExplore):
        options = {
            'request_storage': 'memory'  # Store requests and responses in memory only
        }
        # 打开浏览器 - 不弹出窗口  处理浏览器不被检测到
        chrome_opts = webdriver.ChromeOptions()
        chrome_opts.add_argument("--headless")
        driver = webdriver.Chrome(options=chrome_opts,seleniumwire_options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
                })
            """
        })
    else:
        # 打开浏览器 - 正常命令，弹出窗口
        driver = webdriver.Chrome()
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
                })
            """
        })
    driver.maximize_window()
    # 第一个tab
    firstExploreTab = driver.current_window_handle
    globalData = {'firstExploreTab': firstExploreTab, 'browser': driver}
    return globalData


# 登录用户
def loginUser(username, password, globalData, tryTimes, multiple):
    logger.info('首页地址：%s', homePageUrl)
    browser = globalData['browser']

    # 保存账号
    globalData['username'] = username

    # 打开页面
    try:
        browser.get(homePageUrl)
    except:
        logger.error('登录页面打开失败')
        raise

    # 检查元素  -- 等待页面加载出来  这种未通过直接退出
    waitPageLoad(browser, '', '#button', 10, '登录页加载失败')

    # 输入账号和密码
    browser.find_element(By.ID, "loginName").send_keys(username)
    browser.find_element(By.ID, "password").send_keys(password)

    checkCode = getKaptchaImage(browser, multiple)
    # 验证码 validateCode
    # 这个是WebDriverWait 的一种重要写法，不能删除留作备用，包括  checkCode() 方法
    # WebDriverWait(browser, 10).until(checkCode(), message="验证码填写等待超时")

    logger.info('正在尝试验证码 [' + checkCode + ']')
    while ((not re.match('^\d{4}$', checkCode)) and tryTimes < 5):
        tryTimes += 1
        refreshCheckCode(globalData)
        checkCode = getKaptchaImage(browser, multiple)
        logger.info('正在尝试验证码 [' + checkCode + ']')

    browser.find_element(By.ID, "validateCode").send_keys(checkCode)
    logger.info('验证码填写成功 %s', checkCode)

    # 点击登录
    loginButton = browser.find_element(By.ID, "button")
    loginButton.click()
    # 这个需要做一个优化，不做sleep
    sleep(10)

    # 如果验证码错误，会有这个    div.container_12_btns
    if (len(browser.find_elements(By.CSS_SELECTOR, 'div.container_12_btns')) > 0 and tryTimes < 5):
        tryTimes += 1
        loginUser(username, password, globalData, tryTimes, multiple)

    # 等待页面展开
    waitPageLoad(browser, '', '.ouchnPc_index_course_div', 10, '首页加载失败')

    # 提取token
    # sessionStorageJs = "return JSON.stringify(window.sessionStorage)"
    # sessionStorage = browser.execute_script(sessionStorageJs)
    # sessionStorage = json.loads(sessionStorage)

    # key: oidc.user:http://passport.ouchn.cn/:studentspace
    # token 字段是 access_token
    # sessionStorageStrValue = sessionStorage['oidc.user:http://passport.ouchn.cn/:studentspace']
    # token = json.loads(sessionStorageStrValue)['access_token']
    # print('token:', token)


# 检查验证码
class checkCode(object):
    def __init__(self):
        pass

    def __call__(self, driver):
        logger.info('checking ---')
        value = driver.find_element(By.ID,
                                    "validateCode").get_attribute('value')
        return len(value) == 4


# 查到所有课程
def findAllClass(globalData, classList):
    browser = globalData['browser']
    # 课程列表
    list = browser.find_element(By.CLASS_NAME,
                                'ouchnPc_index_course').find_elements(By.CLASS_NAME, 'ouchnPc_index_course_div')

    currentClassName = ''
    # 遍历处理课程
    for item in list:
        className = item.find_element(By.TAG_NAME, 'p').text.strip()
        classLink = item.find_element(By.TAG_NAME, 'a').get_attribute('href')
        logger.info(className)
        logger.info(classLink)
        currentClassName = className
        existsClass = [
            elem for elem in classList if elem['className'] == className]
        if (len(existsClass) > 0):
            continue
        classList.append({
            'className': className,
            'classLink': classLink,
            'finshExams': [],
            'isFinish': False,
            'examUrl': '',
            'tryTimes': 0
        })
    # button.btn-next:not([disabled])
    nextButtons = browser.find_elements(
        By.CSS_SELECTOR, 'button.btn-next:not([disabled])')
    if (len(nextButtons) > 0):
        nextButtons[0].click()
        sleep(5)
        for i in range(5):
            list = browser.find_element(By.CLASS_NAME, 'ouchnPc_index_course').find_elements(
                By.CLASS_NAME, 'ouchnPc_index_course_div')
            if (len(list) == 0):
                sleep(5)
                continue
            currentClassList = [elem for elem in list if elem.find_element(
                By.TAG_NAME, 'p').text.strip() == currentClassName]
            if (len(currentClassList) > 0):
                sleep(5)
                continue
            break
        classList = findAllClass(globalData, classList)
    return classList


# 处理课程
def dealClassItem(globalData, classData):

    browser = globalData['browser']
    className = classData['className']
    classLink = classData['classLink']

    globalData['className'] = className
    # 测试开关打开并且在测试列表，才会进行。或者测试关闭，正常运行
    if ((testSwitch and className in testClassNames) or not testSwitch):
        logger.info('当前处理的是 %s  %s ', className, classLink)

        # 打开课程页面
        browser.get(classLink)

        # 等待页面加载
        waitPageLoad(browser, classLink, '.activity-tab', 10, "页面加载失败")

        # 分支1，处理基础课程
        if (enableLesson):
            activitiesList = getAllActivity(browser)
            logger.info('刷课 课程数量 : %s' % len(activitiesList))
            cookie = getCookieFromClassPage(browser)
            logger.info('获取cookie : %s' % cookie)
            dealNormalClassItem(activitiesList, cookie)

        # 分支2，处理考试科目
        if (enableExam):
            times = 0
            while not classData['isFinish']:
                try:
                    # 打开课程页面
                    browser.get(classLink)
                    # 等待页面加载完成
                    waitPageLoad(browser, '', ".activity-tab", 10, "页面加载失败")
                    # 切换到形考任务tab
                    switchSuccess = classSwitchExam(globalData)
                    if (not switchSuccess):
                        return
                    classData['examUrl'] = browser.current_url
                    logger.info(classData)
                    # 处理考试
                    dealExamItem(globalData, classData)
                except Exception as e:
                    traceback.print_exc()
                    logger.exception("exec failed, failed msg:" + traceback.format_exc())
                    times += 1
                    if (times > 3):
                        classData['isFinish'] = True
                    pass


# 刷新验证码
def refreshCheckCode(globalData):
    browser = globalData['browser']
    browser.find_element(By.ID, 'kaptchaImage').click()
    sleep(5)


# 等待页面加载成功
def waitPageLoad(browser, customUrl, cssSelector, waitTime, errorMessage):
    # 重试次数控制
    totalTryTimes = 3
    times = 0
    success = False
    confirmPassword = False
    url = ''
    while not success:
        try:
            if (times > 0 and not confirmPassword):
                logger.info('当前加载页面失败，正在重试 %s-%s 次' % (totalTryTimes, times))
                url = customUrl if customUrl else url
                browser.get(url)
            WebDriverWait(browser, waitTime).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, cssSelector)), message=errorMessage)
            success = True
            logger.info('打开成功')
            return
        except Exception as e:
            # 密码失效提醒
            confirmATags = browser.find_elements(By.CSS_SELECTOR,"a.l-btn")
            if (len(confirmATags) > 0 and '确定' in confirmATags[0].text):
                confirmATags[0].click()
                times += 1
                confirmPassword = True
                continue
            else:
                confirmPassword = False

            times += 1
            if (not url):
                url = browser.current_url
            if (times > totalTryTimes):
                raise


# 关闭其他页面回到首页
def closeOtherTabsToHome(globalData):
    browser = globalData['browser']
    firstExploreTab = globalData['firstExploreTab']
    handles = browser.window_handles
    for h in handles:
        if (firstExploreTab != h):
            browser.switch_to.window(h)
            browser.close()
    browser.switch_to.window(firstExploreTab)

# ------------------------------------------------处理考试 开始-----------------------------------------------------


# 切换tabs到形考任务
def classSwitchExam(globalData):
    browser = globalData['browser']

    # 寻找tab
    tabs = browser.find_elements(By.CLASS_NAME, 'activity-tab')
    # sleep(5)
    for tab in tabs:
        if (tab.text == '形考任务'):
            sleep(3)
            tab.click()
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-label='形考任务'].active")), message="切换tab失败")

            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".learning-activity")), message="切换tab失败")
            return True
    return False


# 处理单个形考
def dealExamItem(globalData, classData):
    browser = globalData['browser']

    # 展开隐藏元素
    expendAllEle(browser)
    sleep(5)
    modules = browser.find_elements(By.CSS_SELECTOR,
                                    "div[class='module ng-scope ng-isolate-scope expanded']")

    for module in modules:

        # 单元名称
        moduleName = module.find_element(By.CLASS_NAME, "truncate-text").text
        # 形考列表
        # learning-activity ng-scope
        moduleActivities = module.find_elements(By.CSS_SELECTOR,"div.learning-activity.ng-scope")

        for moduleActivity in moduleActivities:
            globalData['activityName'] = moduleActivity.text
            moduleActivityName = moduleActivity.text
            moduleActivityId = moduleActivity.get_attribute('id')
            classStr = moduleActivity.get_attribute('class')
            moduleActivityClickDiv = moduleActivity.find_element(
                By.CLASS_NAME, 'clickable-area')
            # 本脚本已经处理过的
            if (moduleActivityId in classData['finshExams']):
                continue

            # 查询分数
            scoreDivs = moduleActivity.find_elements(
                By.CSS_SELECTOR, 'div.score.ng-scope')
            score = int(scoreDivs[0].text) if len(scoreDivs) > 0 and re.match(
                '^\d+$', scoreDivs[0].text) else -1

            #  题目已答题或者过分数线
            if ((len(moduleActivity.find_elements(By.CSS_SELECTOR, "span.submitted")) > 0 and score == -1) or score >= examLineScore or 'ng-hide' in classStr):
                logger.info('当前形考 [%s-%s] 分数未知并且已完成  或者  分数大于等于分数线 -----> 跳过  分数为：%s  分数线：%s' % (
                    moduleName, moduleActivityName, score, examLineScore))
                classData['finshExams'].append(moduleActivityId)
                continue
            logger.info('当前形考 [%s-%s] 分数未知并且未完成  或者  分数小于分数线 -----> 答题  分数为：%s  分数线：%s' % (
                moduleName, moduleActivityName, score, examLineScore))

            # 是否开启测试
            if (moduleActivityId not in classData['finshExams'] and ((testSwitch and moduleActivityId in testExamId) or (testSwitch and len(testExamId) == 0) or not testSwitch)):
                # 处理考试
                logger.info('准备完成形考：%s %s %s', moduleName,
                            moduleActivityName, moduleActivityId)
                try:
                    openExam(globalData, classData, moduleActivityClickDiv)
                except:
                    logger.error('考试失败1 - %s' % classData['tryTimes'])
                    classData['tryTimes'] = classData['tryTimes'] + 1
                    logger.error('考试失败2 - %s' % classData['tryTimes'])
                    traceback.print_exc()
                    logger.exception("exec failed, failed msg:" + traceback.format_exc())
                    if (classData['tryTimes'] < 3):
                        return
                classData['tryTimes'] = 0
                classData['finshExams'].append(moduleActivityId)
                return
            classData['finshExams'].append(moduleActivityId)
    classData['isFinish'] = True


# 打开考试并答题
def openExam(globalData, classData, examActity):
    browser = globalData['browser']
    # 打开考试
    examActity.click()
    waitPageLoad(browser, '', 'div.activity-menu-item', 15, '打开考试详情失败')
    # WebDriverWait(browser, 15).until(
    #     EC.presence_of_element_located((By.CSS_SELECTOR, "div.activity-menu-item")), message="打开考试详情失败")

    # 展开基本信息 todo
    # browser.find_elements(By.CSS_SELECTOR,'div.show-advanced')[0].click()

    # 点击开始答题
    startExamButtons = browser.find_elements(By.CSS_SELECTOR, 'a.take-exam')
    if (len(startExamButtons) == 0):
        return
    startExamButtons[0].click()

    # 等待弹窗 && 等待页面加载完成
    WebDriverWait(browser, 10).until(EC.visibility_of_element_located(
        (By.CSS_SELECTOR, '#start-exam-confirmation-popup')), message="考试须知打开失败")

    # 已知晓
    browser.find_element(By.CSS_SELECTOR, "input[name='confirm']").click()

    # 等待已知晓效果完成
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='confirm'].ng-not-empty")), message="已知晓效果失败")

    # 点击确定开始答题
    browser.find_element(By.CSS_SELECTOR,
                         "button.button-green.ng-binding").click()

    # 等待页面加载完成
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "li.subject")), message="考试页面加载失败")

    subjects = browser.find_elements(By.CSS_SELECTOR, "li.subject")

    # 题目总数
    totalCount = 0
    # 有答案的数量
    hasAnswerCount = 0

    for subject in subjects:
        # 慢点让我看看
        sleep(1)

        # 题干
        subjectTitle = getSubjectTitle(subject)

        subjectType = judgeSubjectType(subject.get_attribute('class'))

        if (subjectType not in 'ERROR' and subjectType not in 'text'):
            totalCount += 1

        # 查询答案
        answer = queryAnswer(subjectTitle, subjectType)
        if (not answer):
            answer = queryAnswerFromDb(subjectTitle, subjectType)
        if (not answer):
            # 没有答案或许可以做些社么
            continue
        hasAnswerCount += 1
        # 处理答案
        if (subjectType == 'true_or_false'):
            # 判断题
            lis = subject.find_elements(By.CSS_SELECTOR, "li.option")
            answerOptions = []
            for li in lis:
                itemText = getOptionsValue(li)
                if (itemText and itemText in answer):
                    answerOptions.append(li)

            if (len(answerOptions) > 0 and len(answerOptions[0].find_elements(By.CSS_SELECTOR, "input:checked")) == 0):
                answerOptionLabels = answerOptions[0].find_elements(
                    By.CSS_SELECTOR, "label")
                if (len(answerOptionLabels) > 0):
                    answerOptionLabels[0].click()
                    # 等待页面加载完成
                    WebDriverWait(subject, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".answered-option")), message="选项选择失败")

            pass
        elif (subjectType == 'single_selection'):
            # 单选题
            lis = subject.find_elements(By.CSS_SELECTOR, "li.option")
            answerOptions = []
            for li in lis:
                itemText = getOptionsValue(li)
                if (itemText and itemText in answer):
                    answerOptions.append(li)

            if (len(answerOptions) > 0 and len(answerOptions[0].find_elements(By.CSS_SELECTOR, "input:checked")) == 0):
                answerOptionLabels = answerOptions[0].find_elements(
                    By.CSS_SELECTOR, "label")
                if (len(answerOptionLabels) > 0):
                    answerOptionLabels[0].click()
                    # 等待页面加载完成
                    WebDriverWait(subject, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".answered-option")), message="选项选择失败")
            pass
        elif (subjectType == 'multiple_selection'):
            # 多选题
            lis = subject.find_elements(By.CSS_SELECTOR, "li.option")
            answerOptions = []
            for li in lis:
                itemText = getOptionsValue(li)
                if (itemText and itemText in answer):
                    answerOptions.append(li)

            for answerOption in answerOptions:
                answerOptionLabels = answerOption.find_elements(
                    By.TAG_NAME, "label")
                if (len(answerOptionLabels) > 0 and len(answerOption.find_elements(By.CSS_SELECTOR, "input:checked")) == 0):
                    answerOptionLabels[0].click()
                    WebDriverWait(answerOption, 15).until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "input:checked")), message="选项选择失败")
            pass
        elif (subjectType == 'cloze'):
            # 完形填空
            answerList = answer.split('|')
            selectPs = [elem for elem in subject.find_elements(By.CSS_SELECTOR, 'p') if len(elem.find_elements(
                By.CSS_SELECTOR, 'select')) > 0 and len(elem.find_elements(By.CSS_SELECTOR, 'button')) > 0]
            if (len(answerList) != len(selectPs)):
                hasAnswerCount -= 1
                continue
            for i in range(len(selectPs)):
                print(i)
                selectP = selectPs[i]
                # 检查是否打开了
                try:
                    selectP.find_elements(By.CSS_SELECTOR, 'button')[0].click()
                    WebDriverWait(selectP, 10).until(EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, 'div.ui-multiselect-menu')), message="打开下拉选失败")
                    multiselectCheckboxes = selectP.find_elements(
                        By.CSS_SELECTOR, 'ul.ui-multiselect-checkboxes')
                    if (len(multiselectCheckboxes) > 0):
                        if (not not re.match('^\d$', answerList[i])):
                            correctIndex = int(answerList[i])
                            correctlis = [elem for elem in multiselectCheckboxes[0].find_elements(
                                By.CSS_SELECTOR, 'li') if elem.text]
                            if (len(correctlis) >= correctIndex + 1):
                                correctlis[correctIndex].click()
                                sleep(5)
                            pass
                        else:
                            correctlis = [elem for elem in multiselectCheckboxes[0].find_elements(
                                By.CSS_SELECTOR, 'li') if elem.text in answerList[i] and answerList[i] and elem.text]
                            if (len(correctlis) > 0):
                                correctlis[0].click()
                                sleep(5)
                            pass
                    if (selectP.find_elements(By.CSS_SELECTOR, 'div.ui-multiselect-menu')[0].is_displayed()):
                        selectP.find_elements(By.CSS_SELECTOR, 'button')[
                            0].click()
                except Exception as e:
                    traceback.print_exc()
                    logger.exception("exec failed, failed msg:" + traceback.format_exc())
                    pass
            pass
        #一般填空题（机械制造基础-形考任务一，账号：2161001405429）
        elif (subjectType == 'fill_in_blank'):
            # 完形填空
            answerList = answer.split('|')
            selectPs = [elem for elem in subject.find_elements(By.CSS_SELECTOR, 'p') if len(elem.find_elements(
                By.CSS_SELECTOR, 'select')) > 0 and len(elem.find_elements(By.CSS_SELECTOR, 'button')) > 0]
            if (len(answerList) != len(selectPs)):
                hasAnswerCount -= 1
                continue
            for i in range(len(selectPs)):
                print(i)
                selectP = selectPs[i]
                # 检查是否打开了
                try:
                    selectP.find_elements(By.CSS_SELECTOR, 'button')[0].click()
                    WebDriverWait(selectP, 10).until(EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, 'div.ui-multiselect-menu')), message="打开下拉选失败")
                    multiselectCheckboxes = selectP.find_elements(
                        By.CSS_SELECTOR, 'ul.ui-multiselect-checkboxes')
                    if (len(multiselectCheckboxes) > 0):
                        if (not not re.match('^\d$', answerList[i])):
                            correctIndex = int(answerList[i])
                            correctlis = [elem for elem in multiselectCheckboxes[0].find_elements(
                                By.CSS_SELECTOR, 'li') if elem.text]
                            if (len(correctlis) >= correctIndex + 1):
                                correctlis[correctIndex].click()
                                sleep(5)
                            pass
                        else:
                            correctlis = [elem for elem in multiselectCheckboxes[0].find_elements(
                                By.CSS_SELECTOR, 'li') if elem.text in answerList[i] and answerList[i] and elem.text]
                            if (len(correctlis) > 0):
                                correctlis[0].click()
                                sleep(5)
                            pass
                    if (selectP.find_elements(By.CSS_SELECTOR, 'div.ui-multiselect-menu')[0].is_displayed()):
                        selectP.find_elements(By.CSS_SELECTOR, 'button')[
                            0].click()
                except Exception as e:
                    traceback.print_exc()
                    logger.exception("exec failed, failed msg:" + traceback.format_exc())
                    pass
            pass
        else:
            pass

    # 有答案率不超过80% 不提交
    if (hasAnswerCount / totalCount < examRate):
        logger.info('有答案率为 %s 不足 %s 直接结束不提交',
                    (hasAnswerCount / totalCount), examRate)
        saveFailureRecord(
            globalData['username'], globalData['className'], str(hasAnswerCount / totalCount)+"--"+globalData['activityName'])
        return

    # 遍历问题结束 交卷
    submitTags = browser.find_elements(
        By.CSS_SELECTOR, "div.paper-footer a.button[ng-click]")
    if (len(submitTags) > 0):
        submitTags[0].click()
        WebDriverWait(browser, 15).until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "#submit-exam-confirmation-popup[aria-hidden=false] button[ng-click]")), message="操作交卷按钮失败")
        confirmButtons = browser.find_elements(
            By.CSS_SELECTOR, "#submit-exam-confirmation-popup[aria-hidden=false] button[ng-click]")
        if (len(confirmButtons) > 0):
            confirmButtons[0].click()
            WebDriverWait(browser, 15).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "a.take-exam")), message="交卷失败")


# 获取题干
def getSubjectTitle(elementItem):
    # 题目类型
    subjectType = judgeSubjectType(elementItem.get_attribute('class'))
    # 题干
    subjectTitle = ''
    if (subjectType in ['true_or_false', 'single_selection', 'multiple_selection']):
        # 判断、单选、多选
        # 正常题干
        subjectTitleEles = elementItem.find_elements(
            By.CSS_SELECTOR, 'span.subject-description')
        if (len(subjectTitleEles) == 0):
            return ''
        subjectTitle = elementItem.find_elements(
            By.CSS_SELECTOR, 'span.subject-description')[0].text
        # 带图
        imgs = elementItem.find_elements(
            By.CSS_SELECTOR, 'span.subject-description')[0].find_elements(By.CSS_SELECTOR, 'img')
        for img in imgs:
            imgSrc = img.get_attribute('src')
            if (not imgSrc):
                continue
            imgSrcMd5 = getStrAsMD5(imgSrc)
            subjectTitle = ((subjectTitle + '|')
                            if subjectTitle else '') + imgSrcMd5
    #英语阅读填空题
    elif (subjectType in ['cloze']):
        pElems = [elem for elem in elementItem.find_elements(
            By.CSS_SELECTOR, "div.summary-title span.pre-wrap p.ng-scope")]
        textList = []
        for i in range(len(pElems)):
            pElem = pElems[i]
            if (i == 0):
                continue
            if (i == 1 and len(pElem.find_elements(By.CSS_SELECTOR, 'strong')) > 0):
                continue
            textList.append(pElem.text)
        subjectTitle = ''.join(textList)
        subjectTitle = subjectTitle[0:100]
        pass
    #一般填空题（机械制造基础-形考任务一，账号：2161001405429）
    elif (subjectType in ['fill_in_blank']):
        pElems = [elem for elem in elementItem.find_elements(
            By.CSS_SELECTOR, "div.summary-title span.pre-wrap p.ng-scope")]
        textList = []
        for i in range(len(pElems)):
            pElem = pElems[i]
            textList.append(pElem.text)
        subjectTitle = ''.join(textList)
        #去掉字符串中的所有空格
        subjectTitle = ''.join(subjectTitle.split())
        pass
    return subjectTitle


# 计算MD5
def getStrAsMD5(parmStr):
    # 1、参数必须是utf8
    # 2、python3所有字符都是unicode形式，已经不存在unicode关键字
    # 3、python3 str 实质上就是unicode
    if isinstance(parmStr, str):
        # 如果是unicode先转utf-8
        parmStr = parmStr.encode("utf-8")
    m = hashlib.md5()
    m.update(parmStr)
    return m.hexdigest()


# 获取正确选项
def getOptionsValue(elementItem):
    optionContents = elementItem.find_elements(
        By.CSS_SELECTOR, "div.option-content")
    if (len(optionContents) == 0):
        return ''
    optionText = optionContents[0].text.strip()
    imgs = optionContents[0].find_elements(By.CSS_SELECTOR, 'img')
    for img in imgs:
        imgSrc = img.get_attribute('src')
        if (not imgSrc):
            continue
        imgSrcMd5 = getStrAsMD5(imgSrc)
        optionText = ((optionText + '|') if optionText else '') + imgSrcMd5
    return optionText


# 判断类型
def judgeSubjectType(classValue):
    subjectTypeList = subjectTypeMap.keys()

    for subject in subjectTypeList:
        if (subject in classValue):
            return subject
    return 'ERROR'


# 查询答案
def queryAnswer(title, type):
    if (not title):
        return ''
    payload = json.dumps({
        "tiTitle": title
    })
    headers = {
        'Content-Type': 'application/json'
    }

    response = requests.request(
        "POST", queryAnswerUrl, headers=headers, data=payload)
    if (response.status_code == 200 and not not response.text and len(response.text) > 0):
        logger.info('----' + response.text)
        jsonData = json.loads(response.text)
        success = jsonData.get('success', False)
        tiAnswer = (jsonData.get('result', {}) if jsonData.get(
            'result', {}) else {}).get('tiAnswer', '')
        if (success and not not tiAnswer and len(tiAnswer) > 0):
            return tiAnswer
    return ''


# 查询答案
def queryAnswerFromDb(title, type):
    if (not title):
        return ''
    titleMd5 = getStrAsMD5(title)
    selectIdSql = "SELECT ti_answer FROM " + tikuTable + " WHERE ti_title_md5 = %s"
    ret = db.selectone(sql=selectIdSql, param=titleMd5)
    logger.info('保存题库-查询是否已存在 结果：%s' % ret)
    tiAnswer = ''
    if (ret) :
        tiAnswer = ret[0].decode('utf-8')
    return tiAnswer


# 显示所有
def expendAllEle(browser):
    js = "document.querySelectorAll('i.font-toggle-expanded').forEach(e => e.click())"
    browser.execute_script(js)
    js = "document.querySelectorAll('i.font-toggle-collapsed').forEach(e => e.click())"
    browser.execute_script(js)
    try:
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='module ng-scope ng-isolate-scope expanded']")), message="展开失败")
    except Exception as e:
        raise


# 保存失败记录
def saveFailureRecord(username, className, activityName):
    # SQL 插入语句
    insertSql = "insert into exam_failure_record (id,login_name,class_name,exam_name,create_time) values (UUID(),%s,%s,%s,now())"
    ret = db.insertone(sql=insertSql, param=(
        username, className, activityName))
    logger.info('保存失败记录 结果：%s' % ret)


# ------------------------------------------------处理考试 结束-----------------------------------------------------

# ------------------------------------------------处理常规学习 开始-----------------------------------------------------


# 展开所有的学习活动
def getAllActivity(browser):

    # 所有学习活动列表
    activitiesList = []

    logger.info('---------------------1------------------')

    # 展开
    expandedIcons = browser.find_elements(
        By.CSS_SELECTOR, 'i.font-toggle-all-collapsed')
    logger.info('展开元素1 [%s]' % len(expandedIcons))
    if (len(expandedIcons) > 0):
        expandedIcons[0].click()
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.learning-activity')), message="展开失败")

    logger.info('---------------------2------------------')

    # 关闭
    collapsedIcons = browser.find_elements(
        By.CSS_SELECTOR, 'i.font-toggle-all-expanded')
    logger.info('展开元素2 [%s]' % len(collapsedIcons))
    if (len(collapsedIcons) > 0):
        collapsedIcons[0].click()
        WebDriverWait(browser, 10).until(
            checkAcitvityIsEmpty(), message="合并失败")

    logger.info('---------------------3------------------')

    # 展开
    expandedIcons = browser.find_elements(
        By.CSS_SELECTOR, 'i.font-toggle-all-collapsed')
    logger.info('展开元素3 [%s]' % len(expandedIcons))
    if (len(expandedIcons) > 0):
        expandedIcons[0].click()
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.learning-activity')), message="展开失败")

    # 搜索出来遍历
    learningActivities = browser.find_elements(
        By.CLASS_NAME, 'learning-activity')
    for learningActivity in learningActivities:
        # 如果完成就不再尝试
        completeFlag = len(learningActivity.find_elements(
            By.CSS_SELECTOR, 'div.completeness.none')) == 0
        if (completeFlag):
            continue
        itemId = learningActivity.get_attribute(
            'id').replace('learning-activity-', '')
        notFinish = len(learningActivity.find_elements(
            By.CSS_SELECTOR, 'div.completeness.none')) > 0
        ngSwitchDivs = learningActivity.find_elements(
            By.CSS_SELECTOR, 'div[ng-switch-when]')
        type = ngSwitchDivs[0].get_attribute(
            'ng-switch-when') if len(ngSwitchDivs) > 0 else ''
        titleATags = learningActivity.find_elements(By.CSS_SELECTOR, 'a.title')
        title = titleATags[0].text if len(titleATags) > 0 else ''
        activitiesList.append({
            'activityId': itemId,
            'title': title,
            'notFinish': notFinish,
            'type': type
        })
    return activitiesList


# 检查学习活动为空
class checkAcitvityIsEmpty(object):
    def __init__(self):
        pass

    def __call__(self, driver):
        logger.info('checking ---')
        learningActivities = driver.find_elements(
            By.CSS_SELECTOR, 'div.expanded')
        return len(learningActivities) == 0


# 处理常规学习活动
def dealNormalClassItem(activitiesList, cookie):
    courseIds = []
    for activity in activitiesList:
        courseid = activity['activityId']
        type = activity['type']
        if (type in "online_video"):
            docourse_video(cookie, courseid)
        else:
            docourse(cookie, courseid)
        courseIds.append(courseid)
        sleep(3)


# 处理一般课程
def docourse(c_cookie, courseid):
    try:
        headers = {
            # 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (K。。。。。。。。。。。。。。。。。。。。。。4492.400',
            # 如果不带，无法进去个人主页
            'Cookie': c_cookie
        }
        url = docourseUrl + str(courseid)
        logger.info('刷课请求：%s' % url)
        response = requests.post(url, headers=headers)
        tk = json.loads(response.text)['completeness']
        logger.info('刷课请求响应：%s' % response.text)
    except Exception as e:
        logger.error('处理一般课程报错 %s', e)


# 处理视频课程
def docourse_video(c_cookie, courseid):
    try:
        headers = {
            # 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (K。。。。。。。。。。。。。。。。。。。。。。4492.400',
            # 如果不带，无法进去个人主页
            'Cookie': c_cookie,
            'Content-Type': 'application/json'
        }
        body = {"start": 1, "end": 6000}
        url = docourseUrl+str(courseid)
        logger.info('视频刷课请求 %s' % url)
        response = requests.post(url, headers=headers, data=json.dumps(body))
        tk = json.loads(response.text)['completeness']
        logger.info('视频请求响应：%s' % response.text)
    except Exception as e:
        logger.error('处理视频课程报错 %s', e)


# 获取cookie
def getCookieFromClassPage(browser):
    requests = browser.requests
    total = len(requests)
    for i in range(total):
        request = requests[total-i-1]
        cookie = request.headers['Cookie']
        if ('all-activities' in request.url and 'lms.ouchn.cn' in request.url and cookie is not None and len(cookie) > 0):
            return cookie
    return ''


# ------------------------------------------------处理常规学习 结束-----------------------------------------------------


# 获取验证码图片
def getKaptchaImage(browser, multiple):
    threadId = ''
    if (multiple):
        threadId = threading.current_thread().__getattribute__('threadID')
    screenshotFileName = 'screenshot-%s.png' % (
        threadId if multiple else '0')
    kaptchaImageFileName = 'kaptchaImage-%s.png' % (
        threadId if multiple else '0')
    browser.save_screenshot(screenshotFileName)
    kaptchaImage = browser.find_element(By.ID, 'kaptchaImage')

    """计算页面元素的在整个页面上的坐标"""
    left = kaptchaImage.location['x']
    top = kaptchaImage.location['y']
    right = kaptchaImage.location['x'] + kaptchaImage.size['width']
    bottom = kaptchaImage.location['y'] + kaptchaImage.size['height']
    time.sleep(2)
    """"根据页面元素的坐标，截图元素"""
    im = Image.open(screenshotFileName)
    im = im.crop((left, top, right, bottom))
    im.save(kaptchaImageFileName)
    image = Image.open(kaptchaImageFileName)
    # result = pytesseract.image_to_string(image)
    result = ocr2(kaptchaImageFileName)
    return re.sub(r'\D', "", result.strip())


# ocr第二种识别
def ocr2(imageFileName):
    with open(imageFileName, 'rb') as f:
        image = f.read()
    resp = requests.post(
        "http://121.37.181.45:9898/ocr/file", files={'image': image})
    return resp.text


# 线程定义
class myThread (threading.Thread):
    def __init__(self, threadID, name):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name

    def run(self):
        logger.info("启动线程：%s" % (self.name))
        while not queue.empty():
            try:
                row = queue.get(block=True, timeout=5)
                id = row['id']
                logger.info('线程: %s [%s] 正在处理 >>> 用户名: %s' %
                            (self.name, self.ident, row['userName']))
                # 处理账号
                singleAccountDeal(row['userName'], row['password'], True)
                # 更新记录
                updateAccountFinish(id)
            except:
                continue
        logger.info("线程执行结束：%s" % (self.name))

    def getThreadId(self):
        return self.threadID


# 单账号测试
def singleAccountDeal(username, password, multiple):
    globalData = initExplore()

    # 用户登录
    loginUser(username, password, globalData, 0, multiple)
    browser = globalData['browser']

    # 查到所有课程
    tryTimes = 0
    classList = []
    while tryTimes < 3:
        try:
            # 等待页面展开
            waitPageLoad(
                browser, '', '.ouchnPc_index_course_div', 10, '首页加载失败')
            classList = findAllClass(globalData, classList)
            logger.info(classList)
            break
        except:
            if (tryTimes > 3):
                return
            tryTimes += 1
            traceback.print_exc()
            logger.exception("exec failed, failed msg:" + traceback.format_exc())

    # 保持首页
    globalData['indexUrl'] = browser.current_url

    # 遍历所有课程
    for classItem in classList:
        dealClassItem(globalData, classItem)


# 更新账号状态
def updateAccountFinish(id):
    # 打开数据库连接
    db = pymysql.connect(host=db_host, user=db_user,
                         password=db_password, database=db_database)

    # 使用cursor()方法获取操作游标
    cursor = db.cursor()

    # SQL 更新语句
    updatesql = "update "+userTable+" set job_status=2 where id=%s" % (id)

    try:
        cursor.execute(updatesql)
        db.commit()
        print("all over %s" % ctime())
    except:
        traceback.print_exc()
        logger.exception("exec failed, failed msg:" + traceback.format_exc())
        print("Error: unable to fetch data")

    # 关闭数据库连接
    db.close()


# 查询所有账号
def queryAllAccount():
    # 打开数据库连接
    db = pymysql.connect(host=db_host, user=db_user, port=db_port,
                         password=db_password, database=db_database)
    # 使用cursor()方法获取操作游标
    cursor = db.cursor()
    # SQL 查询语句
    sql = "select id,login_name,password from " + \
        userTable+" where job_status = 0 order by id desc"
    results = []
    userInfoList = []
    try:
        # 执行SQL语句
        cursor.execute(sql)
        # 获取所有记录列表
        # results = cursor.fetchone()
        results = cursor.fetchall()
        for item in results:
            userInfoList.append({
                'id': item[0],
                'userName': item[1],
                'password': item[2]
            })
    except:
        traceback.print_exc()
        logger.exception("exec failed, failed msg:" + traceback.format_exc())
        logger.error("Error: unable to fetch data")
    # 关闭数据库连接
    db.close()
    return userInfoList


# 放入队列
def putQuenue(userInfoList):
    for userInfo in userInfoList:
        queue.put(userInfo)


# 批量处理账号
def batchDealAccount():
    # 查询用户
    userInfoList = queryAllAccount()

    # 放入队列
    putQuenue(userInfoList)

    # 启动足够的线程并且启动
    multipleThreadInit()


# 批量-多线程启动
def multipleThreadInit():
    for i in range(0, threadCount):
        thread = myThread(i, "Thread-%s" % (i))
        threadPool.append(thread)
        thread.start()
        sleep(threadSleepTime)


if __name__ == '__main__':
    # 单账号测试
    singleAccountDeal('2161001453072', 'Ouchn@2021', False)
    # 多账号测试
    #batchDealAccount()
    pass
