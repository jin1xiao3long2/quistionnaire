from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import random
import time
import openpyxl


def run():
    # 获取表格数据
    wb = openpyxl.load_workbook('D:\myxlsx.xlsx')
    sheet = wb['Sheet1']

    # 开始循环

    for i in range(213):
        # 打开浏览器

        while True:
            browser = webdriver.Chrome()
            browser.get("http://baituandazhan.mikecrm.com/zFYSfSi")
            try:
                # 选择社团类别
                while (1):
                    flag = 0
                    # 选择社团
                    la = browser.find_element_by_xpath("//*[@id=\"204125996\"]/div[2]/div/div[1]/p")
                    while (1):
                        la.click()
                        ww = browser.find_element_by_xpath("//*[@id=\"204125996\"]/div[2]/div/div[1]/ul/li[3]")
                        if (ww):
                            ww.click()
                        print(la.text)
                        if ("自律互助类" == la.text):
                            flag = 1
                            break
                    if (flag == 1):
                        break

                # time.sleep(0.3)
                # 选择社团编号
                while (1):
                    flag = 0
                    la = browser.find_element_by_xpath("//*[@id=\"204125996\"]/div[2]/div/div[2]/p")
                    while (1):
                        la.click()
                        browser.find_element_by_xpath("//*[@id=\"204125996\"]/div[2]/div/div[2]/ul/li[8]").click()
                        print(la.text)
                        if ("H007" in la.text):
                            flag = 1
                            break
                    if (flag == 1):
                        break

                # time.sleep(2)
                # 检查数据

                # 填写名字
                Name = sheet['G' + int(i + 2).__str__()].value
                la = browser.find_element_by_xpath("//*[@id=\"204125993\"]/div[2]/div/div[1]/input")
                la.send_keys(Name)

                # 选择性别
                Sex = sheet['H' + int(i + 2).__str__()].value
                if (Sex == 1):
                    lsans = browser.find_element_by_xpath("//*[@id=\"opt203887697\"]/p")
                elif (Sex == 2):
                    lsans = browser.find_element_by_xpath("//*[@id=\"opt203887698\"]/p")
                else:
                    continue
                lsans.click()

                # time.sleep(2)

                # 填写手机号
                phone_number = sheet['I' + int(i + 2).__str__()].value
                la = browser.find_element_by_xpath("//*[@id=\"204125994\"]/div[2]/div/div/div/input")
                la.send_keys(phone_number)

                # time.sleep(2)

                # 选择学院
                college = sheet['L' + int(i + 2).__str__()].value
                flag = 0
                while (1):
                    la = browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/p")
                    while (1):
                        la.click()
                        if ('经济' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[1]").click()
                            if (la.text == '经济学院'):
                                flag = 1
                        elif ('法' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[2]").click()
                            if (la.text == '法学院'):
                                flag = 1
                        elif ('文新' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[3]").click()
                            if (la.text == '文学与新闻学院'):
                                flag = 1
                        elif ('历史' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[6]").click()
                            if (la.text == '历史文化学院(旅游学院)'):
                                flag = 1
                        elif ('管理' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[7]").click()
                            if (la.text == '公共管理学院'):
                                flag = 1
                        elif ('商' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[8]").click()
                            if (la.text == '商学院'):
                                flag = 1
                        elif ('数' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[10]").click()
                            if (la.text == '数学学院'):
                                flag = 1
                        elif ('物理' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[11]").click()
                            if (la.text == '物理学院'):
                                flag = 1

                        elif ('生科' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[13]").click()
                            if (la.text == '生命科学学院'):
                                flag = 1
                        elif ('软件' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[14]").click()
                            if (la.text == '软件学院'):
                                flag = 1
                        elif ('化工' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[15]").click()
                            if (la.text == '化学工程学院'):
                                flag = 1
                        elif ('电信' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[16]").click()
                            if (la.text == '电子信息学院'):
                                flag = 1
                        elif ('材料' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[17]").click()
                            if (la.text == '材料科学与工程学院'):
                                flag = 1
                        elif ('生物' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[18]").click()
                            if (la.text == '生物医学工程学院'):
                                flag = 1
                        elif ('机械' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[19]").click()
                            if (la.text == '机械工程学院'):
                                flag = 1
                        elif ('电工' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[20]").click()
                            if (la.text == '电气工程学院'):
                                flag = 1
                        elif ('计算机' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[21]").click()
                            if (la.text == '计算机学院'):
                                flag = 1
                        elif ('建筑' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[22]").click()
                            if (la.text == '建筑与环境学院'):
                                flag = 1
                        elif ('水利' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[23]").click()
                            if (la.text == '水利水电学院'):
                                flag = 1
                        elif ('轻工' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[24]").click()
                            if (la.text == '轻工科学与工程学院'):
                                flag = 1
                        elif ('高分子' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[25]").click()
                            if (la.text == '高分子科学与工程学院'):
                                flag = 1
                        elif ('匹兹堡' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[28]").click()
                            if (la.text == '匹兹堡学院'):
                                flag = 1
                        elif ('吴玉章' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[29]").click()
                            if (la.text == '吴玉章学院'):
                                flag = 1
                        elif ('基础' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[30]").click()
                            if (la.text == '华西基础医学与法医学院'):
                                flag = 1
                        elif ('临床' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[31]").click()
                            if (la.text == '华西临床医学院'):
                                flag = 1
                        elif ('口腔' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[32]").click()
                            if (la.text == '华西口腔医学院'):
                                flag = 1
                        elif ('公共' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[33]").click()
                            if (la.text == '华西公共卫生学院'):
                                flag = 1
                        elif ('药' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[34]").click()
                            if (la.text == '华西药学院'):
                                flag = 1
                        elif ('网络' == college):
                            browser.find_element_by_xpath("//*[@id=\"204125990\"]/div[2]/div/div/ul/li[36]").click()
                            if (la.text == '网络空间安全学院'):
                                flag = 1
                        else:
                            flag = 2
                        if flag == 1 or flag == 2:
                            break
                    if flag == 1 or flag == 2:
                        break
                if flag == 2:
                    break

                # time.sleep(2)

                # 填入专业
                major = sheet['M' + int(i + 2).__str__()].value
                la = browser.find_element_by_xpath("//*[@id=\"204125992\"]/div[2]/div/div/input")
                la.send_keys(major)

                # time.sleep(2)

                # 填入学号
                id = sheet['K' + int(i + 2).__str__()].value
                la = browser.find_element_by_xpath("//*[@id=\"204125995\"]/div[2]/div/div/input")
                la.send_keys(id)
                print(Name + " " + int(Sex).__str__() + " " + major + " " + int(id).__str__())

                time.sleep(0.5)
            except Exception as e:
                print(e)
                print("this is the " + int(i).__str__() + " time")
                browser.quit()
                continue
            break
        # 提交表格
        #time.sleep(1)
        am = browser.find_element_by_xpath("//*[@id=\"form_submit\"]")
        am.click()
        browser.quit()


if __name__ == "__main__":
    # sheet()
    # for i in range(5):
    run()
   # wb = openpyxl.load_workbook("D:\myxlsx.xlsx")
    #print(wb['Sheet1']['G2'].value)
#     time.sleep(1)
