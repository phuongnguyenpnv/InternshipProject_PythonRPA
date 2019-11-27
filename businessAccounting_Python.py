import pandas as pd, numpy as np, inspect, os 
from pywinauto.application import Application
from xlwt import Workbook
from pandas import ExcelWriter
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from time import sleep
from pyautogui import hotkey
import pyperclip
from pywinauto import keyboard
from xml.dom import minidom
import zipfile
import xml.sax
from datetime import datetime

CurDir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
df_Input = pd.read_excel(CurDir+"\\input\\Input_HDDT.xlsx", dtype='str', sheet_name='Sheet1')
df_Input_file = pd.DataFrame(df_Input)

path_Output = pd.read_excel(CurDir+"\\Output\\Output.xlsx", sheet_name='Sheet1') #Đọc file excel
df = pd.DataFrame(path_Output)
df_link_fdcapcha = pd.read_excel(CurDir+"\\input\\Link_captcha.xlsx", dtype='str')

writer_Output = pd.ExcelWriter(CurDir + "\\Output\Output.xlsx", engine = 'openpyxl')
writer_Input = pd.ExcelWriter(CurDir + "\\input\\Input_HDDT.xlsx", engine = 'openpyxl')

list_URL = df_Input["URL"].values
list_CODE = df_Input["Mã Tra Cứu"].values
list_SoHD = df_Input["Số Hóa Đơn"].values
list_IMG = df_Input["Tên Ảnh"].values

def read_file_excel(path, name, idx):
    value = path[name].values
    value = str(value[idx])
    return value
CurDir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))  #soft link

def CheckExistFile(Extension):
    while True:
        for file in os.listdir(CurDir+"\\download\\"):
            if file.endswith(Extension):
                path = (os.path.join("\\download\\", file))
                return path

def CheckPNGFile(Extension):
    linkServer = read_file_excel(df_link_fdcapcha, "Link", 0)
    print(linkServer)
    while True:
        for file in os.listdir(linkServer):
            print(file)
            if file.endswith(Extension):
                path = (os.path.join(linkServer, file))
                return path

def UnzipFolder(path, new_path=False):
    if os.path.exists(path):
        zipp = zipfile.ZipFile(path)
        if not new_path:
            base_path = "\\".join(path.split("\\")[:-1])
            zipp.extractall(base_path)
        elif os.path.isdir(new_path):
            zipp.extractall(new_path)
        zipp.close()
    return
def write_Excel(idx,inforcty_HD ,inforHH_HD):
    i = 0
    while i < len(df_Input):
        MST_DVB = read_file_excel(df_Input, "Mã Số Thuế", i)
        Ten_Anh = read_file_excel(df_Input, "Tên Ảnh", i)
         
        if inforcty_HD[4] == MST_DVB:
            df.loc[idx,"Tên Ảnh"] = Ten_Anh
            df.loc[idx,"Ngày tháng năm của HĐ"] = inforcty_HD[5]
            df.loc[idx, "Số HĐ"] = inforcty_HD[2]
            df.loc[idx,"MST đơn vị mua hàng"] = inforcty_HD[0]
            df.loc[idx,"Tên đơn vị mua hàng"] = inforcty_HD[1]
            df.loc[idx,"Địa chỉ đơn vị mua hàng"] = str(inforcty_HD[3]).strip().strip(",")
            df.loc[idx, "Tên hàng hóa"] = inforHH_HD[0]
            df.loc[idx, "Đơn vị tính"] = inforHH_HD[2]
            df.loc[idx,"Số lượng"] = int(inforHH_HD[1])
            df.loc[idx,"Thành tiền"] = int(inforHH_HD[3])
            if int(inforHH_HD[4]) == -1:
                df.loc[idx,"Thuế suất GTGT(%)"] = 0
            else:
                df.loc[idx,"Thuế suất GTGT(%)"] = int(inforHH_HD[4])
            df.loc[idx,"Tiền Thuế GTGT"] = int(inforHH_HD[5])
            break
        else:
            i +=1
def get_InforHD(Path, MST, Ten, SoHD, Diachi, ngaythang, MST_DVB):
    file_XML = minidom.parse(str(Path))
    MST_DVMH = file_XML.getElementsByTagName(str(MST))[0].firstChild.data
    ten_DVMH = file_XML.getElementsByTagName(str(Ten))[0].firstChild.data
    so_HD = file_XML.getElementsByTagName(str(SoHD))[0].firstChild.data
    diachi_DVMH = file_XML.getElementsByTagName(str(Diachi))[0].firstChild.data
    MST_DVBH = file_XML.getElementsByTagName(str(MST_DVB))[0].firstChild.data
    Ngay_HD = file_XML.getElementsByTagName(str(ngaythang))[0].firstChild.data
    strNgayHD = Ngay_HD[:10]
    try:
        strNgayHD = datetime.strptime(str(strNgayHD), "%Y-%m-%d")
    except:
        strNgayHD = datetime.strptime(str(strNgayHD), "%d/%m/%Y")
    strNgayHD = str(strNgayHD.strftime("%d-%m-%Y"))
    return(MST_DVMH,ten_DVMH, so_HD, diachi_DVMH, MST_DVBH, strNgayHD)

def get_InforHH(Path, idx, ten, SLuong, dvi, thanhtien, thuesuat, tienthue):
    file_XML = minidom.parse(str(Path))
    ten_HH = file_XML.getElementsByTagName(str(ten))[int(idx)].firstChild.data
    soluong_HH = file_XML.getElementsByTagName(str(SLuong))[int(idx)].firstChild.data
    soluong_HH = str(soluong_HH).split(".")
    soluong_HH = soluong_HH[0]
    donvi_HH = file_XML.getElementsByTagName(str(dvi))[int(idx)].firstChild.data
    if str(thanhtien) == "Total":
        idx +=1
        try:
            thuesuat_HH = file_XML.getElementsByTagName(str(thuesuat))[int(idx -1 )].firstChild.data
        except:
            thuesuat = 0
    else:
        try:
            thuesuat_HH = file_XML.getElementsByTagName(str(thuesuat))[int(idx)].firstChild.data
        except:
            thuesuat = 0
    tien_HH = file_XML.getElementsByTagName(str(thanhtien))[int(idx)].firstChild.data
    tienthue_HH = file_XML.getElementsByTagName(str(tienthue))[int(idx)].firstChild.data
    return(ten_HH, soluong_HH, donvi_HH, tien_HH, thuesuat_HH, tienthue_HH)

def Clear():
    os.system("taskkill /f /im chromedriver.exe")
    os.system("taskkill /f /im chrome.exe")

def Open_Browser():
    chromeOptions = webdriver.ChromeOptions()
    path_DOWNLOAD = CurDir+"\\download\\"
    prefs = {"download.default_directory" : path_DOWNLOAD,"safebrowsing.enabled": "false"}
    chromeOptions.add_experimental_option("prefs",prefs)
    # driver = webdriver.Chrome(executable_path=os.path.abspath("chromedriver.exe"), chrome_options=chromeOptions)
    driver = webdriver.Chrome(CurDir+"\\chromedriver.exe", chrome_options=chromeOptions)
    return driver

def Process_Download():
    linkServer = read_file_excel(df_link_fdcapcha, "Link", 0)
    print(linkServer)
    driver = Open_Browser()
    driver.maximize_window()

    idx = 0
    for str_URL in list_URL:
        driver.get(str_URL)
        if str(str_URL) == "https://www.meinvoice.vn/tra-cuu/":
            driver.find_element_by_xpath('//*[@id="txtCode"]').send_keys(list_CODE[idx])
            sleep(1)
            radiobtn1 = driver.find_element_by_xpath('//*[@id="pnSearch"]/div[1]/div[2]/div/div[2]/div[1]/a[1]')
            radiobtn1.get_attribute("href")
            try:
                driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
                driver.find_element_by_xpath('//*[@id="popup-content-container"]/div[1]/div[2]/div[7]/div').click()
                sleep(2)
                driver.find_element_by_xpath('//*[@id="popup-content-container"]/div[1]/div[2]/div[7]/div/div/div[2]').click()
            except:
                df_Input_file.loc[idx, "Chú thích"] = "Fail"
                df_Input_file.to_excel(writer_Input,sheet_name ="Sheet1", index=False)
            sleep(2)

        elif str(str_URL) == "https://ihoadon.vn/kiem-tra/":
            driver.find_element_by_xpath('//*[@id="nguoi_mua_hang"]/form/div[3]/div[1]/div/input').send_keys(list_SoHD[idx])
            driver.find_element_by_xpath('//*[@id="nguoi_mua_hang"]/form/div[3]/div[2]/div/input').send_keys(list_CODE[idx])
            driver.find_element_by_xpath('//*[@id="nguoi_mua_hang"]/form/div[5]/div/button').click()
            sleep(2)
            driver.find_element_by_xpath('//*[@id="table-data"]/tbody/tr[2]/td[8]/div[1]/a').click()
            sleep(2)
        elif str(str_URL)=="http://nutifood.vinvoice.vn/":
            i= 1
            while True:
                if i == 5:
                    break
                else: 
                    driver.find_element_by_xpath('//*[@id="code"]').send_keys(list_CODE[idx])
                    el = driver.find_element_by_xpath('//*[@id="CaptchaImage"]')
                    el.screenshot(linkServer+str(i)+'_screenshot.png')
                    i+=1
                    sleep(5)
                    sleep(5)
                    read_file_Image = os.listdir(linkServer)
                    str_capcha = str(read_file_Image).split(".")
                    driver.find_element_by_xpath('//*[@id="CaptchaInputText"]').send_keys(str(str_capcha[0]).strip("['"))
                    driver.find_element_by_xpath('/html/body/div[2]/div[2]/div/div[1]/div/form/div/div/div/div[4]/div[2]/button').click()
                    path_Image = CheckPNGFile('.png')
                    os.remove(path_Image)
                    try:
                        driver.find_element_by_xpath('/html/body/div[2]/div[2]/div/div[1]/div/div/div[1]/table/tbody/tr/td[8]/a').click()
                        sleep(2)
                        driver.find_element_by_xpath('//*[@id="btncheckData"]').click()
                        sleep(2)
                        break
                    except: 
                        pass
        idx+=1
    writer_Input.save()       
    driver.quit()

def Process_Handling():
    Path = CurDir + CheckExistFile('.zip')
    UnzipFolder(Path,CurDir+"\\download\\")
    os.remove(Path)
    read_folder = os.listdir(CurDir+"\\download\\")
    x  = dem = 0
    while x <len(read_folder):
        if read_folder[x]=="7_2019_01GTKT0-001_TB-19E_24696.xml":
            file_XML = CurDir+"\\download\\"+read_folder[x]
            inforcty_HD= get_InforHD(file_XML, "CusTaxCode", "CusName", "InvoiceNo","CusAddress", "SignDate", "ComTaxCode")
            file_XML1 = minidom.parse(file_XML)
            tenHH = file_XML1.getElementsByTagName("ProdName")
            i=0
            while i < tenHH.length:
                inforHH_HD = get_InforHH(file_XML, i,"ProdName", "ProdQuantity", "ProdUnit", "Total", "VATRate", "VATAmount")
                i+=1
                write_Excel(dem, inforcty_HD, inforHH_HD)
                dem+=1
            x+=1    
        else:
            file_XML = CurDir+"\\download\\"+read_folder[x]
            inforcty_HD= get_InforHD(file_XML, "inv:buyerTaxCode", "inv:buyerLegalName", "inv:invoiceNumber","inv:buyerAddressLine","inv:signedDate", "inv:sellerTaxCode")
            file_XML1 = minidom.parse(file_XML)
            tenHH = file_XML1.getElementsByTagName("inv:itemName")
            i=0
            while i < tenHH.length -1:
                inforHH_HD = get_InforHH(file_XML, i,"inv:itemName", "inv:quantity", "inv:unitName", "inv:unitPrice", "inv:vatPercentage", "inv:vatAmount")
                i+=1
                write_Excel(dem, inforcty_HD, inforHH_HD)
                dem+=1
            x+=1  
        df.to_excel(writer_Output,sheet_name ="Sheet1", index=False)
        writer_Output.save()
        
def InputProduct():
    keyboard.send_keys("{DOWN 5}")
    keyboard.send_keys("{ENTER}",0.2)
    list_TenHH = [""]
    i = 200
    idx =  0
    while idx < len(path_Output):
        dem =0
        while dem <= len(list_TenHH):
            strtenHH = read_file_excel(path_Output, "Tên hàng hóa", idx)
            if str(strtenHH) == str(list_TenHH[dem]) :
                idx+=1
                break
            elif str(strtenHH) != str(list_TenHH[dem]) and dem < len(list_TenHH)-1:
                dem+=1
            elif str(strtenHH) != str(list_TenHH[dem]) and dem == len(list_TenHH)-1:
                hotkey("alt", "m")
                i+=1
                keyboard.send_keys(str(i))
                keyboard.send_keys('{ENTER}')
                pyperclip.copy(strtenHH)
                hotkey("ctrl", "v")
                keyboard.send_keys('{ENTER}')
                strUnit = read_file_excel(path_Output, "Đơn vị tính",idx)
                keyboard.send_keys(strUnit, with_spaces=True)
                keyboard.send_keys('{ENTER 7}')
                keyboard.send_keys("H")
                keyboard.send_keys('{ENTER 3}')
                intVAT = read_file_excel(path_Output, "Thuế suất GTGT(%)",idx)
                pyperclip.copy(intVAT)
                hotkey("ctrl","v")
                hotkey("alt","m")
                hotkey("alt","q")
                idx += 1
                list_TenHH.append(strtenHH)
                break
    hotkey("alt", "q")      
def inputCustomer():
    print(path_Output)
    keyboard.send_keys("{DOWN 9}",0.1)#các phan he nghiệp vu
    keyboard.send_keys("{ENTER 2}")
    keyboard.send_keys("{DOWN 2}")#danh muc
    keyboard.send_keys("{ENTER}")
    MS = 200
    idx = 0
    while idx < len(path_Output):
        MSHD  = read_file_excel(path_Output, "Số HĐ", idx)
        try:
            MSHD_bf = read_file_excel(path_Output, "Số HĐ", idx-1)
        except:
            MSHD_bf = 0
        if MSHD != MSHD_bf:
            hotkey("alt", "m")
            MS+=1
            keyboard.send_keys(str(MS))
            keyboard.send_keys('{ENTER}')
            Ten_KH = read_file_excel(path_Output, "Tên đơn vị mua hàng", idx)
            pyperclip.copy(Ten_KH)
            hotkey("ctrl","v")
            keyboard.send_keys("{ENTER 2}",0.2)
            Diachi_KH = read_file_excel(path_Output, "Địa chỉ đơn vị mua hàng", idx)
            pyperclip.copy(Diachi_KH)
            hotkey("ctrl", "v")
            keyboard.send_keys("{ENTER}",0.1)
            MST_KH = read_file_excel(path_Output, "MST đơn vị mua hàng", idx)
            pyperclip.copy(MST_KH)
            hotkey("ctrl", "v")
            hotkey("alt","m")
            hotkey("alt", "q")
            sleep(1)
            idx+=1
        else:
            idx+=1
    hotkey("alt", "q")
def InputBills():
    keyboard.send_keys("{TAB}",0.2)
    hotkey("K")
    keyboard.send_keys("{DOWN}",0.2)
    keyboard.send_keys("{ENTER 2}")
    keyboard.send_keys("{DOWN 1}")#Cập nhật số liệu
    keyboard.send_keys("{ENTER}")
    keyboard.send_keys("{UP 12}")#Bỏ qua phân tích chứng từ
    keyboard.send_keys("{ENTER}")
    listMSHD = ["0"]
    idx = 0
    while idx < len(path_Output):
        hotkey("alt","m")
        keyboard.send_keys("{RIGHT}", 0.1)#Di chuyen chuot den ten khach hang
        strTenDVMH = read_file_excel(path_Output,"Tên đơn vị mua hàng", idx)#Nhap ten khách hàng
        pyperclip.copy(strTenDVMH)
        hotkey("ctrl","v")
        keyboard.send_keys("{ENTER 4}", 0.1)
        keyboard.send_keys("Xuất h", with_spaces=True) #xuất hàng bán
        keyboard.send_keys("{ENTER}",0.2)
        strNgayCT = read_file_excel(path_Output,"Ngày tháng năm của HĐ", idx)
        pyperclip.copy(strNgayCT)
        hotkey("ctrl","v")
        keyboard.send_keys("{ENTER}",0.2)
        hotkey("alt","c")
        keyboard.send_keys("{ENTER}",0.2)
        while idx < len(path_Output):
            intMSHD = int(read_file_excel(path_Output,"Số HĐ", idx))
            if idx == 0:
                old_SHD = 0
                newMSHD = int(read_file_excel(path_Output,"Số HĐ", (idx+1)))
            elif idx == (len(path_Output)-1):
                newMSHD = 0
                old_SHD = int(read_file_excel(path_Output,"Số HĐ", (idx-1)))
            else:
                intMSHD = int(read_file_excel(path_Output,"Số HĐ", idx))
                newMSHD = int(read_file_excel(path_Output,"Số HĐ", (idx+1)))
                old_SHD = int(read_file_excel(path_Output,"Số HĐ", (idx-1)))
            strTenHH = read_file_excel(path_Output,"Tên hàng hóa", idx)
            pyperclip.copy(strTenHH)
            hotkey("ctrl","v")
            intSoLuong = int(read_file_excel(path_Output,"Số lượng", idx))
            keyboard.send_keys("{RIGHT 2}", 0.5)
            pyperclip.copy(intSoLuong)
            hotkey("ctrl","v")
            keyboard.send_keys("{RIGHT 2}", 0.5)
            intThanhTien = int(read_file_excel(path_Output,"Thành tiền", idx))
            if intThanhTien != 0:
                pyperclip.copy(intThanhTien)
                hotkey("ctrl","v")
                keyboard.send_keys("{ENTER 2}",0.1)
            else:
                keyboard.send_keys("{ENTER}",0.1)

            if (intMSHD == newMSHD):
                keyboard.send_keys("{RIGHT 10}")
                idx+=1
            else:
                if(intMSHD == old_SHD):
                    hotkey("alt","v")
                    keyboard.send_keys("{UP 5}")
                    intMSHD = int(read_file_excel(path_Output,"Số HĐ", idx))
                    pyperclip.copy(intMSHD)
                    hotkey("ctrl","v")
                    keyboard.send_keys("{ENTER 2}",0.1)
                else:
                    hotkey("alt","v")
                    keyboard.send_keys("{UP 3}")
                    intMSHD = int(read_file_excel(path_Output,"Số HĐ", idx))
                    pyperclip.copy(intMSHD)
                    hotkey("ctrl","v")
                    keyboard.send_keys("{ENTER}")
                hotkey("ctrl","v")
                keyboard.send_keys("{ENTER 2}")
                intThuesuat = read_file_excel(path_Output,"Thuế suất GTGT(%)", idx)
                pyperclip.copy(intThuesuat)
                hotkey("ctrl","v")
                keyboard.send_keys("{ENTER}",0.1)
                hotkey("alt","m")
                hotkey("alt","q")
                idx+=1
                listMSHD.append(intMSHD)
                break

if __name__ == "__main__":
    print ("Bắt đầu quy trình")
    Process_Download()
    Clear()
    Process_Handling()
    path_Output = pd.read_excel(CurDir+"\\Output\\Output.xlsx", sheet_name='Sheet1') #Đọc file excel
    os.startfile(CurDir+"\\TRI VIET Accouting Office 10.lnk")
    keyboard.send_keys("{RIGHT 6}",0.2)
    keyboard.send_keys("{ENTER}")
    inputCustomer()
    InputProduct()
    InputBills()
    print ("Kết thúc quy trình")