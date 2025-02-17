from pathlib import Path
from PIL import Image
import requests
import win32con
import win32gui
import win32ui
import zipfile
import shutil
import json
import sys
import os


def update():
    #Set the download url to what we got from the JSON
    file = f"https://github.com/MapStudioProject/CTR-Studio/releases/download/{name}/CtrStudio-Latest.zip"
    print(f"Updating CTR Studio to {name}...\n")
    #If files or folders already exists
    if os.path.exists(parent_dir / "ConfigGlobal.json") or os.path.exists(parent_dir / "Recent.txt") or os.path.exists(parent_dir / "Themes") or os.path.exists(parent_dir / "Presets"):
        print("ConfigGlobal.json, Recent.txt, Themes and Presets folder are already in the parent folder, remove them before updating CTR Studio\n")
        sys.exit(1)
    #Copy the settings and recent projects files
    if os.path.exists(ctr_studio_dir / "ConfigGlobal.json"):
        shutil.copy(ctr_studio_dir / "ConfigGlobal.json", parent_dir / "ConfigGlobal.json")
    if os.path.exists(ctr_studio_dir / "Recent.txt"):
        shutil.copy(ctr_studio_dir / "Recent.txt", parent_dir / "Recent.txt")
    #Copy the Themes and Presets folders
    if os.path.exists(ctr_studio_dir / "Lib/Themes"):
        shutil.copytree(ctr_studio_dir / "Lib/Themes", parent_dir / "Themes")
    if os.path.exists(ctr_studio_dir / "Lib/Presets"):
        shutil.copytree(ctr_studio_dir / "Lib/Presets", parent_dir / "Presets")
    #Download the file
    os.system(f"curl -L -o ctr.zip {file}")
    #Remove CTR Studio's previous version
    shutil.rmtree(ctr_studio_dir, ignore_errors=True)
    #Extract CTR Studio's newest version
    with zipfile.ZipFile("ctr.zip", 'r') as zip_ref:
        zip_ref.extractall(parent_dir)
    os.rename(parent_dir / f"net6.0", ctr_studio_dir)
    #Move the settings and recent projects files back
    if os.path.exists(parent_dir / "ConfigGlobal.json"):
        shutil.move(parent_dir / "ConfigGlobal.json", ctr_studio_dir / "ConfigGlobal.json")
    if os.path.exists(parent_dir / "Recent.txt"):
        shutil.move(parent_dir / "Recent.txt", ctr_studio_dir / "Recent.txt")
    #Remove the old Themes and Presets folders
    if os.path.exists(ctr_studio_dir / "Lib/Themes"):
        shutil.rmtree(ctr_studio_dir / "Lib/Themes", ignore_errors=True)
    if os.path.exists(ctr_studio_dir / "Lib/Presets"):
        shutil.rmtree(ctr_studio_dir / "Lib/Presets", ignore_errors=True)
    #Move the Themes and Presets folders back
    if os.path.exists(parent_dir / "Themes"):
        shutil.move(parent_dir / "Themes", ctr_studio_dir / "Lib")
    if os.path.exists(parent_dir / "Presets"):
        shutil.move(parent_dir / "Presets", ctr_studio_dir / "Lib")
    #Clean up the downloaded zip
    os.remove("ctr.zip")
    print(f"CTR Studio has been updated to {name}!")


#Reset the variable if no argument hasn't been set
if len(sys.argv) > 1:
    argument = sys.argv[1]
else:
    argument = ""

#Set the directory to the user's directory
user_dir = str(Path.home())

#Set the JSON path
json_path = f"{user_dir}/ctrut.json"

#Create the JSON file if it doesn't exist
if not os.path.exists(json_path):
    with open(json_path, "w") as json_file:
        data = {}
        json.dump(data, json_file)

# Set a variable in a JSON file based on the folder passed by the user
if argument == "set":
    # If there's a second argument
    if len(sys.argv) > 2:
        #Get the argument
        folder = sys.argv[2]
        # If the folder exists and CTR Studio.exe is in it
        if Path(folder + "/CTR Studio.exe").exists():
            with open(json_path, "r") as json_file:
                data = json.load(json_file)
            data["path"] = f"{folder}"
            with open(json_path, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("CTR Studio folder has been set")
            sys.exit(0)
        else:
            print("The folder you provided isn't one where CTR Studio is installed")
            sys.exit(1)
    else:
        print("You did not provide a folder")
        print("\nUse: ctrut set <ctr_studio_folder>")
        sys.exit(1)
#Create a shortcut to the verify action on the start menu
elif argument == "shortcut":
    # If there's a second argument
    if len(sys.argv) > 2:
        #Get the argument
        type = sys.argv[2]
        if getattr(sys, 'frozen', False):
            exe_file = sys.executable
            #Set shortcut attributes
            if type == "desktop":
                shortcut_path = Path(os.environ['USERPROFILE']) / 'Desktop' / 'CTR Studio.lnk'
            elif type == "startmenu":
                shortcut_path = Path(os.environ['APPDATA']) / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs' / 'CTR Studio.lnk'
            else:
                print("You did not provide a shortcut type")
                print("\nUse: ctrut shortcut desktop/startmenu")
                sys.exit(1)
            with open(json_path, "r") as json_file:
                data = json.load(json_file)
                if "path" in data:
                    if  data["path"] == "":
                        print("CTR Studio folder is not set not set")
                        print("\nUse: cput set <ctr_studio_folder>")
                        sys.exit(1)
                    elif data["path"] != "":
                        #Create the icon
                        exe_path = data["path"] + "/CTR Studio.exe"
                        size=(256, 256)

                        large, small = win32gui.ExtractIconEx(exe_path, 0)
                        icon_handle = large[0] if large else small[0]

                        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                        hbmp = win32ui.CreateBitmap()
                        hbmp.CreateCompatibleBitmap(hdc, size[0], size[1])

                        hdc_mem = hdc.CreateCompatibleDC()
                        hdc_mem.SelectObject(hbmp)
                        win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, icon_handle, size[0], size[1], 0, None, win32con.DI_NORMAL)

                        bmpinfo = hbmp.GetInfo()
                        bmpstr = hbmp.GetBitmapBits(True)
                        img = Image.frombuffer(
                            'RGBA',
                            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                            bmpstr, 'raw', 'BGRA', 0, 1
                        )

                        img.save(user_dir + "/ctr.png")

                        image = Image.open(user_dir + "/ctr.png")
                        image = image.convert('RGBA')
                        image.save(user_dir + "/ctr.ico", format='ICO', sizes=[(64, 64), (128, 128), (256, 256)])
                        icon_path = user_dir + "/ctr.ico"
                        #Create the shortcut
                        os.system(f"powershell.exe -Command \"$s=(New-Object -COM WScript.Shell).CreateShortcut('{shortcut_path}');$s.TargetPath='{exe_file}';$s.Arguments='verify';$s.IconLocation='{icon_path}';$s.Save()\"")
                        print("Shortcut has been created")
                        sys.exit(0)
                else:
                    print("CTR Studio folder is not set not set")
                    print("\nUse: ctrut set <ctr_studio_folder>")
                    sys.exit(1)
                json_file.close()
    else:
        print("You did not provide a shortcut type")
        print("\nUse: ctrut shortcut desktop/startmenu")
        sys.exit(1)
#Update CTR Studio to it's latest version
elif argument == "update":
    with open(json_path, "r") as json_file:
        #Load the JSON file
        data = json.load(json_file)
        #If the path value is empty
        if "path" in data:
            if  data["path"] == "":
                print("CTR Studio folder is not set not set")
                print("\nUse: ctrut set <ctr_studio_folder>")
                sys.exit(1)
        else:
            print("CTR Studio folder is not set not set")
            print("\nUse: ctrut set <ctr_studio_folder>")
            sys.exit(1)
        #Get CTR Studio's folder and it's parent
        ctr_studio_dir = Path(data["path"])
        parent_dir = Path(ctr_studio_dir).parent
        #URL to CTR Studio's repo from Github's api
        url = "https://api.github.com/repos/MapStudioProject/CTR-Studio/releases/latest"
        #Parse the returnd JSON
        response = requests.get(url)
        if response.status_code == 200:
            #Get the infos from the JSON
            release_info = json.loads(response.text)
            name = release_info['name']
            #Put the version in the JSON file
            data["version"] = f"{name}"
            with open(json_path, "w") as json_file:
                json.dump(data, json_file, indent=4)
            #Update CTR Studio
            update()
        json_file.close()
        sys.exit(0)
elif argument == "verify":
    #read the version value from the json file
    with open(json_path, "r") as json_file:
        #Load the JSON file
        data = json.load(json_file)
        #If the path value is empty
        if "path" in data:
            if  data["path"] == "":
                print("CTR Studio folder is not set not set")
                print("\nUse: ctrut set <ctr_studio_folder>")
                sys.exit(1)
        else:
            print("CTR Studio folder is not set not set")
            print("\nUse: ctrut set <ctr_studio_folder>")
            sys.exit(1)
        ctr_studio_dir = Path(data["path"])
        parent_dir = Path(ctr_studio_dir).parent
        ctr_studio_ver = data["version"]
        json_file.close()
        #Check for updates
        print(f"Checking for updates...\n")
        url = "https://api.github.com/repos/MapStudioProject/CTR-Studio/releases/latest"
        #Parse the returnd JSON
        response = requests.get(url)
        if response.status_code == 200:
            #Get the infos from the JSON
            release_info = json.loads(response.text)
            name = release_info['name']
            #If the version value is different from the latest version
            if ctr_studio_ver != name:
                #Put the version in the JSON file
                with open(json_path, "r") as json_file:
                    data = json.load(json_file)
                    data["version"] = f"{name}"
                    with open(json_path, "w") as json_file:
                        json.dump(data, json_file, indent=4)
                    json_file.close()
                #Update CTR Studio
                update()
                #Run CTR Studio's executable
                print("\nLaunching CTR Studio...\n")
                os.system(f'start "" "{ctr_studio_dir}/CTR Studio.exe"')
                sys.exit(0)
            elif ctr_studio_ver == name:
                #Run CTR Studio's executable
                print("Launching CTR Studio...\n")
                os.system(f'start "" "{ctr_studio_dir}/CTR Studio.exe"')
                sys.exit(0) 
else:
    #Show help
    print("CTR Studio updater")
    print("\nctrut set <ctr_studio_folder>         Set CTR Studio folder")
    print("ctrut shortcut desktop/startmenu      Create a shortcut to the verify function in the start menu")
    print("ctrut update                          Update CTR Studio to it's latest version")
    print("ctrut verify                          Verify if CTR Studio is up to date and run it")
    sys.exit(1)