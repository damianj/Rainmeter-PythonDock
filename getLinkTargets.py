from os import walk
from os.path import join, basename
from math import floor
import ctypes
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")


def get_targets(folder):
    links = []
    counter = 0

    for (dirpath, dirnames, filenames) in walk(folder):
        links.extend(join(dirpath, filename) for filename in filenames)
        break

    settings = open(r'../settings.inc', "w")
    settings.write("[Variables]\nisDisabled=True\n")

    for target in links:
        if ".lnk" in target or ".url" in target:
            counter += 1
            shortcut = shell.CreateShortCut(target)
            settings.write("mIndex" + str(counter+1) + "Process=" + basename(shortcut.Targetpath) + "\n")
        else:
            continue
    settings.close()
    calculate_parameters(counter)


def calculate_parameters(counter):
    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()
    monitor_height = user32.GetSystemMetrics(1)
    get_min = lambda left, right: left if(left < right) else right

    height = input("\n\nYour main monitor was detected to be {0} px in height.\nBy default the maximum height of the"
                   " dock will be set to 55% of this height ({1}px).\nDo you wish to set a different maximum height? "
                   "([Y]es/[N]o): ".format(monitor_height, round(monitor_height*(595.0/1080))))
    while True:
        if height.lower() == "yes" or height.lower() == "y":
            height = eval(input("\nPlease enter the desired height for the dock in pixels: "))
            break
        elif height.lower() == "no" or height.lower() == "n":
            height = round(monitor_height*(595.0/1080))
            break
        else:
            height = input("\nInvalid option. Please type Yes/Y or No/N to choose: ")

    icon_size = input("\nWould you like to use small (16x16 pixels), medium (32x32 pixels), "
                      "or large (48x48 pixels) icons? ([S]mall/[M]edium/[L]arge): ")
    while True:
        if icon_size.lower() == "s" or icon_size.lower() == "small":
            icon_size = "small"
            calculate_parameters.max_icons = floor((height-28)/18)
            height = (get_min(counter, calculate_parameters.max_icons)*22)+8
            calculate_parameters.y = "22r"
            print("Using {0} icons a maximum of {1} icons can be placed on the dock.\nGiven the number of icons "
                  "detected the height will be {2}px.".format(icon_size, calculate_parameters.max_icons, height))
            break
        elif icon_size.lower() == "m" or icon_size.lower() == "medium":
            icon_size = "medium"
            calculate_parameters.max_icons = floor((height-35)/40)
            height = get_min((40*counter)+15, (40*calculate_parameters.max_icons)+15)
            calculate_parameters.y = "40r"
            print("Using {0} icons a maximum of {1} icons can be placed on the dock.\nGiven the number of icons "
                  "detected the height will be {2}px.".format(icon_size, calculate_parameters.max_icons, height))
            break
        elif icon_size.lower() == "l" or icon_size.lower() == "large":
            icon_size = "large"
            calculate_parameters.max_icons = floor((height-35)/56)
            height = get_min((56*counter)+15, (56*calculate_parameters.max_icons)+15)
            calculate_parameters.y = "56r"
            print("Using {0} icons a maximum of {1} icons can be placed on the dock.\nGiven the number of icons "
                  "detected the height will be {2}px.".format(icon_size, calculate_parameters.max_icons, height))
            break
        else:
            icon_size = input("\nInvalid option. Please type Small/S, Medium/M, or Large/L to choose: ")

    print("\n\n")
    make_settings_file(icon_size, height, get_min(calculate_parameters.max_icons, counter),
                       calculate_parameters.y)


def make_settings_file(size, height, icon_num, y):
    if size == "small":
        # [width,y_init,x,text,t_x,t_y]
        settings = [22, 26, 3, '. . .', 10, 7]
    elif size == "medium":
        settings = [40, 30, 5, 'DOCK', 20, 10]
    else:
        settings = [58, 30, 5, 'D O C K', 28, 10]

    code = '[Metadata]\nName=Simple Dock\nAuthor=damianj and exper1mental (optimized the code)\n' \
           'Information=Creates a simple and minimalistic dock.\nVersion=2.0\nLicense=Creative Commons Attribution-' \
           'NonCommercial-ShareAlike 4.0 International\n\n[Rainmeter]\nUpdate=1000\nMouseActionCursor=1\n\n' \
           '[loadSettings]\nMeasure=Calc\nFormula=loadSettings + 1\nIfCondition=loadSettings = 1\n' \
           'IfTrueAction=["#@#python\launcher.exe"]\nIfFalseAction=[!DisableMeasure loadSettings]\n\n' \
           '[Variables]\n@Include=#@#settings.inc\nfolderpath=#@#Launcher\nSortType=Name\niconSize={0}\n' \
           'fontcolor=237,237,237,150\nhovercolor=237,237,237,250\n\n[IconStyle]\nX={1}\nY={2}\nAntiAlias=1\n' \
           'Group=IconGroup\n\n[mPath]\nMeasure=Plugin\nPlugin=FileView\nPath="#folderpath#"\nCount={3}\n' \
           'HideExtensions=0\nShowFolder=0\nFinishAction=[!UpdateMeasureGroup Children][!UpdateMeterGroup IconGroup]' \
           '[!Redraw]\nSortType=#sorttype#\n\n'.format(size, settings[2], y, icon_num+1)

    for n in range(1, icon_num+1):
        temp = '[MeasureProcess{0}]\nMeasure=Plugin\nPlugin=Process\nProcessName=#mIndex{1}Process#\nUpdate' \
               'Divider=0.3\nGroup=MeasureProcesses\nDynamicVariables=1\nOnChangeAction=[!EnableMeasure Measure' \
               'Condition{0}]\n\n[MeasureCondition{0}]\nMeasure=String\nDynamicVariables=1\nIfCondition=Measure' \
               'Process{0} = 1\nIfTrueAction=[!setoption Index{1}Icon Greyscale "0"][!setoption Index{1}Icon ' \
               'ImageAlpha "255"][!UpdateMeter Index{1}Icon][!Redraw][!DisableMeasure #CURRENTSECTION#]\nIfFalse' \
               'Action=[!setoption Index{1}Icon Greyscale "1"][!setoption Index{1}Icon ImageAlpha "160"][!Update' \
               'Meter Index{1}Icon][!Redraw][!DisableMeasure #CURRENTSECTION#]\nIfConditionMode=1\n\n'.format(n, n+1)
        code += temp

    for n in range(2, icon_num+2):
        temp = '[mIndex{0}Icon]\nMeasure=Plugin\nPlugin=FileView\nPath=[mPath]\nType=Icon\nIconSize=#iconSize#\n' \
               'Index={0}\nGroup=Children\n\n'.format(n)
        code += temp

    code += '[Background]\nMeter=Image\nimagename=#@#images\\background.jpg\nImageAlpha=100\nW={0}\nH={1}\nY=20\n' \
            'hidden=0\n\n[Header]\nMeter=Image\nimagename=#@#images\\background.jpg\nImageAlpha=200\nW={0}\nH=20\n\n' \
            '[Headertext]\nMeter=String\nMeterstyle=textstyle\nText="{2}"\nAntiAlias=1\nStringStyle=Bold\nStringalign' \
            '=CenterCenter\nFontFace=Tahoma\nFontSize=8\nMouseoveraction=[!setoption ' \
            '#CURRENTSECTION# Fontcolor "#Hovercolor#"][!UpdateMeter #CURRENTSECTION#][!Redraw]\nMouseleaveaction=' \
            '[!setoption #CURRENTSECTION# Fontcolor "#Fontcolor#"][!UpdateMeter #CURRENTSECTION#][!Redraw]\nfontcolor' \
            '=#Fontcolor#\nX={3}\nY={4}\nLeftMouseupaction=["#folderpath#"]' \
            '\n\n'.format(settings[0], height, settings[3], settings[4], settings[5])

    for n in range(2, icon_num+2):
        temp = '[Index{0}Icon]\nMeter=Image\nMeasureName=mIndex{0}Icon\nMouseoveraction=[!setoption #CURRENTSECTION# ' \
               'ColorMatrix1 "-0.3086;-0.3086;-0.3086;0;0"][!setoption #CURRENTSECTION# ColorMatrix2 "-0.3086;' \
               '-0.3086;-0.3086;0;0"][!setoption #CURRENTSECTION# ColorMatrix3 "-0.3086;-0.3086;-0.3086;0;0"]' \
               '[!setoption #CURRENTSECTION# ColorMatrix5 "1;1;1;0;1"][!UpdateMeter #CURRENTSECTION#][!Redraw]\n' \
               'Mouseleaveaction=[!setoption #CURRENTSECTION# ColorMatrix1 "1;0;0;0;0"][!setoption #CURRENTSECTION# ' \
               'ColorMatrix2 "0;1;0;0;0"][!setoption #CURRENTSECTION# ColorMatrix3 "0;0;1;0;0"][!setoption ' \
               '#CURRENTSECTION# ColorMatrix5 "0;0;0;0;1"][!UpdateMeasure MeasureProcess{1}][!UpdateMeter ' \
               '#CURRENTSECTION#][!Redraw]\nLeftMouseUpAction=[!CommandMeasure mIndex{0}Icon "FollowPath"]' \
               '[!UpdateMeasure mPath][!UpdateMeasureGroup Children][!UpdateMeter *][!Redraw]\nMeterStyle=' \
               'IconStyle'.format(n, n-1)

        if n < 3:
            code += temp + "\nY={0}\n\n".format(settings[1])
        elif n < icon_num+1:
            code += temp + "\n\n"
        else:
            code += temp

    launcher = open(r'../../Launcher/launcher.ini', "w")
    launcher.write(code)
    launcher.close()

get_targets(r"../Launcher")
