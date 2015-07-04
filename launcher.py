from subprocess import Popen

def launch():
    try:
        file = open(r'../settings.inc', "r")
        for x, line in enumerate(file):
            if x == 1 and ("isDisabled=True" in line):
                file.close()
                break
            elif x > 1:
                file.close()
                process = Popen("getLinkTargets.exe")
                process.wait()
    except OSError:
        process = Popen("getLinkTargets.exe")
        process.wait()

launch()
