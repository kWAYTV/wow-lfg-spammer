import os
try:
    import os, time, colorama
    import win32com.client as comctl
    from colorama import Fore, Back, Style
    print(f"\n{Fore.MAGENTA}[{Fore.RESET}!{Fore.MAGENTA}] {Fore.RESET}Imports successful!")
    time.sleep(1)
except:
    print("\nImports failed! Trying to install...")
    z = "python -m pip install "; os.system('%scolorama' % (z)); os.system('%sos-sys' % (z)); os.system('%spywin32' % (z));
    print(f"\n{Fore.MAGENTA}[{Fore.RESET}!{Fore.MAGENTA}] {Fore.RESET}Imports successful!")
    time.sleep(1)

# Imports
import os, time, colorama
import win32com.client as comctl
from colorama import Fore, Back, Style

# Logo
logo = """
░██╗░░░░░░░██╗░█████╗░░██╗░░░░░░░██╗  ██╗░░░░░███████╗░██████╗░
░██║░░██╗░░██║██╔══██╗░██║░░██╗░░██║  ██║░░░░░██╔════╝██╔════╝░
░╚██╗████╗██╔╝██║░░██║░╚██╗████╗██╔╝  ██║░░░░░█████╗░░██║░░██╗░
░░████╔═████║░██║░░██║░░████╔═████║░  ██║░░░░░██╔══╝░░██║░░╚██╗
░░╚██╔╝░╚██╔╝░╚█████╔╝░░╚██╔╝░╚██╔╝░  ███████╗██║░░░░░╚██████╔╝
░░░╚═╝░░░╚═╝░░░╚════╝░░░░╚═╝░░░╚═╝░░  ╚══════╝╚═╝░░░░░░╚═════╝░"""

# Variables
clear = lambda: os.system("cls" if os.name in ("nt", "dos") else "clear") # Don't touch this
os.system(f"title World of Warcraft - LFG Bot")
print(Fore.MAGENTA + logo + Style.RESET_ALL)
clear()

# Get message
lfg = input(f"{Fore.YELLOW}[{Fore.WHITE}?{Fore.YELLOW}]{Fore.WHITE} Message{Fore.YELLOW} > {Fore.WHITE}")
wsh = comctl.Dispatch("WScript.Shell")

# Start
print(Fore.MAGENTA + logo + Style.RESET_ALL)
clear()

# Loop
while True:
    count = 0
    try:
        print(f"{Fore.GREEN}[{Fore.WHITE}+{Fore.GREEN}]{Fore.WHITE} Focusing World of Warcraft")
        wsh.AppActivate("World of Warcraft")
        print(f"{Fore.CYAN}[{Fore.WHITE}+{Fore.GREEN}]{Fore.CYAN} World of Warcraft Activated")
        time.sleep(1)
        wsh.SendKeys("{ENTER}")
        print(f"{Fore.GREEN}[{Fore.WHITE}+{Fore.GREEN}]{Fore.WHITE} Entering chat")
        time.sleep(1)
        wsh.SendKeys("/join lookingforgroup")
        time.sleep(0.5)
        wsh.SendKeys("{ENTER}")
        print(f"{Fore.GREEN}[{Fore.WHITE}+{Fore.GREEN}]{Fore.WHITE} Joined LFG channel")
        time.sleep(3)
        wsh.SendKeys("{ENTER}")
        time.sleep(1)
        wsh.SendKeys(f"/2 {lfg}")
        time.sleep(0.5)
        wsh.SendKeys("{ENTER}")
        print(f"{Fore.MAGENTA}[{Fore.WHITE}+{Fore.MAGENTA}]{Fore.WHITE} Message sent")
        time.sleep(1.5)
        print(f"{Fore.GREEN}[{Fore.WHITE}+{Fore.GREEN}]{Fore.WHITE} Going back to party channel")
        wsh.SendKeys("{ENTER}")
        time.sleep(1)
        wsh.SendKeys("/p")
        time.sleep(0.5)
        wsh.SendKeys("{ENTER}")
        count += 1
        os.system(f"title World of Warcraft - LFG Bot - Sent {count} messages")
        print(f"{Fore.MAGENTA}[{Fore.WHITE}+{Fore.MAGENTA}]{Fore.WHITE} Waiting 5 seconds")
        time.sleep(5)
        os.system(f"title World of Warcraft - LFG Bot - Sending message number {count}")
    except KeyboardInterrupt:
        print(f"{Fore.RED}[{Fore.WHITE}!{Fore.RED}]{Fore.WHITE} Stopped by user. Exiting...")
        time.sleep(1)
        exit()
    except Exception as e:
        print(f"{Fore.RED}[{Fore.WHITE}!{Fore.RED}]{Fore.WHITE} Error: {str(e)}")
        time.sleep(1)
        exit()