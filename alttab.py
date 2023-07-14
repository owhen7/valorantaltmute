#owen7 on discord.
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import time
from ctypes import windll
import win32gui
import msvcrt


user32 = windll.user32
user32.SetProcessDPIAware() # optional, makes functions return real pixel numbers instead of scaled values
full_screen_rect = (0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))

def is_full_screen():
    try:
        hWnd = user32.GetForegroundWindow()
        rect = win32gui.GetWindowRect(hWnd)
        return rect == full_screen_rect
    except:
        return False
    
def set_mute_state(application_name, state):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == application_name:
            volume.SetMute(state, None)


print('*****/,.             *#%%%(//(##(##%%%%%#%%&%#####(((((#/,.    ....,,*/(##(//(##\n(/**//*.  ....     .,*******/(##((((##########((//****,,..........,*/#%%%%#((#%%\n&#(/((/*,,..    ..,,,...,,*/((#(((///(((#(/////((((//***,.....,***//(#&&&&%##%&&\n@&#(((((/*,**///*,,....,,,/(##%#(((((((/////**/(##%(/******,...,*//(((((((#%&&@@\n@&%#(*//////((//*,.,,,***///(#%%##(/****/***//((#((/*********,,.,,*//**,,,/%&&@@\n@@@&#/,..*///**,,***/////****/(////****,,**///////**,,***//******,,***////#%&&@@\n@@&&%(/*,,... ........,**///********************/*****/((((/*,,....  ..,*(%@@@@@\n@@&%(,.                  .,***,,,,,,*******,,,,**,,,**//*,.               *#&@@@\n@%(*.                      .,*/*,,,,****,,,,,,,,,*//**,,....,,...           ,#&&\n&(,            ....          ,//**,.,*/***,...,*///*,.,,**/***,,..           .,,\n(*.             ....     .   .,*//*.,**///**/((##/,,,,,**,,.   .........        \n*,.  .                         .,**//((########((/******,,..  ....              \n..                              .,*(#%%%%%%#(*,,,/((/*,..,,......               \n..                             ...,*(#%%%%%#*  ..*//*,,......             ......\n(*.                          .,,,*/(#%&&&&&&#//***,,,,,*,,.              .......\n#*.                         ..,,,,*///////(((#/,..,,****,..                  .,,\n%#/*,.                    ..,..,**,.         .,,,,.,/(((*,                  .*//\n%#(/***,,.            ...,...*/(%%(*,.     .,/(#%(*,,.,**,..             .,/(%&&\n#((/,.,,*,,,,......,,.....,/(#%###%&%#(, ,/#%&&%#%%#/,. ..,............,,*//(#%%\n(((/*,.      ......    .*/((###########(*//(##%%%%%%%#(*,   .............,,*/(##\n/*,,,****,,..      ..,*/(##((((#####(/*,..,*/(####%######(/**,..        ..,/(###\n/,....,**,,,......,*((((///(((((((((/,..   .,*/((((((////((((/,.,,,,,,,,,,/(#(((\n*,,.    .,*******,,,,,,*,,**//(//*,...,,*,,.. ..,**////**,,,.,,,,***//////(((///\n,,,..,,,...,***/*,,.  ....,,,,..  ..,***,,,,,,,.....,,,..    .,*******/(##(*,,..\n***,,,,...   .  ......          ..,,***,,,,,,****,,..         .,,**////*,,...,,,\n*****,,,....                    ..   ...,...........      .,,,*****,,.    ..,,**\n*******,,.....                                          .,,,,,,,..     ..,,,,***\n*******,,,,.........                                                   ...,,****\n**,,,,,**,,,,,,.....                                                    ...,,,,,\n,***,.,,,,,,,.....    .                                         .....,,,...,,,,,\n,,****,.............       ..                               .....,,,,,,,,,,,,,,,\n,,,,****,............................                     ............,.........')

print("Thank you for using my Valorant Alt-Mute Tool.")
print("The Alt-Mute Tool is currently running. Press any key inside this window to close the program gracefully.")
print("Note: If you close this forcefully (by clicking the X or ctrl + c) while VALORANT is not the main window, you will have to re-enable sound like this:")
print("Right-click volume icon on toolbar in bottom right -> volume mixer --> Unmute VALORANT and the Riot client.")
print("\nhttps://github.com/owhen7/valorantaltmute")


#Main Loop of the program.
# "VALORANT-Win64-Shipping.exe" is the game sounds.
# "RiotClientServices.exe" appears to sometimes be responsible for the voice chat, though it isn't always and I'm not sure when it is.
#It may actually be unneccessary to mute it. Further testing is needed.
#Sometimes "VALORANT-Win64-Shipping.exe" is responsible for both voice and game sounds.

while(True):

    #If the user inputs any key, break out and end the program.
    if msvcrt.kbhit():
        break

    #If any application is full-screen, unmute valorant.
    if(is_full_screen()):
        set_mute_state("VALORANT-Win64-Shipping.exe", False)  # Unmute Valorant
        set_mute_state("RiotClientServices.exe", False)
    else: #Mute valorant if no application is full screen.
        set_mute_state("VALORANT-Win64-Shipping.exe", True)  # Mute Valorant
        set_mute_state("RiotClientServices.exe", True)
    time.sleep(0.5)

#Make sure to unmute VALORANT when the program is closed.
#These lines don't get ran if the user closes the program with ctrl + c.
set_mute_state("VALORANT-Win64-Shipping.exe", False)  # Unmute Valorant
set_mute_state("RiotClientServices.exe", False)