# The purpose of this script is to convert all images to bytes to ease the bundling of the application
# The script is run one time before converting to exe
# The result of this script will output a file named pic2str.py contains all bytes for all pictures
# The bytes are imported to the main.py and converted back to images and used accordingly.

import base64


def pic2str(file, functionName):
    pic = open(file, "rb")
    content = "{} = {}\n".format(functionName, base64.b64encode(pic.read()))
    pic.close()

    with open("pic2str.py", "a") as f:
        f.write(content)


if __name__ == "__main__":
    """
    pic2str('BLOCK_MIDDLE_EAST.png', 'BLOCK_MIDDLE_EAST')
    pic2str('DESKTOP_ICON.png', 'DESKTOP_ICON')
    pic2str('play_on_eu.png', 'play_on_eu')
    pic2str('PROGRAMABLE_BUTTON.png', 'PROGRAMABLE_BUTTON')
    pic2str('play_on_na_east.png', 'play_on_na_east')
    pic2str('play_on_na_central.png', 'play_on_na_central')
    pic2str('play_on_na_west.png', 'play_on_na_west')
    pic2str('play_on_australia.png', 'play_on_australia')
    pic2str('Frame 1.png', 'SQUARE_BACKGROUND_MINA_TEST')
    pic2str('UNBLOCK_ALL.png', 'UNBLOCK_ALL_MAIN')
    pic2str('LOGO_SMALL_APPLICATION.png', 'LOGO_SMALL_APPLICATION')
    pic2str('CUSTOM_SETTINGS.png', 'CUSTOM_SETTINGS')
    pic2str('RESET_BUTTON.png', 'RESET_BUTTON')
    pic2str('OPEN_IP_LIST_BUTTON.png', 'OPEN_IP_LIST_BUTTON')
    pic2str('APPLY_BUTTON.png', 'APPLY_BUTTON')
    pic2str('Frame 2.png', 'CUSTOM_SETTINGS_BACKGROUND')
    pic2str('Donation Button.png', 'Donation_Button')
    pic2str('DISCORD.png', 'Discord_Button')
    pic2str('INSTALL_UPDATE.png', 'INSTALL_UPDATE')
    pic2str('TEST PING.png', 'TEST_PING')
    """
    pic2str(".img/ow2 v5/CANCEL.png", "CANCEL")
    pic2str(".img/ow2 v5/MANUALLY LOCATE.png", "MANUALLY LOCATE")
