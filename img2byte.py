# The purpose of this script is to convert all images to bytes to ease the bundling of the application
# The script is run one time before converting to exe
# The result of this script will output a file named pic2str.py contains all bytes for all pictures
# The bytes are imported to the main.py and converted back to images and used accordingly.

import base64


def pic2str(file, functionName):
    pic = open(file, 'rb')
    content = '{} = {}\n'.format(functionName, base64.b64encode(pic.read()))
    pic.close()

    with open('pic2str.py', 'a') as f:
        f.write(content)


if __name__ == '__main__':
    pic2str('BLOCK_MIDDLE_EAST.png', 'BLOCK_MIDDLE_EAST')
    pic2str('DESKTOP_ICON.png', 'DESKTOP_ICON')
    pic2str('play_on_eu.png', 'play_on_eu')
    pic2str('play_on_na_east.png', 'play_on_na_east')
    pic2str('play_on_na_west.png', 'play_on_na_west')
    pic2str('SQUARE_BACKGROUND_MINA_TEST.png', 'SQUARE_BACKGROUND_MINA_TEST')
    pic2str('UNBLOCK_ALL_MAIN.png', 'UNBLOCK_ALL_MAIN')
    pic2str('LOGO_SMALL_APPLICATION.png', 'LOGO_SMALL_APPLICATION')

