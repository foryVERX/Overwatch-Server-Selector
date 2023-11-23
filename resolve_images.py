import base64
from io import BytesIO
from PIL import ImageTk, Image


def byteToTkImage(byte):
    image_in_byte = base64.b64decode(byte)  # decoded byte
    image = BytesIO(image_in_byte)  # Make the decoded byte readable for PIL
    return ImageTk.PhotoImage(Image.open(image))  # Resolve byte to ImageTk
