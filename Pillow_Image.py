from PIL import Image


Dimensions = (606, 606)

Sound_Button = Image.open("IST_Project\ProjectV3\sound_button.PNG")

Sound_Button_Final = Sound_Button.resize(Dimensions)

Sound_Button_Final.show()