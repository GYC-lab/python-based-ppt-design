from pptx import Presentation
from utils import *

# 着色1#00BCA7  着色2#6BFFB8  着色3#2CEAA3  着色4#AAEFDF  着色5#9EE37D  着色6#63C132
# 文背深1#00584E 文背浅1#EEFFFD 文背深2#FFFFFF 文背浅1#FDFFFF
colormode = 0
if colormode == 0:
    myThemeColor = {
        'myACCENT_1': '#00BCA7',
        'myACCENT_2': '#6BFFB8',
        'myACCENT_3': '#2CEAA3',
        'myACCENT_4': '#AAEFDF',
        'myACCENT_5': '#9EE37D',
        'myACCENT_6': '#63C132',
        'myDARK_1'  : '#00584E',
        'myLIGHT_1' : '#EEFFFD',
        'myDARK_2'  : '#FFFFFF',
        'myLIGHT_2' : '#FDFFFF',
        'DEFAULT'   : '#000000'
    }
elif colormode == 1:
    myThemeColor = {
        'myACCENT_1': '#FFFFFF',
        'myACCENT_2': '#FFFFFF',
        'myACCENT_3': '#FFFFFF',
        'myACCENT_4': '#FFFFFF',
        'myACCENT_5': '#FFFFFF',
        'myACCENT_6': '#FFFFFF',
        'myDARK_1'  : '#FFFFFF',
        'myLIGHT_1' : '#FFFFFF',
        'myDARK_2'  : '#FFFFFF',
        'myLIGHT_2' : '#FFFFFF',
        'DEFAULT'   : '#FFFFFF'
    }
else:
    myThemeColor = {
        'myACCENT_1': '#000000',
        'myACCENT_2': '#000000',
        'myACCENT_3': '#000000',
        'myACCENT_4': '#000000',
        'myACCENT_5': '#000000',
        'myACCENT_6': '#000000',
        'myDARK_1'  : '#000000',
        'myLIGHT_1' : '#000000',
        'myDARK_2'  : '#000000',
        'myLIGHT_2' : '#000000',
        'DEFAULT'   : '#000000'
    }

NewThemeColor = set_my_theme_color(myThemeColor)

# prs = Presentation('motiongo科技1.pptx')
prs = Presentation('motiongo科技1.pptx')
prs = Presentation('motiongo科技1_sub4.pptx')
prs = Presentation('motiongo科技1_sub5.pptx')
prs = Presentation('motiongo科技1_sub3.pptx')

for slide in prs.slides:

    isGradient = True
    change_all_font_color  (slide,NewThemeColor)
    # change_outline_color   (slide,NewThemeColor)
    change_background_color(slide,NewThemeColor,isGradient)
    change_fill_color      (slide,NewThemeColor,isGradient)

# Save the presentation
prs.save('motiongo科技1_sub2_new.pptx')
