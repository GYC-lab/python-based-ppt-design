from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.dml import MSO_COLOR_TYPE
from utils import *

# 着色1#00BCA7  着色2#6BFFB8  着色3#2CEAA3  着色4#AAEFDF  着色5#9EE37D  着色6#63C132
# 文背深1#00584E 文背浅1#EEFFFD 文背深2#FFFFFF 文背浅1#FDFFFF

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
    'DEFAULT'   : '#FFFFFF'
}
NewThemeColor = set_my_theme_color(myThemeColor)

prs = Presentation('motiongo科技1.pptx')

for slide in prs.slides:
    background = slide.background
    fill = background.fill
    fill.solid()
    
    # change_background_color(slide,NewThemeColor)
    change_all_font_color  (slide,NewThemeColor)
    change_all_table_color (slide,NewThemeColor)
    change_all_chart_color (slide,NewThemeColor)

prs.save('motiongo科技1_new.pptx')