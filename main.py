from pptx import Presentation
from utils import *
from themeColor import *

# 着色1#00BCA7  着色2#6BFFB8  着色3#2CEAA3  着色4#AAEFDF  着色5#9EE37D  着色6#63C132
# 文背深1#00584E 文背浅1#EEFFFD 文背深2#FFFFFF 文背浅1#FDFFFF

myThemeColor = myThemeColor_user
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
