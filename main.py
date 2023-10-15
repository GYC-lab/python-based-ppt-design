from pptx import Presentation
from utils import *
from themeColor import *
from fontColor import * 
from outlineColor import *
from backgroundColor import *
from fillColor import *

# --------------------------------- settings --------------------------------- #
fill_color_gradient = 1
set_single_slile    = 0
slide_index         = 0 # 设置第0张幻灯片（set_single_slile = 1时有效）
input_ppt_path      = './input/'
output_ppt_path     = './output/'

# --------------------- set theme color at themeColor.py --------------------- #
myThemeColor  = myThemeColor_user
NewThemeColor = set_my_theme_color(myThemeColor)

# --------------------------------- load ppt --------------------------------- #
prs = Presentation(input_ppt_path+'motiongo科技1_sub5.pptx')   # 标题页
prs = Presentation(input_ppt_path+'motiongo科技1_sub4.pptx')   # 科技发展现状大纲
prs = Presentation(input_ppt_path+'motiongo科技1_sub2.pptx')   # 科技对商业的影响 - 背景介绍
prs = Presentation(input_ppt_path+'motiongo科技1_sub1.pptx')   # 目录页
prs = Presentation(input_ppt_path+'motiongo科技1.pptx')        # 全部
prs = Presentation(input_ppt_path+'motiongo科技1_sub3.pptx')   # 科技对商业的影响 - 趋势分析

# --------------------------------- main code -------------------------------- #
if set_single_slile:
    slide = prs.slides[slide_index]
    change_all_font_color  (slide,NewThemeColor)
    change_outline_color   (slide,NewThemeColor)
    change_background_color(slide,NewThemeColor,fill_color_gradient)
    change_fill_color      (slide,NewThemeColor,fill_color_gradient)
else:
    for slide in prs.slides:
        change_all_font_color  (slide,NewThemeColor)
        change_outline_color   (slide,NewThemeColor)
        change_background_color(slide,NewThemeColor,fill_color_gradient)
        change_fill_color      (slide,NewThemeColor,fill_color_gradient)

# --------------------------------- save ppt --------------------------------- #
prs.save(output_ppt_path+'motiongo科技1_new.pptx')
