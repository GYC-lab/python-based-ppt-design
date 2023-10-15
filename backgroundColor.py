from utils import *

def change_background_color(slide,NewThemeColor,isGradient=False):
    '''
    change background color of each slide
    '''
    background = slide.background
    fill       = background.fill

    # if isGradient:
    #     set_fill_gradient_color(fill,NewThemeColor,2,90)
    # else:
    #     set_color_by_type(fill_color,NewThemeColor)
    if fill.type in [MSO_FILL.SOLID,MSO_FILL.BACKGROUND]:
        fill.solid()
        if isGradient:
            set_fill_gradient_color(fill,NewThemeColor,2,90)
        else:
            set_fill_solid_color(fill,NewThemeColor)