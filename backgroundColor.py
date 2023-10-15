from utils import *

def change_background_color(slide,NewThemeColor,_change_background_color):
    '''
    change background color of each slide
    '''
    if _change_background_color:
        background = slide.background
        fill       = background.fill

        # if isGradient:
        #     set_fill_gradient_color(fill,NewThemeColor,2,90)
        # else:
        #     set_color_by_type(fill_color,NewThemeColor)

        # print(fill.type)
        # if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
        #     fill = shape.fill
        #     # change SOLID color
        #     # print(fill.type)
        if fill.type in [MSO_FILL.SOLID]:
            set_fill_solid_color(fill,NewThemeColor)
        elif fill.type == MSO_FILL.GRADIENT:
            set_fill_gradient_color(fill,NewThemeColor,2,90)
    else:
        pass