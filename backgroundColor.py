from utils import *

def change_background_color(slide,NewThemeColor,
                            _gradient_angle,
                            _transparency,
                            _change_background_color,):
    '''
    change background color of each slide
    '''
    if _change_background_color:
        background = slide.background
        fill       = background.fill

        # if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
        #     fill = shape.fill
        #     # print(fill.type)
        # change SOLID color
        if fill.type == MSO_FILL.SOLID:
            set_background_solid_color(fill,NewThemeColor,_transparency)
        # change GRADIENT color
        elif fill.type == MSO_FILL.GRADIENT:
            set_fill_gradient_color(fill,NewThemeColor,2,_gradient_angle)
        # change BACKGROUND color
        elif fill.type == MSO_FILL.BACKGROUND:
            set_background_solid_color(fill,NewThemeColor,_transparency)
        # set Nonecolor type to transparent
        else:
            set_fill_solid_color(fill,NewThemeColor,_transparency)

        
    else:
        pass