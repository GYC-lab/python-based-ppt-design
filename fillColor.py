from utils import *

def change_fill_color(slide,NewThemeColor,isGradient=False):
    '''
    change fill color of each shape (not including picture)
    '''
    for shape in slide.shapes:
        # print(shape.shape_type)
        # print("---")
        if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
            fill = shape.fill
            # change SOLID color
            # print(fill.type)
            if fill.type in [MSO_FILL.SOLID,MSO_FILL.BACKGROUND]:
                set_fill_solid_color(fill,NewThemeColor)
            elif fill.type == MSO_FILL.GRADIENT:
                set_fill_gradient_color(fill,NewThemeColor,2,90)
        
    group_shapes = [
        shp for shp in slide.shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP
    ]    
    
    for group_shape in group_shapes:
        for shape in group_shape.shapes:
            if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
                fill = shape.fill
                if fill.type in [MSO_FILL.SOLID,MSO_FILL.BACKGROUND]:
                    set_fill_solid_color(fill,NewThemeColor)
                elif fill.type == MSO_FILL.GRADIENT:
                    set_fill_gradient_color(fill,NewThemeColor,2,90)