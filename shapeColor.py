from utils import *

def change_shape_color(slide,NewThemeColor,
                        _gradient_angle,
                        _transparency,
                        _change_shape_color):
    '''
    change fill color of each shape (not including picture)
    '''
    if _change_shape_color:
        for shape in slide.shapes:
            if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,\
                                    MSO_SHAPE_TYPE.FREEFORM,  \
                                    ]:
                fill = shape.fill

                # change SOLID color
                if fill.type in [MSO_FILL.SOLID]:
                    set_fill_solid_color(fill,NewThemeColor,_transparency)
                # change GRADIENT color
                elif fill.type in [MSO_FILL.GRADIENT]:
                    set_fill_gradient_color(fill,NewThemeColor,2,_gradient_angle)
                # change BACKGROUND color
                elif fill.type in [MSO_FILL.BACKGROUND]:
                    set_fill_solid_color(fill,NewThemeColor,_transparency)   
                # set Nonecolor type to transparent
                else:
                    set_fill_solid_color(fill,NewThemeColor,1)

        group_shapes = [
            shp for shp in slide.shapes
            if shp.shape_type == MSO_SHAPE_TYPE.GROUP
        ]    
        
        for group_shape in group_shapes:
            for shape in group_shape.shapes:
                if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,\
                                        MSO_SHAPE_TYPE.FREEFORM,  \
                                        ]:
                    fill = shape.fill
                    # change SOLID color
                    if fill.type in [MSO_FILL.SOLID]:
                        set_fill_solid_color(fill,NewThemeColor,_transparency)
                    # change GRADIENT color
                    elif fill.type in [MSO_FILL.GRADIENT]:
                        set_fill_gradient_color(fill,NewThemeColor,2,_gradient_angle)
                    # change BACKGROUND color
                    elif fill.type in [MSO_FILL.BACKGROUND]:
                        set_fill_solid_color(fill,NewThemeColor,_transparency)
                    # set Nonecolor type to transparent
                    else:
                        set_fill_solid_color(fill,NewThemeColor,1)                        
    else:
        pass