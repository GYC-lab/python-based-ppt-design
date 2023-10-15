from utils import *

def read_outline_color(shape):
    '''
    read outline color of shape
    '''
    line_fill = shape.line.fill
    print("fill-type == %s" % line_fill.type)
    if line_fill.type != MSO_FILL.SOLID:
        return

    line_color = line_fill.fore_color
    print("color-type == %s" % line_color.type)
    if line_color.type == MSO_COLOR_TYPE.SCHEME:
        print("color == %s" % line_color.theme_color)
    elif line_color.type == MSO_COLOR_TYPE.RGB:
        print("color == %s" % line_color.rgb)
    else:
        print("No line color")

def change_outline_color(slide,NewThemeColor,_change_outline_color):
    '''
    change outline color of each shape
    '''
    if _change_outline_color:
        for shape in slide.shapes:
            # if shape.shape_type in [MSO_SHAPE_TYPE.FREEFORM]:
            if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
                line = shape.line
                line_color = line.color
                set_color_by_type(line_color,NewThemeColor)

        group_shapes = [
            shp for shp in slide.shapes
            if shp.shape_type == MSO_SHAPE_TYPE.GROUP
        ]    
        
        for group_shape in group_shapes:
            for shape in group_shape.shapes:
                # if shape.shape_type in [MSO_SHAPE_TYPE.FREEFORM]:
                if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
                    line = shape.line
                    line_color = line.color
                    set_color_by_type(line_color,NewThemeColor)
    else:
        pass