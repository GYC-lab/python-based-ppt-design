from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from copy import deepcopy

def read_outline_color(shape):
    line_fill = shape.line.fill
    print("fill-type == %s" % line_fill.type)
    # ---we only handle solid, which is most common
    # ---the other common value is "background" which means no-fill
    if line_fill.type != MSO_FILL.SOLID:
        return

    # ---color can be specified as specific RGB color or a theme color
    # ---like ACCENT_1
    line_color = line_fill.fore_color
    print("color-type == %s" % line_color.type)
    if line_color.type == MSO_COLOR_TYPE.SCHEME:
        print("color == %s" % line_color.theme_color)
    elif line_color.type == MSO_COLOR_TYPE.RGB:
        print("color == %s" % line_color.rgb)
    else:
        print("No line color")

def read_font_color(font):
    # font = shape.text_frame.paragraphs[0].runs[0].font
    font_color = font.color
    print("color-type == %s" % font_color.type)
    if font_color.type == MSO_COLOR_TYPE.SCHEME:
        print("color == %s" % font_color.theme_color)
    elif font_color.type == MSO_COLOR_TYPE.RGB:
        print("color == %s" % font_color.rgb)
    else:
        print("No font color")

def hex_to_rgb(hex_color):
    '''
    convert hex color to rgb color
    '''
    # remove #
    hex_color = hex_color.lstrip('#')

    # convert to RGB
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)

    return r, g, b

def set_my_theme_color(myThemeColor):
    '''
    set user-defined theme color based on myThemeColor
    '''
    NewThemeColor = {}
    for key in myThemeColor:
        r,g,b = hex_to_rgb(myThemeColor[key])
        NewThemeColor[key] = RGBColor(r,g,b)
    return NewThemeColor

def change_color(object_color,NewThemeColor):
    '''
    replace original theme color with user-defined theme color
    '''
    if object_color.theme_color == MSO_THEME_COLOR.ACCENT_1:
        return NewThemeColor['myACCENT_1']
    elif object_color.theme_color == MSO_THEME_COLOR.ACCENT_2:
        return NewThemeColor['myACCENT_2']
    elif object_color.theme_color == MSO_THEME_COLOR.ACCENT_3:
        return NewThemeColor['myACCENT_3']
    elif object_color.theme_color == MSO_THEME_COLOR.ACCENT_4:
        return NewThemeColor['myACCENT_4']
    elif object_color.theme_color == MSO_THEME_COLOR.ACCENT_5:
        return NewThemeColor['myACCENT_5']
    elif object_color.theme_color == MSO_THEME_COLOR.ACCENT_6:
        return NewThemeColor['myACCENT_6']
    elif object_color.theme_color == MSO_THEME_COLOR.DARK_1:
        return NewThemeColor['myDARK_1']
    elif object_color.theme_color == MSO_THEME_COLOR.BACKGROUND_1 or object_color.theme_color == MSO_THEME_COLOR.TEXT_1:
        return NewThemeColor['myLIGHT_1']
    elif object_color.theme_color == MSO_THEME_COLOR.DARK_2:
        return NewThemeColor['myDARK_2']
    elif object_color.theme_color == MSO_THEME_COLOR.BACKGROUND_2 or object_color.theme_color == MSO_THEME_COLOR.TEXT_2:
        return NewThemeColor['myLIGHT_2']
    else:
        return NewThemeColor['DEFAULT']
    
def set_color_by_type(object_color,NewThemeColor):
    '''
    set color according to type of object_color
    '''
    if object_color.type == MSO_COLOR_TYPE.SCHEME:
        object_color.rgb = change_color(object_color,NewThemeColor)
    elif object_color.type == MSO_COLOR_TYPE.RGB:
        object_color.rgb = change_color(object_color,NewThemeColor)
    else:
        object_color.rgb = NewThemeColor['DEFAULT']     # set Nonecolor type

def set_fill_solid_color(fill,NewThemeColor):
    fill.solid()
    fill_color = fill.fore_color
    set_color_by_type(fill_color,NewThemeColor)       

def set_fill_gradient_color(fill,NewThemeColor):
    fill.gradient()

def change_all_font_color(slide,NewThemeColor):
    '''
    change all font color of each slide
    '''
    
    # check if title exists and change font color
    title = slide.shapes.title
    if title is not None:
        print(title.text)
        title_color = title.text_frame.paragraphs[0].font.color
        set_color_by_type(title_color,NewThemeColor)
    
    # check if subtitle exists and change font color
    num_of_placeholders = slide.placeholders.__len__()
    for i in range(1,num_of_placeholders):
        subtitle = slide.placeholders[i]
        subtitle_color = subtitle.text_frame.paragraphs[0].font.color
        set_color_by_type(subtitle_color,NewThemeColor)

    # change font color of each shape
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font_color = font.color
                    set_color_by_type(font_color,NewThemeColor)

    # only operate on group shapes
    group_shapes = [
        shp for shp in slide.shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP
    ]
    for group_shape in group_shapes:
        for shape in group_shape.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        font = run.font
                        font_color = font.color
                        set_color_by_type(font_color,NewThemeColor)

def change_background_color(slide,NewThemeColor,isGradient=False):
    '''
    change background color of each slide
    '''
    background = slide.background
    fill = background.fill
    fill.solid()
    fill_color = fill.fore_color
    if isGradient:
        set_fill_gradient_color(fill,5,0)
    else:
        set_color_by_type(fill_color,NewThemeColor)

def change_fill_color(slide,NewThemeColor,isGradient=False):
    '''
    change fill color of each shape
    '''
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            fill = shape.fill
            if isGradient:
                set_fill_gradient_color(fill,5,0)
            else:
                set_fill_solid_color(fill,NewThemeColor)

    group_shapes = [
        shp for shp in slide.shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP
    ]    
    
    for group_shape in group_shapes:
        for shape in group_shape.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                fill = shape.fill
                if isGradient:
                    set_fill_gradient_color(fill,5,0)
                else:
                    set_fill_solid_color(fill,NewThemeColor)

def add_new_gradient_stop(fill):
    '''
    add new gradient stop to fill
    '''
    gsLst = fill.gradient_stops._gsLst
    new_gs = deepcopy(gsLst[0])
    gsLst.append(new_gs)
    new_gradient_stop = fill.gradient_stops[-1]
    return new_gradient_stop

def set_fill_gradient_color(fill,num_colors,angle):
    '''
    fill      : fill object of shape
    num_colors: number of gradient stops
    angle     : angle of gradient
    '''
    fill.gradient()
    num_of_gradient_stops = num_colors

    addition_color = [RGBColor(255, 128, 128),RGBColor(255, 128, 255),RGBColor(128, 255, 255)]
    fill.gradient_stops[0].color.rgb = RGBColor(255, 0, 0)  
    fill.gradient_stops[1].color.rgb = RGBColor(0, 255,0 )  # 

    for i in range(2,num_of_gradient_stops):
        new_gradient_stop = add_new_gradient_stop(fill)
        new_gradient_stop.color.rgb = addition_color[i-2]

    for i in range(0,num_of_gradient_stops):
        fill.gradient_stops[i].position = i/(num_of_gradient_stops-1)

    fill.gradient_angle = angle          

def change_outline_color(slide,NewThemeColor):
    '''
    change outline color of each shape
    '''
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            line = shape.line
            line_color = line.color
            set_color_by_type(line_color,NewThemeColor)

    group_shapes = [
        shp for shp in slide.shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP
    ]    
    
    for group_shape in group_shapes:
        for shape in group_shape.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                line = shape.line
                line_color = line.color
                set_color_by_type(line_color,NewThemeColor)