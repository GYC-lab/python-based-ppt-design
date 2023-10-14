from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.dml.color import RGBColor
from pptx.dml.color import ColorFormat
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.xmlchemy import OxmlElement
from copy import deepcopy

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
    elif object_color.theme_color == MSO_THEME_COLOR.LIGHT_1:
        return NewThemeColor['myLIGHT_1']
    elif object_color.theme_color == MSO_THEME_COLOR.DARK_2:
        return NewThemeColor['myDARK_2']
    elif object_color.theme_color == MSO_THEME_COLOR.LIGHT_2:
        return NewThemeColor['myLIGHT_2']
    else:
        return NewThemeColor['myDARK_1']
    
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

def SubElement(parent, tagname, **kwargs):
        element = OxmlElement(tagname)
        element.attrib.update(kwargs)
        parent.append(element)
        return element

def set_solid_transparency(fill, alpha):
    """ Set the transparency (alpha) of a solid fill"""
    realAlpha = (100-int(alpha*100))*1000
    ts = fill._xPr
    with open('xml.xml', 'a') as f:
        f.write(ts.xml)    
    ts = fill._xPr.solidFill
    sF = ts.get_or_change_to_srgbClr()
    sE = SubElement(sF, 'a:alpha', val=str(realAlpha))

def set_gradient_transparency(fill, alpha, stop_i):
    """ Set the transparency (alpha) of a gradient fill"""
    realAlpha = (100-int(alpha*100))*1000

    ts = fill._xPr.gradFill.gsLst[stop_i]
    sF = ts.get_or_change_to_srgbClr()
    sE = SubElement(sF, 'a:alpha', val=str(realAlpha))

def set_fill_solid_color(fill,NewThemeColor,alpha=1):
    fill.solid()
    fill_color = fill.fore_color
    set_color_by_type(fill_color,NewThemeColor) # set color
    set_solid_transparency(fill, alpha)         # set transparency

def change_all_font_color(slide,NewThemeColor):
    '''
    change all font color of each slide
    '''
    
    # check if title exists and change font color
    title = slide.shapes.title
    if title is not None:
        # print(title.text)
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
                    # print(font_color.type,run.text)
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
                        # print(font_color.type,run.text)
                        set_color_by_type(font_color,NewThemeColor)

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
            if fill.type in [MSO_FILL.SOLID,MSO_FILL.BACKGROUND,MSO_FILL.GRADIENT]:
                if isGradient:
                    set_fill_gradient_color(fill,NewThemeColor,2,90)
                else:
                    set_fill_solid_color(fill,NewThemeColor)
        
    group_shapes = [
        shp for shp in slide.shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP
    ]    
    
    for group_shape in group_shapes:
        for shape in group_shape.shapes:
            if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
                fill = shape.fill
                if fill.type in [MSO_FILL.SOLID,MSO_FILL.BACKGROUND,MSO_FILL.GRADIENT]:
                    # print(fill.type)
                    if isGradient:
                        set_fill_gradient_color(fill,NewThemeColor,2,90)
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

def set_fill_gradient_color(fill,NewThemeColor,num_colors,angle,alpha=0.01):
    '''
    fill      : fill object of shape
    num_colors: number of gradient stops
    angle     : angle of gradient
    '''
    fill.gradient()
    
    num_of_gradient_stops = num_colors
    addition_color = [RGBColor(255, 128, 128),RGBColor(255, 128, 255),RGBColor(128, 255, 255)]
    
    fill.gradient_stops[0].color.rgb = NewThemeColor['myLIGHT_1']  
    fill.gradient_stops[1].color.rgb = NewThemeColor['myACCENT_1']

    for i in range(2,num_of_gradient_stops):
        new_gradient_stop = add_new_gradient_stop(fill)
        new_gradient_stop.color.rgb = addition_color[i-2]

    for i in range(0,num_of_gradient_stops):
        fill.gradient_stops[i].position = i/(num_of_gradient_stops-1)
        set_gradient_transparency(fill, alpha, i)   # set transparency

    fill.gradient_angle = angle          
    

def change_outline_color(slide,NewThemeColor):
    '''
    change outline color of each shape
    '''
    for shape in slide.shapes:
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
            if shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE,MSO_SHAPE_TYPE.FREEFORM]:
                line = shape.line
                line_color = line.color
                set_color_by_type(line_color,NewThemeColor)

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

def read_font_color(font):
    '''
    read font color of font
    '''
    font_color = font.color
    print("color-type == %s" % font_color.type)
    if font_color.type == MSO_COLOR_TYPE.SCHEME:
        print("color == %s" % font_color.theme_color)
    elif font_color.type == MSO_COLOR_TYPE.RGB:
        print("color == %s" % font_color.rgb)
    else:
        print("No font color")                