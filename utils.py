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

def SubElement(parent, tagname, **kwargs):
        element = OxmlElement(tagname)
        element.attrib.update(kwargs)
        parent.append(element)
        return element

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
    
def add_new_gradient_stop(fill):
    '''
    add new gradient stop to fill
    '''
    gsLst = fill.gradient_stops._gsLst
    new_gs = deepcopy(gsLst[0])
    gsLst.append(new_gs)
    new_gradient_stop = fill.gradient_stops[-1]
    return new_gradient_stop
    
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
    '''
    fill         : fill object of shape
    NewThemeColor: user-defined theme color
    alpha        : transparency
    '''
    fill.solid()
    fill_color = fill.fore_color
    set_color_by_type(fill_color,NewThemeColor) # set color
    set_solid_transparency(fill, alpha)         # set transparency

def set_fill_gradient_color(fill,NewThemeColor,num_colors,angle,alpha=1):
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

    try:
        fill.gradient_angle = angle          
    except:
        # print("fail to set gradient angle, maybe is not a linear gradient")
        print("")
    
def report_choice(_change_font_color,_change_outline_color,_change_background_color,_change_shape_color):
    '''
    report choice of user
    '''
    print('-------------------------')
    print('you are going to change:')
    if _change_font_color:
        print('- font color')
    if _change_outline_color:
        print('- outline color')
    if _change_background_color:
        print('- background color')
    if _change_shape_color:
        print('- shape color')
    print('-------------------------')
        