from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.dml import MSO_COLOR_TYPE

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
    
def change_all_font_color(slide,NewThemeColor):
    '''
    change all font color of each slide
    '''
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font_color = font.color
                    if font_color.type == MSO_COLOR_TYPE.SCHEME:
                        font.color.rgb = change_color(font_color,NewThemeColor)
                    elif font_color.type == MSO_COLOR_TYPE.RGB:
                        print("RGB color")
    
def change_all_table_color(slide,NewThemeColor):
    '''
    change all table color of each slide
    '''
    for shape in slide.shapes:
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            font = run.font
                            font_color = font.color
                            if font_color.type == MSO_COLOR_TYPE.SCHEME:
                                font.color.rgb = change_color(font_color,NewThemeColor)
                            elif font_color.type == MSO_COLOR_TYPE.RGB:
                                pass

def change_all_chart_color(slide,NewThemeColor):
    '''
    change all chart color of each slide
    '''
    for shape in slide.shapes:
        if shape.has_chart:
            for series in shape.chart.series:
                for point in series.points:
                    font = point.data_label.font
                    font_color = font.color
                    if font_color.type == MSO_COLOR_TYPE.SCHEME:
                        font.color.rgb = change_color(font_color,NewThemeColor)
                    elif font_color.type == MSO_COLOR_TYPE.RGB:
                        pass
                            
def change_background_color(slide,NewThemeColor):
    '''
    change background color of each slide
    '''
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = change_color(fill.fore_color,NewThemeColor)