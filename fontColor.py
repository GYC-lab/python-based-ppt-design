from utils import *

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