from fontColor import * 
from outlineColor import *
from backgroundColor import *
from fillColor import *

def set_color(slide,NewThemeColor,
              _change_font_color,
              _change_outline_color,
              _change_background_color,
              _change_fill_color):
    '''
    choose which color to change
    '''

    change_font_color      (slide,NewThemeColor,_change_font_color)
    change_outline_color   (slide,NewThemeColor,_change_outline_color)
    change_background_color(slide,NewThemeColor,_change_background_color)
    change_fill_color      (slide,NewThemeColor,_change_fill_color)