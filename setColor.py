from fontColor import * 
from outlineColor import *
from backgroundColor import *
from shapeColor import *

def set_color(slide,NewThemeColor,
              _gradient_angle,
              _transparency,
              _change_font_color,
              _change_outline_color,
              _change_background_color,
              _change_shape_color):
    '''
    choose which color to change
    '''

    change_font_color      (slide,NewThemeColor,_change_font_color)
    change_outline_color   (slide,NewThemeColor,_change_outline_color)
    change_background_color(slide,NewThemeColor,_gradient_angle,_transparency,_change_background_color)
    change_shape_color     (slide,NewThemeColor,_gradient_angle,_transparency,_change_shape_color)