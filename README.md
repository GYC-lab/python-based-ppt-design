<!-- info -->

# 基本信息

- Date  : 2023-10-15
- Description：pptx自动化修改主题颜色

该脚本是按页来修改pptx文件的主题颜色，主要是修改字体颜色、形状填充颜色、形状轮廓颜色。此外也可按形状来修改，具体可参考`utils.py`文件中的函数。

# 代码文件

- `main.py`           : 主程序
- `utils.py`          : 一些工具函数
- `backgroundColor.py`: 用于设置主题颜色
- `fontColor.py`      : 用于设置字体的颜色
- `fillColor.py`      : 用于设置形状的填充颜色
- `outlineColor.py`   : 用于设置形状的轮廓颜色
- `requirement.txt`   : 需要安装的第三方库，通过以下方式安装：

```cmd
pip install -r requirement.txt
```

此外，代码文件中的函数下方均有注释，具体使用方法可参考注释。

# 基本功能

1. 根据自定义的主题颜色修改字体颜色、形状填充颜色、形状轮廓颜色，具体实现方式见`utils.py`文件中的`change_color`函数，是通过if-else判断语句实现的。
目前python的pptx库无法直接修改ppt中的主题颜色，pptx库的作者推荐用自己的模板来改：cf.
  - https://github.com/scanny/python-pptx/issues/917
  - https://stackoverflow.com/questions/57482998/set-theme-with-python-pptx
2. 填充颜色可选择纯色和渐变色，其中渐变色可设置两种以上的颜色，以及修改渐变的角度
3. 可修改填充颜色的透明度，通过`alpha`参数设置，取值范围为`[0, 1]`
4. 可选择是否修改所有幻灯片或者只修改选中的幻灯片

# 使用说明

1. 将需要修改的ppt文件放在`input`文件夹下；
2. 运行`main.py`：`main.py`中的`main`函数中的`myThemeColor`参数为主题颜色，具体在`themeColor.py`文件中设置，可根据需要自行修改。若要关闭某个功能，可在settings中将其设置为`0`；
3. 修改后的ppt文件将会保存在`output`文件夹下。

