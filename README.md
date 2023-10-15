<!-- info -->
- Date  : 2023-10-15
- Author: 标贝数据

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

此外，代码文件中的函数下方均有注释。

# 基本功能

1. 根据自定义的主题颜色修改字体颜色、形状填充颜色、形状轮廓颜色
2. 填充颜色可选择纯色和渐变色，其中渐变色可设置两种以上的颜色
3. 可修改填充颜色的透明度，通过`alpha`参数设置，取值范围为`[0, 1]`
4. 可选择是否修改所有幻灯片或者只修改选中的幻灯片

# 使用说明

1. 将需要修改的ppt文件放在`input`文件夹下；
2. 运行`main.py`：`main.py`中的`main`函数中的`myThemeColor`参数为主题颜色，具体在`themeColor.py`文件中设置，可根据需要自行修改。若要关闭某个功能，可将其注释掉；
3. 修改后的ppt文件将会保存在`output`文件夹下。

