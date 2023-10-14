# 定义一个类
class MyClass:
    def __init__(self):
        self.my_attribute = 42

# 创建对象
obj = MyClass()

# 使用 hasattr 检查对象是否包含某个属性
if hasattr(obj, 'my_attribute'):
    print("对象包含 my_attribute 属性")
else:
    print("对象不包含 my_attribute 属性")