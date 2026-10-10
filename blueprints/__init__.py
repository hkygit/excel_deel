"""Flask 蓝图包。

各蓝图定义在 blueprints/<name>.py 中，由 app.py 显式注册。

本 __init__ 刻意不导入任何蓝图，以避免：
- 循环导入（blueprint 之间可能相互引用工具模块）
- 注册顺序问题（依赖 app.container 初始化完成）
"""
