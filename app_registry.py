_app_instance = None


def set_app(app):
    """设置全局 app 实例（由主入口在创建 GUI 后调用）"""
    global _app_instance
    _app_instance = app


def get_app():
    """获取全局 app 实例"""
    return _app_instance