_app_instance = None


class _NullApp:
    """空操作 app，用于独立测试等不需要 GUI 日志的上下文"""
    @staticmethod
    def log(msg):
        pass


def set_app(app):
    """设置全局 app 实例（由主入口在创建 GUI 后调用）"""
    global _app_instance
    _app_instance = app


def get_app():
    """获取全局 app 实例，未设置时返回安全的空对象"""
    global _app_instance
    if _app_instance is None:
        return _NullApp()
    return _app_instance