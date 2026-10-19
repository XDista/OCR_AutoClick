# OCR_AutoClick Tool by XDista

一个基于 Python 3.12 的泛用型辅助 OCR 点击工具，主要用于《碧蓝航线》和配合 Alas 使用。



## 功能（部分）

 - 点击指定位置、滑动（窗口相对坐标，支持后台）
 - 截图识别（模板匹配 / OCR 文字识别）与多分支处理（AND / OR 逻辑、正向反向匹配）
 - 模拟键盘按键（单键与组合键）、执行关机指令和 CMD 命令
 - 原生或自动拼接 serial 的 ADB 指令
 - 通过系统通知发送消息
 - 支持多种截图方式
 - GPU 硬件加速（OpenCV OpenCL + ONNX DirectML）


## 运行环境

建议在 Windows 10 及以上版本使用，Windows 7、8、8.1 兼容性未测试。


## 项目结构

```
OCR_AutoClick/
├── main.pyw                  # 入口文件，双击启动 GUI
├── config.ini                # 主配置文件
├── gpu.json                  # GPU 检测缓存
├── app_registry.py           # 全局 app 实例注册
├── subscript/                # 核心模块
│   ├── gui.py                # GUI 界面
│   ├── task_worker.py        # 后台任务工作线程
│   ├── action_executor.py    # 动作执行器
│   ├── ocr_engine.py         # OCR 引擎（基于 RapidOCR + ONNX Runtime）
│   ├── image_matcher.py      # 多尺度模板匹配
│   ├── screenshot_capture.py # 截图捕获
│   ├── config_manager.py     # 配置管理
│   ├── window_utils.py       # 窗口工具
│   ├── gpu_accelerator.py    # GPU 硬件加速检测与配置
│   └── project_paths.py      # 项目路径定义
├── tasks/                    # 任务组配置文件
├── refs/                     # 模板匹配参考图文件夹（支持子文件夹）
├── task_ocr/                 # OCR 任务配置文件
├── adb/platform-tools/       # 自带的 ADB 工具
├── install/                  # 安装说明与环境依赖
└── theme/                    # 主题配置
 
```


## Q&A

### Q：为什么会有这个工具？

### A：因为 alas 脚本虽然强大，但有时更新不够及时，不能很好地符合我的需求。

另一方面，在 alas 不同配置之间切换的做法并不符合 alas 的本意，即使提交了 request 也绝不可能被接受。因此我仿照 alas，自己用 AI 搓了一个简易的点击脚本。

<br>

### Q：如何使用？

### A：
0. 配合 Alas（https://github.com/LmeSzinc/AzurLaneAutoScript）使用效果更佳
- 源码版
1. 下载解压 `OCR_AutoClick (版本号).zip` 文件
2. 按照 `install` 文件夹内的安装说明，安装必要环境与依赖库
   - 请确认确实有需要再下载 `install.zip`，这个压缩包储存的是"所有"可能需要用到的工具
3. 双击 `main.pyw` 即可打开脚本 GUI


- 便捷版		
1. 下载 `OCR_AutoClick (版本号)-便捷版.zip` 文件
2. 解压文件
3. 双击 `启动.bat` 即可打开脚本 GUI

<br>

### Q：为什么用 ADB 和 win32gui 等不同的截图识别模式，识别出来的匹配率不一样？

### A：ADB 截图和 win32gui 等方式截图对应的图片色彩是有差异的，可以用附带的 ADB 截图工具和 Windows 自带截图截取同一个部位进行对比以验证这一结论。

<br>

### Q：这个脚本除了那几个任务组里的动作，还能做什么？

### A：你可以根据自己的需要进行编辑和更改任务组配置，但不会超过任务组本身的动作范围。任务组所有支持的动作与示例都已经在 `tasks\1a_json新版任务组说明.txt` 里说明。

<br>

### Q：你的脚本是外挂吗？用你的脚本有被黄鸡封号的风险吗？

### A：对于这个问题，请移步对于 alas 相同问题的讨论并进行独立思考。作为 alas 的辅助脚本，在是否是外挂的问题上，此脚本和 alas 的回答是一致的。如果你认为 alas 是外挂，则这个脚本就是；相反，如果你认为 alas 不是外挂，那这个脚本就不是。

至于封号风险，别明着改客户端改全皮肤改无敌无限伤害无 CD 什么的，一般不会封号。

<br>

> **注意**：经网友提醒，使用此脚本有被其他游戏的反作弊误封的风险（如 Nikke 使用的腾讯 ACE），请有类似环境的网友慎用！对于可能的封号风险，此脚本与作者概不负责！
>
> - （暂不清楚触发 ACE 的是这个脚本还是 alas）
> - （ACE 之间亦有力度之别）

<br>

### Q：你的脚本该不会是伪装成点击器的病毒吧？

### A：😡 脚本是完全开源且没做任何加密的！所有代码都以明文呈现在你眼前了！如果自己看不懂，可以交给 AI 看！👊👊👊（给 AI 看前记得改后缀！部分网页版 AI 未必支持 `.pyw` 格式！）

<br>

### Q：为什么不像 alas 那样做成 exe？

### A：因为真的不会，没这个能力知道吧 😭
	2026.10：嗨嗨嗨，已经实现便捷版了，解压运行即可😊

<br>

### Q：你的源代码怎么这么奇怪？

### A：因为我根本没学过 Python，所有代码都是 AI 写的，我只负责许愿和微调参数。

说实话这份源码被豆包、DeepSeek、GitHub Copilot 三个 AI 转手了十多遍还能正常工作已经谢天谢地了 😋

<br>

### Q：issue："大佬大佬，你能帮我解决 alas 的 xxx 问题吗？"

### A：我真的求你了 😅 
![qaa](https://github.com/user-attachments/assets/b26d67af-68c9-4cef-979b-9f95b8ab8afb)

### 我也不是 alas 项目组的成员，在 bilibili 视频评论区底下问我，我顺手帮着回答一下，对着没有前因后果没有日志信息的截图玩猜猜看已经相当费劲了；都上 GitHub 了为什么不直接去 alas 的项目里提 issue？

