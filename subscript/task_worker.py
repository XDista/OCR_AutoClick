import os
import time
import datetime
import configparser
import json
import cv2
import win32gui
from app_registry import get_app
from project_paths import TASKS_DIR, REFS_DIR, MAIN_CONFIG_PATH
from window_utils import get_window_client_rect, get_target_window_from_config, is_window_visible
from screenshot_capture import capture_window
from image_matcher import template_match
from action_executor import execute_action, send_windows_notification
from config_manager import init_task_config

# 工作线程使用的全局任务索引
current_task_index = 0


def worker(app):
    stop_execution = False
    global current_task_index
    current_task_index = 0
    consecutive_error_count = 0
    scheduled_once_triggered = False
    stop_source = ""
    task_continuous_match = {}
    cached_target_window = None

    worker_gen = app.worker_generation

    while not app.stop_flag:
        if worker_gen != app.worker_generation:
            app.log("🔁 检测到新工作线程启动，旧线程自动退出")
            break

        is_normal_execution = True
        error_msg = ""

        try:
            main_config = configparser.ConfigParser()
            main_config.read(MAIN_CONFIG_PATH, encoding="utf-8")

            freq = float(main_config["GENERAL"]["recognition_frequency"])
            start_time_str = main_config["GENERAL"]["next_start_time"]
            task_group = main_config["GENERAL"]["current_task_group"]
            max_consecutive_errors = int(main_config["GENERAL"]["max_execution_errors"])
            enable_schedule = main_config["GENERAL"].getboolean("enable_schedule")
            schedule_mode = main_config["GENERAL"]["schedule_mode"]

            target_window_title = main_config["GENERAL"].get("target_window_title", "").strip()
            target_program_name = main_config["GENERAL"].get("target_program_name", "").strip()

            screenshot_mode = main_config["GENERAL"]["screenshot_mode"]
            adb_device_serial = main_config["ADBConfig"]["adb_device_serial"]

            try:
                match_step = float(main_config["GENERAL"].get("template_match_step", "0.05"))
                match_step = max(0.001, min(0.5, match_step))
            except ValueError:
                match_step = 0.05

            click_mode = main_config["GENERAL"]["click_mode"]

            adb_usage_mode = main_config["ADBConfig"].get("adb_usage_mode", "1")
            action_config = {
                "adb_device_serial": adb_device_serial,
                "adb_usage_mode": adb_usage_mode,
            }

            now = datetime.datetime.now()
            need_wait = False

            if enable_schedule and not (schedule_mode == "once" and scheduled_once_triggered):
                try:
                    target_time = datetime.datetime.strptime(start_time_str, "%H:%M:%S").replace(
                        year=now.year, month=now.month, day=now.day,
                    )

                    if now < target_time:
                        wait_sec = (target_time - now).total_seconds()
                        app.log(f"【定时等待】{schedule_mode}模式，距离{start_time_str}还有 {wait_sec:.1f} 秒")
                        time.sleep(min(wait_sec, freq))
                        need_wait = True
                    else:
                        need_wait = False
                        if schedule_mode == "once":
                            app.log(f"【定时触发】仅一次模式已触发（{start_time_str}）")
                            scheduled_once_triggered = True
                        else:
                            target_time += datetime.timedelta(days=1)
                            app.log(f"【定时触发】始终模式已触发，下次：{target_time.strftime('%H:%M:%S')}")
                except ValueError:
                    is_normal_execution = False
                    error_msg = f"定时时间格式错误（{start_time_str}）"
                    app.log(f"⚠️ {error_msg}")

            if not need_wait:
                target_window = None
                if cached_target_window is not None:
                    try:
                        hwnd = cached_target_window._hWnd
                        if is_window_visible(hwnd):
                            title_valid = True
                            if target_window_title:
                                current_title = win32gui.GetWindowText(hwnd)
                                if target_window_title not in current_title:
                                    app.log(f"缓存窗口标题已变更（预期包含'{target_window_title}'，当前：'{current_title}'），重新枚举...")
                                    title_valid = False

                            if title_valid:
                                target_window = cached_target_window
                            else:
                                cached_target_window = None
                        else:
                            app.log(f"缓存窗口失效（句柄：{hwnd}），重新枚举目标窗口...")
                            cached_target_window = None
                    except Exception as e:
                        app.log(f"缓存窗口访问异常：{e}，重新枚举...")
                        cached_target_window = None

                if target_window is None:
                    target_window = get_target_window_from_config()
                    cached_target_window = target_window

                if not target_window:
                    is_normal_execution = False
                    error_msg = "未找到目标程序窗口（请检查配置的进程/标题关键词）"
                    app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                else:
                    try:
                        window_text = win32gui.GetWindowText(target_window._hWnd)
                        window_handle = target_window._hWnd
                        client_left, client_top, client_width, client_height = get_window_client_rect(target_window._hWnd)
                        app.log(f"✅ 找到目标窗口 - 标题：{window_text} | 句柄：{window_handle} | 客户区位置/尺寸：({client_left},{client_top}) {client_width}x{client_height}")

                        task_path = os.path.join(TASKS_DIR, f"{task_group}.json")
                        if not os.path.exists(task_path):
                            init_task_config(task_group)
                        with open(task_path, "r", encoding="utf-8") as f:
                            task_config = json.load(f)

                        task_error = False
                        stop_execution = False

                        task_list = list(task_config.keys())

                        while current_task_index < len(task_list) and not stop_execution:
                            screenshot = capture_window(target_window, screenshot_mode=screenshot_mode, adb_device_serial=adb_device_serial)
                            if screenshot is None:
                                is_normal_execution = False
                                error_msg = "窗口截图失败"
                                app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                                task_error = True
                                break
                            else:
                                app.log(f"✅ 任务[{task_list[current_task_index]}]截图成功 - 尺寸：{screenshot.shape[1]}x{screenshot.shape[0]}")
                                screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

                            if app.stop_flag:
                                stop_execution = True
                                break

                            task_name = task_list[current_task_index]

                            try:
                                task = task_config[task_name]
                                ignore_occlusion = task.get("ignore_occlusion", False)

                                ref_images_list = task.get("ref_images", [])
                                if not ref_images_list:
                                    task_error = True
                                    error_msg = f"任务[{task_name}]：ref_images数组为空"
                                    app.log(f"⚠️ {error_msg}")
                                    break

                                matched_branch = None
                                matched_branch_idx = -1

                                if ignore_occlusion:
                                    matched_branch = ref_images_list[0]
                                    matched_branch_idx = 0
                                    app.log(f"✅ 任务[{task_name}]：配置了忽略遮挡，直接执行第一个分支")
                                else:
                                    for branch_idx, ref_branch in enumerate(ref_images_list):
                                        threshold = ref_branch.get("similarity_threshold", 0.9)
                                        match_times = ref_branch.get("match_times", 1)
                                        match_times = match_times if match_times >= 1 else 1
                                        run_on_match = ref_branch.get("run_on_match", True)

                                        branch_key = f"{task_name}_branch{branch_idx}"
                                        if branch_key not in task_continuous_match:
                                            task_continuous_match[branch_key] = 0

                                        image_config = ref_branch["image"]
                                        is_match = False
                                        similarity = 0.0

                                        if isinstance(image_config, str):
                                            ref_path = os.path.join(REFS_DIR, image_config)
                                            match_result = template_match(screenshot_gray, ref_path, threshold, match_step=match_step)
                                            if len(match_result) == 2:
                                                is_match, similarity = match_result
                                            elif len(match_result) == 3:
                                                is_match, similarity, match_loc = match_result
                                            else:
                                                is_match = False
                                                similarity = 0.0
                                                app.log(f"⚠️ 任务[{task_name}]分支{branch_idx}：模板匹配返回值异常")
                                        elif isinstance(image_config, list) and len(image_config) >= 2:
                                            operator = image_config[0].lower()
                                            image_list = image_config[1:]

                                            if operator == "and":
                                                all_matched = True
                                                max_sim = 0.0
                                                unmatched_images = []
                                                for img_name in image_list:
                                                    img_ref_path = os.path.join(REFS_DIR, img_name)
                                                    match_result = template_match(screenshot_gray, img_ref_path, threshold, match_step=match_step)
                                                    if len(match_result) >= 2:
                                                        img_match, img_sim = match_result[0], match_result[1]
                                                        max_sim = max(max_sim, img_sim)
                                                        if not img_match:
                                                            all_matched = False
                                                            unmatched_images.append(f"{img_name}(相似度:{img_sim:.3f})")
                                                            break
                                                    else:
                                                        all_matched = False
                                                        unmatched_images.append(f"{img_name}(匹配失败)")
                                                        break
                                                is_match = all_matched
                                                similarity = max_sim
                                                if all_matched:
                                                    app.log(f"任务[{task_name}]分支{branch_idx}：AND逻辑匹配成功 | 所有图像均匹配")
                                                else:
                                                    app.log(f"任务[{task_name}]分支{branch_idx}：AND逻辑匹配失败 | 未匹配的图像: {', '.join(unmatched_images)}")

                                            elif operator == "or":
                                                any_matched = False
                                                max_sim = 0.0
                                                matched_images = []
                                                for img_name in image_list:
                                                    img_ref_path = os.path.join(REFS_DIR, img_name)
                                                    match_result = template_match(screenshot_gray, img_ref_path, threshold, match_step=match_step)
                                                    if len(match_result) >= 2:
                                                        img_match, img_sim = match_result[0], match_result[1]
                                                        max_sim = max(max_sim, img_sim)
                                                        if img_match:
                                                            any_matched = True
                                                            matched_images.append(f"{img_name}(相似度:{img_sim:.3f})")
                                                            break
                                                    else:
                                                        pass
                                                is_match = any_matched
                                                similarity = max_sim
                                                if any_matched:
                                                    app.log(f"任务[{task_name}]分支{branch_idx}：OR逻辑匹配成功 | 匹配的图像: {', '.join(matched_images)}")
                                                else:
                                                    app.log(f"任务[{task_name}]分支{branch_idx}：OR逻辑匹配失败 | 所有图像均未匹配")

                                            else:
                                                app.log(f"⚠️ 任务[{task_name}]分支{branch_idx}：未知逻辑运算符 '{operator}'，仅支持 'and' 和 'or'")
                                                is_match = False
                                                similarity = 0.0
                                        else:
                                            app.log(f"⚠️ 任务[{task_name}]分支{branch_idx}：无效的图像配置格式")
                                            is_match = False
                                            similarity = 0.0

                                        if run_on_match:
                                            if is_match:
                                                task_continuous_match[branch_key] += 1
                                            else:
                                                task_continuous_match[branch_key] = 0
                                        else:
                                            if not is_match:
                                                task_continuous_match[branch_key] += 1
                                            else:
                                                task_continuous_match[branch_key] = 0

                                        current_continuous = task_continuous_match[branch_key]
                                        match_type = "正向匹配" if run_on_match else "反向匹配"

                                        if isinstance(image_config, str):
                                            app.log(f"任务[{task_name}]分支{branch_idx}({image_config})：{match_type} | 相似度 {similarity:.3f} | 单次匹配: {is_match} | 连续满足次数: {current_continuous}/{match_times}")
                                        else:
                                            app.log(f"任务[{task_name}]分支{branch_idx}({image_config[0].upper()}逻辑)：{match_type} | 结果: {is_match} | 连续满足次数: {current_continuous}/{match_times}")

                                        if current_continuous >= match_times:
                                            matched_branch = ref_branch
                                            matched_branch_idx = branch_idx
                                            if run_on_match:
                                                app.log(f"✅ 任务[{task_name}]：分支{branch_idx}连续匹配成功，已选中此分支")
                                            else:
                                                app.log(f"✅ 任务[{task_name}]：分支{branch_idx}连续匹配失败（反向匹配），已选中此分支")
                                            break

                                    if matched_branch is None:
                                        app.log(f"任务[{task_name}]：所有分支均未达到匹配条件，下一周期将继续检查")
                                        break

                                app.log(f"任务[{task_name}]分支{matched_branch_idx}：开始执行{len(matched_branch.get('actions', []))}个动作...")

                                jump_triggered = False
                                actions = matched_branch.get("actions", [])

                                for action in actions:
                                    if not isinstance(action, dict) or "type" not in action or "params" not in action:
                                        app.log(f"⚠️ 任务[{task_name}]动作格式错误，跳过：{action}")
                                        continue
                                    action_type = action["type"]
                                    params = action["params"]

                                    if app.stop_flag:
                                        stop_execution = True
                                        app.log("  - 检测到停止指令，终止当前动作执行")
                                        break

                                    result = execute_action(target_window, action_type, params, stop_flag=app.stop_flag, click_mode=click_mode, action_config=action_config)

                                    if isinstance(result, tuple):
                                        log_msg = result[0]
                                    else:
                                        log_msg = result
                                    app.log(f"  - {log_msg}")

                                    if isinstance(result, tuple):
                                        if len(result) >= 2 and result[1]:
                                            stop_execution = True
                                            if len(result) >= 3 and result[2] == "stop_action":
                                                stop_source = "stop_action"
                                            app.root.after(0, lambda: app._stop(is_manual=False))
                                            app.stop_flag = True
                                            break

                                        if len(result) == 3 and not result[1]:
                                            current_task_index = result[2]
                                            current_task_index = max(0, min(current_task_index, len(task_list) - 1))
                                            jump_triggered = True
                                            break

                                for branch_idx in range(len(ref_images_list)):
                                    branch_key = f"{task_name}_branch{branch_idx}"
                                    task_continuous_match[branch_key] = 0

                                if not jump_triggered and not stop_execution:
                                    current_task_index += 1

                            except Exception as e:
                                task_error = True
                                error_msg = f"任务[{task_name}]执行失败：{str(e)}"
                                app.log(f"⚠️ 任务执行错误：{error_msg}")
                                break

                        if task_error:
                            is_normal_execution = False
                            app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")
                    except Exception as e:
                        is_normal_execution = False
                        error_msg = f"窗口操作失败：{str(e)}"
                        app.log(f"执行错误：{error_msg}（连续错误：{consecutive_error_count + 1}/{max_consecutive_errors}）")

            if is_normal_execution:
                if consecutive_error_count > 0:
                    app.log(f"✅ 执行正常，连续错误计数器已重置为0（之前：{consecutive_error_count}）")
                consecutive_error_count = 0
            else:
                consecutive_error_count += 1

            if consecutive_error_count >= max_consecutive_errors:
                final_msg = f"连续执行错误达到{max_consecutive_errors}次，任务已停止"
                app.log(final_msg)
                send_windows_notification("自动点击工具 - 任务停止", final_msg)
                app.stop_flag = True
                app.root.after(0, lambda: app._stop(is_manual=False))
                break

            if not need_wait:
                app.stop_event.wait(timeout=freq)

        except Exception as e:
            is_normal_execution = False
            error_msg = str(e)
            consecutive_error_count += 1
            app.log(f"❌ 未预期的执行错误：{error_msg}（连续错误：{consecutive_error_count}/{max_consecutive_errors}）")

            if consecutive_error_count >= max_consecutive_errors:
                final_msg = f"连续执行错误达到{max_consecutive_errors}次，任务已停止"
                app.log(final_msg)
                send_windows_notification("自动点击工具 - 任务停止", final_msg)
                app.stop_flag = True
                app.root.after(0, app._stop)
                break
            if stop_source == "stop_action":
                app.log("🛑 【指令停止】由stop动作指令触发，程序已停止")
            elif app.stop_flag and stop_source == "":
                app.log("🛑 【手动停止】用户手动触发，程序已停止")
            else:
                app.log("🛑 程序已停止")

            time.sleep(1.0)