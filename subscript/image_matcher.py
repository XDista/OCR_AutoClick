import numpy as np
import cv2
from PIL import Image
from app_registry import get_app
from gpu_accelerator import get_gpu_accelerator

_gpu_initialized = False


def _ensure_gpu():
    global _gpu_initialized
    if not _gpu_initialized:
        get_gpu_accelerator().detect()
        _gpu_initialized = True

def template_match(screenshot_gray, ref_image_path, threshold=0.9, match_step=0.05):
    """支持同比例拉伸的多尺度模板匹配

    :param screenshot_gray: 窗口截图（已转换好的灰度图）
    :param ref_image_path: 参考图路径
    :param threshold: 匹配阈值
    :param match_step: 多尺度匹配步长
    :return: (是否匹配成功, 最高相似度)
    """
    _ensure_gpu()
    app = get_app()

    try:
        ref_pil = Image.open(ref_image_path).convert('L')
        ref_img = np.array(ref_pil)
    except FileNotFoundError:
        try:
            if app:
                app.log(f"⚠️ 错误：参考图不存在 → {ref_image_path}")
        except Exception:
            pass
        return False, 0.0
    except Exception as e:
        try:
            if app:
                app.log(f"⚠️ 错误：读取参考图失败 → {e}")
        except Exception:
            pass
        return False, 0.0

    if screenshot_gray is None or screenshot_gray.size == 0:
        return False, 0.0
    ref_h, ref_w = ref_img.shape[:2]
    max_similarity = 0.0

    scales = np.arange(0.1, 4.1, match_step)
    for scale in scales:
        scaled_w = int(ref_w * scale)
        scaled_h = int(ref_h * scale)

        if scaled_w <= 0 or scaled_h <= 0:
            continue
        if scaled_w > screenshot_gray.shape[1] or scaled_h > screenshot_gray.shape[0]:
            continue

        if scale < 1.0:
            scaled_ref = cv2.resize(ref_img, (scaled_w, scaled_h), interpolation=cv2.INTER_AREA)
        else:
            scaled_ref = cv2.resize(ref_img, (scaled_w, scaled_h), interpolation=cv2.INTER_CUBIC)

        result = cv2.matchTemplate(screenshot_gray, scaled_ref, cv2.TM_CCOEFF_NORMED)
        _, current_sim, _, _ = cv2.minMaxLoc(result)

        if current_sim > max_similarity:
            max_similarity = current_sim

        if max_similarity >= threshold:
            break

    is_match = max_similarity >= threshold
    return is_match, max_similarity