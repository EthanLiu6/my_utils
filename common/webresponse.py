from typing import TypeVar
import traceback

T = TypeVar("T")


class WebResponse:
    code: int = 200  # 状态码
    msg: str = "OK"  # 提示信息
    data: T
    success: bool = True

    def __init__(self, data: T = None, code: int = 200, msg: str = "OK", success: bool = True):
        self.code = code
        self.msg = msg
        self.data = data
        self.success = success

        # 如果是 500 错误，并且 msg 中包含异常信息，自动打印 traceback
        if code == 500 and isinstance(msg, str) and "Traceback" not in msg:
            # 尝试打印 traceback
            traceback.print_exc()
