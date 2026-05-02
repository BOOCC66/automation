import os
import time
from datetime import datetime

import xlwings as xw
import win32com.client


class SAPF44Auto:
    def __init__(self, excel_path: str, sheet_name: str = None):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.session = None

    def connect_sap(self):
        """连接已打开的 SAP GUI 会话"""
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            self.session = session
            return session
        except Exception as e:
            raise RuntimeError(
                "无法连接到 SAP GUI。请确认：\n"
                "1. 已登录 SAP\n"
                "2. 已打开一个可操作的会话\n"
                "3. SAP GUI Scripting 已启用\n"
                f"原始错误: {e}"
            )

    def safe_text(self, value, default=""):
        """把 Excel 单元格值转为可安全写入 SAP 的字符串"""
        if value is None:
            return default

        # Excel 日期类型处理
        if isinstance(value, datetime):
            return value.strftime("%d.%m.%Y")

        # Excel 数字转字符串时避免 3401.0 这种情况
        if isinstance(value, float) and value.is_integer():
            value = int(value)

        text = str(value).strip()
        return text if text else default

    def wait_for_element(self, element_id: str, timeout: int = 10, interval: float = 0.2):
        """等待控件出现"""
        start = time.time()
        last_err = None

        while time.time() - start < timeout:
            try:
                obj = self.session.findById(element_id)
                return obj
            except Exception as e:
                last_err = e
                time.sleep(interval)

        raise RuntimeError(f"等待控件超时: {element_id}，原始错误: {last_err}")

    def set_text_safely(self, element_id: str, value: str, timeout: int = 10):
        """更稳地给 SAP 字段赋值，并在失败时给出明确报错"""
        obj = self.wait_for_element(element_id, timeout=timeout)

        try:
            obj.text = value
            return
        except Exception as e:
            obj_type = "unknown"
            current_text = "<unreadable>"

            try:
                obj_type = obj.Type
            except Exception:
                pass

            try:
                current_text = obj.text
            except Exception:
                pass

            raise RuntimeError(
                f"字段不可写: {element_id} | type={obj_type} | current_text={current_text} | error={e}"
            )

    def start_f44(self):
        """进入 F-44"""
        s = self.session
        s.findById("wnd[0]").maximize()
        s.StartTransaction("F-44")
        time.sleep(0.5)

    def read_status_bar(self):
        """读取 SAP 底部状态栏消息"""
        try:
            bar = self.session.findById("wnd[0]/sbar")
            return bar.Text.strip()
        except Exception:
            return ""

    def close_popup_if_any(self):
        """若有弹窗，尝试读取并关闭"""
        try:
            popup = self.session.findById("wnd[1]")
            msg = ""

            try:
                msg = popup.Text.strip()
            except Exception:
                msg = "检测到 SAP 弹窗"

            # 优先按“确定/继续”
            try:
                popup.findById("tbar[0]/btn[0]").press()
            except Exception:
                try:
                    popup.findById("tbar[0]/btn[12]").press()
                except Exception:
                    pass

            return msg
        except Exception:
            return None

    def ensure_no_error_popup(self, prefix: str = ""):
        """检查是否有弹窗或明显错误状态"""
        popup_msg = self.close_popup_if_any()
        if popup_msg:
            raise RuntimeError(f"{prefix}弹窗信息: {popup_msg}")

        status_msg = self.read_status_bar()
        error_keywords = ["错误", "Error", "E:", "不存在", "未定义", "无权限", "not", "cannot"]
        if status_msg and any(k.lower() in status_msg.lower() for k in error_keywords):
            raise RuntimeError(f"{prefix}状态栏报错: {status_msg}")

    def fill_header_and_post(
        self,
        vendor_code: str,
        budat: str,
        monat: str,
        bukrs: str,
        waers: str,
        agums: str,
        header_text: str,
        xblnr: str,
    ):
        """按 F-44 逻辑填写画面并过账"""
        s = self.session

        # 1. 填主画面
        self.set_text_safely("wnd[0]/usr/ctxtRF05A-AGKON", vendor_code)
        self.set_text_safely("wnd[0]/usr/ctxtBKPF-BUDAT", budat)
        self.set_text_safely("wnd[0]/usr/txtBKPF-MONAT", monat)
        self.set_text_safely("wnd[0]/usr/ctxtBKPF-BUKRS", bukrs)
        self.set_text_safely("wnd[0]/usr/ctxtBKPF-WAERS", waers)
        self.set_text_safely("wnd[0]/usr/ctxtRF05A-AGUMS", agums)

        # 2. 回车，等待 SAP 校验完成
        s.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        self.ensure_no_error_popup(prefix="回车后")

        # 3. 进入抬头/附加数据画面
        try:
            s.findById("wnd[0]/mbar/menu[0]/menu[1]").select()
        except Exception as e:
            raise RuntimeError(
                f"无法进入抬头画面，请重新录制确认菜单路径 wnd[0]/mbar/menu[0]/menu[1] 是否正确。原始错误: {e}"
            )

        time.sleep(1.0)
        self.ensure_no_error_popup(prefix="进入抬头画面后")

        # 4. 写抬头字段
        self.set_text_safely("wnd[0]/usr/txtBKPF-XBLNR", xblnr, timeout=10)
        self.set_text_safely("wnd[0]/usr/txtBKPF-BKTXT", header_text, timeout=10)

        # 5. 保存/过账
        try:
            s.findById("wnd[0]/tbar[0]/btn[11]").press()
        except Exception as e:
            raise RuntimeError(f"点击保存/过账失败: {e}")

        time.sleep(1.0)

    def process(self):
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel 文件不存在: {self.excel_path}")

        self.connect_sap()

        with xw.App(visible=False, add_book=False) as app:
            wb = app.books.open(self.excel_path)

            try:
                ws = wb.sheets[self.sheet_name] if self.sheet_name else wb.sheets[0]

                # 自动识别最后一行
                last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row
                if last_row < 2:
                    raise ValueError("Excel 中没有可处理的数据。")

                # 写结果表头
                ws.range("H1").value = "result"
                ws.range("I1").value = "message"
                ws.range("J1").value = "processed_at"

                today = datetime.today()
                default_budat = today.strftime("%d.%m.%Y")
                default_monat = today.strftime("%m")   # 两位期间更稳
                default_bukrs = "3401"
                default_waers = "RMB"
                default_agums = "LKJM"
                default_header_text = "清账"

                for row in range(2, last_row + 1):
                    vendor_code = self.safe_text(ws.range(f"A{row}").value)

                    if not vendor_code:
                        ws.range(f"H{row}").value = "跳过"
                        ws.range(f"I{row}").value = "供应商代码为空"
                        ws.range(f"J{row}").value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        continue

                    budat = self.safe_text(ws.range(f"B{row}").value, default_budat)
                    monat = self.safe_text(ws.range(f"C{row}").value, default_monat)
                    bukrs = self.safe_text(ws.range(f"D{row}").value, default_bukrs)
                    waers = self.safe_text(ws.range(f"E{row}").value, default_waers)
                    agums = self.safe_text(ws.range(f"F{row}").value, default_agums)
                    header_text = self.safe_text(ws.range(f"G{row}").value, default_header_text)
                    xblnr = header_text

                    try:
                        self.start_f44()

                        self.fill_header_and_post(
                            vendor_code=vendor_code,
                            budat=budat,
                            monat=monat,
                            bukrs=bukrs,
                            waers=waers,
                            agums=agums,
                            header_text=header_text,
                            xblnr=xblnr,
                        )

                        popup_msg = self.close_popup_if_any()
                        status_msg = self.read_status_bar()

                        final_msg = popup_msg or status_msg or "执行完成"
                        ws.range(f"H{row}").value = "成功"
                        ws.range(f"I{row}").value = final_msg
                        ws.range(f"J{row}").value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    except Exception as e:
                        # 失败时尽量把 SAP 当前状态也带出来，便于排查
                        status_msg = self.read_status_bar()
                        msg = str(e)
                        if status_msg:
                            msg = f"{msg} | 状态栏: {status_msg}"

                        ws.range(f"H{row}").value = "失败"
                        ws.range(f"I{row}").value = msg
                        ws.range(f"J{row}").value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                wb.save()

            finally:
                wb.close()


if __name__ == "__main__":
    excel_file = r"C:\Users\3006699\Desktop\供应商清账数据.xlsx"  # 改成你的实际路径
    runner = SAPF44Auto(excel_file)
    runner.process()
    print("处理完成。")
