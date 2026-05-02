import win32com.client
import time
import pandas as pd
import logging
import os
from datetime import datetime
import traceback
import re
import xlwings as xw


class SAPF65Automation:
    def __init__(self):
        self.application = None
        self.connection = None
        self.session = None
        self.excel_app = None
        self.workbook = None
        self.worksheet = None
        self.setup_logging()

    def setup_logging(self):
        log_dir = r"C:\SAP_Automation\Logs"
        os.makedirs(log_dir, exist_ok=True)

        log_file = os.path.join(
            log_dir,
            f"sap_f65_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        )

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)

        if logger.handlers:
            logger.handlers.clear()

        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        )

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(
            logging.Formatter("%(levelname)s - %(message)s")
        )

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        logging.info("=== SAP F-65 自动化脚本启动 ===")

    def normalize_value(self, value):
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def format_date(self, value):
        if pd.isna(value):
            return ""
        if isinstance(value, pd.Timestamp):
            return value.strftime("%d.%m.%Y")

        text = str(value).strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(text, fmt)
                return dt.strftime("%d.%m.%Y")
            except Exception:
                pass
        return text

    def format_amount(self, value):
        if pd.isna(value) or str(value).strip() == "":
            return ""
        try:
            num = float(value)
            return f"{num:.2f}"
        except Exception:
            return str(value).strip()

    def wait_ready(self, seconds=0.5):
        time.sleep(seconds)

    def connect_sap(self):
        try:
            logging.info("🔌 尝试连接 SAP GUI...")
            sap_gui = win32com.client.GetObject("SAPGUI")
            self.application = sap_gui.GetScriptingEngine

            if self.application.Children.Count == 0:
                logging.error("❌ 没有检测到已登录的 SAP 连接")
                return False

            self.connection = self.application.Children(0)

            if self.connection.Children.Count == 0:
                logging.error("❌ 当前 SAP 连接下没有可用 Session")
                return False

            self.session = self.connection.Children(0)

            try:
                self.session.findById("wnd[0]").maximize()
            except Exception:
                pass

            logging.info("✅ 成功连接到 SAP")
            return True

        except Exception as e:
            logging.error(f"❌ 连接 SAP 失败: {str(e)}")
            logging.error(traceback.format_exc())
            return False

    def check_element_exists(self, element_id, max_attempts=3, delay=1):
        for attempt in range(max_attempts):
            try:
                element = self.session.findById(element_id, False)
                if element is not None:
                    return element
            except Exception:
                pass

            if attempt < max_attempts - 1:
                time.sleep(delay)

        return None

    def safe_set_text(self, element_id, value, max_attempts=3, delay=1):
        target = self.normalize_value(value)
        logging.info(f"📝 设置文本: {element_id} = {target}")

        for attempt in range(max_attempts):
            try:
                element = self.check_element_exists(element_id, 1, 0)
                if element is None:
                    raise Exception("元素不存在")

                element.text = target
                self.wait_ready(0.2)

                actual = ""
                try:
                    actual = str(element.text).strip()
                except Exception:
                    actual = target

                if actual == target or actual.replace(",", "") == target.replace(",", ""):
                    logging.info(f"✅ 设置成功: {element_id}")
                    return True

            except Exception as e:
                logging.warning(
                    f"⚠️ 设置文本失败，第 {attempt + 1} 次: {element_id}, {str(e)}"
                )

            if attempt < max_attempts - 1:
                time.sleep(delay)

        logging.error(f"❌ 最终设置失败: {element_id} = {target}")
        return False

    def safe_set_checkbox(self, element_id, checked=True, max_attempts=3, delay=1):
        logging.info(f"☑️ 设置复选框: {element_id} = {checked}")

        for attempt in range(max_attempts):
            try:
                element = self.check_element_exists(element_id, 1, 0)
                if element is None:
                    raise Exception("复选框不存在")

                element.selected = bool(checked)
                self.wait_ready(0.2)

                if bool(element.selected) == bool(checked):
                    logging.info(f"✅ 复选框设置成功: {element_id}")
                    return True

            except Exception as e:
                logging.warning(
                    f"⚠️ 设置复选框失败，第 {attempt + 1} 次: {element_id}, {str(e)}"
                )

            if attempt < max_attempts - 1:
                time.sleep(delay)

        logging.error(f"❌ 最终设置复选框失败: {element_id}")
        return False

    def safe_press_button(self, element_id, max_attempts=3, delay=1):
        logging.info(f"🖱️ 点击按钮: {element_id}")

        for attempt in range(max_attempts):
            try:
                element = self.check_element_exists(element_id, 1, 0)
                if element is None:
                    raise Exception("按钮不存在")

                element.press()
                self.wait_ready(0.5)
                logging.info(f"✅ 按钮点击成功: {element_id}")
                return True

            except Exception as e:
                logging.warning(
                    f"⚠️ 点击按钮失败，第 {attempt + 1} 次: {element_id}, {str(e)}"
                )

            if attempt < max_attempts - 1:
                time.sleep(delay)

        logging.error(f"❌ 最终按钮点击失败: {element_id}")
        return False

    def send_enter(self):
        try:
            self.session.findById("wnd[0]").sendVKey(0)
            self.wait_ready(1)
            return True
        except Exception as e:
            logging.error(f"❌ 发送回车失败: {str(e)}")
            return False

    def handle_sap_errors(self):
        try:
            try:
                if self.session.Children.Count > 1:
                    for i in range(1, self.session.Children.Count):
                        try:
                            wnd = self.session.findById(f"wnd[{i}]", False)
                            if wnd is None:
                                continue

                            title = ""
                            msg = ""

                            try:
                                title = wnd.findById("titl").text
                            except Exception:
                                pass

                            for pid in ["usr/lbl50", "usr/txtMESSTXT", "usr/lbl[1,1]"]:
                                try:
                                    obj = wnd.findById(pid, False)
                                    if obj:
                                        msg = obj.text
                                        break
                                except Exception:
                                    pass

                            logging.error(f"🔴 SAP 弹窗错误，标题: {title}，消息: {msg}")

                            try:
                                wnd.findById("tbar[0]/btn[0]").press()
                            except Exception:
                                pass

                            return True, msg or title or "SAP 弹窗错误"
                        except Exception:
                            continue
            except Exception:
                pass

            try:
                sbar = self.session.findById("wnd[0]/sbar")
                msg_type = getattr(sbar, "messageType", "")
                msg_text = sbar.text
                if msg_type in ("E", "A"):
                    logging.error(f"🔴 SAP 状态栏错误: {msg_text}")
                    return True, msg_text
            except Exception:
                pass

            return False, None

        except Exception as e:
            logging.error(f"❌ handle_sap_errors 异常: {str(e)}")
            return False, None

    def handle_rate_warning_by_enter(self):
        """本币金额输入后，若弹出汇率相关提示，直接回车继续"""
        try:
            logging.info("↩️ 尝试处理汇率提示：直接回车继续")
            self.send_enter()

            error_exists, error_msg = self.handle_sap_errors()
            if error_exists:
                logging.warning(f"⚠️ 回车后检测到消息: {error_msg}")

            return True
        except Exception as e:
            logging.warning(f"⚠️ 处理汇率提示时异常: {str(e)}")
            return True

    def extract_doc_number(self, status_text):
        try:
            numbers = re.findall(r"\b\d{8,12}\b", str(status_text))
            return numbers[0] if numbers else ""
        except Exception:
            return ""

    def enter_transaction(self, transaction_code):
        try:
            logging.info(f"🚪 进入事务码: {transaction_code}")
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.text = transaction_code
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)

            error_exists, error_msg = self.handle_sap_errors()
            if error_exists:
                logging.error(f"❌ 进入事务码失败: {error_msg}")
                return False

            return True

        except Exception as e:
            logging.error(f"❌ 进入事务码失败: {str(e)}")
            logging.error(traceback.format_exc())
            return False

    def validate_row(self, row):
        required_fields = [
            "凭证日期",
            "公司代码",
            "货币",
            "借方科目",
            "借方金额",
            "借方本币金额",
            "贷方科目",
            "贷方金额",
            "贷方本币金额",
        ]

        missing = []
        for col in required_fields:
            if col not in row or self.normalize_value(row[col]) == "":
                missing.append(col)

        if missing:
            return False, f"必填字段缺失: {', '.join(missing)}"

        try:
            debit = float(row["借方金额"])
            credit = float(row["贷方金额"])
            if round(debit, 2) != round(credit, 2):
                return False, "借方金额与贷方金额不一致"
        except Exception:
            return False, "外币金额格式错误"

        return True, ""

    def finish_bank_line(
            self,
            posting_key,
            account_code,
            foreign_amount,
            local_amount,
            text_value,
            prctr,
            reason_code
    ):
        try:
            # 1. 科目
            self.safe_set_text("wnd[0]/usr/ctxtRF05V-NEWBS", posting_key)
            self.safe_set_text("wnd[0]/usr/ctxtRF05V-NEWKO", account_code)
            self.send_enter()

            # 2. 金额
            self.safe_set_text("wnd[0]/usr/txtBSEG-WRBTR", self.format_amount(foreign_amount))
            self.safe_set_text("wnd[0]/usr/txtBSEG-DMBTR", self.format_amount(local_amount))

            self.wait_ready(0.5)

            # 3. 文本 + 利润中心
            if text_value:
                self.safe_set_text("wnd[0]/usr/ctxtBSEG-SGTXT", text_value)

            if prctr:
                self.safe_set_text(
                    "wnd[0]/usr/subBLOCK:SAPLKACB:1016/ctxtCOBL-PRCTR",
                    prctr
                )

            # 4. 行确认
            self.safe_press_button("wnd[0]/tbar[1]/btn[7]")

            # 5. 回车进入原因代码
            self.send_enter()

            # 6. 原因代码
            self.safe_set_text("wnd[0]/usr/ctxtBSEG-RSTGR", reason_code)

            # ❗这里绝对不能回车
            return True, ""

        except Exception as e:
            return False, str(e)

    def process_single_document(self, row):
        result = {"success": False, "doc_number": "", "error_message": ""}

        try:
            logging.info("=" * 80)
            logging.info(f"📄 开始处理行: {dict(row)}")

            valid, msg = self.validate_row(row)
            if not valid:
                result["error_message"] = msg
                return result

            if not self.enter_transaction("/nF-65"):
                result["error_message"] = "无法进入 F-65"
                return result

            self.wait_ready(1)

            # 抬头
            self.safe_set_checkbox("wnd[0]/usr/chkVBKPF-XBWAE", False)

            doc_date = row["凭证日期"]
            posting_date = (
                row["过账日期"]
                if "过账日期" in row and self.normalize_value(row["过账日期"]) != ""
                else row["凭证日期"]
            )

            if not self.safe_set_text("wnd[0]/usr/ctxtBKPF-BLDAT", self.format_date(doc_date)):
                result["error_message"] = "无法设置凭证日期"
                return result

            if not self.safe_set_text("wnd[0]/usr/ctxtBKPF-BUDAT", self.format_date(posting_date)):
                result["error_message"] = "无法设置过账日期"
                return result

            if not self.safe_set_text("wnd[0]/usr/ctxtBKPF-BUKRS", row["公司代码"]):
                result["error_message"] = "无法设置公司代码"
                return result

            if "过账期间" in row and self.normalize_value(row["过账期间"]) != "":
                self.safe_set_text("wnd[0]/usr/txtBKPF-MONAT", row["过账期间"])

            if not self.safe_set_text("wnd[0]/usr/ctxtBKPF-WAERS", row["货币"]):
                result["error_message"] = "无法设置货币"
                return result

            if "凭证文本" in row and self.normalize_value(row["凭证文本"]) != "":
                self.safe_set_text("wnd[0]/usr/txtBKPF-BKTXT", row["凭证文本"])

            reason_code = (
                self.normalize_value(row["原因代码"])
                if "原因代码" in row and self.normalize_value(row["原因代码"]) != ""
                else "101"
            )

            text_value = (
                self.normalize_value(row["凭证文本"])
                if "凭证文本" in row and self.normalize_value(row["凭证文本"]) != ""
                else ""
            )

            prctr = (
                self.normalize_value(row["利润中心"])
                if "利润中心" in row and self.normalize_value(row["利润中心"]) != ""
                else ""
            )

            # 借方银行行
            ok, msg = self.finish_bank_line(
                posting_key="40",
                account_code=self.normalize_value(row["借方科目"]),
                foreign_amount=row["借方金额"],
                local_amount=row["借方本币金额"],
                text_value=text_value,
                prctr=prctr,
                reason_code=reason_code
            )
            if not ok:
                result["error_message"] = f"借方处理失败: {msg}"
                return result

            error_exists, error_msg = self.handle_sap_errors()
            if error_exists:
                result["error_message"] = f"借方行错误: {error_msg}"
                return result

            # 贷方银行行
            ok, msg = self.finish_bank_line(
                posting_key="50",
                account_code=self.normalize_value(row["贷方科目"]),
                foreign_amount=row["贷方金额"],
                local_amount=row["贷方本币金额"],
                text_value=text_value,
                prctr=prctr,
                reason_code=reason_code
            )
            if not ok:
                result["error_message"] = f"贷方处理失败: {msg}"
                return result

            error_exists, error_msg = self.handle_sap_errors()
            if error_exists:
                result["error_message"] = f"贷方行错误: {error_msg}"
                return result

            # 保存
            if not self.safe_press_button("wnd[0]/tbar[0]/btn[11]"):
                result["error_message"] = "无法点击保存"
                return result

            self.wait_ready(2)

            error_exists, error_msg = self.handle_sap_errors()
            if error_exists:
                result["error_message"] = f"保存失败: {error_msg}"
                return result

            status_text = ""
            try:
                status_text = self.session.findById("wnd[0]/sbar").text
            except Exception:
                pass

            logging.info(f"📋 保存后状态栏: {status_text}")

            if status_text and (
                "凭证" in status_text or
                "document" in status_text.lower() or
                "beleg" in status_text.lower()
            ):
                result["success"] = True
                result["doc_number"] = self.extract_doc_number(status_text)
                return result

            result["error_message"] = f"保存后未识别成功信息: {status_text}"
            return result

        except Exception as e:
            msg = f"处理凭证异常: {str(e)}"
            logging.error(msg)
            logging.error(traceback.format_exc())
            result["error_message"] = msg
            return result

    def open_excel_by_authorized_office(self, excel_path, sheet_name=None):
        try:
            logging.info(f"📂 使用 Excel 程序打开文件: {excel_path}")

            self.excel_app = xw.App(visible=False, add_book=False)
            self.excel_app.display_alerts = False
            self.excel_app.screen_updating = False

            self.workbook = self.excel_app.books.open(excel_path)

            if sheet_name:
                self.worksheet = self.workbook.sheets[sheet_name]
            else:
                self.worksheet = self.workbook.sheets[0]

            logging.info(f"✅ Excel 打开成功，工作表: {self.worksheet.name}")
            return True

        except Exception as e:
            logging.error(f"❌ Excel 打开失败: {str(e)}")
            logging.error(traceback.format_exc())
            return False

    def read_sheet_to_dataframe(self):
        used_range = self.worksheet.used_range.value
        if not used_range:
            raise Exception("工作表为空")
        if len(used_range) < 2:
            raise Exception("工作表只有表头或无有效数据")

        headers = used_range[0]
        data = used_range[1:]
        df = pd.DataFrame(data, columns=headers)
        logging.info(f"✅ 读取 Excel 数据成功，共 {len(df)} 行")
        return df

    def write_dataframe_to_sheet(self, df):
        self.worksheet.clear_contents()
        self.worksheet.range("A1").value = [df.columns.tolist()] + df.values.tolist()
        logging.info("✅ 数据已回写到工作表")

    def save_excel(self):
        self.workbook.save()
        logging.info("✅ Excel 已保存")

    def process_excel_data(self, excel_path, sheet_name=None):
        try:
            if not os.path.exists(excel_path):
                raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

            if not self.open_excel_by_authorized_office(excel_path, sheet_name):
                raise Exception("无法通过本机 Excel 打开文件，请确认电脑/账号有权限。")

            df = self.read_sheet_to_dataframe()

            for col in ["处理结果", "SAP凭证号", "错误信息", "处理时间"]:
                if col not in df.columns:
                    df[col] = ""

            success_count = 0
            fail_count = 0

            for idx, row in df.iterrows():
                logging.info(f"🚀 开始处理第 {idx + 1} 行")

                if self.normalize_value(df.at[idx, "处理结果"]) == "成功":
                    logging.info(f"⏭️ 第 {idx + 1} 行已成功，跳过")
                    continue

                result = self.process_single_document(row)

                if result["success"]:
                    df.at[idx, "处理结果"] = "成功"
                    df.at[idx, "SAP凭证号"] = result["doc_number"]
                    df.at[idx, "错误信息"] = ""
                    success_count += 1
                else:
                    df.at[idx, "处理结果"] = "失败"
                    df.at[idx, "SAP凭证号"] = ""
                    df.at[idx, "错误信息"] = result["error_message"]
                    fail_count += 1

                df.at[idx, "处理时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.write_dataframe_to_sheet(df)
                self.save_excel()

            return success_count, fail_count, df

        except Exception as e:
            logging.error(f"❌ 处理 Excel 失败: {str(e)}")
            logging.error(traceback.format_exc())
            raise

    def close_connection(self, close_sap=False):
        try:
            if self.workbook:
                try:
                    self.workbook.close()
                except Exception:
                    pass
                self.workbook = None

            if self.excel_app:
                try:
                    self.excel_app.quit()
                except Exception:
                    pass
                self.excel_app = None

            if close_sap and self.session:
                try:
                    self.session.findById("wnd[0]").close()
                    time.sleep(0.5)
                    try:
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    except Exception:
                        pass
                except Exception:
                    pass

            logging.info("✅ 资源清理完成")
        except Exception as e:
            logging.error(f"❌ close_connection 异常: {str(e)}")


if __name__ == "__main__":
    EXCEL_PATH = r"C:\Users\3006699\PycharmProjects\PythonProject\自动化入账\F65_Data.xlsx"
    SHEET_NAME = None

    print("=" * 90)
    print("🚀 SAP F-65 银行外币入账自动化脚本（借贷统一版）")
    print("=" * 90)

    automation = SAPF65Automation()

    try:
        if not automation.connect_sap():
            print("❌ SAP 连接失败")
            input("按 Enter 退出...")
            raise SystemExit

        print("✅ SAP 连接成功")

        if not automation.enter_transaction("/nF-65"):
            print("❌ 无法进入 F-65")
            input("按 Enter 退出...")
            raise SystemExit

        print("✅ F-65 测试成功")
        input("确认当前界面正常后，按 Enter 开始处理...")

        success_count, fail_count, _ = automation.process_excel_data(
            EXCEL_PATH,
            SHEET_NAME
        )

        print("=" * 90)
        print("🎉 批量处理完成")
        print(f"成功: {success_count}")
        print(f"失败: {fail_count}")
        print(f"结果已写回: {EXCEL_PATH}")
        print("=" * 90)

    except Exception as e:
        print(f"❌ 程序运行失败: {str(e)}")
    finally:
        automation.close_connection(close_sap=False)
        input("按 Enter 键退出程序...")
