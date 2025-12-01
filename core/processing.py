import logging
import re


class ExcelProcessor:
    """轻量业务处理服务。

    通过依赖注入接入具体的处理实现，避免跨模块耦合。
    """

    def __init__(self, process_impl):
        self._process_impl = process_impl

    @staticmethod
    def validate_mappings(mappings):
        """验证产品线-对接人映射列表"""
        allowed = re.compile(r"^[\u4e00-\u9fa5A-Za-z0-9 _\-/()]+$")
        seen = set()
        for product, contact in mappings:
            if not product or not contact:
                return False, "产品线和对接人均不能为空"
            if not allowed.match(product) or not allowed.match(contact):
                return False, "存在非法字符，请仅使用中英文、数字和常用符号"
            key = product.strip().casefold()
            if key in seen:
                return False, f"重复的产品线: {product}"
            seen.add(key)
        return True, "OK"

    def process(self, input_file, output_file, start_dt, end_dt, product_contact_list, replace_mode='overwrite', progress_callback=None, cancel_event=None):
        logging.info(
            "开始处理: input=%s, output=%s, range=%s-%s, mappings=%s, mode=%s",
            input_file,
            output_file,
            start_dt,
            end_dt,
            product_contact_list,
            replace_mode,
        )
        return self._process_impl(
            input_file,
            output_file,
            start_dt,
            end_dt,
            product_contact_list=product_contact_list,
            replace_mode=replace_mode,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
        )

