from paddleocr import PaddleOCR
import layoutparser as lp


class FormRecognizer:
    def __init__(self):
        # 初始化OCR引擎（启用表格识别）
        self.ocr_engine = PaddleOCR(
            use_angle_cls=True,
            lang="ch",
            rec_model_dir='./models/ch_ppocr_server_v2.0_rec_infer',
            det_model_dir='./models/ch_ppocr_server_v2.0_det_infer',
            table_model_dir='./models/en_ppocr_mobile_v2.0_table_structure_infer',
            use_gpu=False  # 根据实际情况调整
        )

        # 初始化版面分析模型
        self.layout_model = lp.Detectron2LayoutModel(
            'lp://PubLayNet/faster_rcnn_R_50_FPN_3x/config',
            extra_config=["MODEL.ROI_HEADS.SCORE_THRESH_TEST", 0.8],
            label_map={0: "Text", 1: "Title", 2: "List", 3: "Table", 4: "Figure"}
        )

    def analyze_form(self, image_path):
        # 1. 版面分析
        layout = self.layout_model.detect(image_path)

        # 2. 按区域识别
        results = {}
        for block in layout:
            if block.type == 'Table':
                # 表格特殊处理
                table_result = self.ocr_engine.ocr(
                    block.block.crop_image(image_path),
                    cls=True,
                    rec=True,
                    det=True,
                    table=True
                )
                results[f"table_{block.id}"] = self._parse_table(table_result)
            else:
                text_result = self.ocr_engine.ocr(
                    block.block.crop_image(image_path),
                    cls=True,
                    rec=True,
                    det=True
                )
                results[f"text_{block.id}"] = self._parse_text(text_result)

        return results

    def _parse_table(self, table_result):
        """处理PaddleOCR的表格识别结果"""
        # 示例：将识别结果转为二维数组
        structure = table_result['structure']
        cells = []
        for row in structure:
            row_data = []
            for cell in row:
                text = " ".join([line[-1][0] for line in cell['text']])
                row_data.append(text)
            cells.append(row_data)
        return cells

    # 基于正则表达式的关键字段提取
    import re

    field_rules = {
        'id_card': r'[1-9]\d{5}(18|19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\d{3}[\dXx]',
        'phone': r'1[3-9]\d{9}',
        'date': r'\d{4}-\d{1,2}-\d{1,2}'
    }

    def extract_fields(text):
        results = {}
        for field, pattern in field_rules.items():
            match = re.search(pattern, text)
            if match:
                results[field] = match.group()
        return results