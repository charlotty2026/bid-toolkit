#!/usr/bin/env python3
"""
RFP工具链自动化测试套件 v1.0
覆盖: rfp_structure / rfp_generator / rfp_compliance
"""
import json
import os
import subprocess
import sys
import tempfile
import unittest

RFP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class TestRFPStructure(unittest.TestCase):
    """测试 rfp_structure.py"""

    def test_import_and_basic(self):
        sys.path.insert(0, RFP_DIR)
        import rfp_structure
        self.assertIsNotNone(rfp_structure.STANDARD_CHAPTERS)

    def test_standard_chapters_count(self):
        sys.path.insert(0, RFP_DIR)
        from rfp_structure import get_chapters
        chapters = get_chapters(project_type="services")
        self.assertGreaterEqual(len(chapters), 7, "标准章应至少7章")

    def test_conditional_sections(self):
        sys.path.insert(0, RFP_DIR)
        from rfp_structure import get_chapters
        goods = get_chapters(project_type="goods")
        services = get_chapters(project_type="services")
        engineering = get_chapters(project_type="engineering")
        # goods和services应有质疑投诉
        self.assertTrue(any("质疑" in str(c) for c in goods))
        self.assertTrue(any("质疑" in str(c) for c in services))

    def test_compliance_checklist(self):
        sys.path.insert(0, RFP_DIR)
        from rfp_structure import get_compliance_checklist
        checklist = get_compliance_checklist()
        self.assertGreater(len(checklist), 0, "合规检查清单不应为空")

    def test_legal_basis(self):
        sys.path.insert(0, RFP_DIR)
        from rfp_structure import LEGAL_BASIS
        legal_str = str(LEGAL_BASIS)
        self.assertTrue("87" in legal_str or "政府采购" in legal_str, "法律依据应包含政府采购法/87号令")


class TestRFPGenerator(unittest.TestCase):
    """测试 rfp_generator.py"""

    def setUp(self):
        self.output_dir = tempfile.mkdtemp()

    def test_generate_services_markdown(self):
        output_path = os.path.join(self.output_dir, "test_services.md")
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_generator.py"),
             "--type", "services", "--project", "测试服务项目",
             "--budget", "500000", "-o", output_path],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0, f"生成失败: {result.stderr}")
        self.assertTrue(os.path.exists(output_path), "输出文件应存在")
        content = open(output_path, encoding="utf-8").read()
        self.assertGreater(len(content), 1000, "生成内容应足够长")
        self.assertIn("投标邀请", content, "应包含投标邀请章")

    def test_generate_goods_markdown(self):
        output_path = os.path.join(self.output_dir, "test_goods.md")
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_generator.py"),
             "--type", "goods", "--project", "测试货物项目",
             "--budget", "800000", "-o", output_path],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0, f"生成失败: {result.stderr}")
        content = open(output_path, encoding="utf-8").read()
        self.assertIn("评标办法", content)

    def test_generate_engineering_markdown(self):
        output_path = os.path.join(self.output_dir, "test_engineering.md")
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_generator.py"),
             "--type", "engineering", "--project", "测试工程项目",
             "--budget", "2000000", "-o", output_path],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0, f"生成失败: {result.stderr}")

    def test_generate_with_config(self):
        config = {
            "project_name": "配置文件测试项目",
            "project_id": "TEST-001",
            "purchaser": "测试采购单位",
            "budget": 300000,
            "fund_source": "财政资金"
        }
        config_path = os.path.join(self.output_dir, "test_config.json")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False)
        output_path = os.path.join(self.output_dir, "test_config_output.md")
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_generator.py"),
             "--type", "services", "--config", config_path, "-o", output_path],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0, f"配置文件生成失败: {result.stderr}")
        content = open(output_path, encoding="utf-8").read()
        self.assertIn("配置文件测试项目", content)

    def test_generate_docx(self):
        output_path = os.path.join(self.output_dir, "test_output.docx")
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_generator.py"),
             "--type", "services", "--project", "测试DOCX项目",
             "--budget", "500000", "-o", output_path, "--docx"],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        # docx可能因依赖问题失败，记warning即可
        if result.returncode == 0:
            self.assertTrue(os.path.exists(output_path), "DOCX文件应存在")
        else:
            self.skipTest("docx生成可能缺少依赖")


class TestRFPCompliance(unittest.TestCase):
    """测试 rfp_compliance.py"""

    @classmethod
    def setUpClass(cls):
        """先生成一份测试用RFP文件"""
        cls.temp_dir = tempfile.mkdtemp()
        cls.test_rfp_path = os.path.join(cls.temp_dir, "test_rfp.md")
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_generator.py"),
             "--type", "services", "--project", "合规测试项目",
             "--budget", "500000", "-o", cls.test_rfp_path],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        if result.returncode != 0:
            raise RuntimeError(f"测试RFP生成失败: {result.stderr}")

    def test_run_compliance_check(self):
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_compliance.py"),
             "--rfp", self.test_rfp_path, "--type", "services", "--format", "text"],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0, f"合规检查失败: {result.stderr}")
        output = result.stdout
        # 应有检查结果
        self.assertGreater(len(output), 100, "合规检查应有输出")

    def test_json_output(self):
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_compliance.py"),
             "--rfp", self.test_rfp_path, "--type", "services", "--format", "json"],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0, f"JSON输出失败: {result.stderr}")
        try:
            data = json.loads(result.stdout)
            self.assertIsInstance(data, dict)
        except json.JSONDecodeError:
            self.fail("JSON输出应为有效JSON")

    def test_detection_of_exclusionary_clause(self):
        """测试排他性条款检测"""
        test_content = """# 测试招标文件

本项目必须为华为品牌设备。

近三年业绩不低于500万元。
"""
        test_path = os.path.join(self.temp_dir, "test_exclusion.md")
        with open(test_path, "w", encoding="utf-8") as f:
            f.write(test_content)
        result = subprocess.run(
            [sys.executable, os.path.join(RFP_DIR, "rfp_compliance.py"),
             "--rfp", test_path, "--type", "goods", "--format", "text"],
            capture_output=True, text=True, cwd=RFP_DIR
        )
        self.assertEqual(result.returncode, 0)
        output = result.stdout
        # 应检测到排他性条款（品牌指定/业绩门槛）
        self.assertTrue("品牌" in output or "业绩" in output or "排他" in output,
                        f"应检测到排他性条款，输出: {output[:500]}")


class TestStandardStructureJSON(unittest.TestCase):
    """测试 standard_structure.json"""

    def test_json_valid(self):
        path = os.path.join(RFP_DIR, "standard_structure.json")
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        self.assertIsInstance(data, dict)
        self.assertEqual(data.get("version"), "2.0")

    def test_chapters_complete(self):
        path = os.path.join(RFP_DIR, "standard_structure.json")
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        chapters = data.get("standard_chapters", [])
        self.assertGreaterEqual(len(chapters), 7, "标准章应至少7章")
        names = [c["name"] for c in chapters]
        for required in ["投标邀请", "投标人须知", "资格审查", "采购需求", "评标办法", "合同条款", "投标文件格式"]:
            self.assertIn(required, names, f"缺少必备章节: {required}")

    def test_scoring_templates(self):
        path = os.path.join(RFP_DIR, "standard_structure.json")
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        scoring = data.get("scoring_templates", {})
        for ptype in ["goods", "services", "engineering"]:
            self.assertIn(ptype, scoring, f"缺少{ptype}评分模板")
            total = scoring[ptype].get("total", 0)
            self.assertEqual(total, 100, f"{ptype}评分总分应为100")


class TestComplianceRulesJSON(unittest.TestCase):
    """测试 compliance_rules/ 目录下的规则文件"""

    def setUp(self):
        self.rules_dir = os.path.join(RFP_DIR, "compliance_rules")

    def test_exclusionary_patterns(self):
        path = os.path.join(self.rules_dir, "exclusionary_patterns.json")
        self.assertTrue(os.path.exists(path), "exclusionary_patterns.json应存在")
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        self.assertIsInstance(data, list)
        self.assertGreater(len(data), 20, "排他性规则应至少20条")
        for item in data:
            self.assertIn("pattern", item)
            self.assertIn("reason", item)

    def test_rejection_keywords(self):
        path = os.path.join(self.rules_dir, "rejection_keywords.json")
        self.assertTrue(os.path.exists(path))
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        self.assertIsInstance(data, list)
        self.assertGreater(len(data), 10)

    def test_required_sections(self):
        path = os.path.join(self.rules_dir, "required_sections.json")
        self.assertTrue(os.path.exists(path))
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for ptype in ["goods", "services", "engineering"]:
            self.assertIn(ptype, data)


if __name__ == "__main__":
    unittest.main(verbosity=2)
