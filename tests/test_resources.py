# -*- coding: utf-8 -*-
"""客户端资源完整性测试。"""

import unittest
from pathlib import Path
from xml.etree import ElementTree as ET

from stockwidget.constants import CODE_LIST_FILES


ROOT = Path(__file__).parents[1]


class ResourceTests(unittest.TestCase):
    def test_us_alias_cache_is_not_a_client_download(self):
        self.assertNotIn("cache_us_cn_aliases.json", CODE_LIST_FILES)

    def test_packaged_resources_exist_and_preserve_virtual_names(self):
        root = ET.parse(ROOT / "resources" / "resources.qrc").getroot()
        aliases = {}
        for node in root.iter("file"):
            path = ROOT / "resources" / node.text
            self.assertTrue(path.is_file(), path)
            self.assertIn(path.parent.name, ("icons", "data"))
            self.assertEqual(node.attrib["alias"], path.name)
            aliases[node.attrib["alias"]] = path
        self.assertTrue(set(CODE_LIST_FILES).issubset(aliases))
        self.assertIn("StockWidget.ico", aliases)


if __name__ == "__main__":
    unittest.main()
