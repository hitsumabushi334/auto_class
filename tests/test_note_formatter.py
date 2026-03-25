import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).parent.parent))

from note_formatter import render_markdown_note


class TestRenderMarkdownNote(unittest.TestCase):
    def test_render_markdown_note_uses_template_sections(self):
        summary_data = {
            "title": "サンプル講義",
            "summary": "講義全体の要約です。",
            "topics": [
                {
                    "topic_title": "導入",
                    "topic_time_range": {
                        "start_time": "00:00",
                        "end_time": "05:00",
                    },
                    "topic_keywords": ["キーワードA", "キーワードB"],
                    "topic_summary": "導入パートの要約です。",
                    "topic_points": ["ポイント1", "ポイント2"],
                    "important_knowledge": [
                        {
                            "statement": "重要な結論",
                            "explanation": "重要な理由の説明",
                            "source": "lecture_explicit",
                            "importance": 3,
                            "evidence": {
                                "start_time": "00:30",
                                "end_time": "01:00",
                            },
                        }
                    ],
                    "technical_term": [
                        {
                            "word": "専門用語A",
                            "explanation": "用語の説明",
                            "lecture_definition": "講義での定義文",
                            "pitfall": "よくある誤解",
                            "evidence": [
                                {
                                    "start_time": "00:40",
                                    "end_time": "00:50",
                                }
                            ],
                        }
                    ],
                }
            ],
            "assignments": [
                {
                    "task_title": "レポート提出",
                    "task_detail": "第1章の内容をまとめる",
                    "due_date": "2026-04-01",
                    "evidence": [
                        {
                            "start_time": "04:40",
                            "end_time": "04:55",
                        }
                    ],
                }
            ],
        }

        actual = render_markdown_note(summary_data)

        expected = """# サンプル講義

## 全体要約

講義全体の要約です。

## トピック詳細

### 導入 (00:00 〜 05:00)

キーワード

- キーワードA
- キーワードB

要約

導入パートの要約です。

ポイント

- ポイント1
- ポイント2

重要な知見

- 重要な結論
  詳細: 重要な理由の説明
  根拠: 講義で明示
  重要度: ★★★
  タイムスタンプ: 00:30 〜 01:00

専門用語

- 専門用語A: 用語の説明
  講義での定義: 講義での定義文
  誤解・混同ポイント: よくある誤解
  タイムスタンプ: 00:40 〜 00:50

## 課題、提出物

- レポート提出: 第1章の内容をまとめる
  提出期限: 2026-04-01
  タイムスタンプ: 04:40 〜 04:55
"""

        self.assertEqual(actual, expected)

    def test_render_markdown_note_omits_optional_blocks_when_data_is_missing(self):
        summary_data = {
            "title": "最小講義",
            "summary": "要約",
            "topics": [
                {
                    "topic_title": "基本事項",
                    "topic_keywords": [],
                    "topic_summary": "概要",
                    "topic_points": [],
                    "important_knowledge": [
                        {
                            "statement": "覚えること",
                            "explanation": "理由",
                            "source": "custom_source",
                            "evidence": {
                                "start_time": "01:00",
                                "end_time": "01:20",
                            },
                        }
                    ],
                    "technical_term": [
                        {
                            "word": "用語",
                            "explanation": "説明",
                        }
                    ],
                }
            ],
            "assignments": [],
        }

        actual = render_markdown_note(summary_data)

        self.assertIn("### 基本事項", actual)
        self.assertNotIn("### 基本事項 (", actual)
        self.assertNotIn("## 課題、提出物", actual)
        self.assertNotIn("重要度:", actual)
        self.assertIn("根拠: custom_source", actual)
        self.assertNotIn("講義での定義:", actual)
        self.assertNotIn("誤解・混同ポイント:", actual)
        self.assertNotIn("タイムスタンプ:  \n", actual)


if __name__ == "__main__":
    unittest.main()
