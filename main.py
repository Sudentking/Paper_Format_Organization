import sys
import os
import yaml
from docx import Document
from formatters import (
    HeadingFormatter,
    ParagraphFormatter,
    TableFormatter,
    CaptionFormatter,
    ReferenceFormatter,
    TOCFormatter,
)


class DocumentFormatter:
    def __init__(self, input_path, config_path='config/format_config.yaml'):
        self.input_path = input_path
        self.config_path = config_path
        self.doc = None
        self.config = None
        self.formatters = []
        self.errors = []

        self._load_document()
        self._load_config()
        self._init_formatters()

    def _load_document(self):
        try:
            self.doc = Document(self.input_path)
            print(f"[OK] 成功加载文档: {self.input_path}")
        except FileNotFoundError:
            print(f"[ERROR] 文件不存在: {self.input_path}")
            sys.exit(1)
        except Exception as e:
            print(f"[ERROR] 加载文档失败: {e}")
            sys.exit(1)

    def _load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f)
            print(f"[OK] 成功加载配置: {self.config_path}")
        except FileNotFoundError:
            print(f"[ERROR] 配置文件不存在: {self.config_path}")
            sys.exit(1)
        except yaml.YAMLError as e:
            print(f"[ERROR] 配置文件格式错误: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"[ERROR] 加载配置失败: {e}")
            sys.exit(1)

    def _init_formatters(self):
        self.formatters = [
            TOCFormatter(self.doc, self.config),
            HeadingFormatter(self.doc, self.config),
            ParagraphFormatter(self.doc, self.config),
            TableFormatter(self.doc, self.config),
            CaptionFormatter(self.doc, self.config),
            ReferenceFormatter(self.doc, self.config),
        ]

    def format_all(self):
        print("\n开始格式化文档...")
        success_count = 0

        for formatter in self.formatters:
            name = formatter.__class__.__name__
            try:
                print(f"  -> {name}...", end=" ")
                formatter.format()
                print("完成")
                success_count += 1
            except Exception as e:
                print(f"失败: {e}")
                self.errors.append((name, str(e)))
                import traceback
                traceback.print_exc()

        print(f"\n格式化完成: {success_count}/{len(self.formatters)} 个模块成功")

        if self.errors:
            print("\n失败的模块:")
            for name, err in self.errors:
                print(f"  - {name}: {err}")

    def save(self, output_path=None):
        if output_path is None:
            base, ext = os.path.splitext(self.input_path)
            output_path = f"{base}_formatted{ext}"

        try:
            self.doc.save(output_path)
            print(f"[OK] 文档已保存: {output_path}")
            return output_path
        except PermissionError:
            print(f"[ERROR] 保存失败，文件可能被占用: {output_path}")
            sys.exit(1)
        except Exception as e:
            print(f"[ERROR] 保存文档失败: {e}")
            sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("使用方法: python main.py <输入文档路径> [输出文档路径] [配置文件路径]")
        print("\n示例:")
        print("  python main.py input.docx")
        print("  python main.py input.docx output.docx")
        print("  python main.py input.docx output.docx config/custom_config.yaml")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    config_path = sys.argv[3] if len(sys.argv) > 3 else 'config/format_config.yaml'

    if not os.path.exists(input_path):
        print(f"[ERROR] 文件不存在: {input_path}")
        sys.exit(1)

    formatter = DocumentFormatter(input_path, config_path)
    formatter.format_all()
    formatter.save(output_path)


if __name__ == '__main__':
    main()
