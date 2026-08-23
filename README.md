# 中考英语零基础核心词表（乱序版）

零基础友好的中考英语核心词汇表：乱序排列、高频在前，含音标、词性和中文释义。仓库发布词库数据和构建脚本；学习进度与 SRS 由 [`wordvault`](https://github.com/jim-688/wordvault) 负责。

## 数据规模

- 2820 词条：1883 个单词和 937 个常用短语
- `zhongkao_vocab.json`：完整释义
- `zhongkao_vocab_simple.json`：每个词性保留第一个核心义项
- `zhongkao_vocab.xlsx`：便于人工检查和导入

字段包括 `word`、`phonetic`、`pos`、`meaning`、`meaning_simple`、`frq`。

## 数据特点

- 每 500 词段内使用固定种子乱序，高频词段优先
- 纯单词音标覆盖以校验器输出为准
- 简化释义面向基础薄弱学习者，不替代完整词典

## 使用

```python
import json

with open("zhongkao_vocab_simple.json", encoding="utf-8") as file:
    words = json.load(file)
print(words[0])
```

Excel 结果可按需要导入 Anki、欧路词典或其他支持自定义词表的工具。导入外部工具前请先备份自己的学习进度。

## 构建

安装依赖：

```bash
python -m pip install -r requirements.txt
```

构建脚本默认从仓库内的 `data/` 读取输入，并把结果写回仓库根目录：

```bash
python build_vocab.py
python verify_vocab.py
```

个人输入文件不在公开仓库时，可指定本地目录；输出目录也可以单独指定：

```bash
VOCAB_DATA_DIR=/path/to/private-data \
VOCAB_OUTPUT_DIR=/path/to/output \
python build_vocab.py
```

Windows CMD：

```cmd
set VOCAB_DATA_DIR=C:\path\to\private-data
set VOCAB_OUTPUT_DIR=C:\path\to\output
python build_vocab.py
```

输入文件名必须为 `words_20tian.json`、`sguo4.xlsx` 和 `freq_50k.txt`。缺少输入时构建应失败，而不是产生不完整发布文件。

完整来源和许可边界见 [`SOURCES.md`](SOURCES.md)。

## License

代码按 MIT 发布。数据中的第三方内容仍受其原始许可约束；不能仅因为仓库有 MIT `LICENSE` 就把全部数据视为 MIT。
