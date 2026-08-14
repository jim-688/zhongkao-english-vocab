# 中考英语零基础核心词表（乱序版）

零基础友好的中考英语核心词汇表：**乱序排列、高频在前**，含音标、词性、中文释义。
JSON 和 Excel 双格式，可直接导入背词软件或自测脚本。

## 数据规模
- **2820 词条**（1883 单词 + 937 常用短语），覆盖中考高频核心词汇
- 字段：`word` 单词 | `phonetic` 音标 | `pos` 词性 | `meaning` 中文释义
- 两个版本：
  - `zhongkao_vocab.json`：完整释义（含全部义项）
  - `zhongkao_vocab_simple.json`：零基础精简释义（每个词性保留第一个义项）

## 特色
- **乱序 + 高频在前**：词频排序后每 500 词段内乱序，高频词先背、段内不按字母序
- **音标全覆盖**：纯单词音标覆盖 100%（英音+美音）
- **零基础友好**：`meaning_simple` 只保留核心义项，不吓人

## 数据来源
- 词单：公开中考词表整理（含人工校对修正 12 处释义错误）
- 音标：sguo4.xlsx（四级闪过电子书）+ 有道词典 + Free Dictionary API（Wiktionary，CC BY-SA）
- 词频：freq_50k.txt
- 释义为整理者原创整理，如有版权问题请联系删除

## 用法
```python
import json
words = json.load(open('zhongkao_vocab_simple.json', encoding='utf-8'))
print(words[0])
# {'word': 'feel', 'phonetic': '英[fiːl] 美[fiːl]', 'pos': 'n vi vt', 'meaning': 'n. 感觉；vi. 觉得；vt. 感觉'}
```

Excel 版可直接导入 Anki（自定义笔记模板）、欧路词典、墨墨背单词等。

## 构建
```bash
python build_vocab.py
```
依赖：openpyxl。数据源文件按脚本内路径准备。

## Roadmap
- [ ] 每词一个简单例句（零基础友好）
- [ ] 单词发音音频
- [ ] 按省份考纲差异拆分

## License
MIT（代码）。音标数据部分来源 Wiktionary（CC BY-SA 4.0）。
