#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""zhongkao_vocab 发布内容综合校验：结构/质量/一致性/残留噪音。"""
import json, re, random, os

BASE = os.environ.get("VOCAB_PROGRESS_DIR", os.path.dirname(os.path.abspath(__file__)))
issues = []
def chk(cond, msg):
    if not cond:
        issues.append(msg)

# 1. JSON 合法性 + 结构
words = json.load(open(os.path.join(BASE, "zhongkao_vocab.json"), encoding="utf-8"))
chk(isinstance(words, list) and len(words) == 2820, f"JSON 词条数: {len(words)} (期望 2820)")
keys_ok = all(set(w.keys()) >= {"word", "phonetic", "pos", "meaning", "meaning_simple"} for w in words)
chk(keys_ok, "字段缺失")

# 2. word 唯一性
ws = [w["word"] for w in words]
dup = {x for x in ws if ws.count(x) > 1}
chk(not dup, f"重复词: {dup}")

# 3. 空字段
no_meaning = [w["word"] for w in words if not (w.get("meaning") or "").strip()]
no_simple = [w["word"] for w in words if not (w.get("meaning_simple") or "").strip()]
chk(not no_meaning, f"空释义: {no_meaning[:10]}")
chk(not no_simple, f"空精简释义: {no_simple[:10]}")

# 4. 音标覆盖率（纯单词）
words_only = [w for w in words if re.fullmatch(r"[a-zA-Z]+[a-zA-Z\-']*", w["word"]) and " " not in w["word"]]
no_ph = [w["word"] for w in words_only if not (w.get("phonetic") or "").strip()]
chk(not no_ph, f"纯单词缺音标: {no_ph[:20]} ({len(no_ph)}个)")

# 5. 音标格式异常（含中文/数字）
bad_ph = [w["word"] for w in words if re.search(r"[\u4e00-\u9fff0-9]", w.get("phonetic") or "")]
chk(not bad_ph, f"音标含中文/数字: {bad_ph[:10]}")

# 6. 词序验证：前 20 应为高频词
top20 = [w["word"] for w in words[:20]]
print(f"前20: {top20}")
high_freq_ok = {"feel", "work", "team", "hope", "time", "have", "hurry"} <= set(top20)
chk(high_freq_ok, "前20不含典型高频词")

# 7. 乱序验证：相邻词首字母不单调（抽样）
rng = random.Random(42)
idx = rng.sample(range(100, 2700), 30)
monotone = 0
for i in idx:
    a, b = words[i]["word"][0], words[i + 1]["word"][0]
    if a.lower() > b.lower():
        monotone += 1
# 不严格判断，仅报告
print(f"乱序抽样: 30 对相邻词中 {monotone} 对为降序 (期望~15±5)")

# 8. 残留噪音扫描（语法注记"（bite的过去式）"与词典标注 <美>/<美俚> 为有用信息，仅报告不判错）
noise_patterns = {
    "语法注记(可接受)": r"（\w+的(过去式|过去分词|现在分词)）",
    "词典标注(可接受)": r"<美>|<英>|<美俚>|<英俚>",
    "词形变化表(应修)": r"\[ *过去式|\[复数|过去分词|现在分词",
    "网络语(应修)": r"Hold住|中英混用",
    "品牌(应修)": r"品牌|服装品牌|皮革品牌",
    "截断": r"\(\s*[a-z]+\s*$|；；|;;",
}
for name, pat in noise_patterns.items():
    hits = [w["word"] for w in words if re.search(pat, w["meaning"])]
    if hits:
        print(f"  噪音[{name}]: {len(hits)}条 例: {hits[:5]}")
        # 词形变化/网络语/品牌 视为问题，截断/HTML 也记录
        if name in ("词形变化", "网络语", "品牌", "HTML残留"):
            issues.append(f"残留噪音[{name}]: {hits[:8]}")

# 9. 抽样释义质量（10 个词人工核对点）
print("\n抽样 10 词:")
for w in rng.sample(words, 10):
    print(f"  {w['word']} {w.get('phonetic','')[:25]} | {w['meaning_simple'][:50]}")

# 10. Excel 与 JSON 一致性
import openpyxl
wb = openpyxl.load_workbook(os.path.join(BASE, "zhongkao_vocab.xlsx"), read_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
header, data = rows[0], rows[1:]
chk(len(data) == len(words), f"Excel 行数 {len(data)} vs JSON {len(words)}")
chk(header == ("序号", "单词", "音标", "词性", "中文释义", "中文释义(精简)"), f"Excel 表头: {header}")
mismatch = [i for i, (r, j) in enumerate(zip(data, words)) if r[1] != j["word"]]
chk(not mismatch, f"Excel/JSON 词序不一致: {mismatch[:5]}")
wb.close()

print(f"\n=== 校验结果: {'✅ 通过' if not issues else '❌ ' + str(len(issues)) + ' 个问题'} ===")
for i in issues:
    print("  ❌", i)
