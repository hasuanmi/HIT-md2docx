# 哈工大硕士学位论文 · 范例 MD 样板（知识库②：范例 Markdown）

> 用途：作为"MD 转结构化富文本"Skill 的结构范本；也是生成器 `md2docx` 的输入最小完整结构。
> 入口：上传你自己的论文 MD（保持同样的前 matter 标题与层级），一键生成标准 docx。

```markdown
# 人工智能与产业技术融合对制造业产业链韧性的影响研究

## 封面信息
论文题目：人工智能与产业技术融合对制造业产业链韧性的影响研究
英文题目：RESEARCH ON THE IMPACT OF ARTIFICIAL INTELLIGENCE AND INDUSTRIAL TECHNOLOGY CONVERGENCE ON THE RESILIENCE OF MANUFACTURING INDUSTRY CHAINS
作者：【姓名】
学号：【学号】
所在单位：深圳研究生院
学科专业：【学科】
指导教师：【导师】
答辩日期：2026 年 6 月
学位类别：Master Degree in Engineering
学校代码：10213
密级：公开

---

## 声明
（原创性声明 + 使用授权说明正文）

---

## 摘要
中文摘要正文（500~1000 字，独立短文，不用"首先/其次/(1)(2)"，主要反映背景、问题、方法、结果）。

关键词：人工智能；产业链韧性；产业技术融合；制造业

---

## ABSTRACT
English abstract (no "First, ...", short sentences, consistent with Chinese abstract).

Keywords: artificial intelligence, industrial chain resilience, industrial technology convergence, manufacturing

---

## 目录
（导出器自动写入目录域，编到 3 级 X.X.X）

---

# 1 绪论
## 1.1 研究背景
正文（宋体小4，Times New Roman 数字/英文，多倍行距 1.25）。

本文的主要研究内容如下：
(1)……；
(2)……；
(3)……。

# 2 相关研究
## 2.1 国内外研究现状
如图 2-1 所示，……

![图 2-1 框架示意图](img/framework.png)

## 2.2 本章小结
（各章末节"本章小结"，第一章除外）

# 参考文献
[1] 李晓东，张庆红，叶瑾琳. 气候学研究的若干理论问题[J]. 北京大学学报，1999，35（1）：101-106.
[2] Helfat C E, Raubitschek R S. Product Sequencing: Co Evolution of Knowledge, Capabilities and Products[J]. Strategic Management Journal, 2000（21）：961-979.

# 致谢
……
```

## 结构规则速记

- 前 Matter 用二级标题：`## 封面信息 / ## 声明 / ## 摘要 / ## ABSTRACT / ## 目录`。
- 正文一级 `# 1 绪论`（导出为「第1章 绪论」，居中黑体二号）；层级 `# / ## / ###` 对应章/节/条。
- 图题 `图 2-1 …`、表题 `表 2-1 …` 单独成行；引用标注 `[1]`、`[1-3]`、`[3]92`。
- 参考文献 `# 参考文献`、致谢 `# 致谢`、附录 `# 附录`。
- 关键词：中文用"；"、英文用","且全小写（见 R01）。
