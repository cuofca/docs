# Excel函数公式整理

{% tabs %}
{% tab title="在线预览" %}
## [点我在线预览\(Excel函数公式整理\)](https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fjxjjxy-my.sharepoint.com%2Fpersonal%2Fon_mail_mzr_me%2F_layouts%2F15%2Fdownload.aspx%3FUniqueId%3D4ced169a-a44f-4e06-a76b-8d664c185673%26Translate%3Dfalse%26tempauth%3DeyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvanhqanh5LW15LnNoYXJlcG9pbnQuY29tQDFjZGYxZDdjLWNhODUtNDVjNy1iMWE2LTBkMTk3YWEwN2RjNSIsImlzcyI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMCIsIm5iZiI6IjE1NzA5NjQ4NzciLCJleHAiOiIxNTcwOTY4NDc3IiwiZW5kcG9pbnR1cmwiOiJDMTZ3cGRJVGVtcWtTSFJoR0Z3TjIvOU9tTFJ3ZC91ZUNndmx1bVAxTjVZPSIsImVuZHBvaW50dXJsTGVuZ3RoIjoiMTQ0IiwiaXNsb29wYmFjayI6IlRydWUiLCJjaWQiOiJOVEkzWmpsaU9UQXRNakkwTUMwMFl6azJMVGxtWW1RdE5qRTRNV05rTTJKa05EbG0iLCJ2ZXIiOiJoYXNoZWRwcm9vZnRva2VuIiwic2l0ZWlkIjoiTjJVeFpXWTRaVGt0TkdVd015MDBOakpoTFRobU1Ea3ROVFl5T0RZM05ESTRabVEwIiwiYXBwX2Rpc3BsYXluYW1lIjoiT0xBSU5ERVgiLCJnaXZlbl9uYW1lIjoi5rC46L6JIiwiZmFtaWx5X25hbWUiOiLpn6kiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiI0NGU4YzdkZC02YTdhLTQ4YWMtYWIwZC04YWU4NWQyYWI3YTUiLCJ0aWQiOiIxY2RmMWQ3Yy1jYTg1LTQ1YzctYjFhNi0wZDE5N2FhMDdkYzUiLCJ1cG4iOiJvbkBtYWlsLm16ci5tZSIsInB1aWQiOiIxMDAzN0ZGRUFFQkY5OURFIiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzdmZmVhZWJmOTlkZUBsaXZlLmNvbSIsInNjcCI6ImFsbGZpbGVzLndyaXRlIGFsbHByb2ZpbGVzLnJlYWQiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsfQ.WU1iY2h5MzhZQ0w1dDlqeGtZaTc3OHgrZW43MlZXUVJZQ0xsOXZaU2R2az0%26ApiVersion%3D2.0)
{% endtab %}

{% tab title="在线下载" %}
## [点我在线下载\(Excel函数公式整理\)](https://dev.onti.net/down/CDN/Files/2019/10/13/Excel%E5%87%BD%E6%95%B0%E5%85%AC%E5%BC%8F%E6%95%B4%E7%90%86%20%281%29%281%29.docx)
{% endtab %}
{% endtabs %}

1、提取员工生日： 【 **注：" " 为小写 】**

＝ MID（ F3，7，4 ）& "年" & MID（ F3，11，2 ）& "月" & MID（ F3，13，2）& "日"

2、计算员工工龄： 【 注： **★** 为入职时间 】

＝ INT（ （TODAY（）－ **★** ） **/** 365 ）

3、计算工龄工资：

＝ J3 \* 工龄工资！$B $3

4、统计"项目经理"基本工资总额：

＝ SUMIF（ 员工档案！E3:E37， "项目经理"， 员工档案！K3:K37 ）

5、统计"本科生"平均基础工资：

＝ AVERAGEIF（ 员工档案！H3:H37， "本科"， 员工档案！K3:K37 ）

6、图书单价的填充：

＝ VLOOKUP（ D3， 图书定价！A3：C19 ， 3 ， FALSE ）

```text
      图书编号                    返回第三列的值
```

7、提取每个学生所在班级，并按下列对应关系填写在"班级"列中：

＝ LOOKUP **（** MID（A2，3，2）， **｛**"01"，"02"，"03" **｝** ， **｛**"1班"，"2班"，"3班" **｝\*\*** ）\*\*

8、"方向"列中只能有借，贷，平三种选择：

```text
选定G2：G6区域  →  数据  →  数据工具  →  数据有效性  →  设置允许：序列  →  来源：借，贷，平（英文小写逗号）  →   勾选&quot;忽略空值&quot;、&quot;提供下拉箭头&quot;       →  输入信息：请在此选择
```

G2 ＝ IF **（** H2＝0， "平"， IF **（** H2＞0， "借"， "贷" **）\*\*** ）\*\*

为数据列表自动套用格式，并将其转换为区域 ：

**——** 选定套用表格格式后的区域 → 表格工具：设计 → 工具：转为区域

通过"分类汇总"按"日"计算"借、贷方发生额总计"，并将汇总行放于明细数据下方：

选定区域 → 数据 → 分级显示：分类汇总 → "日"、"求和"、"本期借贷方"

9、创建"数据透视表"：

①书店名称为列标签（ 点击拖入"列标签" ）

②日期、图书名称为行标签（ 点击拖入"行标签" ）

③销售额求和（ 点击拖入"∑数值" ）

