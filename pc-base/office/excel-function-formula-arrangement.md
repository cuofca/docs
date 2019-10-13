# <center>Excel函数公式整理下载</center>

## [Excel函数公式整理源文件下载](https://dev.onti.net/down/CDN/Files/2019/10/13/Excel%E5%87%BD%E6%95%B0%E5%85%AC%E5%BC%8F%E6%95%B4%E7%90%86%20%281%29%281%29.docx"Excel函数公式整理源文件下载")

1、提取员工生日：   【 **注：&quot; &quot; 为小写 】**

＝ MID（ F3，7，4 ）&amp; &quot;年&quot; &amp; MID（ F3，11，2 ）&amp; &quot;月&quot; &amp; MID（ F3，13，2）&amp; &quot;日&quot;

2、计算员工工龄：   【 注： **★** 为入职时间 】

＝ INT（ （TODAY（）－ **★** ） **/** 365  ）

3、计算工龄工资：

＝ J3 \*  工龄工资！$B $3

4、统计&quot;项目经理&quot;基本工资总额：

＝ SUMIF（ 员工档案！E3:E37， &quot;项目经理&quot;， 员工档案！K3:K37 ）

5、统计&quot;本科生&quot;平均基础工资：

＝ AVERAGEIF（ 员工档案！H3:H37， &quot;本科&quot;， 员工档案！K3:K37 ）

6、图书单价的填充：

＝ VLOOKUP（ D3，   图书定价！A3：C19 ，   3 ，   FALSE ）

          图书编号                    返回第三列的值

7、提取每个学生所在班级，并按下列对应关系填写在&quot;班级&quot;列中：

  ＝ LOOKUP **（** MID（A2，3，2）， **｛**&quot;01&quot;，&quot;02&quot;，&quot;03&quot; **｝** ， **｛**&quot;1班&quot;，&quot;2班&quot;，&quot;3班&quot; **｝**** ）**

8、&quot;方向&quot;列中只能有借，贷，平三种选择：

    选定G2：G6区域  →  数据  →  数据工具  →  数据有效性  →  设置允许：序列  →  来源：借，贷，平（英文小写逗号）  →   勾选&quot;忽略空值&quot;、&quot;提供下拉箭头&quot;       →  输入信息：请在此选择

G2 ＝ IF **（** H2＝0， &quot;平&quot;， IF **（** H2＞0， &quot;借&quot;， &quot;贷&quot; **）**** ）**

   为数据列表自动套用格式，并将其转换为区域 ：

**——** 选定套用表格格式后的区域  →  表格工具：设计   →  工具：转为区域

   通过&quot;分类汇总&quot;按&quot;日&quot;计算&quot;借、贷方发生额总计&quot;，并将汇总行放于明细数据下方：

   选定区域  →  数据   →  分级显示：分类汇总   →  &quot;日&quot;、&quot;求和&quot;、&quot;本期借贷方&quot;

9、创建&quot;数据透视表&quot;：

①书店名称为列标签（ 点击拖入&quot;列标签&quot; ）

②日期、图书名称为行标签（ 点击拖入&quot;行标签&quot; ）

③销售额求和（ 点击拖入&quot;∑数值&quot; ）
