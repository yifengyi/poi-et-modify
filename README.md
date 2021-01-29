# poi-et
#### 备注
 该项目在`https://gitee.com/heibaixiong/poi-et` 基础上进行了微调。
#### 介绍
由于经常需要操作web导出excel，使用了poi组件来进行操作。后来在偶然中遇到了poi-tl[https://github.com/Sayi/poi-tl]，使得自己受到了启发，于是有了poi-et。

#### 软件说明
软件说明请参照源码doc目录下的文档

#### 更新说明
1.0.2: 扩展了NiceXSSFWorkbook。NiceXSSFWorkbook是针对POI中XSSFWorkbook功能的进一步封装和完善，以便更好的帮助java开发者操作excel。它继承类XSSFWorkbook中的所有的功能，并扩展了一些功能（包括excel表格的插入行、删除行、插入列、删除列等等功能）。XSSFTemplate可以通过getXSSFWorkbook()方法来获取这个对象，以便进行操作。
       更正poi-et文档中的maven地址不能正常使用的问题，修改了maven地址。
1.0.3: TextRenderData支持已数值方式写入数据，详见poi-et文档
