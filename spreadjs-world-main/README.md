# spreadjs-world
SpreadJS+组件化设计器+GCExcel Java API


# 使用指南

### 运行环境
* 安装IntelliJ IDEA
* IDEA 安装插件 Spring Assistant 


### 试用版升级正式版
* 升级SpreadJS和Designer Lib至V14，并引入SpreadJS新增透视表插件
* 更新config
  * 通过GC.Spread.Sheets.Designer.DefaultConfig获取默认config。无需引入config json或js
  * 在默认config基础上新增修改内容
  * 所有自定义config中commands必须使用commandMap注册，config中command为string类型，否则config序列化后execute会丢失。
  * Command type “combobox”修改为“comboBox”


### 功能介绍
示例使用SpreadJS、SpreadJS Designer lib和GCExcel。实现了如下功能：

* 自定义设计器

|  模块   | 功能  |  状态 |  参考功能| 
|  ----  | ----  |  ---  |  --- |
| Ribbon  | 自定义按钮（Button） |  已完成 |  加载、上传 |
| Ribbon  | 自定义Checkbox |    已完成 |页面布局-打印预览、数据透视-字段列表  |   
| Ribbon | 自定义Checkbox（大图标）  |  已完成 | 视图-显示滚动条|  
| Ribbon | 自定义Combo  |  已完成 | 页面布局-宽度/高度|  
| Ribbon | 自定义DropDown  |  已完成 | 文件操作-导出图片，插入-图表|  
| Ribbon | 修改默认Combo选项  |  已完成 | 开始字体| 
| Ribbon | 按钮状态，是否可选可点击 |  已完成 | 操作-返回模板| 
| Ribbon | 按钮状态，显示隐藏|  未完成|
| 文件菜单| 自定义文件菜单模板| 已完成| 技术博客（删除，暂不支持增加）
| Dialog | 自定义Dialog - 区域选择 | 已完成| 数据透视-新建
| Dialog | 自定义Dialog - 文本、Combo、Checkbox | 已完成 | 页面布局-页面设置
| Dialog | 修改默认Dialog内容  | 已完成| 设置单元格格式 - 字体
| Dialog | 自定义布局示例 | 已完成 | 页面布局-页面设置
| SidePanel | 自定义SidePanel | 已完成 | 点击图片 （仅做展示用）
| ContentMenu | 自定义右键菜单 | 已完成 | 右键-文字旋转

* 操作

|  模块   | 功能  |  状态 |  参考功能| 
|  ----  | ----  |  ---  |  --- |
|文件IO| 上传文件 | 已完成| 操作 - 上传
|文件IO| GCExcel 加载Excel| 已完成| 操作 - 加载
|文件IO| GCExcel 导出PDF| 已完成| 操作 - PDF
|文件IO| GCExcel 导出图片| 已完成| 操作 - 导出图片
|文件IO| GCExcel 导出HTML| 已完成| 操作 - HTML
|文件IO|  SpreadJS 导出HTML| 未完成
|文件IO|  GCExcel 导出Excel| 未完成
|数据绑定|  SpreadJS 自动加载模板结构| 已完成 | 操作 - 加载
|数据绑定|  SpreadJS 数据绑定| 已完成 | 操作 - 前端绑定
|数据绑定|  SpreadJS 模板通过GCExcel绑定| 已完成 | 操作 - GCExcel BindingPath 绑定
|数据绑定|  GCExcel 模板| 已完成 | 操作 - GCExcel 模板绑定
|数据绑定|表格绑定设置列公式|未完成
|数据绑定|表格绑定设置插入模式|未完成


*页面布局

|  模块   | 功能  |  状态 |  参考功能| 
|  ----  | ----  |  ---  |  --- |
|分页预览| 分页线 | 已完成 |页面布局-打印预览
|分页符| 插入分页符 | 已完成 |页面布局-插入分页符
|分页符| 删除分页符 | 已完成 |页面布局-删除分页符
|分页| 自适应分页数量 | 已完成 |页面布局-高度、宽度
|纸张| 纸张方向 | 已完成 | 页面布局-打印设置
|纸张| 纸张大小 | 已完成 | 页面布局-打印设置（仅做展示用）

*数据透视

|  模块   | 功能  |  状态 |  参考功能| 
|  ----  | ----  |  ---  |  --- |
|  透视表  | 新建  |  已完成  |  数据透视 - 新建 |
|  透视表  | 展示字段列表  |  已完成  |  数据透视 - 字段列表 |
|  透视表  | 属性设置  |  未完成

*开始

|  模块   | 功能  |  状态 |  参考功能| 
|  ----  | ----  |  ---  |  --- |
|  格式刷  | 复制格式 | 已完成|

