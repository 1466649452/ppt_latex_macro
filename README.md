# ppt_latex_macro

# 开发过程
1. 开发 macro.bas 的代码处理逻辑
2. 开发 ppt.xml 的宏组件显示逻辑
3. 创建 latex_ppt.pptm 的宏文件PPT
4. 在 Windows 环境下，使用 RibbonXmlEditor 软件打开 ppt.xml 文件，并邮件点击 callback 获取 xml 返回得到的回调函数（需要成功，不成功则xml文件写错）
5. 在 RibbonXmlEditor 里选择点击保存，保存的时候保存到 latex_ppt.pptm 这个文件，等待完成该 .pptm 里就有插件的内容了。
6. 新建宏，添加模块，里面写入回调函数，并绑定对应的 macro.bas 里的 VBA 处理函数。