# FineReportDemo
1、直接git项目作为你的maven项目（当然你可以自己去FineReport帮助文档—>二次开发看web项目如何嵌入FineReport）
2、FineReport客户端负责创建cpt模板，保存后必须复制到本项目下的WEB-INF/reportlets下
3、XXXServlet：java导出excel （当然你本地没有我的数据库，你是运行不出来的^_^）

ps.关于cpt模板中，动态where的编写：where 1=1 ${if(len(订单类型) == 0,""," and o.title like '%" + 订单类型 + "%'")}

视频教程：http://bbs.fanruan.com/plugin.php?id=threed_video:view&tid=67986&cid=2

帮助文档：http://help.finereport.com/

FineReport报表和J2EE应用的集成： http://www.blogjava.net/fannie/archive/2013/05/08/398985.html
