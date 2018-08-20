
<%
Const SystemName = "新闻出版审读系统"         '系统名称
Const JccTitleName = "新闻出版审读系统"       '系统通用标题
Const SiteUrl = "http://127.0.0.1"            'ip地址
Const Copyright = "锦华科技"                  '版权所有
Const Vsersion = "Version 1.0 Beta"           '系统版本号
Const SessionTimeout = 120                    'session过期时间为120分钟
Const CookieTimeout = 2                       'Cookie过期时间为2小时
Const AdminPath = "/Site/"                    '数字报后台地址
Const PaperFilePath = "/SLB/"                 '数字报文件存放路径

Rem-------------------------------------------- 数据库设置
Const SqlDatabaseName = "Jcs_db"            '数据库名称
Const SqlUsername     = "jcs_paper"                '数据库用户名
Const SqlPassword     = "111111"                '数据库密码
Const SqlLocalName    = "(local)"              '数据库地址
Const SqlNowString    = "GetDate()"		'数据库时间参数
'数据库设置完毕
'数据库设置完毕

Const PaperSize = 12                        '分页每页的记录个数

Rem 定义当前运行模式，测试为1，正常运行为0,不输出错误信息有利于安全
Const IsDebug = 0
Const CookiesName		= "JcsCookie"	    'Cookies名称
Const SessionName	    = "Jcs20090315"		'Session名称
Const PassLength		= 4					'随机字符串长度，程序中多处用到，请勿修改
Const ArrSafeCode		= "Jcs"			    '字符串和数组修正参数
Const JcsDomain         = "127.0.0.1"       '网站域名或ip
%>