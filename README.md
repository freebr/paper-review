# 专业学位论文评阅系统 Speciality Degree Paper Review System
+ 版本列表
    * 20201105
    * 20201104
    * 20180403
    * 20170912

+ 一、组件安装
    - 安装 URLRewrite 组件（IIS 服务器 URL 重写模块），下载地址：[https://www.microsoft.com/zh-CN/download/details.aspx?id=7435]。

+ 二、数据库配置
    - 数据库部署到服务器后，创建一个登录名 `PaperReviewSystem`，密码为 `HgggLwpy@87114057`，配置如下用户映射：数据库 `PaperReviewSystem`、`TutorRecruitSys`、`Jiaowu`，用户 `PaperReviewSystem`，默认架构 `dbo`，权限为 `db_owner`、`public`。
    - 根据所安装的 SQL Server 版本，修改 `inc\database.inc` 第 51 行中的 `Provider` 属性值，若为 2008，则值为 `SQLNCLI10`，若为 2012，则值为 `SQLNCLI11`

+ 三、系统配置
    * 1.配置本系统数据库地址：修改 `/inc/config.inc` 中 `uriDatabaseServer` 的值
    * 2.配置教务系统数据库地址：修改 `/inc/config.inc` 中 `uriJWDatabaseServer` 的值，默认为 `116.57.68.162,14033`

+ 四、服务器配置
    * 1.创建一个 Administrators 组的账户 `IUSR_OFFICE`；
    * 2.在“组件服务”管理工具中，选择 `计算机/我的电脑/DCOM 配置/Microsoft Word 97 * 2003 文档`，在其“属性”窗口-“安全”选项卡中，为 `IUSR_OFFICE` 账户添加“启动和激活权限”、“访问权限”、“配置权限”，在“标识”选项卡中，选择“下列用户”并输入 `IUSR_OFFICE` 的账号和密码，单击“确定”；
    * 3.在 IIS 管理器中对下列文件配置用户标识为 `IUSR_OFFICE` 的匿名身份验证：`/student/fillInTable.asp`, `/admin/updatePaper.asp`, `/expert/doReview.asp`, `/tutor/updatePaper.asp`, `/admin/extra/doReview.asp`, `/admin/batchExportProfile.asp`, `/admin/batchExportProfileById.asp`, `/admin/updateReviewApp.asp`, `/admin/genReviewApp.asp`, `/admin/importDetectResult.asp`, `/admin/importReviewResult.asp`, `/admin/importReviewResultByExcel.asp`, `/admin/importReviewResult.asp`，用于实现后台导出 PDF 等功能

+ 五、系统入口
    * 1.教务端：`/?usertype=admin`（以 `ouyangquan` 身份登录）
    * 2.学生端：`/?usertype=student&no=XXX`（XXX为学生学号）
    * 3.教师端：`/?usertype=tutor&name=XXX`（XXX为教师账号）
    * 4.评阅专家端：`/?usertype=expert&name=XXX`（XXX为评阅专家账号）