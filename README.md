# 优雅地将课表导入日历

## 为什么要把课表导入日历呢？

当然是为了**优雅**。

- 可以方便地自定义和修改课程；
- 可以将课表作为小部件放到桌面上，方便而实用；
- 可以准确地管理上课时间，并将课程和其他日程安排统一管理；
- 可以很容易地处理调课等突发情况；
- 可以设置上课前N分钟提醒，避免忘记上课；
- 脱离臃肿的充斥着广告的大学生课表应用；
- 可以借助Google Calendar或Outlook Calendar做到多平台无缝融合，在电脑、手机、PAD甚至Web网页都可以管理课程；
- 可以借助滴答清单等To-Do-List应用安排日程，iOS用户还有好用而漂亮的Sorted3可以使用。

那么，何乐而不为呢？

### 效果展示


![Windows日历应用](https://qn.rneko.com/20202007/2002-F.png)

![Windows任务栏](https://qn.rneko.com/20202007/1948-p.png)

![iOS日历应用 (from SunsetYeu)](https://cdn.sspai.com/2020/03/27/a5542bf3ae121cba462ee8387c1e9428.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

![Sorted3 (from SunsetYeu)](https://cdn.sspai.com/2020/03/27/5621d33ea3e1a9da1adba462065c2fd4.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

![MIUI系统日历](https://qn.rneko.com/20202007/1954-b.png)

![课程详情](https://qn.rneko.com/20202007/1954-Q.png)

## 如何将课表导入日历

在[@Triple-Z](https://github.com/Triple-Z)、[@MiaoTony](https://github.com/miaotony)等同学的努力下，我校本科生已经可以很方便地从教务系统，获取课表并生成 iCal 日历文件，以导入Google Calendar、Outlook或系统日历。

具体项目参见[NUAA_ClassSchedul](https://github.com/miaotony/NUAA_ClassSchedule)，该项目还有一个漂亮又易用的[在线版本](https://anyknew.a2os.club/)。

而遗憾的是截至目前，研究生信息管理系统还尚未对接成功，所以无法一键导出=_=。但得益于[@陈某豪](https://sspai.com/u/iobi2v7m/updates)和[@SunsetYe66](https://github.com/SunsetYe66)的贡献，我们依然可以较为优雅地将课表转换为ics日历文件。

## 生成ics文件

### 部署项目

从 GitHub 拉取程序

```
git clone https://github.com/Xm798/ClasstableToIcalforNUAA.git
```

确保本机已安装**Python 3环境**，然后安装依赖

```
pip install uuid xlrd
```

### 准备Excel课表文件

登录[南京航空航天大学研究生信息管理系统](http://graduate.nuaa.edu.cn/gmis5/home/stulogin)，进入到“课程学习”-“选课结果查询”模块之中，点击“导出数据”，获得一份选课结果的Excel文件。

![导出课表](https://qn.rneko.com/20202007/2010-M.png)

在交给 ClasstableToIcal 处理之前，我们需要先对该课表文件做一些预先处理。

#### 拆分课表

首先，针对“时间”中存在多个上课时间的，例如`星期二 下午5-下午6,星期五 下午5-下午6`，需要先拆分为单条记录，确保每行记录为一节课。

借助Power Query，可以很轻松的实现这项工作。

> Power Query 在Microsoft Office 2016及以上版本中为内置组件，Office 2010或2013版本需要手动下载安装，不支持2010以下版本和WPS。
>
> 若无PQ组件或不想使用PQ操作，由于数据量非常小，**手动拆分完全可行*，只要将每行拆分为单独的一节课即可。

打开下载的课表文件，在Excel中选择“数据-获取数据-来自文件-从工作簿”，选择下载的课表文件并打开，选中sheet1，点击“加载”。

双击右侧“查询&连接”中的sheet1，打开Power Query编辑器。

删除无用的“学分、阶段、选课人数”三列，然后选中“时间”一列，点击“转换-拆分列-按分隔符”，分隔符选择“逗号”。

![拆分列](https://qn.rneko.com/20202007/1856-S.png)

按住Ctrl键多选除了时间之外的其他列，点击“转换-逆透视列-逆透视其他列”。

![逆透视列](https://qn.rneko.com/20202007/1856-G.png)

删除多出的“属性”列 ，点击“主页-关闭并上载”，得到拆分后的表格。



#### 预处理课表文件

按照模板文件 temp_classInfo.xlsx 的要求，需要将待导入的文件处理成模板文件的样子。

时间设置 conf_classTime.json 文件我已按照我校教学日历调整完毕，无需再做修改，这部分主要针对 classInfo.xlsx 的制作进行预处理。

在拆分完成的表格中继续操作，选中时间一列，点击“数据-分列”，依次选择“分隔符号、空格、完成”，将星期和节次拆分。

![拆分星期](https://qn.rneko.com/20202007/1859-V.png)

而后，在节次右侧的两个单元格分别使用以下两个公式，并下拉填充，将星期和节次转换为对应的数字。

```
=IF(RIGHT(H2,1)="","",MATCH(RIGHT(H2,1),{"一";"二";"三";"四";"五";"六";"日"},))

=IF(I2="","",MATCH(I2,{"上午1-上午2";"上午1-上午3";"上午1-上午4";"上午2-上午3";"上午2-上午4";"上午3-上午4";"下午5-下午6";"下午5-下午7";"下午5-下午8";"下午6-下午7";"下午6-下午8";"下午7-下午8";"晚上9-晚上10";"晚上9-晚上11"},))
```

![转换星期和节次](https://qn.rneko.com/20202007/1901-c.png)

到这里，对classInfo.xlsx的预处理已经完成，可以将对应字段粘贴到classInfo.xlsx之中了。

#### 制作Classinfo.xlsx

根据作者介绍，classInfo.xlsx中的字段含义为：

```
className - 课程名称
startWeek - 开始周数
endWeek - 结束周数
weekday - 课程日期（周几）
classTime - conf_classTime.json 中定义的时间段代号
classroom - 教室
weekStatus - 是否单双周排课：正常排课 = 0，单周排课 = 1，双周排课 = 2
classSerial - 可选，课程序号
classTeacher - 可选，教师名

最后两个可选字段如果不需要可以关闭，只需在 excel_reader.py 中的 27 和 28 行将：
self.config["isClassSerialEnabled"] = [1, 7]
self.config["isClassTeacherEnabled"] = [1, 8]
后方的方框中，要关闭的功能的 1 改成 0 即可（即 [0, 7] 或 ​[0, 8]​）。

注意：若课程有不同排课方式或一周有多节课，需要分多条记录录入。
```

**将之前表格中的对应字段，粘贴至 classInfo.xlsx** 。`weekStatus`填写0，并补充线上课的`classroom`为线上即可。

![制作完成的classinfo.xlsx](https://qn.rneko.com/20202007/1914-t.png)

### 生成文件

> 该部分直接引用脚本作者教程。
>
> **我校2020-2021学年第一学期第一周的日期输入20200831。**

打开命令提示符或终端，定位到该项目根目录下。

![img](https://cdn.sspai.com/2020/03/27/0a6dccddce3cfb1c5092abaa95ade2a4.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

然后执行 main.py：

```
python main.py
或
python3 main.py
```

![img](https://cdn.sspai.com/2020/03/27/fc4aea7f78366215bea3bdbf11fccc4c.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

首先**选 2 进入课程信息读取工具**：若未修改过 Excel 文件结构，直接回车即可；若提示成功，文件夹中将多出一个新的 JSON 文件。

![img](https://cdn.sspai.com/2020/03/27/aa879429832ef4987b4f314dc6b93465.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

此时，再输入 3 进入 iCal 生成工具，按提示输入必需信息后即可生成最终文件。请注意此处输入的日期必须严格按照提示输入**开学第一周周一的日期**，且该日期的格式是 `YYYYMMDD`，即 **年年年年月月日日**，中间不加符号。

![img](https://cdn.sspai.com/2020/03/27/fb9ccfb6880a34247b17aeb94a1a183a.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

## 导入设备

注意：导入任何日历之前，都建议新建一个新的日历用于导入操作，便于出错时批量修改，以及不同颜色更容易区分。

### iOS设备

直接将生成的文件传输到手机后点击打开，选择“导入日历”即可。

![img](https://cdn.sspai.com/2020/03/27/3cb27efb119e6e979a1f3b70f2ce00eb.png?imageView2/2/w/1120/q/90/interlace/1/ignore-error/1)

### Android与Windows设备

部分Android设备的原生日历并不支持ics文件的导入功能，其实更加推荐的方法是**借助于Google Calendar或者Outlook来完成日程管理**，因为无论是谷歌日历还是微软的Outlook，都可以很方便的完成手机、PAD与Windows设备之间的日历同步。

如果有科学手段，可以选用Google Calendar，如果没有，Outlook也是一个极为不错的选择，全凭个人喜好。

#### Google Calendar导入

电脑端打开[Google Calendar](https://calendar.google.com/)，选择添加其他日历-导入，记得选择一个单独的日历，如果没有，可以新建一个。点击导入，即可完成。

![电脑端导入Google Calendar](https://qn.rneko.com/20202007/1923-Q.png)

也可以在手机端完成导入，在安装好谷歌日历的情况下，直接点击接收到的ics文件，选择使用“日历”打开，即可。

![手机端导入Google Calendar](https://qn.rneko.com/20202007/1925-G.png)

#### Outlook导入

电脑端访问[Outlook](https://office.live.com/start/Outlook.aspx)，使用微软账号登录。

点击左侧的“日历”，添加日历-创建空白日历，新建一个课表日历。

![创建日历](https://qn.rneko.com/20202007/1934-t.png)

然后选择“从文件上传”，选择生成的ics文件，导入到创建的课表日历中即可。

![Outlook导入日历](https://qn.rneko.com/20202007/1936-I.png)



或者，Windows10系统可以直接在本地右键ics文件，选择使用系统日历打开。

![使用日历打开](https://qn.rneko.com/20202007/1938-s.png)

然后导入到日历即可。首次使用可能需要登录账号，此处登录微软账号或Gmail或其他任何支持的账号都可以。

![导入到日历](https://qn.rneko.com/20202007/1939-O.png)

## 跨平台同步

在手机端、电脑端、PAD端登录同一账号之后，即可实现全平台同步。

### Android端

若使用Google Calendar，手机端需要安装谷歌日历应用；若使用Outlook，手机端需要安装Outlook应用。当然，部分安卓操作系统的系统日历可以直接订阅谷歌日历和Outlook的日历，不安装App使用系统日历也完全OK。

例如，小米日历在“日历-设置-日历账号管理”中，可以添加Gmail日历。

![添加日历账号](https://qn.rneko.com/20202007/1944-R.png)



### 电脑端

Windows10系统使用系统的“邮件与日历”应用，登录对应账号，即可完成电脑端平台的日历同步。

如果是使用Google服务，需要先解除UWP应用的代理限制，请自行搜索“解除UWP应用网络隔离”，或参阅[Win10解除UWP应用网络隔离允许访问代理](https://blog.rneko.com/archives/12/)。推荐使用**Fiddler**中的**WinConfig**模块，全部解除隔离。



---

# 源项目README

Convert Classtable to iCal using Pything and Excel as data source.

该工具可以方便地将课程表转换为 `.ics` 格式以导入各种设备的「日程」中。

## Usage

以下为简略说明，更详细教程请参看[少数派](https://sspai.com/post/59694)。

先安装依赖：

```shell
pip install uuid xlrd 
```

然后执行 `main.py`：

```shell
python main.py
# or
python3 main.py
```

测试环境为：Python 3.7.2，Windows 10 x64.

## 文件中格式解释

### temp_classInfo.xlsx

课程的名称、起始周数等在文件里已标示清楚，weekStatus 是单双周标记。

> 0=不分单双周，1=单周，2=双周

是否显示教师、是否开启单双周功能可在 `excel_reader.py` 中更改。

### conf_classTime.json

```json
"1": {
    "name": "第 1、2 节", 
    "startTime": "082000",
    "endTime": "095500"
}
```

该文件为 JSON 格式，一开始的数字是**时段编号**，对应 `temp_classinfo.xlsx` 里的 `classTime` 字段；`startTime` 与 `endTime` 采用 `%H%M%S` 格式，即时、分、秒去掉分隔符。

## Feature

现在支持：

- 单双周排课
- 课前n分钟提醒（待进一步测试）
- 不同教室（添加多个条目）

## License

LGPLv3