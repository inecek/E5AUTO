# OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

## 说明 ##
* E5自动续期程序，但是**不保证续期**
* 设置了**周六日(UTC时间)不启动**自动调用，周1-5每6小时自动启动一次 （修改看教程）
* 调用api保活：
     * 查询系api：onedrive,outkook,notebook,site等
     * 创建系api: 自动发送邮件，上传文件，修改excel等
 
 * 教程地址: https://www.vvhan.com/officeE5-AutoApi.html
 * 冲冲冲！！！
E5自动续期程序，但是不保证续期
设置了周六日(UTC时间)不启动自动调用，周1-5每6小时自动启动一次 （修改看教程）
调用api保活：
查询系api：onedrive,outkook,notebook,site等
创建系api: 自动发送邮件，上传文件，修改excel等
步骤
准备工具：

E5开发者账号（非个人/私人账号）
管理员号 ———— 必选
子号 ———— 可选 （不清楚微软是否会统计子号的活跃度，想弄可选择性补充运行）
rclone软件，下载地址  前往下载
步骤大纲：

微软方面的准备工作 （获取应用id、密码、密钥）
GIHTHUB方面的准备工作 （获取Github密钥、设置secret）
试运行
微软方面的准备工作
第一步，注册应用，获取应用id、secret

首先去  E5应用注册 注册一个应用

先用e5管理员账号登录网站,然后在主页找到Azure Active Directory点进去

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

再在左侧目录找到点击应用注册

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

再点上方的新注册就会跳出一个新建应用的界面，应用名字随意填写,然后选择任何组织目录(任何 Azure AD 目录 – 多租户)中的帐户，重定向url选web，填入http://localhost:53682/,最后点注册即可

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

复制应用程序（客户端）ID到记事本备用(获得了应用程序ID！) 记录ID 下面会用到

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

点击左边管理的证书和密码，点击+新客户端密码，点击添加，复制新客户端密码的值 记录这个值 下面会用到

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

点击左边管理的API权限，点击+添加权限，点击常用Microsoft API里的Microsoft Graph(就是那个蓝色水晶)， 点击委托的权限，然后在下面的条例搜索以下12个 最后点击底部添加权限

Calendars.ReadWrite    、    Contacts.ReadWrite    、    Directory.ReadWrite.All

Files.ReadWrite.All    、    MailboxSettings.ReadWrite    、    Mail.ReadWrite

Mail.Send    、    Notes.ReadWrite.All    、    People.Read.All

Sites.ReadWrite.All    、    Tasks.ReadWrite    、    User.ReadWrite.All
OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

添加完自动跳回到权限首页，点击代表授予管理员同意

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

现在 把前期准备的rclone软件掏出来
打开 rclone.exe 所在文件夹，shift+右键，在此处打开powershell，输入下面修改后的内容，回车后跳出浏览器，登入e5账号，点击接受，回到powershell窗口，看到一串东西。（MAC打开终端，一个道理）

./rclone authorize "onedrive" "应用程序(客户端)ID 去上面找，让你保存的" "应用程序密码 去上面找，让你保存的"
执行完毕后 找到 "refresh_token"：" ，从双引号开始选中到 ","expiry":2021 为止（就是refresh_token后面双引号里那一串，不要双引号），如下图，右键复制保存（获得了微软密钥）

OFFICE365 E5调用api使E5开发者续订 修复版AutoApi （不使用服务器）

GITHUB方面的准备工作
 Github项目地址

第一步，fork本项目

登陆/新建github账号，回到本项目页面，点击右上角fork本项目的代码到你自己的账号，然后你账号下会出现一个一模一样的项目，接下来的操作均在你的这个项目下进行。

第二步，新建github密钥

进入你的个人设置页面 (右上角头像 Settings，不是仓库里的 Settings)，选择 Developer settings -> Personal access tokens -> Generate new token

设置名字为 GH_TOKEN , 然后勾选repo，点击 Generate token ，最后复制保存生成的github密钥（获得了github密钥，一旦离开页面下次就看不到了！）

第三步，新建secret

依次点击页面上栏右边的 Setting -> 左栏 Secrets -> 右上 New repository secret，新建6个secret： GH_TOKEN、MS_TOKEN、CLIENT_ID、CLIENT_SECRET、CITY、EMAIL

(以下填入内容注意前后不要有空格空行)

GH_TOKEN

github密钥 (第三步获得)，例如获得的密钥是abc...xyz，则在secret页面直接粘贴进去，不用做任何修改，只需保证前后没有空格空行
MS_TOKEN

微软密钥（第二步获得的refresh_token）
CLIENT_ID

应用程序ID (第一步获得)
CLIENT_SECRET

应用程序密码 (第一步获得)
CITY

城市 (例如Beijing,自动发送天气邮件要用到)
EMAIL

收件邮箱 (自动发送天气邮件要用到)
试运行
点击上栏中间的Action进入运行日志页面，中间应该有个绿色按钮（I understand my workflow...），点击。

自动刷新后，会看到左边有三个流程，一个Run api.Read，一个Run api.Write，一个Update Token。

工作流程说明
    Run api.Write：创建系api，一天自动运行一次
    Run api.Read:  查询系api，每6小时自动运行一次
    Update Token： 微软密钥更新，每2天运行一次
这三个流程名字前面应该是都有黄色感叹号的

分别点进去，然后会看到有个黄条（this schedule was disabled......），点击 enable workflow 按钮，三个流程都要按这个！

（不确定是否都需要进行这一步，我自己做视频教程的时候发现有的。如果你没有，直接忽略并往下进行，能正常运行就可以了 ）

点击两次右上角的星星（star，就是fork按钮的隔壁）启动action

再点击上面的Action选择Run api.Read或者api.Write流程 -> build -> run api 就能看到每次的运行日志

（必需点进去build里面的run api.XXX看下，api有没有调用到位，操作有没有成功，有没有出错）

image

再点两次星星，查看是否能再次成功运行

然后点击Action里的 update token 流程 -> build -> update token ，日志里显示“微软密钥上传成功”。

同时，依次点击页面上栏右边的 Setting -> 左栏 Secrets（也就是Github方面准备的第三步的secret页面），应该能看到MS_TOKEN显示刚刚update了

（这一步是为了保证重新上传到secret的token是正确的）
