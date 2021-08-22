# WindowsBot手册
## 创建/关闭WindowsBot
### 主函数示例
````javascript
const windowsBot = require('WindowsBot');//引用WindowsBot模块
async function mian(){
	let windowsBotPort = 9999;
	let proxyServer = "127.0.0.1:888";
	//初始化WindowsBot
	let driver = await windowsBot.InitWindowsBot(windowsBotPort, proxyServer = null);
}
mian();
//主函数，初始化WindowsBot服务
//windowsBotPort 整型，WindowsBot服务端口,多开进程，端口不可重复
//proxyServer 字符串，可选参数。设置WebBot内置浏览器代理ip
````
### 关闭WindowsBot
````javascript
await driver.CloseWindowsBot();
//关闭WindowsBot服务
````
## 等待超时
````javascript
await driver.Sleep(3000);
//显示等待
//参数一 整型，等待时间，单位毫秒

await driver.SetImplicitTimeout(waitMS, intervalMS = 10);
//隐式等待
//参数一 整型，等待时间，单位毫秒
//参数二 整型，心跳间隔，单位毫秒。可选参数，默认10毫秒
````
## 操作模式
````javascript
await driver.SetCurMode(true);
//参数一 布尔类型，true前台,false后台,WindowsBot默认前台操作
````
## 查找窗口句柄
````javascript
await driver.FindWindowHandle("className", "windowName");
//查找窗口句柄
//参数一 字符串型，窗口类名
//参数二 字符串型，窗口标题
//成功返回句柄,失败返回null

await driver.FindWindowsHandle("className", "windowName");
//查找窗口句柄数组
//参数一 字符串型，窗口类名
//参数二 字符串型，窗口标题
//成功返回句柄数组,失败返回null

await driver.FindSubWindowHandle("hwndParent", "className", "windowName");
//查找子窗口句柄
//参数一 整型，父窗口句柄
//参数二 字符串型，窗口类名
//参数三 字符串型，窗口标题
//成功返回句柄,失败返回null

await driver.FindWebBotHandle();
//查找WebBot内置浏览器句柄
//成功返回句柄,失败返回null
````
## 鼠标操作
````javascript
await driver.MoveMouse(hwnd, x, y);
//移动鼠标
//参数一 整型，窗口句柄
//参数二 整型，x坐标
//参数三 整型，y坐标

await driver.RollMouse(hwnd, x, y, dwData);
//滚动鼠标
//参数一 整型，窗口句柄
//参数二 整型，x坐标
//参数三 整型，y坐标
//参数四 整型，鼠标滚动次数,负数下滚鼠标,正数上滚鼠标

await driver.ClickMouse(hwnd, x, y, msg);
//点击鼠标,后台模式下,控件独句柄与hwnd不一致时会失败
//参数一 整型，窗口句柄
//参数二 整型，x坐标
//参数三 整型，y坐标
//参数四，单击左键:1 单击右键:2 按下左键:3 弹起左键:4 按下右键:5 弹起右键:6 双击左键:7 双击右键:8
````
## 键盘操作
````javascript
await driver.InputText(hwnd, text);
//输入文本,杀毒软件可能会拦截
//参数一 整型，窗口句柄
//参数二 字符串型，输入内容

await driver.InputKeyborad(bVk, msg);
//输入键盘VK值
//参数一 整型，VK键值
//参数二 整型，按下弹起:1 按下:2 弹起:3
````
## 截图保存
````javascript
await driver.SaveScreenshot(hwnd, imagePath, x = 0, y =0, width = 0, height = 0);
//后台截图保存
//参数一 整型，窗口句柄
//参数二 字符串型，图片保存的路径，png格式
//参数三 整型，可选参数,顶点x坐标
//参数四 整型，可选参数,顶点y坐标
//参数五 整型，可选参数,图片宽度
//参数六 整型，可选参数,图片高度
//成功返回true,失败返回false

await driver.FrontSaveScreenshot(imagePath, x = 0, y =0, width = 0, height = 0);
//前台截图保存,当SaveScreenshot截图失效时,可用此函数
//参数一 字符串型，图片保存的路径，png格式
//参数二 整型，可选参数,顶点x坐标
//参数三 整型，可选参数,顶点y坐标
//参数四 整型，可选参数,图片宽度
//参数五 整型，可选参数,图片高度
//成功返回true,失败返回false
````
## 图色定位
````javascript
await driver.FindColorCount(imagePath, R, G, B);
//查找图片颜色点的数量和坐标
//参数一 字符串型，图片的路径
//参数二 整型，红
//参数三 整型，绿
//参数四 整型，蓝
//返回指定rgb点的数量和坐标

await driver.FindColor(hwnd, imagePath, R, G, B, index = 0);
//查找指定色值在图片上的位置
//参数一 整型，窗口句柄
//参数二 字符串型，图片的路径
//参数三 整型，红
//参数四 整型，绿
//参数五 整型，蓝
//参数六 整型，可选参数,指定出现的次数,默认第一次查找到的坐标
//成功返回指定rgb点的坐标,失败返回 {x:-1, y:-1}

await driver.FindImage(hwnd, sourceImagePath, destImagePath, sim);
//查找图片,大图找小图
//参数一 整型，窗口句柄
//参数二 字符串型，大图路径
//参数三 字符串型，小图路径
//参数四 浮点型，相似度0.0-1.0
//成功返回指定rgb点的坐标,失败返回 {x:-1, y:-1}
````
## 系统剪切板
````javascript
await driver.SetClipboardText(text);
//设置剪切板文本
//参数一 字符串型，设置的文本

await driver.GetClipboardText();
//获取剪切板文本
//返回剪切板文本
````
## 启动程序
````javascript
await driver.StartProcess(commandLine, showWindow = true, isWait = false);
//启动指定程序
//参数一 字符串型，启动命令行
//参数二 布尔型，是否显示窗口。可选参数,默认显示窗口
//参数三 布尔型，是否等待程序结束。可选参数,默认不等待
//成功返回true,失败返回false
````
## 下载文件
````javascript
await driver.DownloadFile(url, filePath, isWait);
//启动指定程序
//参数一 字符串型，文件地址
//参数二 字符串型，文件保存的路径
//参数三 布尔型，是否等待.为true时,等待下载完成
````
## ocr系统
````javascript
await driver.InitOcr(appId, apiKey, secretKey);
//初始化百度ocr
//参数一 字符串型，百度ocr提供
//参数二 字符串型，百度ocr提供
//参数三 字符串型，百度ocr提供

await driver.GetImageWords(imagePath, ocrType);
//百度ocr获取图片上的文字
//参数一 字符串型，图片路径
//参数二 整型，识别类型,通用文字识别:1 高精度版:2 网络图片文字识别:3 数字识别:4 手写文字识别:5 二维码识别:6
//成功返回获取的文字

await driver.FindImageWords(hwnd, imagePath, words, index = 0);
//百度ocr获取图片文字坐标
//参数一 整型，图片所属窗口的句柄,手机图片识别一般为null
//参数二 字符串型，图片路径
//参数三 字符串型，识别的文字
//参数四 整型，可选参数。文字出现的次数,默认为0,查找第一次出现的坐标
//成功返回文字坐标,失败返回{x:-1, y:-1}
````
## 验证码识别系统
````javascript
await driver.InitCheckCode(userName, passWord);
//初始化验证码系统
//参数一 字符串型，用户名,第三方平台提供
//参数二 字符串型，用户密码,第三方平台提供

await driver.GetCheckCode(imagePath, codeType);
//获取验证码
//参数一 字符串型，验证码图片路径
//参数二 字符串型，验证码类型,参考第三方平台
//成功返回验证码,失败返回错误信息

await driver.SendErrorCode();
//发送验证码识出错
//成功返回"true",失败返回错误信息
````
## 窗口元素操作
注意：窗口元素操作未默认释放，必须手动释放，否则内存泄漏！
### 查找元素
````javascript
await driver.FindRootElement(hwnd);
//查找窗口根元素
//参数一 整型，窗口句柄
//成功返回元素对象，失败返回null

await driver.FindFocusElement();
//查找焦点元素
//成功返回元素对象，失败返回null

await driver.FindElementByAutomationId(rootElement, automationId);
//通过automationId查找元素
//参数一 对象，根元素
//参数二 字符串，要查找元素的automationId值
//成功返回元素对象，失败返回null

await driver.FindElementsByAutomationId(rootElement, automationId);
//通过automationId查找元素数组
//参数一 对象，根元素
//参数二 字符串，要查找元素的automationId值
//成功返回元素对象数组，失败返回null

await driver.FindElementByName(rootElement, name);
//通过name查找元素
//参数一 对象，根元素
//参数二 字符串，要查找元素的name值
//成功返回元素对象，失败返回null

await driver.FindElementsByName(rootElement, name);
//通过name查找元素数组
//参数一 对象，根元素
//参数二 字符串，要查找元素的name值
//成功返回元素对象数组，失败返回null

await driver.FindElementByClassName(rootElement, className);
//通过className查找元素
//参数一 对象，根元素
//参数二 字符串，要查找元素的className值
//成功返回元素对象，失败返回null

await driver.FindElementsByClassName(rootElement, className);
//通过className查找元素数组
//参数一 对象，根元素
//参数二 字符串，要查找元素的className值
//成功返回元素对象数组，失败返回null

await driver.FindElementByControlType(rootElement, controlType);
//通过controlType查找元素
//参数一 对象，根元素
//参数二 字符串，要查找元素的controlType值
//成功返回元素对象，失败返回null

await driver.FindElementsByControlType(rootElement, controlType);
//通过controlType查找元素数组
//参数一 对象，根元素
//参数二 字符串，要查找元素的controlType值
//成功返回元素对象数组，失败返回null

await driver.FindElementByXpath(rootElement, xpath);
//通过xpath查找元素
//参数一 对象，根元素
//参数二 字符串，要查找元素的xpath值
//成功返回元素对象，失败返回null

await driver.FindParentElement(currentElement, isRelease = true);
//查找父元素
//参数一 对象，当前元素
//参数二 布尔型，可选参数。是否释放当前元素
//currentElement不再使用时必须释放,否则内存泄漏.默认为true释放当前元素
//成功返回元素对象，失败返回null

await driver.FindNextSiblingElement(currentElement, isRelease = true);
//查找下一个兄弟元素
//参数一 对象，当前元素
//参数二 布尔型，可选参数。是否释放当前元素
//currentElement不再使用时必须释放,否则内存泄漏.默认为true释放当前元素
//成功返回元素对象，失败返回null

await driver.FindPreviousSiblingElement(currentElement, isRelease = true);
//查找上一个兄弟元素
//参数一 对象，当前元素
//参数二 布尔型，可选参数。是否释放当前元素
//currentElement不再使用时必须释放,否则内存泄漏.默认为true释放当前元素
//成功返回元素对象，失败返回null

await driver.FindFirstChildElement(currentElement, isRelease = true);
//查找第一个子元素
//参数一 对象，当前元素
//参数二 布尔型，可选参数。是否释放当前元素
//currentElement不再使用时必须释放,否则内存泄漏.默认为true释放当前元素
//成功返回元素对象，失败返回null

await driver.FindLastChildElement(currentElement, isRelease = true);
//查找最后一个子元素
//参数一 对象，当前元素
//参数二 布尔型，可选参数。是否释放当前元素
//currentElement不再使用时必须释放,否则内存泄漏.默认为true释放当前元素
//成功返回元素对象，失败返回null
````
### 元素方法
````javascript
await driver.GetElementName(element, isRelease = true);
//获取指定元素name属性值
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回元素name属性值

await driver.GetElementBoundingRectangle(element, isRelease = true);
//获取指定元素矩形大小
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回元素矩形大小,失败返回{left:-1, top:-1, right:-1, bottom:-1}

await driver.GetElementValueValue(element, isRelease = true);
//获取指定元素ValueValue属性值
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回元素ValueValue属性值

await driver.GetControlType(element, isRelease = true);
//获取指定元素ControlType属性值
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回元素ControlType属性值

await driver.SetElementValue(element, value, isRelease = true);
//设置指定元素value
//参数一 对象，指定元素
//参数二 字符串型，设置的值
//参数三 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false

await driver.SetElementValueEx(element, value, isRelease = true);
//通过元素窗口句柄设置指定元素value
//参数一 对象，指定元素
//参数二 字符串型，设置的值
//参数三 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false.

await driver.InvokeElement(element, isRelease = true);
//执行元素默认操作,一般用作点击
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false

await driver.SetElementFocus(element, isRelease = true);
//设置指定元素作为焦点
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false

await driver.SetElementScroll(element, horizontalPercent, verticalPercent, isRelease = true);
//设置指定元素滚动条
//参数一 对象，指定元素
//参数二 浮点型，水平百分比0.0-1.0,为-1时不滚动
//参数三 浮点型，垂直百分比0.0-1.0,为-1时不滚动
//参数四 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false

await driver.ClickElement(element, msg, isRelease = true);
//通过元素窗口句柄点击控件
//参数一 对象，指定元素
//参数二 整型，单击左键:1 单击右键:2 按下左键:3 弹起左键:4 按下右键:5 弹起右键:6 双击左键:7 双击右键:8
//参数三 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false.后台模式无法点击非客户端区域

await driver.CloseWindow(element, isRelease = true);
//关闭指定元素窗口
//参数一 对象，指定元素
//参数二 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false.

await driver.SetWindowState(element, state, isRelease = true);
//设置指定元素窗口状态
//参数一 对象，指定元素
//参数二 整型，正常:0 最大化:1 最小化:2
//参数三 布尔型，可选参数。element不再使用时必须释放,否则内存泄漏.默认为true释放
//成功返回true,失败返回false.

await driver.ReleaseElement(element);
//手动释放指定元素,windows元素操作必须要释放元素,否则内存泄漏
//参数一 对象，指定元素
//成功返回true,失败返回false.

await driver.ReleaseElements(elements);
//手动释放指定元素数组,元素数组不再使用时,必须调用此函数手动释放,否则大量内存泄漏
//参数一 对象，指定元素数组
````
## 内置浏览器(WebBot)
### WebBot操作
````javascript
await driver.openWebBot(url);
//打开内置浏览器WebBot
//参数一 字符串型，指定网址

await driver.switchPage(index);
//切换WebBot页面
//参数一 整型，索引号,0为起首页

await driver.closePage(index);
//关闭WebBot页面
//参数一 整型，索引号,0为起首页
````

### 获取元素
````javascript
await driver.getElementById(id, iframe = null);
//通过id获取元素
//参数一 字符串型，元素id值
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象,失败返回null

await driver.getElementsByClassName(classNames, iframe = null);
//通过class获取元素数组
//参数一 字符串型，元素class值
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象数组,失败返回null

await driver.getElementsByName(elementName, iframe = null);
//通过name获取元素数组
//参数一 字符串型，元素name值
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象数组,失败返回null

await driver.getElementsByTagName(localName, iframe = null);
//通过tagName获取元素数组
//参数一 字符串型，元素tagName值
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象数组,失败返回null

await driver.getElementByXpath(xpath, iframe = null);
//通过xpath获取元素
//参数一 字符串型，元素xpath路径
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象,失败返回null

await driver.querySelector(selector, iframe = null);
//通过ccs selector获取元素
//参数一 字符串型，元素selector路径
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象,失败返回null

await driver.querySelectorAll(selectors, iframe = null);
//通过ccs selector获取元素数组
//参数一 字符串型，元素selector路径
//参数二 对象，可选参数,元素所属的iframe
//成功返回元素对象数组,失败返回null
````
### 元素方法
````javascript
await driver.clickElement(element);
//点击元素
//参数一 对象，要点击的元素

await driver.clickElementEx(element, iframe);
//点击元素(windows api方式点击)
//参数一 对象，要点击的元素。 参数二 对象，可选参数,元素所属的iframe

await driver.setValue(element, value);
//设置元素value
//参数一 对象，设置的元素
//参数二 字符串型，设置的值

await driver.getValue(element);
//获取元素value
//参数一 对象，指定的元素
//成功返回value值

await driver.setTextConten(element, value);
//设置元素textConten
//参数一 对象，设置的元素
//参数二 字符串型，设置的值

await driver.getTextConten(element);
//获取元素textConten
//参数一 对象，指定的元素
//成功返回textConten值

await driver.setOuterText(element, value);
//设置元素outerText
//参数一 对象，设置的元素
//参数二 字符串型，设置的值

await driver.getOuterText(element);
//获取元素outerText
//参数一 对象，指定的元素
//成功返回outerText值
await driver.setOuterHTML(element, value);
//设置元素setOuterHTML
//参数一 对象，设置的元素
//参数二 字符串型，设置的值

await driver.getOuterHTML(element);
//获取元素outerText
//参数一 对象，指定的元素
//成功返回outerHTML值

await driver.setInnerText(element, value);
//设置元素setInnerText
//参数一 对象，设置的元素
//参数二 字符串型，设置的值

await driver.getInnerText(element);
//获取元素innerHTML
//参数一 对象，指定的元素
//成功返回innerHTML值

await driver.setAttribute(element, value, name);
//设置元素属性
//参数一 对象，设置的元素
//参数二 字符串型，设置的值
//参数三 字符串型，属性名称

await driver.getAttribute(element, name);
//获取元素属性值
//参数一 对象，指定的元素
//参数二 字符串型，属性名称
//成功返回name对应的属性值

await driver.setSelect(element, checkValue);
//选择下拉框值
//参数一 对象，指定的元素
//参数二 字符串型，选中的值
//成功返回name对应的属性值

await driver.getBoundingClientRect(element);
//获取元素矩形大小
//参数一 对象，指定的元素
//成功返矩形大小，失败返回null
````
### 注入JavaScript
````javascript
await driver.executeScript(scriptCode);
//注入JavaScript代码在v8执行
//参数一 字符串型，JavaScript代码
````
### node与V8数据共享
````javascript
await driver.nodePipeScript(shareData = 'null');
//WebBot V8有同名称扩展函数，可在executeScript调用。shareData更新需要10-100毫秒
//参数一 字符串型，设置共享数据值,可选参数。
//返回node/v8 设置的shareData值。
````
### Cookie操作
````javascript
await driver.ExportCookie(filePath, url);
//导出指定url的cookie
//参数一 字符串型，cookie文件路径
//参数二 字符串型，url 格式：http://www.ai-bot.net
//成功返回true，失败返回false

await driver.ImportCookie(filePath, url);
//导入cookie到指定url网页
//参数一 字符串型，cookie文件路径
//参数二 字符串型，url 格式：http://www.ai-bot.net
//成功返回true，失败返回false

await driver.ClearCookie(filePath, url);
//清除指定url网页内的cookie， 不会删除本地cookie文件
//参数一 字符串型，cookie文件路径
//参数二 字符串型，url 格式：http://www.ai-bot.net
//成功返回true，失败返回false
````
## Word文档
````javascript
await driver.wordCreate(options = {type: 'docx'});
/*创建Wrod文档
参数一 对象，Word文档格式，可选参数，options列表:
	type（字符串）'docx' 必要参数值
	author（字符串）-文档的作者
	creator（字符串）-别名。文档的作者
 	description'（字符串）-文档的属性注释
 	keywords（字符串）-文档的关键字
 	orientation（字符串）-横向'landscape'或纵向'portrait'。默认值为'portrait'
 	pageMargins（对象）-设置文档页边距。默认值为{top: 1800, right: 1440, bottom: 1800, left: 1440}
 	pageSize（字符串|对象）-设置文档页面大小。默认值为A4（支持值：'A4', 'A3', 'letterpaper'）。或使用{width: 11906, height: 16838}设置自定义尺寸
 	subject （字符串）-文档的主题
 	title（字符串）-文档的标题
 	columns（整型）-每页中的列数。默认值为1列。
返回word对象*/

await driver.wordAddPage(docxObject);
//添加word分页
//参数一 对象，word对象

await driver.wordCreateTable(docxObject, table, tableStyle);
//创建word表格
//参数一 对象，word对象
//参数二 对象，表格
//参数三 对象，表格样式

await driver.wordAddHeader(docxObject, text, options = '');
//添加页眉
//参数一 对象，word对象
//参数二 字符串，要添加的文本
//参数三 字符串，可选参数，要更改的页面 'even'更改偶数页。'first'仅更改首页

await driver.wordAddFooter(docxObject, text, options = '');
//添加页脚
//参数一 对象，word对象
//参数二 字符串，要添加的文本
//参数三 字符串，可选参数，要更改的页面 'even'更改偶数页。'first'仅更改首页

await driver.wordCreateParagraph(docxObject, options = {align: 'left'});
/*创建段落
  参数一 对象，word对象
  参数二 对象， 可选参数，段落格式，options列表:
	align（字符串）-水平对齐，可以是'left'（认），'right'，'center'或'justify'
	textAlignment（字符串）-垂直对齐方式，'center', 'top', 'bottom'或'baseline'
	indentLeft（数字）- 向左缩进
	indentFirstLine（数字）- 缩进第一行
	backline（字符串）-颜色代码，例如：'ffffff'（白色）或'000000'黑色）
成功返回段落对象*/

await driver.wordCreateListOfDots(docxObject, options = {align:'left'});
//创建无作序号的段落
//参数一 对象，word对象
//参数二 对象， 可选参数，段落格式，options列表:参考wordCreateParagraph
//成功返回段落对象

await driver.wordCreateListOfNumbers(docxObject, options = {align: 'left'});
//创建有序号数字的段落
//参数一 对象，word对象
//参数二 对象， 可选参数，段落格式，options列表:参考wordCreateParagraph
//成功返回段落对象

await driver.wordCreateNestedUnOrderedList(docxObject, options = {"level":2});
//创建有等级的无序号段落
//参数一 对象，word对象
//参数二 对象， 可选参数，段落格式，options列表:参考wordCreateParagraph
//成功返回段落对象

await driver.wordCreateNestedOrderedList(docxObject, options = {"level":2});
//创建有等级的序号段落
//参数一 对象，word对象
//参数二 对象， 可选参数，段落格式，options列表:参考wordCreateParagraph
//成功返回段落对象

await driver.wordAddText(paragraphObject, text, sytle = {});
/*将文本添加到段落
  参数一 对象，段落对象对象
  参数二 字符串， 要添加的文本
  参数三 对象，可选参数，文本样式，sytle列表:
	back（字符串）-背景颜色代码，例如：'ffffff'（白色）或'000000'黑色）
	shdType（字符串）-要使用的可选模式代码：'clear'（无模
式），'pct10'，'pct12'，'pct15'，'diagCross'，'diagStripe'，'horzCross'，'horzStripe'，'nil' ，'thinDiagCross'，'solid'等
	shdColor（字符串）-模式的前部颜色（与shdType一起使用）
	bold（布尔）-为true时文本变为粗体
	border（字符串）-边框类型：'single'，'dashDotStroked'，'dashed'，
'dashSmallGap'，'dotDash'，'dotDotDash'，'dotted'，'double'，'thick'
	color（字符串）-字体颜色代码，例如：“ ffffff”（白色）或“ 000000”（黑色）
	italic（布尔值）-设置为斜体时为true
	underline（布尔值）-为true表示要添加下划线
	font_face（字符串）-要使用的字体,例如：'Arial'
	font_face_east（字符串）-高级设置：用于东亚的字体。您还必须设置font_face
	font_face_cs（字符串）-高级设置：要使用的字体（cs）。您还必须设置font_face
	font_face_h（字符串）-高级设置：要使用的字体（hAnsi）。您还必须设置font_face
	font_hint（字符串）-可选。'ascii'（默认），'eastAsia'，'cs'或'hAnsi'
	font_size（数字）-以磅为单位的字体大小
	rtl（布尔值）-将其添加到rtl语言的任何文本中
	highlight (字符串) - 高亮颜色。'black', 'blue', 'cyan', 'darkBlue', 'darkCyan','darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white'，'yellow'
	strikethrough（布尔值）-添加删除线时为true。
	superscript（布尔值）-如果可以使用较小的尺寸，则为true可以将此次运行中的文本降低到基线以下，并将其更改为较小的尺寸。
 * subscript（布尔值）-为true时，如果可以使用较小的尺寸，则将运行中的文本升高到基线以上，并将其更改为较小的尺寸
	link（字符串）-超连接
	hyperlink（字符串）向书签的超链接*/
	
await driver.wordAddImage(paragraphObject, imagePath, area = {});
//将图像添加到段落
//参数一 对象，段落对象
//参数二 字符串， 图片路径
//参数三 对象， 可选参数, 图片宽高 {cx: 300, cy: 200}

await driver.wordBreak(paragraphObject);
//段落换行
//参数一 对象，段落对象

await driver.wordAddHorizontalLine(paragraphObject);
//添加一条水平线,看上去与换行相似？
//参数一 对象，段落对象

await driver.wordStartBookmark(paragraphObject, markName);
//添加书签
//参数一 对象，段落对象
//参数二 字符串，标签名

await driver.wordEndBookmark(paragraphObject);
//书签结束位置
//参数一 对象，段落对象

await driver.wordSave(docxObject, filePath);
//保存word文件
//参数一 对象，word对象
//参数二 字符串，保存的路径

await driver.wordReadText(filePath);
//读取word内容
//参数一 字符串，保存的路径
//成功返回读取到的内容，失败返回null
````
## Excel文档
````javascript
await driver.OpenExcel(excelPath);
//打开excel文档
//参数一 字符串，excle路径
//成功返回excel对象，失败返回null

await driver.OpenExcelSheet(excelObject, sheetName);
//打开excel表格
//参数一 对象，excel对象
//参数二 字符串，表名
//成功返回sheet对象，失败返回null

await driver.SaveExcel(excelObject);
//保存excel文档
//参数一 对象，excel对象
//成功返回true，失败返回false

await driver.WriteExcelNum(sheetObject, row, col, value);
//写入数字到excel表格
//参数一 对象，sheet对象
//参数二 整型，行
//参数三 整型，列
//参数四 数字型，写入的值
//成功返回true，失败返回false

await driver.WriteExcelStr(sheetObject, row, col, strValue);
//写入字串到excel表格
//参数一 对象，sheet对象
//参数二 整型，行
//参数三 整型，列
//参数四 字符串，写入的值
//成功返回true，失败返回false

await driver.ReadExcelNum(sheetObject, row, col);
//读取excel表格数字
//参数一 对象，sheet对象
//参数二 整型，行
//参数三 整型，列
//返回读取到的数字

await driver.ReadExcelStr(sheetObject, row, col);
//读取excel表格字串
//参数一 对象，sheet对象
//参数二 整型，行
//参数三 整型，列
//返回读取到的字符
````
## 邮件收发
````javascript
await driver.SendMail(smtpServer, smtpPort, authCode, sendMailer, sendName, recvMailer, subject, text, html = null, attachments = null);
/*发送邮件
  参数一 字符串，SMTP服务地址
  参数二 整型，SMTP服务端号
  参数三 字符串，第三方邮箱平台授权码
  参数四 字符串，发送人邮箱
  参数五 字符串，发送人名称
  参数六 字符串，接收人邮箱,多人逗号分开
  参数七 字符串，邮箱主题
  参数八 字符串，发送的正文 文本格式
  参数九 字符串，发送的正文 html格式，可选参数，默认为null
  参数十 字符串， 发送的附件,可选参数，默认为null 示例：
  let attachments = [
      { // utf-8字符串作为附件
        filename: 'text1.txt',
        content: 'hello world!'
      },
      { //磁盘上的文件作为附件
        filename: 'text3.txt',
        path: '/path/to/file.txt'
      },
      { //文件名和内容类型是从路径生成
        path: '/path/to/file.txt'
      },
      { //使用URL作为附件
        filename: 'WindowsDoc.pdf',
        path: 'http://www.ai-bot.net/WindowsDoc.pdf'
      }
  ]
成功返回"邮件发送成功!"，失败返回错误信息*/

await driver.GetMail(imapServer, imapPort, authCode, mailAccount,criteria);
//获取邮件
//参数一 字符串，IMAP服务地址
//参数二 整型，IMAP服务端号
//参数三 字符串，第三方邮箱平台授权码
//参数四 字符串，邮箱账号
//参数五 字符串，获取哪类邮件，所有邮件'ALL' 未读邮件'UNSEEN' 已读邮件'SEEN' 第1-3封'1:3'
//成功返回邮件信息对象，失败返回错误信息
````
## PDF转PNG图片
````javascript
await driver.PdfToImage(pdfPath, quality = 100);
//获取邮件
//参数一 字符串，pdf路径
//参数二 整型，图片品质，可选参数，默认100
//成功返回true,失败返回false.图片生成在当前目录下，以pdf文件命名，多页Pdf会生成多个图片
````
## 获取主板识别码
````javascript
await driver.GetBoisId();
//一般用作脚本授权
//成功返回BoisId,失败返回null
````
## 语音合成
````javascript
await driver.Tts(saveFile, appkey, keyId, keySecret, text, voice = "ruoxi",ttsArgs = {"format":"wav", "sampleRate":24000, "volume":50, "speechRate":0, "pitchRate":0});
/*支持一次性合成300字符以内的文字，超过300个字符的内容会被截断。
参数1 字符串类型，音频保存文件路径
参数2 字符串类型，appkey
参数3 字符串类型，accessKeyId
参数4 字符串类型，accessKeySecret
参数5 字符串类型，待合成的文本，需要为UTF-8编码，支持SSML，例如：<speak>我<break
 time="2000ms"/>停顿了两秒</speak> time范围[50, 10000]的整数。
参数6 字符串类型，发音人，可选参数，默认"ruoxi"。支持以下音种：
 ruoxi(温柔女声) siqi(温柔女声) sijia(标准女声) sicheng(标准男声)ninger(标准女声)
 ruilin(标准女声) siyue(温柔女声) xiaomei(甜美女声) yina(浙普女声) sijing(严厉女声)
 sitong(儿童音) xiaobei(萝莉女声) wendy(英音女声) william(英音男声) olivia(英音女声) shanshan(粤语女声)
参数7 JSON类型,可选，默认值为：
 {"format":"wav", "sampleRate":24000, "volume":50, "speechRate":0, "pitchRate":0}
 "format"：音频编码格式，支持pcm/wav/mp3格式，默认"wav"。
 "sampleRate"：音频采样率，支持24000/16000/8000(Hz)，默认24000。
 "volume"：音量，范围是0~100，默认50。
 "speechRate"：语速，范围是-500~500，默认是0。
 "pitchRate"：语调，范围是-500~500，默认是0。
成功返回true,失败返回false*/

await driver.LongTts(saveFile, appkey, keyId, keySecret, text, voice = "zhiqi", ttsArgs = {"format":"wav", "sampleRate":48000, "volume":50, "speechRate":0, "pitchRate":0});
/*长语音合成，支持一次性合成最高10万字符，支持精品音质
参数1 字符串类型，音频保存文件路径
参数2 字符串类型，appkey
参数3 字符串类型，accessKeyId
参数4 字符串类型，accessKeySecret
参数5 字符串类型，待合成的文本，需要为UTF-8编码，支持SSML，例如：<speak>我<break
 time="2000ms"/>停顿了两秒</speak> time范围[50, 10000]的整数。
参数6 字符串类型，发音人，可选参数，默认"zhiqi"。支持以下音种：
超高清场景
zhiqi(温柔女声) zhichu(舌尖男声) zhixiang(磁性男声) zhijia(标准女声) zhinan(广告男声)
zhiqian(资讯女声) zhiru(新闻女声) zhide(新闻男声) zhifei(激昂解说) zhilun(悬疑解说)
zhiwei(萝莉女声)
文学场景
aiyuan(知心姐姐) aiying(软萌童声) aixiang(磁性男声) aimo(情感男声) aiye(青年男声)
aiting(电台女声) aifan(情感女声) aide(新闻男声) ainan(广告男声) aihao(资讯男声)
aiming(诙谐男声) aixiao(资讯女声) aichu(舌尖男声) aiqian(资讯女声) aishu(资讯男声)
airu(新闻女声) 除以上值外，还支持语音合成所有的发音人
参数7 JSON类型,可选，默认值为：
 {"format":"wav", "sampleRate":48000, "volume":50, "speechRate":0, "pitchRate":0}
 "format"：音频编码格式，支持pcm/wav/mp3格式，默认"wav"。
 "sampleRate"：音频采样率，支持48000/24000/16000/8000(Hz)，默认48000。
 "volume"：音量，范围是0~100，默认50。
 "speechRate"：语速，范围是-500~500，默认是0。
 "pitchRate"：语调，范围是-500~500，默认是0。
成功返回true,失败返回false*/
````
## 坐标转换
````javascript
await driver.ClientToScreen(hwnd, x, y);
//客户端坐标转换屏幕坐标(忽略标题栏的高度)
//参数1整型，窗口句柄
//参数2整型，x坐标
//参数3整型，y坐标
//返回客户端对应的屏幕坐标{x:number, y:number}
````
## 视频流操作
````javascript
await driver.YoutubeDownload(url, filePath, maxSize, options);
/*视频流资源下载
参数1 字符串，视频资源链接
参数2 字符串，保存的mp4文件
参数3 整型，最大分辨率（高度） 1920
参数4 JSON对象，
{"subLang":'zh-Hans', "thumbnail":true, "proxy": 'https://127.0.0.1:19180/',  "embed": false}
 "subLang" 视频字幕 zh-Hans中文
 "thumbnail" 是否下载视频封面
 "proxy" 代理服务器
 "embed" 是否嵌入字幕*/
 
await driver.MergeVideo(srcVideos, destVideo);
//合并视频
//参数一 字符串数组，要合并的视频文件集合
//参数二 字符串，合并完成，要保存的文件名

await driver.SplitVideo(srcVideo, destVideo, startTime, endTime);
//分割视频
//参数一 字符串，视频源文件
//参数二 字符串，分割完成，要保存的文件名。格式必须和 srcVideo 一致
//参数三 字符串，分割起始时间
//参数四 字符串，分割 持续/结束时间。"23.18"(23.18秒)为持续时间，0:0:50则为结束时间

await driver.ExportAudio(videoFile, audioFile);
//导出音频
//参数一 字符串，视频文件
//参数二 字符串，导出的音频文件

await driver.ExportVideo(videoFile, naVideoFile);
//导出视频(无声音)
//参数一 字符串，视频文件
//参数二 字符串，导出的视频文件(无声音)，导出格式必须和 videoFile 一致

await driver.ConvertVideo(srcVideo, destVideo);
//转换视频格式
//参数一 字符串，源视频
//参数二 字符串，转换后的视频

await driver.ConvertAudio(srcAudio, destAudio);
//转换音频格式
//参数一 字符串，源音频
//参数二 字符串，转换后的音频

await driver.ConvertSubtitle(srcSubtitle, destSubtitle);
//转换字幕格式
//参数一 字符串，源字幕
//参数二 字符串，转换后的字幕

await driver.EmbedSubtitle(srcVideo, subtitleFile, destVideo);
//视频嵌入字幕
//参数一 字符串，要嵌入字幕的视频文件
//参数二 字符串，字幕文件
//参数三 字符串，嵌入字幕完成的文件

await driver.VideoToImage(srcVideo, destImage, time);
//视频提取图片
//参数一 字符串，视频文件
//参数二 字符串，要提取的图片
//参数三 字符串，截图时间点 00:01:20或者80.09

await driver.DelogoVideo(srcVideo, destVideo, rect);
//视频去水印
//参数一 字符串，源视频文件
//参数二 字符串，去水印后的视频文件
//参数三 JSON对象，去水印的矩形位置 {x:10, y:20, width:100, height:200}

await driver.AdlogoVideo(srcVideo, destVideo, logoImage, point);
//视频去水印
//参数一 字符串，源视频文件
//参数二 字符串，添加印后的视频文件
//参数三 字符串，水印图片
//参数四 JSON对象，水印摆放的位置{x:10, y:20}
````
