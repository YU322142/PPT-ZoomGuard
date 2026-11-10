# 开发与验证记录

实现为 .NET Framework COM 加载项，使用 Office 自带 Add-in Designer 类型库定义的 `_IDTExtensibility2`。类型库以 `/sysarray` 导入，并嵌入互操作类型；PowerPoint 的回调数组可能具有非零下限，不能使用普通零下标 `object[]` 替代。

## 最终实现

- 发现当前 PowerPoint 进程的 `screenClass` 放映窗口。
- 正常模式使用无按钮、无文字的输入过滤窗口，过滤多点触摸，处理单指点击、滑动和长按。
- 原生 `PPTOpticalZoom` 选区出现时，移除过滤窗口，允许 PowerPoint 处理双指缩放。
- Esc／原生右键退出缩放／换页后恢复过滤；切换到编辑窗口或其他应用时隐藏。
- 过滤窗口的黑色填充以 1/255 的透明度显示，以保留输入命中能力。它不承担显示提示的功能。
- 原生笔输入转发需要真实触摸设备继续验证；此项不视为已通过。

## 用户实机反馈与保留限制

用户已确认双指缩放屏蔽成功，同时左侧菜单栏也会被屏蔽。按照用户要求，本版本不修复该行为；公开发布时必须保留这项说明。

输入过滤窗口覆盖整个放映窗口，没有为左侧菜单栏单独设置命中穿透区域。因此不能把整体输入过滤描述为只影响双指缩放、完全不影响所有菜单控件。需要使用原生缩放时，可采用已测试的长按菜单入口。

`SetGestureConfig` 作为传统手势路径的补充；在所测 PowerPoint 2024 中，它本身不能拦住通过现代指针输入触发的缩放，因此不将接口返回成功视为功能通过。

本版本的原生窗口类名、内部退出命令与 Office 2024 2408 的实际行为有关。其他版本需要重新验证。没有进程注入、没有全局鼠标或键盘输入拦截，不修改课件文件。

## 已验证

测试环境：Windows 11，PowerPoint 2024 2408，16.0.17932.21056，64 位；使用 Windows 触摸注入接口操作临时测试演示文稿。

| 场景 | 结果 |
|---|---|
| 正常全屏放映时双指捏合 | 平铺总览未出现，保留当前页 |
| 单指点击 | 第 1 页进入第 2 页 |
| 单指向右滑动 | 第 2 页返回第 1 页 |
| 长按后手指点选原有“放大幻灯片”菜单 | 原生选区出现，过滤关闭 |
| 选定放大区域后双指展开 | PowerPoint 接收缩放，过滤保持关闭 |
| 原生右键退出缩放 | 自动重新过滤 |
| 再次双指捏合 | 平铺总览未出现 |
| 切换到本测试课件的编辑窗口 | 过滤隐藏 |
| 返回全屏放映 | 过滤恢复 |
| Esc 退出缩放 | 前序同实现路径实测恢复过滤 |

实机批注、压感、第三方批注、演示者视图、多屏和其他版本待验证。不能将模拟输入验证替代真实大屏验收。

用户后续的单项实机反馈仅更新对应项目，不视为整个实机兼容性测试矩阵已经通过。

## 重新编译

需有 Office 桌面版、.NET Framework 编译器及 Windows SDK 的 TlbImp 工具。运行 `Build.ps1`；生成的 DLL 不依赖额外的互操作 DLL 或 VSTO Runtime。

若要替换安装包中的 DLL，应同时更新 `manifest.json` 中的 SHA-256，然后关闭 PowerPoint 后重新安装。不要在已加载 DLL 的运行实例中直接替换程序集。

接口参考：

- https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshowbegin
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setgestureconfig
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getpointerinfo
