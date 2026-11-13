# 开发与验证记录

## v0.2.1：独立于输入路径的选页兜底

用户报告 v0.2.0 在大屏上仍可进入选页。补测“组件先加载、再创建放映窗口”的顺序，本机仍能过滤，尚未据此确认实机根因。

在 Connect 中增加独立保护：枚举当前应用的全部 SlideShowWindows，读取各自的 SlideNavigation.Visible；若为 true，设置为 false，并通知触摸过滤器重新锁定。关闭选页时不调用 GotoSlide，避免重新加载当前幻灯片和影响视频状态。该保护在临时缩放模式中也执行。

保护由当前进程窗口显示事件和 60 毫秒 UI 定时器调用。重入标志避免 COM 状态变化再次触发自身。它是选页显示后的关闭措施，可能闪现；PowerPoint 繁忙时定时器也可能延迟，不能承诺零帧显示或固定响应上限。菜单和快捷键手动打开的总览同样关闭。

GetStatus 新增 navigationClosed 和 navigationError；日志记录 navigation closed。导航关闭后保存当前选区可见性，避免把残留选区误认为新一次主动开启放大镜。

验证包括：先连接加载项再创建放映；通过 COM 接口直接把导航设置为可见，检查其自动关闭、页码不变；退出原生放大镜后重复直接打开并确认关闭。此外视频点击、连续多指后视频恢复、原生单指滑动和放大镜路径均回归通过。独立测试标识对应相同运行实现。真实大屏根因及 v0.2.1 的实机效果仍待确认。

## v0.2.0：保留原生单指输入

实现为 .NET Framework COM 加载项，使用 Office Add-in Designer 类型库定义的 `_IDTExtensibility2`。类型库以 `/sysarray` 导入并嵌入互操作类型，保留 PowerPoint 回调数组的非零下限兼容性。

- 在当前 PowerPoint 进程、当前 UI 线程中发现并附加到 `screenClass` 放映窗口。
- 在窗口消息实际分派时读取 `WM_POINTER*` 和 `GetPointerFrameInfo`。在队列钩子中读取完整触摸帧不足以可靠拦截，因此不使用队列钩子作为发布实现。
- 同时检查完整帧的接触数量与已落下的触点，避免两个手指同时落下时先放过一次点击。
- 出现多指后，直接调用原窗口处理过程发送 `WM_POINTERCAPTURECHANGED`，取消已经送达的单指操作。取消消息必须绕过自身过滤，否则 PowerPoint 收不到清理通知，后续视频点击可能失效。
- 消耗该多指序列，直到所有触点结束。普通单指、笔、鼠标及其拖动保持原消息路径，不重建点击，不模拟方向键。
- 对传统 `WM_GESTURE` 路径补充过滤 `GID_ZOOM`，保留其他传统手势。
- `PPTOpticalZoom` 原生选区出现后恢复多指；Esc、原生右键退出或换页时恢复限制。
- 移除旧版 `InputShield` 透明窗口。不使用全局输入钩子、不注入其他进程、不更改课件文件。

当前正常模式下会取消包含多个触摸接触点的操作，不能声称只取消捏合而保留全部其他多指手势。单指滑动由 PowerPoint 自己识别。

## 验证记录

Windows 11、64 位 PowerPoint 2024 2408，Build 16.0.17932.21056。使用自建两页临时课件，内嵌本地生成的 12 秒测试视频；通过 Windows 触摸注入接口输入真实触摸序列。

为避免在用户打开的 PowerPoint 进程中复用 CLR 已加载的旧程序集，候选实现使用独立测试类标识和程序集名加载。发布源码恢复正式 COM 标识；触摸过滤实现相同，并重新编译、校验发布 DLL。测试仅关闭自己创建的演示文稿，结束时断开测试加载项并移除临时注册。

| 场景 | 结果 |
|---|---|
| 无加载项的双指缩小基线 | 出现幻灯片平铺总览 |
| 正常放映时双指缩小 | 未出现总览，保持当前页 |
| 单指滑动 | 第 2 页返回第 1 页 |
| 单指点击内嵌视频 | 播放进度由 0 前进到约 900 毫秒 |
| 视频和空白区域双指展开 | 保持当前页 |
| 两指先后落下、先后抬起，重复操作 | 未出现总览或意外翻页 |
| 多指操作后重新点击视频 | 播放进度正常前进 |
| 原生右键菜单选择放大幻灯片、鼠标选定区域 | 自动放开双指缩放 |
| 放大状态下双指展开 | PowerPoint 收到原生双指输入 |
| 原生右键退出缩放 | 恢复锁定，无残留抑制状态 |

纯触屏长按菜单、放大区域选定、左侧菜单栏、视频暂停与进度条拖动、压感批注、第三方批注、多屏和其他版本仍需实机验证。不能以模拟输入代替真实大屏验收。

## v0.1.0 历史限制

旧版使用几乎透明的输入覆盖窗口，重新转发单指点击、滑动和长按。用户确认能屏蔽缩放，但左侧菜单栏被一并挡住，随后反馈内嵌视频点击无反应。前一次按用户要求保留并记录了菜单栏限制；v0.2.0 根据后续需求撤掉覆盖层，改用原生消息过滤。

## 重新编译

运行 `source/Build.ps1`。需安装 Office 桌面版、.NET Framework 编译器以及 Windows SDK 的 TlbImp 工具；生成的 DLL 不依赖独立互操作 DLL 或 VSTO Runtime。

运行 `scripts/Package.ps1` 自动更新安装包清单中的 DLL 与源码 SHA-256。不要直接替换已被运行中的 PowerPoint 加载的 DLL。

接口参考：

- https://learn.microsoft.com/en-us/windows/win32/inputmsg/wm-pointerdown
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getpointerframeinfo
- https://learn.microsoft.com/en-us/windows/win32/inputmsg/wm-pointercapturechanged
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setgestureconfig
