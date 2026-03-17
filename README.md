# WorkLogLite (离线工作记录)

一个尽量“轻、离线、好部署”的 Windows PC 工作记录工具。

目标约束:
- 离线使用, 不依赖网络
- 便携部署(尽量单个 exe)
- 体积控制在 15MB 以内(取决于编译方式, 建议 MSVC `/MT` 静态链接运行时)
- 不依赖 Electron/.NET/Java 等运行时框架(仅用 Win32 原生控件)

当前实现(MVP):
- 左侧月历(Month Calendar), 右侧当天条目列表 + 编辑区
- 支持自定义分类(例如: 工作/会议/目标/沟通/研发/复盘...), 并可在程序内管理分类
- 正文支持富文本编辑: 加粗/斜体/下划线/项目符号/编号/段落对齐/缩进(并保存为 RTF)
- 支持季度/年度汇总报告(生成 Markdown, 可复制/导出), 报告包含每条记录正文
- 数据本地保存为文本文件, 便于备份和迁移

## 数据存储

默认会优先使用程序同目录下的 `data/`(便携模式)。如果该目录不可写(例如放在 `Program Files`)，会自动回退到:

- `%LOCALAPPDATA%\\WorkLogLite\\data`

文件目录结构:

`data/YYYY/MM/YYYY-MM-DD.wlr`

文件格式是“可读的文本文件”(WLR1)，包含条目基础字段 + 正文纯文本(`body_plain`) + 富文本(`body_rtf_b64`, RTF 的 base64)。
说明:
- 你不需要手工编辑这些文件，建议通过程序界面操作。
- 需要迁移/备份时，直接拷贝整个 `data/` 目录即可。

分类配置:
- 在数据目录下的 `categories.txt`, 一行一个分类

长期任务:
- `tasks.wlt` (WLT2)

周期会议:
- `recurring_meetings.wlrp` (WLRP1)

## 运行 vs 编译

你日常使用时只需要运行 `WorkLogLite.exe`。

项目里的 `build.bat` 是“从源码编译生成 exe”的脚本，只在你需要自己构建程序时才用到。你现在遇到的“本地运行都难”，其实是“本地没有编译器所以编译不出来”，不是程序运行本身难。

为了避免混淆，我加了一个运行脚本:

- `run.bat`: 运行 `build\\WorkLogLite.exe`（如果 exe 不存在会提示原因）

如果你只想直接用(不关心编译):
- `portable\\WorkLogLite.exe` 是已经编译好的可执行文件，可以直接双击运行
- `WorkLogLite-portable.zip` 可以直接拷贝/发给同事解压使用

## 编译(Windows / MSVC)

需要:
- Visual Studio 或 Visual Studio Build Tools (包含 MSVC 编译器)

推荐方式:
1. 打开 “x64 Native Tools Command Prompt for VS”
2. 进入项目目录:

```bat
cd /d D:\Project\WorkLogLite
```

3. 编译:

```bat
build.bat
```

输出:
- `build\WorkLogLite.exe`

如果你希望体积更小:
- 使用 `/O1` 或 `/Os`
- 继续保持 `/MT`(静态链接)可免安装 VC 运行时

## 使用

运行 `WorkLogLite.exe` 后:
- 选中某一天即可查看/编辑该日期的条目
- “新增”创建条目, “保存”写入当天文件, “删除”移除条目
- 菜单 “报告” 可生成本季度/本年度汇总(可复制/导出)
- 菜单 “工具 -> 管理分类” 可新增/修改/删除分类(快捷键 `Ctrl+K`)
- 菜单 “工具 -> 管理长期任务” 可维护跨多日的任务，并在每天自动生成任务进展占位(快捷键 `Ctrl+T`)
- 菜单 “工具 -> 管理周期会议” 可维护例会/技术审查会等周期会议，并支持会议模板(快捷键 `Ctrl+R`)
- 菜单 “工具 -> 生成示例数据(演示)” 可生成一批符合真实工作流的示例数据，便于快速体验
- 菜单 “查看 -> 预览” 可预览 Markdown 渲染效果(当前条目或当天全部)
- 快捷键: `Ctrl+S` 保存, `Ctrl+N` 新增, `Ctrl+P` 预览当前条目, `Ctrl+Shift+P` 预览当天全部, `Delete` 删除(会二次确认), `F1` 帮助
