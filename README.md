# WorkLogLite

`WorkLogLite` 是一个原生 `Win32` 的离线工作日志工具，定位是单机、轻量、可便携分发。

当前版本重点解决的是：
- 按天记录工作内容
- 富文本编辑与预览
- 本地材料目录管理
- 常用办公小工具集成
- 数据导入、导出、清空与恢复

## 当前功能

### 日志与编辑

- 左侧日历按天切换
- 右侧条目列表查看当天记录
- 支持分类、日期范围、状态、标题、正文
- 正文使用 `RichEdit`，支持加粗、斜体、下划线、编号、对齐等基础格式
- 支持插入“时间分隔”行，格式为 `----YYYYMMDD----`
- 支持 `Ctrl + 鼠标滚轮` 调整正文显示缩放，`Ctrl + 0` 恢复 `100%`
- 支持编辑区最大化
- 支持 `Ctrl` 多选条目并批量删除
- 条目列表支持按开始时间、分类、状态排序

### 工作区

左下工作区集成了常用本地工具：

- 截图
- 常用回复
- 复制时间戳
- 专注计时
- 单位/带宽计算器
- 数据目录
- 截图目录
- 日志目录
- 便签
- 取色器
- 纯贴
- 录系统声
- 录麦克风

### 录音

- 录系统声和录麦克风分开支持
- 保留保存对话框
- 默认文件名按时间生成
- 最长录制 `5` 小时，超时自动停止
- 输出为系统兼容性更好的 `WAV`

### 数据管理

- 导出全部数据
- 导入全部数据
- 清空全部数据并自动备份
- 对错误导入源、危险目录位置有显式拦截

### 报表

- Markdown 汇总
- 按分类汇总
- CSV 导出

## 数据目录

程序优先使用可执行文件旁边的 `data/` 目录。

如果当前目录不可写，会回退到：

`%LOCALAPPDATA%\WorkLogLite\data`

主要文件格式：

- `data/YYYY/MM/YYYY-MM-DD.wlr`
- `data/categories.txt`
- `data/tasks.wlt`
- `data/recurring_meetings.wlrp`

说明：
- 所有数据都保存在本地文本文件中，便于备份和迁移。
- 当前版本已取消密码功能，不再有登录或改密流程。

## 运行

直接运行：

- `build\WorkLogLite.exe`
- 或便携版 `portable\WorkLogLite.exe`

如果你只需要分发成品，优先使用：

- `WorkLogLite-portable.zip`

## 构建

需要 `MSVC` 或仓库内准备好的 `MinGW` 工具链。

推荐在 Windows 下执行：

```bat
build.bat
```

输出：

```text
build\WorkLogLite.exe
```

## 测试

执行：

```bat
test.bat
```

当前 smoke tests 覆盖：

- 分类读写
- 日记录 roundtrip
- 富文本字段持久化
- 材料路径持久化
- 非法日期范围拒绝
- 旧格式兼容读取
- 任务数据 roundtrip
- 报表生成
- 导入/导出/清空数据
- 坏 `.wlr` / `.wlt` 文件容错
- 非法导入目录拒绝

## 当前发版状态

当前仓库已经完成一轮发版前校验：

- `build.bat --no-pause` 通过
- `test.bat --no-pause` 通过

正式产物：

- `build\WorkLogLite.exe`
- `portable\WorkLogLite.exe`
- `dist_portable\WorkLogLite.exe`
- `WorkLogLite-portable.zip`
