# 云端验证与使用说明

验证日期：2026-10-06（Asia/Tokyo）。验证输入代码：main，6a648090f1796e6270268f54e5dfba92dbcacb33。仅处理 interior-design-workspace。

## 使用

当前脚本面向此 Codex Debian 13 / Python 3.12 环境，仓库固定在 `/workspace/interior-design-workspace`，辅助工具固定在 `/workspace/design-tools`。使用现有隔离检出，不创建 worktree。

```bash
cd /workspace/interior-design-workspace
bash scripts/install-cloud.sh
source /workspace/design-tools/activate.sh
bash scripts/verify-cloud.sh
.venv/bin/python your_script.py
```

安装脚本包含完整本地工具配置、验证程序和 `bash setup.sh`。系统目录不可写，LibreOffice 包通过 Debian 签名索引及包校验下载后解包到工作区；未关闭 TLS 或签名校验。依赖范围仍由既有 requirements.txt 定义，实际版本记录在 work/installed-packages.txt。这是可复现工作流程，不承诺未来仓库软件包版本不变。

当前实例已有 Noto CJK 与 Poppler，脚本保留按需下载要求；其他基础镜像须先确认这些工具可用。激活脚本设置可写缓存与 LibreOffice 包装器。没有常驻服务，进程不作为恢复前提。

验证会覆盖 work/onboarding 中命名固定的合成测试样例；转换前删除该验证自建的三个 PDF，防止旧产物掩盖失败。勿在此目录放客户文件。inputs 原件不改动，work 和 outputs 用于中间产物及交付；这些目录、.venv 与工具缓存均不上传 GitHub。

## 本次实际通过

- 已通过平台 Git 代理 fetch GitHub main，与原检出一致；无需合并。原跟踪文件无未提交修改；旧验证产物已另存 work/sync-20261006/previous-onboarding。
- 原保存安装脚本成功复跑；新 scripts/install-cloud.sh 完整执行成功；独立启动验证在同等批准执行范围也成功。默认沙箱中 Git 代理不可达，LibreOffice 转换提前退出；经批准执行范围复测成功，后续任务可能仍需相同权限。这是已诊断的执行限制，未关闭沙箱或改动私有权限。
- Python 3.12.14；ezdxf 1.4.4、openpyxl 3.1.5、python-docx 1.2.0、python-pptx 1.0.2、pypdf 6.19.0、Pillow 12.3.0、matplotlib 3.11.2 导入通过；pip check 无冲突。
- R2007 DXF 往返 AC1021，合成样例显式毫米，闭合四边形与中文文字保留，audit 无错误。未知输入单位仍须询问，不能套用样例单位。
- Noto CJK 中文字体匹配，Poppler 可用，LibreOffice 25.2.3.2 Writer/Calc/Impress 可执行。
- 中文 DOCX/XLSX/PPTX 创建并重读；LibreOffice 转三份 PDF，Matplotlib 生成中文图表 PDF。
- 四份 PDF 均单页，正确中文文本、Noto 嵌入字体、Poppler PNG 渲染通过；逐张目视检查无乱码、裁切或溢出。实际客户交付仍逐份检查。
- 日志保留在 work/sync-20261006；合成 PDF/PNG 不上传。

## 限制与发布

LibreOffice 有 Java 组件和 liblangtag 系统数据路径警告；已测普通办公转换通过，Java 宏和复杂语言处理未验证。

AutoCAD 2007 仍需本地复核单位、尺寸、图层、中文、块、外参、线型和打印；ezdxf 不直接处理 DWG。SketchUp 的具体版本、许可、交换格式和单位一致性待本地实测。Carat 的产品版本、导出能力及原生工程处理均待确认。未安装这些商业桌面软件。

当前环境运行且保存的安装/启动字段可读，基础版本与首次配置时已不同；本次验证证明当前实例能力，不等于新环境冷恢复成功。环境状态工具的网络策略 state=unknown，网络能力依据实际 Git 与安装请求验证。

本次同步不修改私有权限或共享范围。现有工具不能核实“仅本人可用”，也没有发布操作；新配置可保存为草稿，但本次不能代为发布。请在环境设置核实仅本人访问，保存草稿并 Publish，随后在新任务执行 scripts/verify-cloud.sh。未把更新的发布或新任务恢复描述为已通过。
