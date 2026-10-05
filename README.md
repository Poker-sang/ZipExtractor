# ZipExtractor

自动解压各种各样的压缩包

## 引擎与结构

Windows x64 / .NET 10。运行时直接调用完整 `7z.dll`（ZIP、7z）和官方 `UnRAR.dll`（RAR4、RAR5），不启动 WinRAR/7z 命令行，也不依赖 SharpCompress。

- `Extraction/IArchiveExtractor`（定义在 `Contracts.cs`）：一次密码尝试、一个全新输出目录，返回结构化状态和实际校验成功的卷。替换后端只需实现接口，在 `Program.cs` 注入。
- `ArchiveDiscovery`：读文件头确认格式；文件名编号用于关联分卷，可读的 RAR5 卷头编号用于排序。它只负责候选发现，最终有效性由引擎打开及完整解压校验确认。
- 7-Zip 接口和回调使用 .NET 源生成 COM（`GeneratedComInterface` / `GeneratedComClass`），关闭传统运行时 COM 支持；独立 COM 包装器确定性释放。互操作层使用 SDK 枚举、明确类型的缓冲区指针以及内置字符串/接口封送。
- `IVolumeResolver`：把引擎请求的卷名映射到物理文件。默认支持发现阶段建立的别名；也可注入自定义映射。
- `ExtractionPipeline`：密码重试、嵌套解压、临时目录、结果提交和清理。
- 保留 `FileSystemHelper.cs` 的空目录清理、“新建文件夹”整理和重复目录合并规则；解压与整理共用游戏目录保护判断。

## 安装、运行

在仓库根目录运行：

```powershell
.\scripts\Install-NativeLibraries.ps1
dotnet build
dotnet run --project .\ZipExtractor
```

安装脚本下载并检查固定 SHA-256 的官方 7-Zip 26.03 / UnRAR 7.20 安装包，提取 x64 DLL 和许可证到 `ZipExtractor/native`。脚本需要已有 `7z.exe` 来解开安装包，可用 `-SevenZipExe` 指定。运行程序只需 DLL，无需安装 WinRAR。DLL 和下载缓存不加入 Git，build/publish 自动复制 DLL 及许可证。

默认沿用 `D:\TempExt`、`D:\CompleteExt`、`D:\ErrorExt` 和原有候选密码。示例：

```powershell
dotnet run --project .\ZipExtractor -- --input C:\Archives --output C:\Extracted --keep-sources --non-interactive
```

参数：

| 参数 | 含义 |
|---|---|
| `--input` / `--output` | 输入、完成目录，两者不能互相包含 |
| `--password` | 增加一个候选密码；终端当前状态行显示序号及实际候选密码 |
| `--keep-sources` | 成功后保留源包；默认把已校验的原始卷送入回收站 |
| `--max-depth` | 嵌套解压层数上限，默认 20，包括最外层 |
| `--non-interactive` | 不等待按键，也不执行原来的最后一次强制唯一子项合并 |
| `--no-normalize` | 不整理完成目录的文件夹层级 |
| `--7zip-dll` / `--unrar-dll` | 指定替代的 x64 DLL 路径 |

Ctrl+C 取消解压。退出码：0 完成，1 有失败，2 取消。

日志使用 Spectre.Console：彩色状态、任务序号、嵌套层级、实时进度条、处理速度、耗时及完成汇总。7-Zip 显示原生百分比；UnRAR 总量未知时显示动态进度条和已处理字节，不伪造百分比。候选密码直接显示，方括号等特殊字符按原文输出。所有密码尝试共用一条临时状态行：密码失败时原地显示“注意”，下一次尝试覆盖它，结束后清除；历史密码不占用日志。重定向或关闭 ANSI 时仅输出最终任务结果，设置 `NO_COLOR` 可禁用颜色。

## 分卷与递归行为

- 隐写 ZIP/7z 由 `7z.dll` 原生探测：无法通过外层文件头识别的文件交给引擎打开，`maxCheckStartPosition` 设置为输入长度，允许定位大前缀后的压缩包。直接读取原文件，不自行解析嵌入 ZIP 目录或复制切片；解压仍执行原生 CRC 校验，内层压缩包独立尝试密码。
- 单包即使叫 `archive.txt`，只要文件头是 ZIP/7z/RAR，也会发现。
- 支持 `name.part01.rar.fake`、`.rar/.r00`、`name.7z.001.txt`、`name.zip.001`、`name.z01.fake + name.zip.fake` 等常见命名。标准 ZIP 分卷与把 ZIP 整体按字节切分的方式分别处理。
- 对可读卷头的 RAR5，同一文件名前缀下，即使后缀完全换成 `.a/.b/.c`，也可利用卷头编号确定顺序。RAR 加密头可能隐藏这些编号，此时必须有可靠的文件名编号或显式卷映射。
- 所有卷原名、原位置不变。缺卷、不连续编号及不确定顺序会失败并保留源文件；不按任意目录顺序拼接。
- 不要求“目录里全是压缩包”才递归。每层扫描真实内容，混合目录里的普通图片、说明文件等保留，改后缀的内层包仍会继续解压。
- 游戏目录优先保护：目录直属文件中含 `.exe`、`.dll` 或沿用原规则的 `.bin`（大小写不敏感），就停止扫描/解压该目录及其全部子目录。游戏目录不会参与文件夹合并（包括被父目录折叠），内部空目录也保留。这同时适用于初始输入和每一层解压结果；邻近的非游戏目录仍正常处理。
- 每次密码尝试使用独立目录。只在整个外层任务及其嵌套包全部成功后提交结果，再清理源包。嵌套失败、达到深度上限或取消时不提交半成品，原始源包保留。
- 失败包留在原位置，控制台显示最终失败原因，不生成失败报告 txt。去掉了原先按解压前后大小比例判断成功的逻辑。
- 清理只针对发现的卷与引擎实际校验卷的交集。重复输出文件名明确报错，避免静默覆盖。

限制：无法从无格式头的裸分片中恢复已经完全丢失的卷顺序。RAR/7z 的 SFX 探测目前扫描 PE 文件前 1 MiB；符号链接/硬链接提取明确返回不支持。恢复记录修复、恢复卷重建不在本解压层范围内。加密数据 CRC 错误有时不能区分密码错误和文件损坏，会保留这种不确定性。

## 验证

```powershell
dotnet run --project .\ZipExtractor.Tests
```

集成测试只在仓库 `work/tests-*` 生成和解压样本，不扫描默认的用户目录。生成 ZIP/7z/RAR5 样本需要 7-Zip 和 RAR 命令行，可用 `TEST_7ZIP_EXE`、`TEST_RAR_EXE` 环境变量指定；RAR4 固定样本来自 libarchive，出处和许可证在测试 Fixtures 中。测试程序退出非零表示失败。

覆盖改名、中文文件名及密码、AES ZIP、加密头、固实、两类 ZIP 分卷、RAR4/RAR5 分卷、缺卷、CRC 损坏、取消、路径越界、空包、嵌套混合内容、递归上限、源文件保留和原有目录合并。

原生依赖来源与许可证：

- [7-Zip 源码](https://github.com/ip7z/7zip) / [许可证](https://www.7-zip.org/license.txt)：LGPL 等；发布应用时应一并遵守相应源码及再分发要求。
- [官方 UnRAR SDK 与源码](https://www.rarlab.com/rar_add.htm)：UnRAR 许可，含禁止利用源码重建 RAR 压缩算法的限制。
