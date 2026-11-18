# .pi/wheels —— 离线依赖库

**目标解释器：CPython 3.12 / win_amd64。** 这个目录里的二进制 wheel 全是 cp312，
3.13 / 3.14 的解释器装不上（numpy、pillow、lxml 这些带 ABI 标记）。

## 怎么用

两种装法，别混：

| 场景 | 命令 | 网络 |
|---|---|---|
| pi 启动自举（默认路径） | 什么都不用做，pi 起来时自动跑 | **只走内网 server**（`indexUrl`），不看本目录 |
| 手工兜底 / 环境救急 | `python -m pip install --no-index --find-links=.pi\wheels\py312 -r requirements.txt` | 全程不联网，只用本目录 |

`setup.bat` 走第二条：先按解释器 minor 版本找 `py312/` 这种子目录（pip 的 `--find-links`
不递归，**必须指到子目录**），找不到才转内网镜像。

pi 启动时的自举（`.pi/extensions/deps-bootstrap.ts`）**只认 `.pi/deps.json` 的 `indexUrl`**，
不回退公网、也不读本目录——依赖统一由内网 server 供应，本目录是另一条要手工维护的路径。

## 内网 server 上必须齐的 20 个文件

server 端要能解析出这**全部 20 个** wheel。只放 7 个直接依赖不够：pip 会顺着依赖树
去 server 找那 13 个传递依赖，缺一个就可能退化成找 sdist、进而编译失败（离线机没有编译链）。

直接依赖 7 个（`requirements.txt` 里写的）：

    python_pptx-1.0.2-py3-none-any.whl
    lxml-6.1.3-cp312-cp312-win_amd64.whl
    pydantic-2.13.5-py3-none-any.whl
    matplotlib-3.11.2-cp312-cp312-win_amd64.whl
    numpy-2.4.6-cp312-cp312-win_amd64.whl
    pillow-12.3.0-cp312-cp312-win_amd64.whl
    pyyaml-6.0.3-cp312-cp312-win_amd64.whl

传递依赖 13 个（不写进 requirements.txt，pip 顺着依赖树自己取）：

    annotated_types-0.8.0-py3-none-any.whl
    contourpy-1.3.3-cp312-cp312-win_amd64.whl
    cycler-0.12.1-py3-none-any.whl
    fonttools-4.65.0-cp312-cp312-win_amd64.whl
    kiwisolver-1.5.1-cp312-cp312-win_amd64.whl
    packaging-26.3-py3-none-any.whl
    pydantic_core-2.46.5-cp312-cp312-win_amd64.whl
    pyparsing-3.3.2-py3-none-any.whl
    python_dateutil-2.9.0.post0-py2.py3-none-any.whl
    six-1.17.0-py2.py3-none-any.whl
    typing_extensions-4.16.0-py3-none-any.whl
    typing_inspection-0.4.4-py3-none-any.whl
    xlsxwriter-3.2.9-py3-none-any.whl

合计 20 个 / 约 38 MB。**文件名不必一模一样**（server 只按包名+版本+ABI 标签匹配），
但包名、版本号、`cp312-win_amd64` / `py3-none-any` 这两类标签必须对得上。

> **server 端必须是 pip 能读的索引**，光把 whl 扔进一个目录不够：
> `http://<ip>:<port>/simple/` 要能列出索引页，`/simple/<包名>/` 要能列出该包的文件链接
> （PEP 503）。用 `python -m http.server` 手搓或者用 devpi / pypiserver / nginx autoindex
> 都可以，但目录结构得符合规范。索引页 404 时 pip 报的是
> `No matching distribution found`，看起来像缺包、实际是索引没建对——所以先用上面的预检确认。

## 怎么确认 server 上真的齐了

在**能连内网 server 的那台离线 PC** 上，仓库根执行：

```bat
set PI_DEPS_VERIFY=1
npx jiti .pi/extensions/deps-bootstrap.ts
```

只查询、不安装、不改本地环境，退出码 0 = 20 个包全都能取到：

```
[deps] 预检：向 http://172.30.39.34:8765/simple 查询 7 个直接依赖及其传递依赖（只查询，不安装）...
[deps] ✓ 预检通过：requirements.txt 的直接依赖与传递依赖都能从内网 server 取到
[deps]   解析出 20 个包：numpy-2.4.6  pillow-12.3.0  ...
```

失败会分两种，别混：

- `✗ 预检失败：内网 server 拒绝连接（...）` → 网络/服务问题，不是缺包
- `✗ 预检失败：内网 server 上没有 xxx` + `内网 server 上缺这些包：...` → 去 server 补包

**不要在本机开发机（非离线 PC）上跑这条**，它连不到内网 IP，只会白等。

## 补一个包 / 换解释器版本

从内网镜像拉（`172.30.39.34:8765`，pip 标准 `/simple` 索引）：

```bat
REM 只下 wheel 不安装，落到指定子目录
python -m pip download -d .pi\wheels\py312 --only-binary=:all: ^
    --index-url http://172.30.39.34:8765/simple --trusted-host 172.30.39.34 ^
    -r requirements.txt

REM 换 3.14 的解释器就改版本号
python -m pip download -d .pi\wheels\py314 --only-binary=:all: ^
    --python-version 3.14 --platform win_amd64 ^
    --index-url http://172.30.39.34:8765/simple --trusted-host 172.30.39.34 ^
    -r requirements.txt
```

## 已经清掉的包

2026-11 做过一次瘦身：原来 45 个 wheel / 87 MB，混着 pdf（pdfplumber / pymupdf /
pdf2image / PyPDF2 / pypdfium2）、plotly、markdown、svgwrite、requests 整条链、
python-docx，以及 contourpy / fonttools / pyparsing 各自的旧版本。

这些**当前代码一个都没 import**（`.pi/renderer/docx_in.py` 读 Word 用的是
`zipfile` + `lxml` 直接解 OOXML，不需要 python-docx）。

删掉的 25 个包暂时存放在 `.pi/work/_wheels-backup/`（该目录不入库），确认镜像能正常
补包后可以直接删。真要恢复某个能力，按上面「补一个包」的命令拉回来即可。

## 别做的事

- 不要把 `.pi/wheels` 整个目录当 `--find-links` 传进去（不递归，找不到包）。
- 不要把 wheel 提交进 git（`.gitignore` 已排除，86 MB 不该进库）。
- 不要在 requirements.txt 里写中文注释（pip 按 GBK 解码该文件，非 ASCII 会崩）。
