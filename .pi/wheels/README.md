# .pi/wheels —— 离线依赖库

**目标解释器：CPython 3.12 / win_amd64。** 这个目录里的二进制 wheel 全是 cp312，
3.13 / 3.14 的解释器装不上（numpy、pillow、lxml 这些带 ABI 标记）。

## 怎么用

```bat
REM 离线机（3.12）上装依赖，全程不联网
python -m pip install --no-index --find-links=.pi\wheels\py312 -r requirements.txt
```

或者直接跑仓库根的 `setup.bat`——它会先按当前解释器的 minor 版本找 `py312/` 这样的子目录；
pip 的 `--find-links` 不递归，所以**必须指到子目录，不能只指 `.pi/wheels`**。

pi 启动时会自动做同一件事（`.pi/extensions/deps-bootstrap.ts`），私有源和解释器版本
配在 `.pi/deps.json`：`{"python": "3.12", "indexUrl": "http://172.30.39.34:8765/simple"}`。

## 目录里有什么（20 个）

7 个直接依赖（`requirements.txt` 里显式写的）：

    python-pptx  lxml  pydantic  matplotlib  numpy  Pillow  PyYAML

13 个传递依赖（pip 顺着依赖树自己取，不用写进 requirements.txt）：

    annotated-types  contourpy  cycler  fonttools  kiwisolver  packaging
    pyparsing  python-dateutil  six  typing-extensions  typing-inspection
    pydantic-core  XlsxWriter

合计约 38 MB。验证方式（在任意有 pip 的机器上跑，不需要 3.12）：

```bat
python -m pip install --dry-run --no-index --only-binary=:all: ^
    --find-links=.pi/wheels/py312 --python-version 3.12 --platform win_amd64 ^
    --target %TEMP%\dryrun -r requirements.txt
```

末行 `Would install` 应该正好列出 20 个包，少一个就说明 wheel 缺了。

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
