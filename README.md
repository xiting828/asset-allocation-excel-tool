# 粘贴表格转 Excel

一个可在本地或 Streamlit Community Cloud 运行的中文网页工具。用户可以粘贴从 Excel、网页或聊天软件复制的 Tab/换行分隔表格，在浏览器中预览，并下载使用 `openpyxl` 在内存中生成的 `.xlsx` 文件。

## 主要功能

- 识别 Tab 分列、换行分行的数据
- 将所有 `*` 字符替换为 `0`
- 只裁剪表格外围全空白行和列，保留内部空白单元格
- 各行列数不一致时自动在末尾补空单元格并提示
- 带前导零的编号和超长数字优先按文本保存
- 所有 Excel 单元格水平、垂直居中
- 表头淡蓝色并加粗，自动设置列宽、行高、换行、边框、冻结首行和筛选
- 精确识别“资产大类”列：“固定收益类”整行淡绿色，“权益类”整行淡蓝色
- 现金类、另类投资、保障类、合计及其他内容保持原背景
- Excel 文件仅在内存中生成，不写入服务器永久目录

## 隐私和输入限制

应用不会主动保存或记录用户粘贴的表格内容，也不收集姓名、联系方式等用户信息。生成的 Excel 以浏览器下载数据的形式临时提供，不会永久保存到服务器；下载触发后会清除会话内的生成文件数据。

请勿输入银行卡号、身份证号、密码等高度敏感信息。

为避免单次异常输入消耗过多云端资源，应用限制为：

- 文本大小不超过 2 MB
- 最多 10,000 行
- 最多 200 列
- 最多 500,000 个单元格
- 单个单元格最多 32,000 个字符

## 项目结构

```text
.
├── app.py
├── table_tool.py
├── requirements.txt
├── runtime.txt
├── README.md
├── .gitignore
└── tests/
    ├── __init__.py
    └── test_table_tool.py
```

`outputs/`、`work/`、虚拟环境和生成的 `.xlsx` 文件已被 `.gitignore` 排除，不应上传到公开仓库。

## 本地运行

建议使用 Python 3.12，与云端部署选择和部署前验证环境保持一致。

macOS/Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run app.py
```

Windows PowerShell：

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run app.py
```

浏览器打开 <http://localhost:8501>。按 `Ctrl+C` 停止应用。

## 运行测试

测试使用 Python 自带的 `unittest`，无需额外安装测试框架：

```bash
python -m unittest discover -s tests -v
```

## 上传至 GitHub

以下命令会创建一个全新的仓库历史，不会覆盖其他项目。先在 GitHub 创建一个空仓库，例如 `asset-allocation-excel-tool`，不要勾选自动创建 README、License 或 `.gitignore`。

```bash
git init
git branch -M main
git add app.py table_tool.py requirements.txt runtime.txt README.md .gitignore tests
git commit -m "Prepare Streamlit Excel tool for deployment"
git remote add origin https://github.com/你的用户名/asset-allocation-excel-tool.git
git push -u origin main
```

推送前建议执行：

```bash
git status
git ls-files
```

确认没有 `.venv/`、`outputs/`、`work/`、生成的 Excel、密钥或真实客户数据。

## 部署到 Streamlit Community Cloud

1. 打开 <https://share.streamlit.io/>。
2. 使用 GitHub 登录，并授权 Streamlit 访问刚创建的仓库。
3. 点击右上角 **Create app**。
4. 选择 **Yup, I have an app**。
5. 填写：
   - **Repository**：`你的用户名/asset-allocation-excel-tool`
   - **Branch**：`main`
   - **Main file path**：`app.py`
6. 打开 **Advanced settings**，Python version 选择 **3.12**。本项目不需要填写 Secrets。
7. 可选填写一个未被占用的 App URL，然后点击 **Deploy**。
8. 等待依赖安装完成。部署成功后会得到 `https://...streamlit.app` 公共链接。

部署后必须实际粘贴测试数据、点击生成并下载文件，确认下载的 Excel 可以打开，才能视为部署验证完成。

## 后续更新和重新部署

修改并通过本地测试后提交、推送：

```bash
git add app.py table_tool.py requirements.txt README.md tests
git commit -m "Update application"
git push origin main
```

Community Cloud 会监控 GitHub 仓库。普通代码提交通常会自动更新；`requirements.txt` 发生变化时会重新安装依赖，耗时稍长。

如果页面没有更新：

1. 打开 <https://share.streamlit.io/>。
2. 找到应用，点击右侧三点菜单。
3. 点击 **Reboot** 并确认。
4. 也可以在公开应用右下角点击 **Manage app**，从日志菜单选择 **Reboot app**。

## 常见部署错误

### `ModuleNotFoundError`

检查缺失包是否写在仓库根目录的 `requirements.txt` 中，提交并推送后等待依赖重新安装。本项目运行依赖只有 Streamlit、openpyxl 和 pandas。

### 找不到 `app.py`

部署页面的 **Main file path** 必须填写 `app.py`，并确认文件位于仓库根目录。

### GitHub 仓库无法选择

重新登录 <https://share.streamlit.io/> 并授权 Streamlit Community Cloud 访问目标仓库。私有仓库需要授予相应权限；公开工具建议使用不含客户数据的公开仓库。

### Python 或依赖安装失败

在 **Advanced settings** 选择 Python 3.12，并确认 `requirements.txt` 没有系统模块或仅限某个平台的包。修改 Python 版本通常需要删除应用后按原 Repository、Branch 和 Main file path 重新部署。

### 页面显示资源超限或运行缓慢

缩小粘贴数据量并分批处理。应用本身会阻止超过安全限制的输入。必要时在 **Manage app** 中查看日志并重启应用。

### 部署后页面是旧版本

确认最新提交已经推送到部署所选的 `main` 分支。等待自动更新；仍未更新时执行 **Reboot**。

### 应用长时间无人访问后休眠

Streamlit Community Cloud 可能让无流量应用休眠。再次访问公开链接并按页面提示唤醒即可。
