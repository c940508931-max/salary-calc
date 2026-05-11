# README - 路易小姐薪资计算工具

一款轻量级薪资计算工具，适用于路易小姐直播间薪资核算。

## Mac 版（.app 应用）

### 使用方式
- 将 `dist/SalaryCalc.app` 发给任何人
- 双击运行即可
- 浏览器打开 `http://localhost:5001`

### 注意事项
- macOS 首次运行可能提示"无法验证开发者"，去 **系统设置 → 隐私与安全性 → 仍要打开**
- 支持 Apple Silicon (M1/M2/M3/M4) 和 Intel Mac

## Windows 版（打包方式）

### 方式一：直接运行（需要 Python）
1. 安装 Python 3.8+（[下载](https://www.python.org/downloads/windows/)）
2. 双击 `start.bat` 即可运行
3. 浏览器打开 `http://localhost:5001`

### 方式二：打包成 exe（推荐分享）
1. 以**管理员身份**打开 PowerShell
2. 执行：
   ```
   Set-ExecutionPolicy Unrestricted -Scope Process
   .\package-windows.ps1
   ```
3. 等待打包完成，生成 `dist/薪资计算工具.exe`
4. 把这个 exe 发给同事即可

> ⚠️ 首次运行 .exe 可能被 Windows Defender 拦截，点击"更多信息"→"仍要运行"

## 项目结构
```
salary-calc/
├── app.py                  # Flask 应用主程序
├── templates/              # 前端模板
│   ├── base.html
│   ├── index.html
│   └── preview.html
├── start.bat               # Windows 直接启动脚本
├── package-windows.ps1     # Windows 打包脚本
├── dist/                   # 打包输出目录
│   └── SalaryCalc.app      # Mac 应用
└── README.md
```
