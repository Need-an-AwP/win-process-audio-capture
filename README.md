# win-process-audio-capture

Windows 进程音频捕获的 Node.js 原生插件。

## 描述

这个模块提供了一个预编译的原生插件，用于捕获 Windows 特定进程的音频

## 安装

```bash
npm install win-process-audio-capture
```

## 使用方法

参考https://github.com/Need-an-AwP/Capture-Audio-from-Process---javascript-addon

## 重要说明

只考虑提供预编译版本

仅支持windows

## 对于开发者

如果你需要修改并重新构建这个模块，请注意：

- 编译需要以下工具：
  - node-gyp
  - Visual Studio Build Tools
  - Windows SDK
  - WIL (Windows Implementation Library)

- 构建过程可以通过以下命令启动：

```bash
npm run build
```

然而，除非你在为模块的开发做贡献，否则不建议重新构建。

## 故障排除

运行 `npm install` 可能会出错，这是因为 `npm install` 会自动运行 `node-gyp rebuild`，而 `node-gyp rebuild` 可能会导致未知的编译错误。本模块设计为无需重新构建即可工作，所以如


```
"build": "node-gyp configure && node-gyp build",
```

